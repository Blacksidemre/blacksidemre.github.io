# -*- coding: utf-8 -*-
"""
===============================================================================
SAMSUN / ATAKUM KONUT FİYAT TAHMİNİ - BİTİRME TEZİ ANALİZ SCRİPTİ
===============================================================================
Tez çalışmasındaki tüm istatistiksel testleri (4.1 - 4.8) ve makine öğrenmesi
modellerini (4.10.1) tek parça, baştan sona çalıştırılabilir bir dosyada toplar.

Kullanım:
    python bitirme_tezi_analiz.py

Gerekli veri dosyası (script ile aynı klasörde olmalı):
    bitirmetezi_veriseti.xlsx

Gerekli kütüphaneler:
    pip install pandas numpy matplotlib seaborn scipy scikit-learn statsmodels
    pip install xgboost tensorflow tabulate openpyxl

statsmodels / xgboost / tensorflow kurulu değilse ilgili bölümler otomatik
olarak atlanır, script yine de hatasız tamamlanır.
===============================================================================
"""

import warnings
warnings.filterwarnings("ignore")

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

from scipy.stats import chi2_contingency, shapiro, ttest_ind, t as t_dist

from sklearn.model_selection import train_test_split, GridSearchCV
from sklearn.linear_model import LinearRegression, Ridge, Lasso, ElasticNet
from sklearn.preprocessing import StandardScaler, OneHotEncoder, PolynomialFeatures
from sklearn.compose import ColumnTransformer
from sklearn.pipeline import Pipeline
from sklearn.metrics import mean_squared_error, r2_score, mean_absolute_error
from sklearn.neighbors import KNeighborsRegressor
from sklearn.ensemble import RandomForestRegressor, GradientBoostingRegressor

# Opsiyonel kütüphaneler: kurulu değilse ilgili bölüm otomatik atlanır
try:
    import statsmodels.api as sm
    from statsmodels.formula.api import ols
    from statsmodels.stats.multicomp import pairwise_tukeyhsd
    from statsmodels.multivariate.manova import MANOVA
    from statsmodels.stats.outliers_influence import variance_inflation_factor
    HAS_STATSMODELS = True
except ImportError:
    HAS_STATSMODELS = False
    print("[UYARI] statsmodels bulunamadı -> ANOVA / MANOVA / VIF bölümleri atlanacak.")
    print("        Kurulum: pip install statsmodels\n")

try:
    import xgboost as xgb
    HAS_XGBOOST = True
except ImportError:
    HAS_XGBOOST = False
    print("[UYARI] xgboost bulunamadı -> XGBoost bölümü atlanacak.")
    print("        Kurulum: pip install xgboost\n")

try:
    from tensorflow.keras.models import Sequential
    from tensorflow.keras.layers import Dense, Dropout
    from tensorflow.keras.optimizers import Adam
    from tensorflow.keras.callbacks import EarlyStopping
    HAS_TENSORFLOW = True
except ImportError:
    HAS_TENSORFLOW = False
    print("[UYARI] tensorflow bulunamadı -> Yapay Sinir Ağı (ANN) bölümü atlanacak.")
    print("        Kurulum: pip install tensorflow\n")

pd.set_option("display.width", 160)
pd.set_option("display.max_columns", 40)
RANDOM_STATE = 42
np.random.seed(RANDOM_STATE)

VERI_DOSYASI = "bitirmetezi_veriseti.xlsx"


def tablo_yazdir(df_):
    """DataFrame'i markdown tablo olarak yazdırır, tabulate yoksa düz yazdırır."""
    try:
        print(df_.to_markdown(numalign="left", stralign="left"))
    except ImportError:
        print(df_.to_string())


# ==============================================================================
# 1. VERİ YÜKLEME VE SÜTUN ADLARININ SADELEŞTİRİLMESİ
# ==============================================================================
print("=" * 80)
print("1. VERİ YÜKLEME")
print("=" * 80)

raw_df = pd.read_excel(VERI_DOSYASI)

TR_HARF_HARITASI = str.maketrans({
    "ç": "c", "Ç": "C", "ı": "i", "İ": "I",
    "ş": "s", "Ş": "S", "ğ": "g", "Ğ": "G",
    "ü": "u", "Ü": "U", "ö": "o", "Ö": "O",
    "²": "2",
})


def sutun_adi_temizle(ad: str) -> str:
    ad = ad.translate(TR_HARF_HARITASI)
    ad = ad.replace("(", "").replace(")", "")
    ad = ad.strip().replace(" ", "_")
    while "__" in ad:
        ad = ad.replace("__", "_")
    return ad.strip("_")


raw_df.columns = [sutun_adi_temizle(c) for c in raw_df.columns]
print(f"Veri seti boyutu: {raw_df.shape}")
print("Sütunlar:", list(raw_df.columns))

# ==============================================================================
# 2. VERİ TİPİ DÖNÜŞÜMÜ VE EKSİK DEĞER İŞLEME
# ==============================================================================
print("\n" + "=" * 80)
print("2. VERİ TİPİ DÖNÜŞÜMÜ VE EKSİK DEĞERLER")
print("=" * 80)

numerical_cols = [
    "Fiyat_TL", "Brut_m2", "Net_m2", "Banyo_Sayisi",
    "Bina_Yasi_Ortalama", "Bulundugu_Kat_Donusturulmus",
    "Oda_Sayisi_Numeric", "Kat_Sayisi_Numeric",
]

for col in numerical_cols:
    raw_df[col] = pd.to_numeric(raw_df[col], errors="coerce")
    if raw_df[col].isnull().any():
        raw_df[col] = raw_df[col].fillna(raw_df[col].median())

# NOT: Sütun ADLARI yukarıda ASCII'ye çevrildi, ancak hücre İÇERİKLERİ orijinal
# Türkçe karakterleriyle kalıyor (ör. "Hayır", "Boş", "Kiracılı"). Bu yüzden
# aşağıdaki doldurma ve eşleştirme (map) işlemlerinde orijinal Türkçe yazımı kullanıyoruz.
kategorik_bosluk_doldur = {
    "Mahalle": "Belirtilmemiş",
    "Esyali": "Belirtilmemiş",
    "Krediye_Uygun": "Belirtilmemiş",
    "Takas": "Hayır",
    "Site_Adi": "Belirtilmemiş",
}
for col, deger in kategorik_bosluk_doldur.items():
    if col in raw_df.columns:
        raw_df[col] = raw_df[col].fillna(deger)

print("Eksik değer sayıları (doldurma sonrası):")
print(raw_df[numerical_cols].isnull().sum())

# ==============================================================================
# 3. KATEGORİK DEĞİŞKENLERİN KODLANMASI
# ==============================================================================
raw_df["Esyali_Kod"] = raw_df["Esyali"].map({"Evet": 1, "Hayır": 0, "Belirtilmemiş": 2}).fillna(2)
raw_df["Krediye_Uygun_Kod"] = raw_df["Krediye_Uygun"].map({"Evet": 1, "Hayır": 0, "Belirtilmemiş": 2}).fillna(2)
raw_df["Takas_Kod"] = raw_df["Takas"].map({"Evet": 1, "Hayır": 0}).fillna(0)
raw_df["Kullanim_Durumu_Kod"] = raw_df["Kullanim_Durumu"].map({"Boş": 0, "Kiracılı": 1, "Mülk Sahibi": 2}).fillna(0)
raw_df["Site_Icerisinde_Kod"] = raw_df["Site_Icerisinde"].map({"Evet": 1, "Hayır": 0}).fillna(0)

original_df = raw_df.copy()  # aykırı değer temizliği öncesi referans veri seti

# ==============================================================================
# 4. TANIMLAYICI İSTATİSTİKLER (Bulgular 4.1.1)
# ==============================================================================
print("\n" + "=" * 80)
print("4.1.1. TANIMLAYICI İSTATİSTİKLER")
print("=" * 80)

target_numerical_cols = numerical_cols

desc_stats = original_df[target_numerical_cols].describe().T
desc_stats["median"] = original_df[target_numerical_cols].median()
desc_stats["var"] = original_df[target_numerical_cols].var()
desc_stats["IQR"] = desc_stats["75%"] - desc_stats["25%"]
desc_stats["N"] = original_df[target_numerical_cols].count()
desc_stats = desc_stats[["N", "mean", "median", "var", "std", "min", "25%", "75%", "max", "IQR"]]
desc_stats.columns = ["N", "Ort.", "Med.", "Var.", "SS", "Min", "Q1", "Q3", "Max", "IQR"]
tablo_yazdir(desc_stats.round(2))

# ==============================================================================
# 5. AYKIRI DEĞER TESPİTİ VE TEMİZLİĞİ (Bulgular 4.1.2)
# ==============================================================================
print("\n" + "=" * 80)
print("4.1.2. AYKIRI DEĞER TESPİTİ (IQR ve Z-SKORU)")
print("=" * 80)

iqr_outliers = []
zscore_outliers = []
for col in target_numerical_cols:
    Q1, Q3 = original_df[col].quantile(0.25), original_df[col].quantile(0.75)
    IQR = Q3 - Q1
    alt_sinir, ust_sinir = Q1 - 1.5 * IQR, Q3 + 1.5 * IQR
    iqr_sayisi = original_df[(original_df[col] < alt_sinir) | (original_df[col] > ust_sinir)].shape[0]
    iqr_outliers.append({"Değişken": col, "Aykırı Say. (IQR)": iqr_sayisi})

    ortalama, std = original_df[col].mean(), original_df[col].std()
    z_skorlari = np.abs((original_df[col] - ortalama) / std) if std > 0 else pd.Series(0, index=original_df.index)
    z_sayisi = (z_skorlari > 3).sum()
    zscore_outliers.append({"Değişken": col, "Aykırı Say. (Z-Skoru)": z_sayisi})

tablo_yazdir(pd.DataFrame(iqr_outliers))
tablo_yazdir(pd.DataFrame(zscore_outliers))

# Nihai temizlenmiş veri seti: IQR yöntemiyle aykırı değerlerin çıkarılması
df = original_df.copy()
for col in target_numerical_cols:
    Q1, Q3 = df[col].quantile(0.25), df[col].quantile(0.75)
    IQR = Q3 - Q1
    alt_sinir, ust_sinir = Q1 - 1.5 * IQR, Q3 + 1.5 * IQR
    df = df[(df[col] >= alt_sinir) & (df[col] <= ust_sinir)]

print(f"\nTemizlik öncesi satır sayısı : {original_df.shape[0]}")
print(f"Temizlik sonrası satır sayısı: {df.shape[0]}")
print(f"Çıkarılan satır sayısı       : {original_df.shape[0] - df.shape[0]}")

comparison_data = []
for col in target_numerical_cols:
    comparison_data.append({
        "Değişken": col,
        "Ort. (Temiz Öncesi)": round(original_df[col].mean(), 2),
        "Ort. (Temiz Sonrası)": round(df[col].mean(), 2),
        "SS (Temiz Öncesi)": round(original_df[col].std(), 2),
        "SS (Temiz Sonrası)": round(df[col].std(), 2),
    })
print("\n4.1.2.3. Aykırı Değerlerin Veriye Etkisi:")
tablo_yazdir(pd.DataFrame(comparison_data))

# ==============================================================================
# 6. LOGARİTMİK DÖNÜŞÜM (Bulgular 4.1.3)
# ==============================================================================
print("\n" + "=" * 80)
print("4.1.3. LOGARİTMİK DÖNÜŞÜM (log1p)")
print("=" * 80)

log_cols = ["Fiyat_TL", "Net_m2"]
log_comparison_data = []
for col in log_cols:
    orijinal = df[col].dropna()
    log_donusumlu = np.log1p(orijinal)
    df[f"{col}_log"] = log_donusumlu

    log_comparison_data.append({
        "Değişken": col,
        "Orijinal Ort.": round(orijinal.mean(), 2),
        "Orijinal SS": round(orijinal.std(), 2),
        "Orijinal Çarpıklık": round(orijinal.skew(), 4),
        "Log Ort.": round(log_donusumlu.mean(), 4),
        "Log SS": round(log_donusumlu.std(), 4),
        "Log Çarpıklık": round(log_donusumlu.skew(), 4),
    })

print("Logaritmik Dönüşüm Öncesi ve Sonrası İstatistikler:")
tablo_yazdir(pd.DataFrame(log_comparison_data))

fig, axes = plt.subplots(2, 2, figsize=(13, 9))
for i, col in enumerate(log_cols):
    sns.histplot(df[col], kde=True, ax=axes[i, 0], color="steelblue")
    axes[i, 0].set_title(f"Orijinal Dağılım: {col}")
    sns.histplot(df[f"{col}_log"], kde=True, ax=axes[i, 1], color="darkorange")
    axes[i, 1].set_title(f"log1p Dönüşümü Sonrası: {col}")
plt.tight_layout()
plt.show()

# ==============================================================================
# 7. GÖRSEL ANALİZLER (Bulgular 4.1.4)
# ==============================================================================
print("\n" + "=" * 80)
print("4.1.4. GÖRSEL ANALİZLER")
print("=" * 80)

fig, axes = plt.subplots(len(target_numerical_cols), 2, figsize=(12, 4 * len(target_numerical_cols)))
fig.suptitle("Aykırı Değer Temizliği Öncesi ve Sonrası Dağılım (Kutu Grafikleri)", y=1.0)
for i, col in enumerate(target_numerical_cols):
    sns.boxplot(y=original_df[col], ax=axes[i, 0])
    axes[i, 0].set_title(f"Orijinal: {col}")
    sns.boxplot(y=df[col], ax=axes[i, 1])
    axes[i, 1].set_title(f"Temizlenmiş: {col}")
plt.tight_layout()
plt.show()

top_10_mahalle = df["Mahalle"].value_counts().nlargest(10).index
plt.figure(figsize=(14, 7))
sns.boxplot(x="Mahalle", y="Fiyat_TL", data=df[df["Mahalle"].isin(top_10_mahalle)])
plt.title("İlk 10 Mahalledeki Konut Fiyatlarının Dağılımı")
plt.xticks(rotation=45, ha="right")
plt.tight_layout()
plt.show()

# ==============================================================================
# 8. BAĞIMSIZ ÖRNEKLEM T-TESTİ (Bulgular 4.2)
# ==============================================================================
print("\n" + "=" * 80)
print("4.2. BAĞIMSIZ ÖRNEKLEM T-TESTİ")
print("=" * 80)

numeric_vars_t_test = ["Fiyat_TL", "Brut_m2", "Net_m2", "Bina_Yasi_Ortalama"]


def t_testi_uygula(grup_kolonu, grup1_adi, grup0_adi):
    print(f"\n--- {grup_kolonu} Değişkenine Göre T-Testi ---")
    for var in numeric_vars_t_test:
        g1 = df[df[grup_kolonu] == 1][var].dropna()
        g0 = df[df[grup_kolonu] == 0][var].dropna()
        if len(g1) > 1 and len(g0) > 1:
            _, p = ttest_ind(g1, g0, equal_var=False)
            print(f"{var:28s} | {grup1_adi} Ort.: {g1.mean():>14,.2f} | {grup0_adi} Ort.: {g0.mean():>14,.2f} | "
                  f"p={p:.4f} | {'Anlamlı fark var' if p < 0.05 else 'Anlamlı fark yok'}")
        else:
            print(f"{var}: yeterli veri bulunamadı.")


t_testi_uygula("Takas_Kod", "Takas Yapan", "Takas Yapmayan")
t_testi_uygula("Site_Icerisinde_Kod", "Site İçinde", "Site Dışında")

# ==============================================================================
# 9. KORELASYON ANALİZİ (Bulgular 4.3)
# ==============================================================================
print("\n" + "=" * 80)
print("4.3. KORELASYON ANALİZİ")
print("=" * 80)

correlation_cols = ["Fiyat_TL", "Brut_m2", "Net_m2", "Bina_Yasi_Ortalama", "Oda_Sayisi_Numeric", "Kat_Sayisi_Numeric"]
corr_matrix = df[correlation_cols].corr(method="pearson")
tablo_yazdir(corr_matrix.round(3))

plt.figure(figsize=(9, 7))
sns.heatmap(corr_matrix, annot=True, cmap="coolwarm", fmt=".2f", linewidths=0.5)
plt.title("Korelasyon Isı Haritası")
plt.tight_layout()
plt.show()

# ==============================================================================
# 10. Kİ-KARE TESTİ (Bulgular 4.4)
# ==============================================================================
print("\n" + "=" * 80)
print("4.4. Kİ-KARE BAĞIMSIZLIK TESTİ")
print("=" * 80)


def ki_kare_uygula(baslik, kolon1, kolon2):
    print(f"\n--- {baslik} ---")
    tablo = pd.crosstab(df[kolon1], df[kolon2])
    chi2, p, dof, _ = chi2_contingency(tablo)
    tablo_yazdir(tablo)
    print(f"Chi2={chi2:.2f} | p-değeri={p:.6f} | {'Anlamlı ilişki var' if p < 0.05 else 'Anlamlı ilişki yok'}")


ki_kare_uygula("Krediye Uygunluk vs Eşyalı Olma Durumu", "Krediye_Uygun_Kod", "Esyali_Kod")
ki_kare_uygula("Takas Durumu vs Kullanım Durumu", "Takas_Kod", "Kullanim_Durumu_Kod")
ki_kare_uygula("Takas Durumu vs Eşyalı Olma Durumu", "Takas_Kod", "Esyali_Kod")
ki_kare_uygula("Kullanım Durumu vs Site İçerisinde Olma Durumu", "Kullanim_Durumu_Kod", "Site_Icerisinde_Kod")

# ==============================================================================
# 11. NORMALLİK TESTLERİ VE GÜVEN ARALIĞI (Bulgular 4.5 / 4.6)
# ==============================================================================
print("\n" + "=" * 80)
print("4.5 / 4.6. NORMALLİK TESTLERİ VE GÜVEN ARALIĞI")
print("=" * 80)

normallik_kolonlari = ["Fiyat_TL", "Brut_m2", "Net_m2", "Bina_Yasi_Ortalama"]
normallik_sonuclari = []
for col in normallik_kolonlari:
    veri = df[col].dropna()
    if len(veri) >= 3:
        stat, p = shapiro(veri.sample(min(len(veri), 5000), random_state=RANDOM_STATE))
        normallik_sonuclari.append({
            "Değişken": col, "İstatistik (W)": round(stat, 3), "p-değeri": f"{p:.2E}",
            "Sonuç": "Normal dağılıma uymuyor" if p < 0.05 else "Normal dağılıma uyuyor",
        })
tablo_yazdir(pd.DataFrame(normallik_sonuclari))

ortalama_fiyat = df["Fiyat_TL"].mean()
standart_hata = df["Fiyat_TL"].std() / np.sqrt(len(df["Fiyat_TL"]))
serbestlik_derecesi = len(df["Fiyat_TL"]) - 1
t_kritik = t_dist.ppf(0.975, serbestlik_derecesi)
hata_payi = t_kritik * standart_hata

print(f"\nFiyat için %95 Güven Aralığı: [{ortalama_fiyat - hata_payi:,.0f} TL, {ortalama_fiyat + hata_payi:,.0f} TL]")
print(f"Ortalama Fiyat: {ortalama_fiyat:,.0f} TL")

# ==============================================================================
# 12. ANOVA VE TUKEY HSD (Bulgular 4.7)
# ==============================================================================
print("\n" + "=" * 80)
print("4.7. ANOVA VE TUKEY HSD")
print("=" * 80)

if HAS_STATSMODELS:
    min_mahalle_sayisi = 30
    sik_mahalleler = df["Mahalle"].value_counts()[df["Mahalle"].value_counts() >= min_mahalle_sayisi].index
    df_mahalle = df[df["Mahalle"].isin(sik_mahalleler)].copy()

    if not df_mahalle.empty:
        model = ols('Q("Fiyat_TL") ~ C(Mahalle)', data=df_mahalle).fit()
        anova_tablosu = sm.stats.anova_lm(model, typ=2)
        print("\nANOVA Tablosu:")
        tablo_yazdir(anova_tablosu.round(4))

        p_deger = model.f_pvalue
        print(f"\nF-istatistiği: {anova_tablosu.loc['C(Mahalle)', 'F']:.2f} | p-değeri: {p_deger:.2E}")
        print("Sonuç:", "Mahalleler arası anlamlı fiyat farkı var." if p_deger < 0.05 else "Anlamlı fark yok.")

        en_kalabalik_3 = df["Mahalle"].value_counts().nlargest(3).index.tolist()
        df_tukey = df_mahalle[df_mahalle["Mahalle"].isin(en_kalabalik_3)]
        if df_tukey["Mahalle"].nunique() > 1:
            tukey_sonuc = pairwise_tukeyhsd(endog=df_tukey["Fiyat_TL"], groups=df_tukey["Mahalle"], alpha=0.05)
            print("\nTukey HSD Sonuçları (en kalabalık 3 mahalle):")
            print(tukey_sonuc)
    else:
        print("ANOVA için yeterli mahalle verisi bulunamadı.")
else:
    print("statsmodels kurulu olmadığı için bu bölüm atlandı.")

# ==============================================================================
# 13. MANOVA ANALİZİ (Bulgular 4.8)
# ==============================================================================
print("\n" + "=" * 80)
print("4.8. MANOVA ANALİZİ (WILKS' LAMBDA)")
print("=" * 80)

if HAS_STATSMODELS:
    bagimli_degiskenler = ["Fiyat_TL", "Brut_m2", "Net_m2"]
    bagimsiz_kategorik_degiskenler = ["Krediye_Uygun_Kod", "Takas_Kod", "Kullanim_Durumu_Kod", "Site_Icerisinde_Kod"]
    manova_sonuclari = []

    for ind_var in bagimsiz_kategorik_degiskenler:
        formul = f"Q('{bagimli_degiskenler[0]}') + Q('{bagimli_degiskenler[1]}') + Q('{bagimli_degiskenler[2]}') ~ C({ind_var})"
        try:
            gecici = df[[ind_var] + bagimli_degiskenler].dropna()
            if gecici[ind_var].nunique() > 1:
                manova_model = MANOVA.from_formula(formul, data=gecici)
                sonuc_tablosu = manova_model.mv_test().get_anova_table(iterms=[ind_var])
                manova_sonuclari.append({
                    "Değişken": ind_var,
                    "Wilks Lambda": round(sonuc_tablosu.loc[ind_var, "Wilks' lambda"], 4),
                    "F Değeri": round(sonuc_tablosu.loc[ind_var, "F Value"], 2),
                    "p-Değeri": f"{sonuc_tablosu.loc[ind_var, 'PR(>F)']:.3E}",
                    "Anlamlı mı?": "Evet" if sonuc_tablosu.loc[ind_var, "PR(>F)"] < 0.05 else "Hayır",
                })
        except Exception as hata:
            manova_sonuclari.append({"Değişken": ind_var, "Wilks Lambda": "-", "F Değeri": "-",
                                      "p-Değeri": "-", "Anlamlı mı?": f"Hata: {hata}"})

    tablo_yazdir(pd.DataFrame(manova_sonuclari))
else:
    print("statsmodels kurulu olmadığı için bu bölüm atlandı.")

# ==============================================================================
# 14. MAKİNE ÖĞRENMESİ İÇİN DEĞİŞKEN HAZIRLIĞI (Bulgular 4.10.1)
# ==============================================================================
print("\n" + "=" * 80)
print("4.10.1. MAKİNE ÖĞRENMESİ MODELLERİ İÇİN VERİ HAZIRLIĞI")
print("=" * 80)

numerical_features = ["Brut_m2", "Net_m2", "Oda_Sayisi_Numeric", "Kat_Sayisi_Numeric",
                       "Banyo_Sayisi", "Bina_Yasi_Ortalama", "Bulundugu_Kat_Donusturulmus"]
categorical_features = ["Mahalle", "Isitma", "Krediye_Uygun", "Tapu_Durumu", "Kimden",
                         "Takas", "Site_Adi", "Esyali_Kod", "Kullanim_Durumu_Kod"]
categorical_features = [c for c in categorical_features if c in df.columns]

preprocessor_reg = ColumnTransformer(transformers=[
    ("num", StandardScaler(), numerical_features),
    ("cat", OneHotEncoder(handle_unknown="ignore"), categorical_features),
])

model_karsilastirma = []  # tüm modellerin sonuçlarını toplayacağımız liste


def performans_hesapla(y_gercek, y_tahmin):
    mse = mean_squared_error(y_gercek, y_tahmin)
    return {
        "RMSE": np.sqrt(mse),
        "MAE": mean_absolute_error(y_gercek, y_tahmin),
        "R2": r2_score(y_gercek, y_tahmin),
    }


# ==============================================================================
# 15. BASİT VE ÇOKLU DOĞRUSAL REGRESYON (Bulgular 4.10.1.1)
# ==============================================================================
print("\n--- 4.10.1.1. Basit ve Çoklu Doğrusal Regresyon ---")

# Basit doğrusal regresyon (Brüt m² -> Fiyat)
X_basit, y_basit = df[["Brut_m2"]], df["Fiyat_TL"]
X_train_b, X_test_b, y_train_b, y_test_b = train_test_split(X_basit, y_basit, test_size=0.2, random_state=RANDOM_STATE)
model_basit = LinearRegression().fit(X_train_b, y_train_b)
perf_basit = performans_hesapla(y_test_b, model_basit.predict(X_test_b))
print(f"Basit Doğrusal Regresyon (Brüt m²): R²={perf_basit['R2']:.3f}")
model_karsilastirma.append({"Model": "Basit Doğrusal Regresyon", **perf_basit})

# Çoklu doğrusal regresyon (tüm sayısal değişkenler)
X_coklu, y_coklu = df[numerical_features], df["Fiyat_TL"]
X_train_m, X_test_m, y_train_m, y_test_m = train_test_split(X_coklu, y_coklu, test_size=0.2, random_state=RANDOM_STATE)
model_coklu = LinearRegression().fit(X_train_m, y_train_m)
perf_coklu = performans_hesapla(y_test_m, model_coklu.predict(X_test_m))
print(f"Çoklu Doğrusal Regresyon (Tüm Sayısal Özellikler): R²={perf_coklu['R2']:.3f}")
model_karsilastirma.append({"Model": "Çoklu Doğrusal Regresyon", **perf_coklu})

# ==============================================================================
# 16. GERİYE ELEME (STEPWISE) REGRESYONU (Bulgular 4.10.1.2)
# ==============================================================================
print("\n--- 4.10.1.2. Geriye Eleme (Backward Elimination) Regresyonu ---")

mevcut_ozellikler = list(numerical_features)
en_iyi_r2 = -np.inf
secilen_ozellikler = list(mevcut_ozellikler)

while len(mevcut_ozellikler) > 1:
    en_iyi_cikarilan, gecici_en_iyi_r2 = None, -np.inf
    for cikarilacak in mevcut_ozellikler:
        deneme_ozellikler = [f for f in mevcut_ozellikler if f != cikarilacak]
        X_t = df[deneme_ozellikler]
        X_tr, X_te, y_tr, y_te = train_test_split(X_t, df["Fiyat_TL"], test_size=0.2, random_state=RANDOM_STATE)
        r2_deneme = r2_score(y_te, LinearRegression().fit(X_tr, y_tr).predict(X_te))
        if r2_deneme > gecici_en_iyi_r2:
            gecici_en_iyi_r2, en_iyi_cikarilan = r2_deneme, cikarilacak

    if gecici_en_iyi_r2 < en_iyi_r2:
        break
    en_iyi_r2 = gecici_en_iyi_r2
    mevcut_ozellikler.remove(en_iyi_cikarilan)
    secilen_ozellikler = list(mevcut_ozellikler)
    print(f"Çıkarılan: {en_iyi_cikarilan:28s} | Yeni R²: {en_iyi_r2:.4f} | Kalan: {secilen_ozellikler}")

print(f"\nGeriye eleme ile seçilen nihai değişkenler: {secilen_ozellikler}")

X_nihai = df[secilen_ozellikler]
X_train_f, X_test_f, y_train_f, y_test_f = train_test_split(X_nihai, df["Fiyat_TL"], test_size=0.2, random_state=RANDOM_STATE)
model_nihai = LinearRegression().fit(X_train_f, y_train_f)
y_pred_f = model_nihai.predict(X_test_f)
perf_nihai = performans_hesapla(y_test_f, y_pred_f)
print(f"Nihai Model Performansı -> RMSE={perf_nihai['RMSE']:,.0f} | MAE={perf_nihai['MAE']:,.0f} | R²={perf_nihai['R2']:.3f}")
model_karsilastirma.append({"Model": "Geriye Eleme (Seçilmiş Değişkenler)", **perf_nihai})

if HAS_STATSMODELS:
    X_nihai_sm = sm.add_constant(X_nihai)
    ols_ozet = sm.OLS(df["Fiyat_TL"], X_nihai_sm).fit()
    print("\nKatsayıların Anlamlılığı (OLS Özeti):")
    print(ols_ozet.summary())

    print("\n--- 4.10.1.7. Çoklu Bağlantı (VIF) Analizi ---")
    vif_df = pd.DataFrame()
    vif_df["Değişken"] = X_coklu.columns
    vif_df["VIF"] = [variance_inflation_factor(X_coklu.values, i) for i in range(len(X_coklu.columns))]
    tablo_yazdir(vif_df.round(2))
    print("Not: VIF > 5-10 olması çoklu bağlantı sorununa işaret eder.")

# ==============================================================================
# 17. GERÇEK vs TAHMİN GÖRSELİ (Bulgular 4.10.1.5)
# ==============================================================================
plt.figure(figsize=(8, 6))
plt.scatter(y_test_f, y_pred_f, alpha=0.5)
plt.plot([y_test_f.min(), y_test_f.max()], [y_test_f.min(), y_test_f.max()], "r--", label="İdeal Uyum (y=x)")
plt.title("Gerçek vs Tahmin Edilen Fiyatlar (Geriye Eleme Modeli)")
plt.xlabel("Gerçek Fiyat (TL)")
plt.ylabel("Tahmin Edilen Fiyat (TL)")
plt.legend()
plt.grid(True)
plt.tight_layout()
plt.show()

# ==============================================================================
# 18. POLİNOMİK REGRESYON (Bulgular 4.10.1.2 - polinom kısmı)
# ==============================================================================
print("\n--- 4.10.1.2b. Polinomik Regresyon (1., 2. ve 3. derece) ---")

X_poly_ham, y_poly_ham = df[["Brut_m2"]], df["Fiyat_TL"]
X_train_p, X_test_p, y_train_p, y_test_p = train_test_split(X_poly_ham, y_poly_ham, test_size=0.2, random_state=RANDOM_STATE)

poly_sonuclari = []
plt.figure(figsize=(15, 5))
for i, derece in enumerate([1, 2, 3]):
    poly = PolynomialFeatures(degree=derece, include_bias=False)
    X_train_poly = poly.fit_transform(X_train_p)
    X_test_poly = poly.transform(X_test_p)

    model_poly = LinearRegression().fit(X_train_poly, y_train_p)
    y_pred_poly = model_poly.predict(X_test_poly)
    perf_poly = performans_hesapla(y_test_p, y_pred_poly)
    poly_sonuclari.append({"Polinom Derecesi": derece, **perf_poly})

    plt.subplot(1, 3, i + 1)
    plt.scatter(X_test_p["Brut_m2"], y_test_p, alpha=0.5, label="Gerçek")
    X_cizgi = np.linspace(X_poly_ham["Brut_m2"].min(), X_poly_ham["Brut_m2"].max(), 300).reshape(-1, 1)
    plt.plot(X_cizgi, model_poly.predict(poly.transform(X_cizgi)), color="red", label=f"{derece}. derece")
    plt.title(f"{derece}. Derece Polinom (R²={perf_poly['R2']:.3f})")
    plt.xlabel("Brüt m²")
    plt.ylabel("Fiyat (TL)")
    plt.legend()

plt.tight_layout()
plt.show()
tablo_yazdir(pd.DataFrame(poly_sonuclari).round(3))
en_iyi_poly = {k: v for k, v in poly_sonuclari[1].items() if k != "Polinom Derecesi"}
model_karsilastirma.append({"Model": "Polinomik Regresyon (2. derece)", **en_iyi_poly})

# ==============================================================================
# 19. ÜSTEL REGRESYON VE DÜZENLİLEŞTİRME (Bulgular 4.10.1.3)
# ==============================================================================
print("\n--- 4.10.1.3. Üstel Regresyon ve Düzenlileştirme (Ridge / Lasso / ElasticNet) ---")

X_exp = df[["Brut_m2"]]
y_exp_log = np.log1p(df["Fiyat_TL"])
X_train_e, X_test_e, y_train_e, y_test_e = train_test_split(X_exp, y_exp_log, test_size=0.2, random_state=RANDOM_STATE)
model_exp = LinearRegression().fit(X_train_e, y_train_e)
y_pred_exp = np.expm1(model_exp.predict(X_test_e))
y_test_exp_gercek = np.expm1(y_test_e)
perf_exp = performans_hesapla(y_test_exp_gercek, y_pred_exp)
print(f"Üstel Regresyon -> R²={perf_exp['R2']:.3f} | RMSE={perf_exp['RMSE']:,.0f} | MAE={perf_exp['MAE']:,.0f}")
model_karsilastirma.append({"Model": "Üstel Regresyon", **perf_exp})

X_reg, y_reg = df[numerical_features + categorical_features], df["Fiyat_TL"]
X_train_r, X_test_r, y_train_r, y_test_r = train_test_split(X_reg, y_reg, test_size=0.2, random_state=RANDOM_STATE)

duzenlilestirme_modelleri = {
    "Ridge": (Ridge(random_state=RANDOM_STATE), {"regressor__alpha": [0.1, 1, 10, 100]}),
    "Lasso": (Lasso(random_state=RANDOM_STATE, max_iter=3000), {"regressor__alpha": [0.01, 0.1, 1, 10]}),
    "Elastic Net": (ElasticNet(random_state=RANDOM_STATE, max_iter=3000),
                    {"regressor__alpha": [0.1, 1, 10], "regressor__l1_ratio": [0.1, 0.5, 0.9]}),
}

for ad, (tahminci, param_grid) in duzenlilestirme_modelleri.items():
    pipeline = Pipeline([("preprocessor", preprocessor_reg), ("regressor", tahminci)])
    arama = GridSearchCV(pipeline, param_grid, cv=5, scoring="r2", n_jobs=-1)
    arama.fit(X_train_r, y_train_r)
    perf = performans_hesapla(y_test_r, arama.best_estimator_.predict(X_test_r))
    print(f"{ad:12s} -> En iyi parametreler: {arama.best_params_} | R²={perf['R2']:.3f}")
    model_karsilastirma.append({"Model": ad, **perf})

# ==============================================================================
# 20. K-EN YAKIN KOMŞU (K-NN) REGRESYONU (Bulgular 4.10.1.4)
# ==============================================================================
print("\n--- 4.10.1.4. K-En Yakın Komşu (K-NN) Regresyonu ---")

X_knn, y_knn = df[numerical_features + categorical_features], df["Fiyat_TL"]
X_train_k, X_test_k, y_train_k, y_test_k = train_test_split(X_knn, y_knn, test_size=0.2, random_state=RANDOM_STATE)

knn_pipeline = Pipeline([("preprocessor", preprocessor_reg), ("regressor", KNeighborsRegressor())])
knn_arama = GridSearchCV(knn_pipeline, {"regressor__n_neighbors": range(1, 21)}, cv=5, scoring="r2", n_jobs=-1)
knn_arama.fit(X_train_k, y_train_k)
en_iyi_k = knn_arama.best_params_["regressor__n_neighbors"]
perf_knn = performans_hesapla(y_test_k, knn_arama.best_estimator_.predict(X_test_k))
print(f"En iyi K değeri: {en_iyi_k} | R²={perf_knn['R2']:.3f} | RMSE={perf_knn['RMSE']:,.0f} | MAE={perf_knn['MAE']:,.0f}")
model_karsilastirma.append({"Model": f"K-NN (k={en_iyi_k})", **perf_knn})

# ==============================================================================
# 21. RANDOM FOREST REGRESYONU (Bulgular 4.10.1.5)
# ==============================================================================
print("\n--- 4.10.1.5. Random Forest Regresyonu ---")

X_rf, y_rf = df[numerical_features + categorical_features], df["Fiyat_TL"]
X_train_rf, X_test_rf, y_train_rf, y_test_rf = train_test_split(X_rf, y_rf, test_size=0.2, random_state=RANDOM_STATE)

rf_pipeline = Pipeline([("preprocessor", preprocessor_reg),
                         ("regressor", RandomForestRegressor(n_estimators=200, random_state=RANDOM_STATE, n_jobs=-1))])
rf_pipeline.fit(X_train_rf, y_train_rf)
y_pred_rf = rf_pipeline.predict(X_test_rf)
perf_rf = performans_hesapla(y_test_rf, y_pred_rf)
print(f"Random Forest -> R²={perf_rf['R2']:.3f} | RMSE={perf_rf['RMSE']:,.0f} | MAE={perf_rf['MAE']:,.0f}")
model_karsilastirma.append({"Model": "Random Forest", **perf_rf})

ohe_ozellikleri = rf_pipeline.named_steps["preprocessor"].named_transformers_["cat"].get_feature_names_out(categorical_features)
tum_ozellikler = list(numerical_features) + list(ohe_ozellikleri)
onem_siralamasi = pd.Series(rf_pipeline.named_steps["regressor"].feature_importances_, index=tum_ozellikler)
print("\nEn Önemli 10 Özellik (Random Forest):")
tablo_yazdir(onem_siralamasi.nlargest(10).to_frame("Önem"))

plt.figure(figsize=(8, 6))
plt.scatter(y_test_rf, y_pred_rf, alpha=0.5, color="seagreen")
plt.plot([y_test_rf.min(), y_test_rf.max()], [y_test_rf.min(), y_test_rf.max()], "r--", label="İdeal Uyum (y=x)")
plt.title("Gerçek vs Tahmin Edilen Fiyatlar (Random Forest)")
plt.xlabel("Gerçek Fiyat (TL)")
plt.ylabel("Tahmin Edilen Fiyat (TL)")
plt.legend()
plt.grid(True)
plt.tight_layout()
plt.show()

# ==============================================================================
# 22. BOOSTING ALGORİTMALARI: GRADIENT BOOSTING VE XGBOOST (Bulgular 4.10.1.6)
# ==============================================================================
print("\n--- 4.10.1.6. Boosting Algoritmaları ---")

gb_pipeline = Pipeline([("preprocessor", preprocessor_reg),
                         ("regressor", GradientBoostingRegressor(n_estimators=100, learning_rate=0.1,
                                                                  max_depth=3, random_state=RANDOM_STATE))])
gb_pipeline.fit(X_train_rf, y_train_rf)
perf_gb = performans_hesapla(y_test_rf, gb_pipeline.predict(X_test_rf))
print(f"Gradient Boosting -> R²={perf_gb['R2']:.3f} | RMSE={perf_gb['RMSE']:,.0f} | MAE={perf_gb['MAE']:,.0f}")
model_karsilastirma.append({"Model": "Gradient Boosting", **perf_gb})

if HAS_XGBOOST:
    xgb_pipeline = Pipeline([("preprocessor", preprocessor_reg),
                              ("regressor", xgb.XGBRegressor(objective="reg:squarederror", n_estimators=100,
                                                              learning_rate=0.1, max_depth=3,
                                                              random_state=RANDOM_STATE, n_jobs=-1))])
    xgb_pipeline.fit(X_train_rf, y_train_rf)
    perf_xgb = performans_hesapla(y_test_rf, xgb_pipeline.predict(X_test_rf))
    print(f"XGBoost -> R²={perf_xgb['R2']:.3f} | RMSE={perf_xgb['RMSE']:,.0f} | MAE={perf_xgb['MAE']:,.0f}")
    model_karsilastirma.append({"Model": "XGBoost", **perf_xgb})
else:
    print("xgboost kurulu olmadığı için bu model atlandı.")

# ==============================================================================
# 23. YAPAY SİNİR AĞLARI (ANN) (Bulgular 4.10.1)
# ==============================================================================
print("\n--- 4.10.1. Yapay Sinir Ağları (ANN) ---")

if HAS_TENSORFLOW:
    X_ann, y_ann = df[numerical_features + categorical_features], df["Fiyat_TL"]
    X_train_a, X_test_a, y_train_a, y_test_a = train_test_split(X_ann, y_ann, test_size=0.2, random_state=RANDOM_STATE)

    X_train_islenmis = preprocessor_reg.fit_transform(X_train_a)
    X_test_islenmis = preprocessor_reg.transform(X_test_a)

    ann_model = Sequential([
        Dense(128, activation="relu", input_shape=(X_train_islenmis.shape[1],)),
        Dropout(0.3),
        Dense(64, activation="relu"),
        Dropout(0.3),
        Dense(32, activation="relu"),
        Dense(1),
    ])
    ann_model.compile(optimizer=Adam(learning_rate=0.001), loss="mse", metrics=["mae"])
    durdurma = EarlyStopping(monitor="val_loss", patience=10, restore_best_weights=True)

    gecmis = ann_model.fit(X_train_islenmis, y_train_a, epochs=100, batch_size=32,
                            validation_data=(X_test_islenmis, y_test_a), callbacks=[durdurma], verbose=0)

    y_pred_ann = ann_model.predict(X_test_islenmis).flatten()
    perf_ann = performans_hesapla(y_test_a, y_pred_ann)
    print(f"ANN -> R²={perf_ann['R2']:.3f} | RMSE={perf_ann['RMSE']:,.0f} | MAE={perf_ann['MAE']:,.0f}")
    model_karsilastirma.append({"Model": "Yapay Sinir Ağı (ANN)", **perf_ann})

    plt.figure(figsize=(9, 5))
    plt.plot(gecmis.history["loss"], label="Eğitim Kaybı")
    plt.plot(gecmis.history["val_loss"], label="Doğrulama Kaybı")
    plt.title("ANN Eğitim ve Doğrulama Kaybı")
    plt.xlabel("Epok")
    plt.ylabel("Kayıp (MSE)")
    plt.legend()
    plt.grid(True)
    plt.tight_layout()
    plt.show()
else:
    print("tensorflow kurulu olmadığı için bu bölüm atlandı.")

# ==============================================================================
# 24. TÜM MODELLERİN KARŞILAŞTIRMALI PERFORMANS TABLOSU
# ==============================================================================
print("\n" + "=" * 80)
print("24. TÜM MODELLERİN KARŞILAŞTIRMALI PERFORMANSI")
print("=" * 80)

karsilastirma_df = pd.DataFrame(model_karsilastirma)
karsilastirma_df["RMSE"] = karsilastirma_df["RMSE"].map(lambda x: f"{x:,.0f}")
karsilastirma_df["MAE"] = karsilastirma_df["MAE"].map(lambda x: f"{x:,.0f}")
karsilastirma_df["R2"] = karsilastirma_df["R2"].map(lambda x: f"{x:.3f}")
tablo_yazdir(karsilastirma_df)

print("\nAnaliz tamamlandı.")
