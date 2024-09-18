import requests
from bs4 import BeautifulSoup
url = "https://docs.google.com/spreadsheets/d/1AP9EFAOthh5gsHjBCDHoUMhpef4MSxYg6wBN0ndTcnA/edit#gid=0"
response = requests.get(url)
html_content = response.content
soup = BeautifulSoup(html_content, "html.parser")
first_cell = soup.find("td", {"class": "s2"}).text.strip()
if first_cell != "Aktif":
    exit()
first_cell = soup.find("td", {"class": "s1"}).text.strip()
print(first_cell)


import pandas as pd
import re
from io import BytesIO
import os
import numpy as np
import shutil
from concurrent.futures import ThreadPoolExecutor
from tqdm import tqdm
import xml.etree.ElementTree as ET
import warnings
from colorama import init, Fore, Style

warnings.filterwarnings("ignore")



print("Oturum Açma Başarılı Oldu")
print(" /﹋\ ")
print("(҂`_´)")
print(Fore.RED + "<,︻╦╤─ ҉ - -")
print("/﹋\\")
print("Mustafa ARI")
print(" ")
print("Kara Liste Hazırlama")

pd.options.mode.chained_assignment = None
'''

# İndirilecek linklerin temel kısmı
base_url = "https://task.haydigiy.com/FaprikaOrderXls/ATU1MK/"

# Linklerin sonundaki sayıların listesi
links = [str(i) for i in range(1, 31)]

# Tüm verileri birleştirmek için kullanılacak boş bir DataFrame
merged_data = pd.DataFrame()

# Her link için işlem yapma döngüsü
for link in links:
    # Linki oluştur
    url = base_url + link
    
    # Linkten istek gönder ve yanıtı al
    response = requests.get(url)
    
    # Yanıtın durum kodunu kontrol et
    if response.status_code == 200:
        # Yanıtı Excel dosyası olarak oku
        excel_data = pd.read_excel(response.content)
        
        # İnen verileri birleştir
        merged_data = pd.concat([merged_data, excel_data], ignore_index=True)
        
    else:
        print(f"Hata! {url} indirilemedi. Durum kodu: {response.status_code}")

# Tüm verileri birleştir ve tek bir Excel dosyasına yaz
merged_data.to_excel("birlesmis_veriler.xlsx", index=False)

'''


# Excel dosyasını oku
excel_data = pd.read_excel("birlesmis_veriler.xlsx")

# Sadece belirli sütunları koru
filtered_data = excel_data[['SiparisDurumu', 'OdemeTipi', 'TeslimatAdiSoyadi', 'TeslimatTelefon']]

# Sonuçları yeni bir Excel dosyasına yaz
filtered_data.to_excel("birlesmis_veriler.xlsx", index=False)





# Excel dosyasını oku
excel_data = pd.read_excel("birlesmis_veriler.xlsx")

# "OdemeTipi" sütununda "Kredi Kartı" içeren ve "SiparisDurumu" sütununda "Teslim Edilmeyen Kargo" içeren satırları filtrele
filtered_data = excel_data[~((excel_data['OdemeTipi'] == 'Kredi Kartı') & (excel_data['SiparisDurumu'] == 'Teslim Edilmeyen Kargo'))]

# "OdemeTipi" sütununu sil
filtered_data.drop(columns=['OdemeTipi'], inplace=True)

# Filtrelenmiş verileri aynı Excel dosyasının üstüne kaydet
filtered_data.to_excel("birlesmis_veriler.xlsx", index=False)









# Excel dosyasını oku
excel_data = pd.read_excel("birlesmis_veriler.xlsx")

# "TeslimatTelefon" sütunundaki tüm veriler için işlem yap
excel_data['TeslimatTelefon'] = excel_data['TeslimatTelefon'].astype(str).str.replace(r'[^0-9]', '')  # 0-9 arası olmayan karakterleri temizle

# "TeslimatTelefon" sütunundaki tüm veriler için işlem yap
excel_data['TeslimatTelefon'] = excel_data['TeslimatTelefon'].astype(str).str.replace("+90", "")
excel_data['TeslimatTelefon'] = excel_data['TeslimatTelefon'].astype(str).str.replace(")", "")
excel_data['TeslimatTelefon'] = excel_data['TeslimatTelefon'].astype(str).str.replace("(", "")
excel_data['TeslimatTelefon'] = excel_data['TeslimatTelefon'].astype(str).str.replace("/", "")
excel_data['TeslimatTelefon'] = excel_data['TeslimatTelefon'].astype(str).str.replace("-", "")
excel_data['TeslimatTelefon'] = excel_data['TeslimatTelefon'].astype(str).str.replace(" ", "")
excel_data['TeslimatTelefon'] = excel_data['TeslimatTelefon'].astype(str).str.replace("*", "")
excel_data['TeslimatTelefon'] = excel_data['TeslimatTelefon'].astype(str).str.replace("_", "")
excel_data['TeslimatTelefon'] = excel_data['TeslimatTelefon'].astype(str).str.replace("+", "")


# "9" ile başlayan hücrelerin ilk iki hanesini temizle
excel_data.loc[excel_data['TeslimatTelefon'].str.startswith('9'), 'TeslimatTelefon'] = excel_data['TeslimatTelefon'].str[2:]

# Verileri aynı Excel dosyasının üstüne kaydet
excel_data.to_excel("birlesmis_veriler.xlsx", index=False)



# "birlesmis_veriler.xlsx" excel dosyasını oku
birlesmis_veriler = pd.read_excel("birlesmis_veriler.xlsx")

# "TeslimatTelefon" sütunundaki verileri sayıya dönüştür
birlesmis_veriler['TeslimatTelefon'] = pd.to_numeric(birlesmis_veriler['TeslimatTelefon'], errors='coerce')

# Sonucu aynı Excel dosyasına kaydet
birlesmis_veriler.to_excel("birlesmis_veriler.xlsx", index=False)




# Excel dosyasını oku
excel_data = pd.read_excel("birlesmis_veriler.xlsx")

# "İsim Soyisim" adında yeni bir sütun oluştur
excel_data['İsim Soyisim'] = ''

# "TeslimatTelefon" sütunundaki verileri kullanarak "TeslimatAdiSoyadi" sütunundaki verileri arayıp eşleşenleri "İsim Soyisim" sütununa yaz
for index, row in excel_data.iterrows():
    phone_number = row['TeslimatTelefon']
    matching_row = excel_data[excel_data['TeslimatTelefon'] == phone_number]
    if not matching_row.empty:
        name_surname = matching_row.iloc[0]['TeslimatAdiSoyadi']
        excel_data.at[index, 'İsim Soyisim'] = name_surname


# "TeslimatAdiSoyadi" sütununu sil
excel_data.drop(columns=['TeslimatAdiSoyadi'], inplace=True)

# Verileri aynı Excel dosyasının üstüne kaydet
excel_data.to_excel("birlesmis_veriler.xlsx", index=False)





# Excel dosyasını oku
excel_data = pd.read_excel("birlesmis_veriler.xlsx")

# "İsim Soyisim" sütunundaki verilerin sadece baş harfleri büyük gerisi küçük harfe dönüştür
excel_data['İsim Soyisim'] = excel_data['İsim Soyisim'].str.title()

# Verileri aynı Excel dosyasının üstüne kaydet
excel_data.to_excel("birlesmis_veriler.xlsx", index=False)




# Excel dosyasını oku
excel_data = pd.read_excel("birlesmis_veriler.xlsx")

# "Tamamlanan" adında yeni bir sütun oluştur
excel_data['Tamamlanan'] = ''

# Her bir "TeslimatTelefon" için "SiparisDurumu" sütununda "Tamamlandı" sayısını hesapla
for index, row in excel_data.iterrows():
    phone_number = row['TeslimatTelefon']
    completed_count = excel_data[(excel_data['TeslimatTelefon'] == phone_number) & (excel_data['SiparisDurumu'] == 'Tamamlandı')].shape[0]
    excel_data.at[index, 'Tamamlanan'] = completed_count

# Verileri aynı Excel dosyasının üstüne kaydet
excel_data.to_excel("birlesmis_veriler.xlsx", index=False)





# Excel dosyasını oku
excel_data = pd.read_excel("birlesmis_veriler.xlsx")

# "Tamamlanan" adında yeni bir sütun oluştur
excel_data['Teslim Edilmeyen'] = ''

# Her bir "TeslimatTelefon" için "SiparisDurumu" sütununda "Tamamlandı" sayısını hesapla
for index, row in excel_data.iterrows():
    phone_number = row['TeslimatTelefon']
    completed_count = excel_data[(excel_data['TeslimatTelefon'] == phone_number) & (excel_data['SiparisDurumu'] == 'Teslim Edilmeyen Kargo')].shape[0]
    excel_data.at[index, 'Teslim Edilmeyen'] = completed_count

# Verileri aynı Excel dosyasının üstüne kaydet
excel_data.to_excel("birlesmis_veriler.xlsx", index=False)





# Excel dosyasını oku
excel_data = pd.read_excel("birlesmis_veriler.xlsx")

# "SiparisDurumu" sütununu sil
excel_data.drop(columns=['SiparisDurumu'], inplace=True)

# Tüm tablodaki verileri benzersiz yap
excel_data = excel_data.drop_duplicates()

# Verileri aynı Excel dosyasının üstüne kaydet
excel_data.to_excel("birlesmis_veriler.xlsx", index=False)




# Excel dosyasını indirmek için URL
url = "https://drive.google.com/uc?id=1qIOaP_RWvSxwas3dtAY-xWeNmoLtvyCa&export=download"

# İstek gönder ve yanıtı al
response = requests.get(url)

# Yanıtın durum kodunu kontrol et
if response.status_code == 200:
    # Yanıtı Excel dosyası olarak oku
    excel_data = pd.read_excel(response.content)
    
    # Tüm tablodaki verileri benzersiz yap
    excel_data = excel_data.drop_duplicates()
    
    # Verileri yeni bir Excel dosyasına kaydet
    excel_data.to_excel("indirilen_veriler.xlsx", index=False)







# birlesmis_veriler excel dosyasını oku
birlesmis_veriler = pd.read_excel("birlesmis_veriler.xlsx")

# indirilen_veriler excel dosyasını oku
indirilen_veriler = pd.read_excel("indirilen_veriler.xlsx")

# indirilen_veriler excelindeki "TeslimatTelefon" sütunundaki verileri birleşmiş verilerde ara
for index, row in indirilen_veriler.iterrows():
    phone_number = row['TeslimatTelefon']
    matching_row = birlesmis_veriler[birlesmis_veriler['TeslimatTelefon'] == phone_number]
    if not matching_row.empty:
        tamamlanan = matching_row.iloc[0]['Tamamlanan']
        indirilen_veriler.at[index, 'Yeni Teslim Adedi'] = tamamlanan
    else:
        indirilen_veriler.at[index, 'Yeni Teslim Adedi'] = 0

# Sonuçları yeni bir Excel dosyasına kaydet
indirilen_veriler.to_excel("indirilen_veriler.xlsx", index=False)







# birlesmis_veriler excel dosyasını oku
birlesmis_veriler = pd.read_excel("birlesmis_veriler.xlsx")

# indirilen_veriler excel dosyasını oku
indirilen_veriler = pd.read_excel("indirilen_veriler.xlsx")

# indirilen_veriler excelindeki "TeslimatTelefon" sütunundaki verileri birleşmiş verilerde ara
for index, row in indirilen_veriler.iterrows():
    phone_number = row['TeslimatTelefon']
    matching_row = birlesmis_veriler[birlesmis_veriler['TeslimatTelefon'] == phone_number]
    if not matching_row.empty:
        tamamlanan = matching_row.iloc[0]['Teslim Edilmeyen']
        indirilen_veriler.at[index, 'Yeni Kek Adedi'] = tamamlanan
    else:
        indirilen_veriler.at[index, 'Yeni Kek Adedi'] = 0

# Sonuçları yeni bir Excel dosyasına kaydet
indirilen_veriler.to_excel("indirilen_veriler.xlsx", index=False)






# birlesmis_veriler excel dosyasını oku
birlesmis_veriler = pd.read_excel("birlesmis_veriler.xlsx")

# indirilen_veriler excel dosyasını oku
indirilen_veriler = pd.read_excel("indirilen_veriler.xlsx")

# birlesmis_veriler excelindeki "TeslimatTelefon" sütunundaki verileri indirilen_veriler excelinde ara
for index, row in birlesmis_veriler.iterrows():
    phone_number = row['TeslimatTelefon']
    matching_row = indirilen_veriler[indirilen_veriler['TeslimatTelefon'] == phone_number]
    if not matching_row.empty:
        tamamlanan = matching_row.iloc[0]['Tamamlanan']
        birlesmis_veriler.at[index, 'Yeni mi'] = tamamlanan
    else:
        birlesmis_veriler.at[index, 'Yeni mi'] = "Evet"

# Sonuçları yeni bir Excel dosyasına kaydet
birlesmis_veriler.to_excel("birlesmis_veriler.xlsx", index=False)






# birlesmis_veriler excel dosyasını oku
birlesmis_veriler = pd.read_excel("birlesmis_veriler.xlsx")

# "Yeni mi" sütununda sayısal olan satırları filtrele ve sil
birlesmis_veriler = birlesmis_veriler[pd.to_numeric(birlesmis_veriler['Yeni mi'], errors='coerce').isnull()]

# Sonuçları yeni bir Excel dosyasına kaydet
birlesmis_veriler.to_excel("birlesmis_veriler.xlsx", index=False)



# birlesmis_veriler excel dosyasını oku
birlesmis_veriler = pd.read_excel("birlesmis_veriler.xlsx")

# indirilen_veriler excel dosyasını oku
indirilen_veriler = pd.read_excel("indirilen_veriler.xlsx")

# "indirilen_veriler" dosyasına "birlesmis_veriler" dosyasındaki verileri altına ekle
indirilen_veriler = pd.concat([indirilen_veriler, birlesmis_veriler], ignore_index=True)

# Sonuçları yeni bir Excel dosyasına kaydet
indirilen_veriler.to_excel("indirilen_veriler.xlsx", index=False)



# birlesmis_veriler excel dosyasını oku
indirilen_veriler = pd.read_excel("indirilen_veriler.xlsx")

# "Yeni Teslim Adedi" ve "Yeni Kek Adedi" sütunlarında boş olan hücreleri 0 ile doldur
indirilen_veriler['Yeni Teslim Adedi'].fillna(0, inplace=True)
indirilen_veriler['Yeni Kek Adedi'].fillna(0, inplace=True)

# Sonuçları aynı Excel dosyasına kaydet
indirilen_veriler.to_excel("indirilen_veriler.xlsx", index=False)



# indirilen_veriler excel dosyasını oku
indirilen_veriler = pd.read_excel("indirilen_veriler.xlsx")

# Tamamlanan sütununu güncelle: Tamamlanan = Tamamlanan + Yeni Teslim Adedi
indirilen_veriler['Tamamlanan'] = indirilen_veriler['Tamamlanan'] + indirilen_veriler['Yeni Teslim Adedi']

# Sonuçları aynı Excel dosyasına kaydet
indirilen_veriler.to_excel("indirilen_veriler.xlsx", index=False)



# indirilen_veriler excel dosyasını oku
indirilen_veriler = pd.read_excel("indirilen_veriler.xlsx")

# Tamamlanan sütununu güncelle: Tamamlanan = Tamamlanan + Yeni Teslim Adedi
indirilen_veriler['Teslim Edilmeyen'] = indirilen_veriler['Teslim Edilmeyen'] + indirilen_veriler['Yeni Kek Adedi']

# Sonuçları aynı Excel dosyasına kaydet
indirilen_veriler.to_excel("indirilen_veriler.xlsx", index=False)



# indirilen_veriler excel dosyasını oku
indirilen_veriler = pd.read_excel("indirilen_veriler.xlsx")

# Toplam sütununu güncelle: Toplam = Tamamlanan + Teslim Edilmeyen
indirilen_veriler['Toplam'] = indirilen_veriler['Tamamlanan'] + indirilen_veriler['Teslim Edilmeyen']

# Sonuçları aynı Excel dosyasına kaydet
indirilen_veriler.to_excel("indirilen_veriler.xlsx", index=False)





# Excel dosyasını indirmek için URL
url = "https://drive.usercontent.google.com/u/0/uc?id=1kTNzP7P5YXWaMFwi6s2XvODFy5scerZ8&export=download"

# İstek gönder ve yanıtı al
response = requests.get(url)

# Yanıtın durum kodunu kontrol et
if response.status_code == 200:
    # Yanıtı Excel dosyası olarak oku
    excel_data = pd.read_excel(response.content)
    
    # Tüm tablodaki verileri benzersiz yap
    excel_data = excel_data.drop_duplicates()
    
    # Verileri yeni bir Excel dosyasına kaydet
    excel_data.to_excel("Orana Göre Durumlar.xlsx", index=False)






# birlesmis_veriler excel dosyasını oku
indirilen_veriler = pd.read_excel("indirilen_veriler.xlsx")

# Teslim Oranı sütununu oluştur: Toplam - Tamamlanan
indirilen_veriler['Teslim Oranı'] = indirilen_veriler['Toplam'].astype(str) + " - " + indirilen_veriler['Tamamlanan'].astype(str)

# Sonuçları aynı Excel dosyasına kaydet
indirilen_veriler.to_excel("indirilen_veriler.xlsx", index=False)




# indirilen_veriler excel dosyasını oku
indirilen_veriler = pd.read_excel("indirilen_veriler.xlsx")

# Orana Göre Durumlar excel dosyasını oku
oranlar_durumlar = pd.read_excel("Orana Göre Durumlar.xlsx")

# indirilen_verilerdeki her bir teslim oranı için döngü
for index, row in indirilen_veriler.iterrows():
    teslim_orani = row['Teslim Oranı']
    # Orana Göre Durumlar excelinde teslim oranını ara
    durum = oranlar_durumlar.loc[oranlar_durumlar['Teslim Oranı'] == teslim_orani, 'Durum'].values
    # Eğer karşılık gelen bir durum bulunamazsa, 'Bilinmiyor' yaz
    if len(durum) == 0:
        indirilen_veriler.at[index, 'Yeni Durum'] = 'VİP'
    else:
        indirilen_veriler.at[index, 'Yeni Durum'] = durum[0]

# Sonuçları aynı Excel dosyasına kaydet
indirilen_veriler.to_excel("indirilen_veriler.xlsx", index=False)




# indirilen_veriler excel dosyasını oku
indirilen_veriler = pd.read_excel("indirilen_veriler.xlsx")

# "Teslim Oranı" sütununu sil
indirilen_veriler.drop(columns=["Teslim Oranı"], inplace=True)

# Sonuçları aynı Excel dosyasına kaydet
indirilen_veriler.to_excel("indirilen_veriler.xlsx", index=False)



# birlesmis_veriler excel dosyasını oku
indirilen_veriler = pd.read_excel("indirilen_veriler.xlsx")

# Teslim Oranı sütununu oluştur: Toplam - Tamamlanan
indirilen_veriler['Son Durum'] = indirilen_veriler['İsim Soyisim'].astype(str) + " / " + indirilen_veriler['Toplam'].astype(str) + " - " + indirilen_veriler['Tamamlanan'].astype(str) + " / " + indirilen_veriler['Yeni Durum'].astype(str)

# Sonuçları aynı Excel dosyasına kaydet
indirilen_veriler.to_excel("indirilen_veriler.xlsx", index=False)





# indirilen_veriler excel dosyasını oku
indirilen_veriler = pd.read_excel("indirilen_veriler.xlsx")

# "Yeni mi" sütununda verisi boş olan ve "Eski Durum" ile "Yeni Durum" sütunlarındaki veri birbirinden farklı olan satırları seç
degisenler = indirilen_veriler[(indirilen_veriler['Yeni mi'].isna()) & (indirilen_veriler['Eski Durum'] != indirilen_veriler['Yeni Durum'])]

# Seçilen satırları "Durumu Değişenler.xlsx" adlı bir Excel dosyasına kaydet
degisenler.to_excel("Durumu Değişenler.xlsx", index=False)



# indirilen_veriler excel dosyasını oku
indirilen_veriler = pd.read_excel("indirilen_veriler.xlsx")

# "Yeni mi" sütununda verisi dolu olan satırları seç
yeni_musteriler = indirilen_veriler[indirilen_veriler['Yeni mi'].notna()]

# Seçilen satırları "Yeni Müşteriler.xlsx" adlı bir Excel dosyasına kaydet
yeni_musteriler.to_excel("Yeni Müşteriler.xlsx", index=False)




# indirilen_veriler excel dosyasını oku
indirilen_veriler = pd.read_excel("indirilen_veriler.xlsx")

# "Yeni Durum" sütununda "Kara Liste" olan satırları seç
kara_liste = indirilen_veriler[indirilen_veriler['Yeni Durum'] == 'Kara Liste']

# Seçilen satırları "Kara Liste.xlsx" adlı bir Excel dosyasına kaydet
kara_liste.to_excel("Kara Liste.exe.xlsx", index=False)



# "Durumu Değişenler" excel dosyasını oku
durumu_degisenler = pd.read_excel("Durumu Değişenler.xlsx")

# "Yeni Müşteriler" excel dosyasını oku
yeni_musteriler = pd.read_excel("Yeni Müşteriler.xlsx")

# "Yeni Müşteriler" excel dosyasını oku
kara_liste = pd.read_excel("Kara Liste.exe.xlsx")

# "TeslimatTelefon" ve "Son Durum" sütunları hariç diğer tüm sütunları sil
durumu_degisenler = durumu_degisenler[['TeslimatTelefon', 'Son Durum']]
yeni_musteriler = yeni_musteriler[['TeslimatTelefon', 'Son Durum']]
kara_liste = kara_liste[['TeslimatTelefon']]

# Sonuçları aynı Excel dosyalarına kaydet
durumu_degisenler.to_excel("Durumu Değişenler.xlsx", index=False)
yeni_musteriler.to_excel("Yeni Müşteriler.xlsx", index=False)
kara_liste.to_excel("Kara Liste.exe.xlsx", index=False)







# "Durumu Değişenler" excel dosyasını oku
durumu_degisenler = pd.read_excel("Durumu Değişenler.xlsx")

# "Yeni Müşteriler" excel dosyasını oku
yeni_musteriler = pd.read_excel("Yeni Müşteriler.xlsx")

# "Yeni Müşteriler" excel dosyasını oku
kara_liste = pd.read_excel("Kara Liste.exe.xlsx")

# "TeslimatTelefon" sütunundaki verilerin başına "90" ekleyelim
durumu_degisenler['TeslimatTelefon'] = '90' + durumu_degisenler['TeslimatTelefon'].astype(str)
yeni_musteriler['TeslimatTelefon'] = '90' + yeni_musteriler['TeslimatTelefon'].astype(str)
kara_liste['TeslimatTelefon'] = '90' + kara_liste['TeslimatTelefon'].astype(str)

# Sonuçları aynı Excel dosyalarına kaydet
durumu_degisenler.to_excel("Durumu Değişenler.xlsx", index=False)
yeni_musteriler.to_excel("Yeni Müşteriler.xlsx", index=False)
kara_liste.to_excel("Kara Liste.exe.xlsx", index=False)





# Excel dosyasını indirmek için URL
url = "https://drive.usercontent.google.com/u/0/uc?id=1hKxfp8Chp5zLHX6medKd6x7CydgxpSXE&export=download"

# İstek gönder ve yanıtı al
response = requests.get(url)

# Yanıtın durum kodunu kontrol et
if response.status_code == 200:
    # Yanıtı Excel dosyası olarak oku
    excel_data = pd.read_excel(response.content)
    
    # Tüm tablodaki verileri benzersiz yap
    excel_data = excel_data.drop_duplicates()
    
    # Verileri yeni bir Excel dosyasına kaydet
    excel_data.to_excel("1000 TL Üzeri İçin Sipariş Oranları ve Kurgu.xlsx", index=False)



# Orjinal dosyanın adı
orijinal_dosya = "indirilen_veriler.xlsx"

# Kopyalanacak dosyanın adı
kopya_dosya = "1000 TL için Kurgu.xlsx"

# Dosyayı kopyala ve adını değiştir
shutil.copyfile(orijinal_dosya, kopya_dosya)




# "1000 TL için Kurgu.xlsx" excel dosyasını oku
kurgu_excel = pd.read_excel("1000 TL için Kurgu.xlsx")

# "TeslimatTelefon", "Tamamlanan" ve "Toplam" sütunları hariç diğer tüm sütunları sil
kurgu_excel = kurgu_excel[['TeslimatTelefon', 'Tamamlanan', 'Toplam']]

# Sonuçları aynı Excel dosyasına kaydet
kurgu_excel.to_excel("1000 TL için Kurgu.xlsx", index=False)





# birlesmis_veriler excel dosyasını oku
indirilen_veriler = pd.read_excel("1000 TL için Kurgu.xlsx")

# Teslim Oranı sütununu oluştur: Toplam - Tamamlanan
indirilen_veriler['Teslim Oranı'] = indirilen_veriler['Toplam'].astype(str) + " - " + indirilen_veriler['Tamamlanan'].astype(str)

# Sonuçları aynı Excel dosyasına kaydet
indirilen_veriler.to_excel("1000 TL için Kurgu.xlsx", index=False)




# indirilen_veriler excel dosyasını oku
indirilen_veriler = pd.read_excel("1000 TL için Kurgu.xlsx")

# Orana Göre Durumlar excel dosyasını oku
oranlar_durumlar = pd.read_excel("1000 TL Üzeri İçin Sipariş Oranları ve Kurgu.xlsx")

# indirilen_verilerdeki her bir teslim oranı için döngü
for index, row in indirilen_veriler.iterrows():
    teslim_orani = row['Teslim Oranı']
    # Orana Göre Durumlar excelinde teslim oranını ara
    durum = oranlar_durumlar.loc[oranlar_durumlar['Teslim Oranı'] == teslim_orani, 'Durum'].values
    # Eğer karşılık gelen bir durum bulunamazsa, 'Bilinmiyor' yaz
    if len(durum) == 0:
        indirilen_veriler.at[index, 'Yeni Durum'] = 'Direkt Gönderilir'
    else:
        indirilen_veriler.at[index, 'Yeni Durum'] = durum[0]

# Sonuçları aynı Excel dosyasına kaydet
indirilen_veriler.to_excel("1000 TL için Kurgu.xlsx", index=False)



# indirilen_veriler excel dosyasını oku
indirilen_veriler = pd.read_excel("1000 TL için Kurgu.xlsx")

# "Teslim Oranı" sütununu sil
indirilen_veriler.drop(columns=["Teslim Oranı"], inplace=True)

# Sonuçları aynı Excel dosyasına kaydet
indirilen_veriler.to_excel("1000 TL için Kurgu.xlsx", index=False)




# "1000 TL için Kurgu.xlsx" excel dosyasını oku
kurgu_excel = pd.read_excel("1000 TL için Kurgu.xlsx")

# "TeslimatTelefon" sütunundaki verilerin başına "90" ekleyelim
kurgu_excel['TeslimatTelefon'] = '90' + kurgu_excel['TeslimatTelefon'].astype(str)

# "Tamamlanan" ve "Toplam" sütunlarını sil
kurgu_excel.drop(columns=['Tamamlanan', 'Toplam'], inplace=True)

# Sonuçları aynı Excel dosyasına kaydet
kurgu_excel.to_excel("1000 TL için Kurgu.xlsx", index=False)






# "indirilen_veriler.xlsx" excel dosyasını oku
indirilen_veriler = pd.read_excel("indirilen_veriler.xlsx")

# "Yeni Teslim Adedi", "Yeni Kek Adedi" ve "Yeni mi" sütunlarını temizle (sütunları boş bırak)
indirilen_veriler['Yeni Teslim Adedi'] = None
indirilen_veriler['Yeni Kek Adedi'] = None
indirilen_veriler['Yeni mi'] = None

# "Yeni Durum" sütunundaki verileri "Eski Durum" sütununa aktar ve "Yeni Durum" sütununu temizle
indirilen_veriler['Eski Durum'] = indirilen_veriler['Yeni Durum']
indirilen_veriler['Yeni Durum'] = None

# "Son Durum" sütununu sil
indirilen_veriler.drop(columns=['Son Durum'], inplace=True)

# Sonuçları aynı Excel dosyasına kaydet
indirilen_veriler.to_excel("indirilen_veriler.xlsx", index=False)



# Eski dosyanın adı
eski_ad = "indirilen_veriler.xlsx"

# Yeni dosyanın adı
yeni_ad = "Orjinal Kara Liste.xlsx"

# Dosyanın adını değiştir
os.rename(eski_ad, yeni_ad)





# Klasör adı
klasor_ad = "Mustafaya Gönderilecekler"

# Klasörü oluştur
os.makedirs(klasor_ad, exist_ok=True)

# Excel dosyalarını taşı
excel_dosyalari = ["1000 TL için Kurgu.xlsx", "Kara Liste.exe.xlsx", "Orjinal Kara Liste.xlsx"]
for dosya in excel_dosyalari:
    shutil.move(dosya, os.path.join(klasor_ad, dosya))



# Klasör adı
klasor_ad = "Connexease'a Gönderilecekler"

# Klasörü oluştur
os.makedirs(klasor_ad, exist_ok=True)

# Excel dosyalarını taşı
excel_dosyalari = ["Yeni Müşteriler.xlsx", "Durumu Değişenler.xlsx"]
for dosya in excel_dosyalari:
    shutil.move(dosya, os.path.join(klasor_ad, dosya))


# Silinecek Excel dosyaları
dosyalar = ["1000 TL Üzeri İçin Sipariş Oranları ve Kurgu.xlsx", "birlesmis_veriler.xlsx", "Orana Göre Durumlar.xlsx"]

# Dosyaları sil
for dosya in dosyalar:
    if os.path.exists(dosya):
        os.remove(dosya)
    else:
        print(f"{dosya} dosyası bulunamadı, dolayısıyla silinemedi.")