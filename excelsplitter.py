#2024.10.13
#generat

import pandas as pd

def ayir_ve_yaz(input_file, sutun, output_path):
    # Excel dosyasını oku
    df = pd.read_excel(input_file)
    
    # Seçilen sütundaki benzersiz değerleri al
    unique_values = df[sutun].unique()
    
    # Her benzersiz değer için bir filtre uygulayıp, yeni Excel dosyası oluştur
    for value in unique_values:
        # Sütun değerine göre filtrele
        df_filtered = df[df[sutun] == value]
        
        # Yeni dosya ismi oluştur
        output_file = f"{output_path}/{sutun}_{value}.xlsx"
        
        # Filtrelenen veriyi yeni Excel dosyasına yaz
        df_filtered.to_excel(output_file, index=False)
        print(f"Dosya kaydedildi: {output_file}")

# Örnek kullanım
input_file = "veri.xlsx"  # Ana Excel dosyanız
sutun = "büro"  # Ayrıştırılacak sütun adı
output_path = "."  # Çıkış dosyalarının kaydedileceği klasör

ayir_ve_yaz(input_file, sutun, output_path)
