import pandas as pd

def ayir_ve_yaz(input_file, sutun, output_path):
    # Excel dosyasını oku
    df = pd.read_excel(input_file)
    
    # Seçilen sütundaki benzersiz değerleri al
    unique_values = df[sutun].unique()
    
    # Her benzersiz değer için bir filtre uygulayıp, yeni Excel dosyası oluştur
    for value in unique_values:
        # Sütun değerine göre filtrele
        df_filtered = df[df[sutun] == value].copy()
        
        # Filtrelenmiş dosyadan seçilen sütunu kaldır
        df_filtered.drop(columns=[sutun], inplace=True)
        
        # Yeni dosya ismi oluştur
        output_file = f"{output_path}/{sutun}_{value}.xlsx"
        
        # Excel dosyasına yazma işlemini başlat
        with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
            # Yazıcı objesini alın ve ilk hücreye büro değerini yazın
            workbook = writer.book
            worksheet = workbook.add_worksheet()
            worksheet.write('A1', f"{sutun}: {value}")
            
            # Filtrelenmiş veriyi ikinci satırdan itibaren yaz
            df_filtered.to_excel(writer, sheet_name="Sheet1", startrow=2, index=False)
            
        print(f"Dosya kaydedildi: {output_file}")

# Örnek kullanım
input_file = "veri.xlsx"  # Ana Excel dosyanız
sutun = "büro"  # Ayrıştırılacak sütun adı
output_path = "."  # Çıkış dosyalarının kaydedileceği klasör

ayir_ve_yaz(input_file, sutun, output_path)