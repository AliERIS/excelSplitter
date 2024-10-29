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
            # Filtrelenmiş veriyi ikinci satırdan itibaren yaz
            df_filtered.to_excel(writer, sheet_name="Sheet1", startrow=2, index=False)
            
            # Çalışma kitabı ve çalışma sayfası nesnelerini alın
            workbook = writer.book
            worksheet = writer.sheets["Sheet1"]
            
            # İlk satırda hücreleri birleştirme ve değer yazma
            num_columns = len(df_filtered.columns)
            merge_range = f"A1:{chr(65 + num_columns - 1)}1"  # A1'den son sütuna kadar birleştirme
            
            # Birleştirilmiş hücre için format oluştur
            merge_format = workbook.add_format({
                'align': 'center', 
                'valign': 'vcenter', 
                'bold': True
            })
            
            # Hücreleri birleştir ve değer yaz
            worksheet.merge_range(merge_range, f"{sutun}: {value}", merge_format)
        
        print(f"Dosya kaydedildi: {output_file}")

# Örnek kullanım
input_file = "veri.xlsx"  # Ana Excel dosyanız
sutun = "büro"  # Ayrıştırılacak sütun adı
output_path = "."  # Çıkış dosyalarının kaydedileceği klasör

ayir_ve_yaz(input_file, sutun, output_path)