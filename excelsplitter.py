import pandas as pd

def ayir_ve_yaz(input_file, sutun, output_path):
    # Excel dosyasını oku
    df = pd.read_excel(input_file)
    
    # Seçilen sütundaki benzersiz değerleri al
    unique_values = df[sutun].unique()
    print(df.columns)
    
    for value in unique_values:
        print(value)
    # Her benzersiz değer için bir filtre uygulayıp, yeni Excel dosyası oluştur
    for value in unique_values:
        # Sütun değerine göre filtrele
        df_filtered = df[df[sutun] == value].copy()
        
        # Filtrelenmiş dosyadan seçilen sütunu kaldır
        df_filtered.drop(columns=[sutun], inplace=True)
        
        # Yeni dosya ismi oluştur
        output_file = f"{output_path}/{value}.xlsx"
        
        # Excel dosyasına yazma işlemini başlat
        with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
            # Filtrelenmiş veriyi ikinci satırdan itibaren yaz
            df_filtered.to_excel(writer, sheet_name="PERSONEL", startrow=1, index=False)
            
            # Çalışma kitabı ve çalışma sayfası nesnelerini alın
            workbook = writer.book
            worksheet = writer.sheets["PERSONEL"]
            
            # 1. Başlık hücresi formatı
            header_format = workbook.add_format({
                'align': 'center',
                'valign': 'vcenter',
                'bold': True,
                'font_name': 'Times New Roman',
                'font_size': 13,
                'bg_color': 'yellow'
            })
            
            # 2. Veri hücresi formatı
            cell_format = workbook.add_format({
                'align': 'center',
                'valign': 'vcenter',
                'font_name': 'Times New Roman',
                'font_size': 11,
                'border': 1,
                'bg_color': 'white',
                'font_color': 'black'
            })
            
            # İlk satırda hücreleri birleştirme ve başlık ekleme
            num_columns = len(df_filtered.columns)
            merge_range = f"A1:{chr(65 + num_columns - 1)}1"  # A1'den son sütuna kadar birleştirme
            worksheet.merge_range(merge_range, f" {value} ", header_format)
            
            # Veri hücrelerine kenarlık ve diğer formatları uygulama
            for row in range(2, len(df_filtered) + 2):  # Başlık sonrası satırları seçiyoruz
                for col in range(num_columns):
                    worksheet.write(row, col, df_filtered.iloc[row - 2, col], cell_format)
        
        print(f"Dosya kaydedildi: {output_file}")

# Örnek kullanım
input_file = "veri2.xlsx"  # Ana Excel dosyanız
sutun = "büro"  # Ayrıştırılacak sütun adı
output_path = "./output/"  # Çıkış dosyalarının kaydedileceği klasör

ayir_ve_yaz(input_file, sutun, output_path)
