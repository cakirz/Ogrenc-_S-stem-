import pandas as pd


input_file = r"C:\Users\----\OneDrive\Masaüstü\Data.xlsx"  # Data dosyamız
weighted_file = r"C:\Users\----\OneDrive\Masaüstü\AğırlıklıDeğerlendirme.xlsx"  # Ağırlıklı değerlendirme tablosu için dosya
output_file = r"C:\Users\----\OneDrive\Masaüstü\ÖğrenciSonuçlar.xlsx"  # Öğrenci ders ve program çıktıları için dosya

# Excel dosyasını okumaya yarayan kod parçası
excel_data = pd.ExcelFile(input_file)
sheet2_data = excel_data.parse("Tablo2")
tablo_not_data = excel_data.parse("TabloNot")
sheet1_data = excel_data.parse("Tablo1", header=1)


relation_column_name = "İlişki Değ."

#Ağırlıklı değerlendirme tablosunu oluştur
weights_raw = sheet2_data.iloc[0, 1:6]
weights = pd.to_numeric(weights_raw.replace(r'[^\d.]', '', regex=True), errors='coerce').fillna(0).astype(float)

grades = sheet2_data.iloc[2:, 1:6].apply(pd.to_numeric, errors='coerce').fillna(0)
weighted_grades = grades * weights.values / 100
weighted_grades['Ağırlıklı Toplam'] = weighted_grades.sum(axis=1)
weighted_table = weighted_grades.copy()

column_headers = list(tablo_not_data.columns[1:6])
weighted_table.columns = column_headers + ['Ağırlıklı Toplam']
weighted_table = weighted_table.round(1)
weighted_table.to_excel(weighted_file, index=False, sheet_name="Ağırlıklı Tablo")
print(f"Ağırlıklı değerlendirme tablosu '{weighted_file}' dosyasına kaydedildi.")

#Tablo 4 ve Tablo 5 oluşturur
student_data = tablo_not_data.iloc[:, 1:6].apply(pd.to_numeric, errors='coerce').fillna(0)
weights_data = weighted_table.iloc[:, :-1]
students = [f"Öğrenci {i + 1}" for i in range(len(tablo_not_data))]

# Tablo 4 ve Tablo 5'i yazar
try:
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for i, student in enumerate(students):
            student_grades = student_data.iloc[i].values
            weighted_grades = weights_data.values * student_grades
            weighted_grades_sum = weighted_grades.sum(axis=1)
            max_values = weights_data.sum(axis=1).values * 100
            success_percentage = (weighted_grades_sum / max_values) * 100

            # Tablo 4: Ders çıktıları
            student_result = pd.DataFrame(weighted_grades, columns=column_headers)
            student_result['TOPLAM'] = weighted_grades_sum
            student_result['MAX'] = max_values
            student_result['% Başarı'] = success_percentage
            student_result = student_result.round(1)
            student_result.to_excel(writer, sheet_name=f"{student}_Tablo4", index=False)

            # Tablo 5: Program çıktıları
            relationship_values = sheet1_data[relation_column_name].apply(pd.to_numeric, errors='coerce').fillna(1)
            program_outcomes = sheet1_data.iloc[:, 0]
            success_percentages = student_result['% Başarı'].values

            tablo5_rows = []
            tablo5_values = []

            for j in range(len(program_outcomes)):
                related_outcomes = sheet1_data.iloc[j, 1:6].apply(pd.to_numeric, errors='coerce').fillna(0)
                related_success = success_percentages * related_outcomes
                tablo5_rows.append(related_success.round(1).tolist())

                num_outcomes = sheet1_data.shape[1] - 2
                mean_success = related_success.sum() / num_outcomes if num_outcomes > 0 else 0
                success_rate = mean_success / relationship_values[j] if relationship_values[j] != 0 else 0
                tablo5_values.append(success_rate)

            tablo5_df = pd.DataFrame(tablo5_rows, columns=[f"Ders Çıktısı {k+1}" for k in range(5)])
            tablo5_df['Başarı Oranı'] = tablo5_values
            tablo5_df.index = program_outcomes
            tablo5_df = tablo5_df.round(1)

            tablo5_df.to_excel(writer, sheet_name=f"{student}_Tablo5", index=True)

    print(f"Tüm öğrenci sonuçları '{output_file}' dosyasına kaydedildi.")
except Exception as e:
    print(f"Hata oluştu: {e}")
