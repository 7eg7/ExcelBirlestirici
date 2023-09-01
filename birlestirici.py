import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog

def merge_excel_to_single_text_file():
    excel_file_paths = filedialog.askopenfilenames(title="Excel Dosyalarını Seçin", filetypes=[("Excel Dosyaları", "*.xlsx")])
    
    if excel_file_paths:
        selected_sheet_name = combo_sheet_name.get()  # Seçilen sayfa adını al
        
        output_file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Metin Dosyaları", "*.txt")])
        if output_file_path:
            try:
                with open(output_file_path, 'w') as f:
                    for excel_file_path in excel_file_paths:
                        try:
                            data_frame = pd.read_excel(excel_file_path, sheet_name=selected_sheet_name)
                            for index, row in data_frame.iterrows():
                                for column in data_frame.columns:
                                    f.write(str(row[column]) + '\n')
                                f.write('\n')
                            f.write('\n')
                            print(f"Veriler ({excel_file_path}) metin dosyasına yazıldı.")
                        except Exception as e:
                            print(f"Hata ({excel_file_path}):", e)
                
                print("Tüm dosyalar birleştirildi ve veriler metin dosyasına kaydedildi.")
                root.destroy()  # Programı otomatik olarak kapat
                
            except Exception as e:
                print("Hata:", e)

# Arayüzü oluşturma
root = tk.Tk()
root.title("Excel Birleştirici")

# Stil ayarları
style = ttk.Style()
style.configure("TButton", padding=10)
style.configure("TLabel", padding=10)

frame_sheet = ttk.Frame(root)
frame_sheet.pack(padx=20, pady=10, fill="both")

frame_buttons = ttk.Frame(root)
frame_buttons.pack(padx=20, pady=10, fill="both")

label_sheet_name = ttk.Label(frame_sheet, text="Sayfa Adı:")
label_sheet_name.pack(side="left")

# Mevcut sayfa adları için bir örnek liste
available_sheet_names = ["Sheet1", "Sheet2", "Sheet3"]  # Gerçek sayfa adlarınızı buraya ekleyin

combo_sheet_name = ttk.Combobox(frame_sheet, values=available_sheet_names)
combo_sheet_name.pack(side="left")

button_merge = ttk.Button(frame_buttons, text="Excel Dosyalarını Birleştir ve Kaydet", command=merge_excel_to_single_text_file)
button_merge.pack()

root.mainloop()
