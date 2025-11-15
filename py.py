import pandas as pd
import json
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pathlib import Path

class ExcelToJSONLDConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel to JSON-LD Converter")
        self.root.geometry("800x600")
        
        # Variables
        self.selected_files = []
        
        self.setup_ui()
    
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Title
        title_label = ttk.Label(main_frame, text="Excel to JSON-LD Converter", 
                               font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # File selection section
        file_frame = ttk.LabelFrame(main_frame, text="File Selection", padding="10")
        file_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Buttons for file selection
        ttk.Button(file_frame, text="Pilih File Excel", 
                  command=self.select_files).grid(row=0, column=0, padx=(0, 10))
        
        ttk.Button(file_frame, text="Pilih Folder", 
                  command=self.select_folder).grid(row=0, column=1, padx=(0, 10))
        
        ttk.Button(file_frame, text="Hapus Semua", 
                  command=self.clear_files).grid(row=0, column=2)
        
        # Selected files list
        self.files_listbox = tk.Listbox(file_frame, height=6, width=80)
        self.files_listbox.grid(row=1, column=0, columnspan=3, pady=(10, 0), sticky=(tk.W, tk.E))
        
        # Output folder section
        output_frame = ttk.LabelFrame(main_frame, text="Output Settings", padding="10")
        output_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(output_frame, text="Folder Output:").grid(row=0, column=0, sticky=tk.W)
        
        self.output_var = tk.StringVar(value="result")
        ttk.Entry(output_frame, textvariable=self.output_var, width=50).grid(row=0, column=1, padx=(10, 10))
        
        ttk.Button(output_frame, text="Browse", 
                  command=self.browse_output_folder).grid(row=0, column=2)
        
        # Conversion options
        options_frame = ttk.LabelFrame(main_frame, text="Conversion Options", padding="10")
        options_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.base_uri_var = tk.StringVar(value="http://example.org/data/")
        ttk.Label(options_frame, text="Base URI:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(options_frame, textvariable=self.base_uri_var, width=50).grid(row=0, column=1, padx=(10, 0))
        
        # Convert button
        self.convert_btn = ttk.Button(main_frame, text="Konversi ke JSON-LD", 
                                     command=self.convert_files, state=tk.DISABLED)
        self.convert_btn.grid(row=4, column=0, columnspan=3, pady=20)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='determinate')
        self.progress.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Log output
        log_frame = ttk.LabelFrame(main_frame, text="Log Output", padding="10")
        log_frame.grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=15, width=80)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(6, weight=1)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
    
    def log(self, message):
        """Add message to log"""
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update()
    
    def select_files(self):
        """Select multiple Excel files"""
        files = filedialog.askopenfilenames(
            title="Pilih file Excel",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if files:
            self.selected_files.extend(files)
            self.update_files_list()
            self.convert_btn.config(state=tk.NORMAL)
    
    def select_folder(self):
        """Select folder containing Excel files"""
        folder = filedialog.askdirectory(title="Pilih folder berisi file Excel")
        
        if folder:
            excel_files = []
            for ext in ['*.xlsx', '*.xls']:
                excel_files.extend(Path(folder).glob(ext))
            
            if excel_files:
                self.selected_files.extend([str(f) for f in excel_files])
                self.update_files_list()
                self.convert_btn.config(state=tk.NORMAL)
            else:
                messagebox.showwarning("Peringatan", "Tidak ditemukan file Excel di folder tersebut")
    
    def browse_output_folder(self):
        """Browse for output folder"""
        folder = filedialog.askdirectory(title="Pilih folder output")
        if folder:
            self.output_var.set(folder)
    
    def clear_files(self):
        """Clear all selected files"""
        self.selected_files = []
        self.files_listbox.delete(0, tk.END)
        self.convert_btn.config(state=tk.DISABLED)
    
    def update_files_list(self):
        """Update the files listbox"""
        self.files_listbox.delete(0, tk.END)
        for file in self.selected_files:
            self.files_listbox.insert(tk.END, file)
    
    def excel_to_jsonld_fuseki(self, excel_file, output_file, base_uri):
        """Convert Excel to JSON-LD"""
        try:
            self.log(f"📖 Membaca file: {os.path.basename(excel_file)}")
            df = pd.read_excel(excel_file, engine='openpyxl')
            self.log(f"✅ Berhasil membaca {len(df)} rows, {len(df.columns)} columns")
            
        except Exception as e:
            self.log(f"❌ Error membaca Excel: {e}")
            return None
        
        # Buat JSON-LD structure
        jsonld_data = {
            "@context": {
                "schema": "https://schema.org/",
                "ex": base_uri,
                "rdf": "http://www.w3.org/1999/02/22-rdf-syntax-ns#",
                "rdfs": "http://www.w3.org/2000/01/rdf-schema#",
                "xsd": "http://www.w3.org/2001/XMLSchema#"
            },
            "@graph": []
        }
        
        # Process semua rows
        total_records = len(df)
        processed_records = 0
        
        for index, row in df.iterrows():
            try:
                if row.isna().all():
                    continue
                    
                record_id = f"record_{index + 1}"
                record = {
                    "@id": f"ex:{record_id}",
                    "@type": "schema:Product",
                    "schema:position": index + 1
                }
                
                # Add properties
                for col_name in df.columns:
                    value = row[col_name]
                    
                    if pd.isna(value):
                        continue
                        
                    clean_col = col_name.replace(' ', '_').replace('/', '_').replace('(', '').replace(')', '').lower()
                    
                    property_mapping = {
                        'web_scraper_order': 'schema:identifier',
                        'web_scraper_start_url': 'schema:url',
                        'link': 'schema:url',
                        'nama_barang': 'schema:name',
                        'harga_barang': 'schema:price',
                        'kondisi': 'schema:itemCondition',
                        'stok': 'schema:inventoryLevel',
                        'detail': 'schema:description',
                        'logo_toko': 'schema:logo',
                        'gambar_barang': 'schema:image',
                        'lokasi': 'schema:location',
                        'rating': 'schema:aggregateRating'
                    }
                    
                    property_uri = property_mapping.get(clean_col, f"ex:{clean_col}")
                    
                    if isinstance(value, (int, float)):
                        record[property_uri] = {
                            "@value": value,
                            "@type": "xsd:decimal" if isinstance(value, float) else "xsd:integer"
                        }
                    elif isinstance(value, bool):
                        record[property_uri] = {
                            "@value": str(value).lower(),
                            "@type": "xsd:boolean"
                        }
                    else:
                        str_value = str(value)
                        if str_value.startswith(('http://', 'https://')):
                            record[property_uri] = {"@id": str_value}
                        else:
                            record[property_uri] = {
                                "@value": str_value,
                                "@type": "xsd:string"
                            }
                
                jsonld_data["@graph"].append(record)
                processed_records += 1
                
            except Exception as e:
                self.log(f"❌ Error processing row {index}: {e}")
                continue
        
        # Buat folder output jika belum ada
        output_folder = self.output_var.get()
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)
            self.log(f"📁 Folder '{output_folder}' dibuat")
        
        # Simpan file
        try:
            output_path = os.path.join(output_folder, output_file)
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(jsonld_data, f, indent=2, ensure_ascii=False)
            
            self.log(f"✅ JSON-LD berhasil disimpan: {output_path}")
            self.log(f"📝 Total records: {len(jsonld_data['@graph'])}")
            
            return jsonld_data
            
        except Exception as e:
            self.log(f"❌ Error menyimpan file: {e}")
            return None
    
    def convert_files(self):
        """Convert all selected files"""
        if not self.selected_files:
            messagebox.showwarning("Peringatan", "Tidak ada file yang dipilih")
            return
        
        # Reset progress
        self.progress['value'] = 0
        self.progress['maximum'] = len(self.selected_files)
        
        success_count = 0
        
        for i, excel_file in enumerate(self.selected_files):
            self.log(f"\n{'='*50}")
            self.log(f"🔄 Memproses file {i+1}/{len(self.selected_files)}: {os.path.basename(excel_file)}")
            
            # Generate output filename
            base_name = os.path.splitext(os.path.basename(excel_file))[0]
            output_file = f"{base_name}_complete.jsonld"
            
            # Convert
            result = self.excel_to_jsonld_fuseki(excel_file, output_file, self.base_uri_var.get())
            
            if result:
                success_count += 1
                self.log(f"✅ SUCCESS: {os.path.basename(excel_file)} -> {output_file}")
            else:
                self.log(f"❌ FAILED: {os.path.basename(excel_file)}")
            
            # Update progress
            self.progress['value'] = i + 1
            self.root.update()
        
        # Show summary
        self.log(f"\n{'='*50}")
        self.log(f"🎉 KONVERSI SELESAI!")
        self.log(f"📁 Total file diproses: {len(self.selected_files)}")
        self.log(f"✅ Berhasil: {success_count}")
        self.log(f"❌ Gagal: {len(self.selected_files) - success_count}")
        self.log(f"📂 Output disimpan di: {self.output_var.get()}")
        
        messagebox.showinfo("Selesai", f"Konversi selesai!\nBerhasil: {success_count}/{len(self.selected_files)} file")

def main():
    root = tk.Tk()
    app = ExcelToJSONLDConverter(root)
    root.mainloop()

if __name__ == "__main__":
    main()