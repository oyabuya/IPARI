import json
import csv
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, Canvas, Frame
from datetime import datetime
from PIL import Image, ImageTk, ImageDraw, ImageFont
import os
import webbrowser
import sys

# Handle path untuk PyInstaller
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class IuranApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Aplikasi Rekap Iuran IPARI Mentawai")
        self.root.geometry("800x600")
        self.data_anggota = []
        self.current_anggota = None
        self.data_file = "data_iuran.json"

	# Status bar
        self.status_var = tk.StringVar()
        self.status_bar = ttk.Label(root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        self.update_status("Siap. Silakan tambahkan data anggota.")
        
        # Muat data yang tersimpan
        self.load_data()
        
        # Path logo IPARI
        self.logo_path = resource_path("logo_ipari.png")
        
        # Styling
        self.style = ttk.Style()
        self.style.configure("TFrame", background="#f0f0f0")
        self.style.configure("TLabel", background="#f0f0f0", font=("Arial", 10))
        self.style.configure("TButton", font=("Arial", 10))
        self.style.configure("Header.TLabel", font=("Arial", 12, "bold"))
        self.style.configure("Subheader.TLabel", font=("Arial", 11, "bold"))
        self.style.configure("Lunas.TLabel", foreground="green")
        self.style.configure("Belum.TLabel", foreground="red")
        
        # Create main frame
        self.main_frame = ttk.Frame(root)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.create_search_frame()
        
        # Form input
        self.create_input_form()
        
        # Results display
        self.create_results_display()
        
        # Bottom buttons
        self.create_action_buttons()
        
        # Handle aplikasi ditutup
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
    
    def load_data(self):
        """Memuat data dari file JSON jika ada"""
        if os.path.exists(self.data_file):
            try:
                with open(self.data_file, 'r') as f:
                    self.data_anggota = json.load(f)
                self.update_status(f"Data berhasil dimuat. Total anggota: {len(self.data_anggota)}")
            except Exception as e:
                messagebox.showerror("Error", f"Gagal memuat data: {str(e)}")
                self.data_anggota = []
    
    def save_data(self):
        """Menyimpan data ke file JSON"""
        try:
            with open(self.data_file, 'w') as f:
                json.dump(self.data_anggota, f, indent=4)
        except Exception as e:
            messagebox.showerror("Error", f"Gagal menyimpan data: {str(e)}")
    
    def on_close(self):
        """Ditangani saat aplikasi ditutup"""
        self.save_data()
        self.root.destroy()

    def create_search_frame(self):
        search_frame = ttk.Frame(self.main_frame)
        search_frame.pack(fill=tk.X, padx=5, pady=(0, 10))
    
        ttk.Label(search_frame, text="Cari Anggota:").pack(side=tk.LEFT, padx=(0, 5))
    
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=40)
        self.search_entry.pack(side=tk.LEFT, padx=5)
        self.search_entry.bind("<Return>", lambda e: self.cari_anggota())
    
        ttk.Button(search_frame, text="Cari", command=self.cari_anggota).pack(side=tk.LEFT, padx=5)
    
    def create_input_form(self):
        form_frame = ttk.LabelFrame(self.main_frame, text="Input Data Anggota")
        form_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Nama
        ttk.Label(form_frame, text="Nama Anggota:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.nama_entry = ttk.Entry(form_frame, width=40)
        self.nama_entry.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        # Status 2024
        ttk.Label(form_frame, text="Status 2024:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.status_2024_var = tk.StringVar()
        status_options = ["Penyuluh PNS", "Penyuluh PPPK", "Penyuluh Non ASN"]
        self.status_2024_combo = ttk.Combobox(form_frame, textvariable=self.status_2024_var, 
                                            values=status_options, state="readonly", width=37)
        self.status_2024_combo.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
        self.status_2024_combo.bind("<<ComboboxSelected>>", self.on_status_change)
        
        # Kondisi 2025 (initially hidden)
        self.kondisi_frame = ttk.Frame(form_frame)
        self.kondisi_frame.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)
        self.kondisi_frame.grid_remove()
        
        ttk.Label(self.kondisi_frame, text="Perubahan Status 2025:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.kondisi_var = tk.StringVar()
        self.kondisi_combo = ttk.Combobox(self.kondisi_frame, textvariable=self.kondisi_var, 
                                        state="readonly", width=35)
        self.kondisi_combo.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        # Pembayaran
        pembayaran_frame = ttk.Frame(form_frame)
        pembayaran_frame.grid(row=3, column=0, columnspan=2, padx=5, pady=5, sticky=tk.W)
        
        ttk.Label(pembayaran_frame, text="Pembayaran:").grid(row=0, column=0, padx=(0, 10), pady=5, sticky=tk.W)
        
        ttk.Label(pembayaran_frame, text="2024:").grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        self.bayar_2024_var = tk.StringVar()
        self.bayar_2024_entry = ttk.Entry(pembayaran_frame, textvariable=self.bayar_2024_var, width=10)
        self.bayar_2024_entry.grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
        
        ttk.Label(pembayaran_frame, text="2025:").grid(row=0, column=3, padx=5, pady=5, sticky=tk.W)
        self.bayar_2025_var = tk.StringVar()
        self.bayar_2025_entry = ttk.Entry(pembayaran_frame, textvariable=self.bayar_2025_var, width=10)
        self.bayar_2025_entry.grid(row=0, column=4, padx=5, pady=5, sticky=tk.W)
        
        # Add button
        self.add_button = ttk.Button(form_frame, text="Tambah / Perbarui Anggota", command=self.add_anggota)
        self.add_button.grid(row=4, column=1, padx=5, pady=10, sticky=tk.E)
    
    def on_status_change(self, event):
        if self.status_2024_var.get() == "Penyuluh Non ASN":
            kondisi_options = [
                "SK tidak dikeluarkan pemerintah (berhenti)",
                "Diangkat PPPK Tahap I Penyuluh Agama",
                "Diangkat PPPK Tahap I NON Penyuluh",
                "Sedang berjuang PPPK Tahap II"
            ]
            self.kondisi_combo["values"] = kondisi_options
            self.kondisi_frame.grid()
        else:
            self.kondisi_frame.grid_remove()
    
    def create_results_display(self):
        self.results_frame = ttk.LabelFrame(self.main_frame, text="Ringkasan Pembayaran")
        self.results_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.results_frame.pack_forget()  # Sembunyikan awalnya
        
        # Header
        ttk.Label(self.results_frame, text="Status Pembayaran Anggota", style="Header.TLabel").pack(pady=10)
        
        # Nama anggota
        self.nama_label = ttk.Label(self.results_frame, text="", font=("Arial", 11, "bold"))
        self.nama_label.pack(pady=5)
        
        # Container for status boxes
        status_container = ttk.Frame(self.results_frame)
        status_container.pack(fill=tk.X, padx=10, pady=10)
        
        # Status 2024
        frame_2024 = ttk.LabelFrame(status_container, text="Tahun 2024")
        frame_2024.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        ttk.Label(frame_2024, text="Status:").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        self.status_2024_label = ttk.Label(frame_2024, text="", style="Subheader.TLabel")
        self.status_2024_label.grid(row=0, column=1, padx=5, pady=2, sticky=tk.W)
        
        ttk.Label(frame_2024, text="Iuran Wajib:").grid(row=1, column=0, padx=5, pady=2, sticky=tk.W)
        self.iuran_2024_label = ttk.Label(frame_2024, text="")
        self.iuran_2024_label.grid(row=1, column=1, padx=5, pady=2, sticky=tk.W)
        
        ttk.Label(frame_2024, text="Sudah Dibayar:").grid(row=2, column=0, padx=5, pady=2, sticky=tk.W)
        self.dibayar_2024_label = ttk.Label(frame_2024, text="")
        self.dibayar_2024_label.grid(row=2, column=1, padx=5, pady=2, sticky=tk.W)
        
        ttk.Label(frame_2024, text="Kekurangan:").grid(row=3, column=0, padx=5, pady=2, sticky=tk.W)
        self.kurang_2024_label = ttk.Label(frame_2024, text="")
        self.kurang_2024_label.grid(row=3, column=1, padx=5, pady=2, sticky=tk.W)
        
        ttk.Label(frame_2024, text="Status Pembayaran:").grid(row=4, column=0, padx=5, pady=2, sticky=tk.W)
        self.status_bayar_2024_label = ttk.Label(frame_2024, text="", font=("Arial", 10, "bold"))
        self.status_bayar_2024_label.grid(row=4, column=1, padx=5, pady=2, sticky=tk.W)
        
        # Status 2025
        frame_2025 = ttk.LabelFrame(status_container, text="Tahun 2025")
        frame_2025.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        ttk.Label(frame_2025, text="Status:").grid(row=0, column=0, padx=5, pady=2, sticky=tk.W)
        self.status_2025_label = ttk.Label(frame_2025, text="", style="Subheader.TLabel")
        self.status_2025_label.grid(row=0, column=1, padx=5, pady=2, sticky=tk.W)
        
        ttk.Label(frame_2025, text="Iuran Wajib:").grid(row=1, column=0, padx=5, pady=2, sticky=tk.W)
        self.iuran_2025_label = ttk.Label(frame_2025, text="")
        self.iuran_2025_label.grid(row=1, column=1, padx=5, pady=2, sticky=tk.W)
        
        ttk.Label(frame_2025, text="Sudah Dibayar:").grid(row=2, column=0, padx=5, pady=2, sticky=tk.W)
        self.dibayar_2025_label = ttk.Label(frame_2025, text="")
        self.dibayar_2025_label.grid(row=2, column=1, padx=5, pady=2, sticky=tk.W)
        
        ttk.Label(frame_2025, text="Kekurangan:").grid(row=3, column=0, padx=5, pady=2, sticky=tk.W)
        self.kurang_2025_label = ttk.Label(frame_2025, text="")
        self.kurang_2025_label.grid(row=3, column=1, padx=5, pady=2, sticky=tk.W)
        
        ttk.Label(frame_2025, text="Status Pembayaran:").grid(row=4, column=0, padx=5, pady=2, sticky=tk.W)
        self.status_bayar_2025_label = ttk.Label(frame_2025, text="", font=("Arial", 10, "bold"))
        self.status_bayar_2025_label.grid(row=4, column=1, padx=5, pady=2, sticky=tk.W)
        
        # Action buttons
        btn_frame = ttk.Frame(self.results_frame)
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="Lihat Rincian Lengkap", command=self.show_detail).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Edit Pembayaran", command=self.edit_payment).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Cetak Nota", command=self.cetak_nota).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Cari Anggota Lain", command=self.clear_results).pack(side=tk.LEFT, padx=5)
    
    def create_action_buttons(self):
        button_frame = ttk.Frame(self.main_frame)
        button_frame.pack(fill=tk.X, padx=5, pady=10)
        
        ttk.Button(button_frame, text="Lihat Semua Anggota", command=self.show_all_members).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Ekspor CSV", command=self.export_to_csv).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Statistik", command=self.show_statistics).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cetak Semua Nota", command=self.cetak_semua_nota).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Bersihkan Data", command=self.clear_all).pack(side=tk.LEFT, padx=5)
    
    def cari_anggota(self):
        query = self.search_var.get().strip().lower()
        if not query:
            messagebox.showinfo("Pencarian", "Masukkan nama anggota untuk dicari")
            return
    
        results = []
        for anggota in self.data_anggota:
            if query in anggota['nama'].lower():
                results.append(anggota)
    
        if not results:
            messagebox.showinfo("Pencarian", f"Tidak ditemukan anggota dengan nama '{query}'")
            return
    
        if len(results) == 1:
            self.current_anggota = results[0]
            self.show_summary()
            self.update_status(f"Menampilkan data {results[0]['nama']}")
        else:
            self.tampilkan_dialog_pilihan(results)

    def tampilkan_dialog_pilihan(self, results):
        dialog = tk.Toplevel(self.root)
        dialog.title("Pilih Anggota")
        dialog.geometry("400x300")
        dialog.transient(self.root)
        dialog.grab_set()
    
        ttk.Label(dialog, text="Ditemukan beberapa anggota. Pilih salah satu:").pack(pady=10)
    
        list_frame = ttk.Frame(dialog)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
    
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
        listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, width=50)
        for anggota in results:
            status = "Lunas" if anggota['status_bayar_2024'] == "Lunas" and anggota['status_bayar_2025'] == "Lunas" else "Belum lunas"
            listbox.insert(tk.END, f"{anggota['nama']} ({status})")
        listbox.pack(fill=tk.BOTH, expand=True)
    
        scrollbar.config(command=listbox.yview)
    
        def on_select():
            selected_index = listbox.curselection()
            if selected_index:
                idx = selected_index[0]
                self.current_anggota = results[idx]
                self.show_summary()
                self.update_status(f"Menampilkan data {results[idx]['nama']}")
                dialog.destroy()
    
        ttk.Button(dialog, text="Pilih", command=on_select).pack(pady=10)

    def format_rupiah(self, nilai):
        if isinstance(nilai, (int, float)):
            return f"Rp {nilai:,.0f}".replace(',', '.')
        return nilai
    
    def parse_rupiah(self, rupiah_str):
        """Convert 'Rp 100.000' to 100000"""
        if isinstance(rupiah_str, (int, float)):
            return rupiah_str
            
        try:
            clean_str = rupiah_str.replace('Rp', '').replace('.', '').replace(',', '.').strip()
            if clean_str == '':
                return 0
            return int(float(clean_str))
        except:
            return 0
    
    def add_anggota(self):
        nama = self.nama_entry.get().strip()
        if not nama:
            messagebox.showerror("Error", "Nama anggota harus diisi!")
            return
            
        # Cek apakah anggota sudah ada
        existing_index = -1
        for idx, anggota in enumerate(self.data_anggota):
            if anggota['nama'].lower() == nama.lower():
                existing_index = idx
                break
        
        status_2024 = self.status_2024_var.get()
        if not status_2024:
            messagebox.showerror("Error", "Status 2024 harus dipilih!")
            return
            
        dibayar_2024 = self.parse_rupiah(self.bayar_2024_var.get())
        dibayar_2025 = self.parse_rupiah(self.bayar_2025_var.get())
        
        # Prepare anggota data
        anggota = {
            "nama": nama,
            "status_2024": status_2024,
            "status_2025": status_2024,  # Default sama dengan 2024
            "kondisi": "Tidak ada perubahan",
            "iuran_2024": 200000,
            "iuran_2025": 200000,
            "dibayar_2024": dibayar_2024,
            "dibayar_2025": dibayar_2025,
            "kurang_2024": 0,
            "kurang_2025": 0,
            "status_bayar_2024": "Belum",
            "status_bayar_2025": "Belum"
        }
        
        # Handle Non ASN cases
        if status_2024 == "Penyuluh Non ASN":
            kondisi = self.kondisi_var.get()
            if not kondisi:
                messagebox.showerror("Error", "Kondisi perubahan 2025 harus dipilih untuk Non ASN!")
                return
            
            anggota["kondisi"] = kondisi
            
            if kondisi == "SK tidak dikeluarkan pemerintah (berhenti)":
                anggota["status_2025"] = "Bukan anggota"
                anggota["iuran_2024"] = 0
                anggota["iuran_2025"] = 0
            elif kondisi == "Diangkat PPPK Tahap I Penyuluh Agama":
                anggota["status_2025"] = "Penyuluh PPPK"
                anggota["iuran_2024"] = 100000
                anggota["iuran_2025"] = 200000
            elif kondisi == "Diangkat PPPK Tahap I NON Penyuluh":
                anggota["status_2025"] = "PPPK Non Penyuluh"
                anggota["iuran_2024"] = 100000
                anggota["iuran_2025"] = 100000
            elif kondisi == "Sedang berjuang PPPK Tahap II":
                anggota["status_2025"] = "Penyuluh Non ASN"
                anggota["iuran_2024"] = 100000
                anggota["iuran_2025"] = 100000
        
        # Hitung kekurangan
        anggota["kurang_2024"] = max(0, anggota["iuran_2024"] - anggota["dibayar_2024"])
        anggota["kurang_2025"] = max(0, anggota["iuran_2025"] - anggota["dibayar_2025"])
        
        # Update status pembayaran
        anggota["status_bayar_2024"] = "Lunas" if anggota["kurang_2024"] == 0 else "Belum"
        anggota["status_bayar_2025"] = "Lunas" if anggota["kurang_2025"] == 0 else "Belum"
        
        # Update or add data
        if existing_index >= 0:
            self.data_anggota[existing_index] = anggota
            self.update_status(f"Data {nama} berhasil diperbarui")
        else:
            self.data_anggota.append(anggota)
            self.update_status(f"Data {nama} berhasil ditambahkan. Total anggota: {len(self.data_anggota)}")
        
        # Tampilkan ringkasan
        self.current_anggota = anggota
        self.show_summary()
        
        # Clear form
        self.nama_entry.delete(0, tk.END)
        self.status_2024_var.set("")
        self.kondisi_var.set("")
        self.bayar_2024_var.set("")
        self.bayar_2025_var.set("")
        self.kondisi_frame.grid_remove()
        
        # Simpan data
        self.save_data()
    
    def show_summary(self):
        if not self.current_anggota:
            return
            
        anggota = self.current_anggota
        
        # Tampilkan frame hasil
        self.results_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Update labels
        self.nama_label.config(text=anggota["nama"])
        
        # Tahun 2024
        self.status_2024_label.config(text=anggota["status_2024"])
        self.iuran_2024_label.config(text=self.format_rupiah(anggota["iuran_2024"]))
        self.dibayar_2024_label.config(text=self.format_rupiah(anggota["dibayar_2024"]))
        self.kurang_2024_label.config(text=self.format_rupiah(anggota["kurang_2024"]))
        
        # Status pembayaran 2024
        status_2024 = anggota["status_bayar_2024"]
        self.status_bayar_2024_label.config(text=status_2024)
        if status_2024 == "Lunas":
            self.status_bayar_2024_label.config(style="Lunas.TLabel")
        else:
            self.status_bayar_2024_label.config(style="Belum.TLabel")
        
        # Tahun 2025
        self.status_2025_label.config(text=anggota["status_2025"])
        self.iuran_2025_label.config(text=self.format_rupiah(anggota["iuran_2025"]))
        self.dibayar_2025_label.config(text=self.format_rupiah(anggota["dibayar_2025"]))
        self.kurang_2025_label.config(text=self.format_rupiah(anggota["kurang_2025"]))
        
        # Status pembayaran 2025
        status_2025 = anggota["status_bayar_2025"]
        self.status_bayar_2025_label.config(text=status_2025)
        if status_2025 == "Lunas":
            self.status_bayar_2025_label.config(style="Lunas.TLabel")
        else:
            self.status_bayar_2025_label.config(style="Belum.TLabel")
    
    def show_detail(self):
        if not self.current_anggota:
            return
            
        anggota = self.current_anggota
        
        detail = f"RINCIAN PEMBAYARAN UNTUK: {anggota['nama']}\n"
        detail += "=" * 50 + "\n"
        detail += f"Status 2024: {anggota['status_2024']}\n"
        detail += f"Status 2025: {anggota['status_2025']}\n"
        detail += f"Kondisi Perubahan: {anggota['kondisi']}\n\n"
        
        detail += "TAHUN 2024:\n"
        detail += f"  Iuran Wajib: {self.format_rupiah(anggota['iuran_2024'])}\n"
        detail += f"  Sudah Dibayar: {self.format_rupiah(anggota['dibayar_2024'])}\n"
        detail += f"  Kekurangan: {self.format_rupiah(anggota['kurang_2024'])}\n"
        detail += f"  Status: {anggota['status_bayar_2024']}\n\n"
        
        detail += "TAHUN 2025:\n"
        detail += f"  Iuran Wajib: {self.format_rupiah(anggota['iuran_2025'])}\n"
        detail += f"  Sudah Dibayar: {self.format_rupiah(anggota['dibayar_2025'])}\n"
        detail += f"  Kekurangan: {self.format_rupiah(anggota['kurang_2025'])}\n"
        detail += f"  Status: {anggota['status_bayar_2025']}"
        
        messagebox.showinfo("Rincian Pembayaran", detail)
    
    def edit_payment(self):
        if not self.current_anggota:
            return
            
        anggota = self.current_anggota
        
        # Create edit dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("Edit Pembayaran")
        dialog.geometry("400x250")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text=f"Edit Pembayaran untuk: {anggota['nama']}", font=("Arial", 10, "bold")).pack(pady=10)
        
        # Frame for 2024
        frame_2024 = ttk.Frame(dialog)
        frame_2024.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(frame_2024, text="Tahun 2024:").grid(row=0, column=0, padx=5, sticky=tk.W)
        ttk.Label(frame_2024, text=f"Iuran Wajib: {self.format_rupiah(anggota['iuran_2024'])}").grid(row=0, column=1, padx=5, sticky=tk.W)
        
        ttk.Label(frame_2024, text="Sudah Dibayar:").grid(row=1, column=0, padx=5, sticky=tk.W)
        bayar_2024_var = tk.StringVar(value=self.format_rupiah(anggota["dibayar_2024"]))
        bayar_2024_entry = ttk.Entry(frame_2024, textvariable=bayar_2024_var, width=15)
        bayar_2024_entry.grid(row=1, column=1, padx=5, sticky=tk.W)
        
        kurang_2024 = ttk.Label(frame_2024, text=f"Kurang: {self.format_rupiah(anggota['kurang_2024'])}")
        kurang_2024.grid(row=1, column=2, padx=10, sticky=tk.W)
        
        # Frame for 2025
        frame_2025 = ttk.Frame(dialog)
        frame_2025.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(frame_2025, text="Tahun 2025:").grid(row=0, column=0, padx=5, sticky=tk.W)
        ttk.Label(frame_2025, text=f"Iuran Wajib: {self.format_rupiah(anggota['iuran_2025'])}").grid(row=0, column=1, padx=5, sticky=tk.W)
        
        ttk.Label(frame_2025, text="Sudah Dibayar:").grid(row=1, column=0, padx=5, sticky=tk.W)
        bayar_2025_var = tk.StringVar(value=self.format_rupiah(anggota["dibayar_2025"]))
        bayar_2025_entry = ttk.Entry(frame_2025, textvariable=bayar_2025_var, width=15)
        bayar_2025_entry.grid(row=1, column=1, padx=5, sticky=tk.W)
        
        kurang_2025 = ttk.Label(frame_2025, text=f"Kurang: {self.format_rupiah(anggota['kurang_2025'])}")
        kurang_2025.grid(row=1, column=2, padx=10, sticky=tk.W)
        
        def update_kurang():
            # Update kurang labels
            dibayar_2024 = self.parse_rupiah(bayar_2024_var.get())
            kurang = max(0, anggota["iuran_2024"] - dibayar_2024)
            kurang_2024.config(text=f"Kurang: {self.format_rupiah(kurang)}")
            
            dibayar_2025 = self.parse_rupiah(bayar_2025_var.get())
            kurang = max(0, anggota["iuran_2025"] - dibayar_2025)
            kurang_2025.config(text=f"Kurang: {self.format_rupiah(kurang)}")
        
        bayar_2024_entry.bind("<KeyRelease>", lambda e: update_kurang())
        bayar_2025_entry.bind("<KeyRelease>", lambda e: update_kurang())
        
        def save_changes():
            # Update data
            anggota["dibayar_2024"] = self.parse_rupiah(bayar_2024_var.get())
            anggota["dibayar_2025"] = self.parse_rupiah(bayar_2025_var.get())
            
            # Hitung kekurangan
            anggota["kurang_2024"] = max(0, anggota["iuran_2024"] - anggota["dibayar_2024"])
            anggota["kurang_2025"] = max(0, anggota["iuran_2025"] - anggota["dibayar_2025"])
            
            # Update status pembayaran
            anggota["status_bayar_2024"] = "Lunas" if anggota["kurang_2024"] == 0 else "Belum"
            anggota["status_bayar_2025"] = "Lunas" if anggota["kurang_2025"] == 0 else "Belum"
            
            # Refresh summary
            self.show_summary()
            dialog.destroy()
            self.update_status(f"Pembayaran untuk {anggota['nama']} berhasil diperbarui")
            
            # Simpan data
            self.save_data()
        
        # Save button
        ttk.Button(dialog, text="Simpan Perubahan", command=save_changes).pack(pady=10)
    
    def cetak_nota(self):
        if not self.current_anggota:
            return
            
        anggota = self.current_anggota
        
        # Buat jendela khusus untuk nota
        nota_window = tk.Toplevel(self.root)
        nota_window.title("Nota Pembayaran IPARI")
        nota_window.geometry("600x800")
        
        # Buat canvas dengan scrollbar
        canvas = Canvas(nota_window)
        scrollbar = ttk.Scrollbar(nota_window, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Header nota
        header_frame = ttk.Frame(scrollable_frame)
        header_frame.pack(fill=tk.X, padx=20, pady=10)
        
        # Logo IPARI - versi gambar
        try:
            # Muat gambar logo
            logo_img = Image.open(self.logo_path)
            
            # Resize dengan menjaga aspek rasio
            max_size = (180, 180)
            logo_img.thumbnail(max_size, Image.LANCZOS)
            
            # Konversi ke format yang kompatibel dengan Tkinter
            self.logo_photo = ImageTk.PhotoImage(logo_img)
            
            # Buat label untuk menampilkan logo
            logo_label = ttk.Label(header_frame, image=self.logo_photo)
            logo_label.image = self.logo_photo  # Simpan referensi untuk mencegah garbage collection
            logo_label.pack(pady=5)
        except Exception as e:
            print(f"Gagal memuat logo: {e}")
            # Fallback ke teks jika gambar tidak ada
            logo_label = ttk.Label(header_frame, text="LOGO IPARI MENTAWAI", 
                                  font=("Arial", 14, "bold"), foreground="blue")
            logo_label.pack(pady=5)
        
        # Judul
        ttk.Label(header_frame, text="IKATAN PENYULUH AGAMA REPUBLIK INDONESIA", 
                 font=("Arial", 12, "bold")).pack()
        ttk.Label(header_frame, text="Kabupaten Kepulauan Mentawai", 
                 font=("Arial", 11)).pack()
        ttk.Label(header_frame, text="NOTA PEMBAYARAN IURAN", 
                 font=("Arial", 12, "bold"), foreground="navy").pack(pady=10)
        
        # Garis pemisah
        ttk.Separator(scrollable_frame, orient='horizontal').pack(fill='x', padx=20, pady=5)
        
        # Informasi anggota
        info_frame = ttk.Frame(scrollable_frame)
        info_frame.pack(fill=tk.X, padx=20, pady=10)
        
        ttk.Label(info_frame, text="Nama Anggota:", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky=tk.W, pady=2)
        ttk.Label(info_frame, text=anggota['nama'], font=("Arial", 10)).grid(row=0, column=1, sticky=tk.W, pady=2)
        
        ttk.Label(info_frame, text="Status 2024:", font=("Arial", 10, "bold")).grid(row=1, column=0, sticky=tk.W, pady=2)
        ttk.Label(info_frame, text=anggota['status_2024'], font=("Arial", 10)).grid(row=1, column=1, sticky=tk.W, pady=2)
        
        ttk.Label(info_frame, text="Status 2025:", font=("Arial", 10, "bold")).grid(row=2, column=0, sticky=tk.W, pady=2)
        ttk.Label(info_frame, text=anggota['status_2025'], font=("Arial", 10)).grid(row=2, column=1, sticky=tk.W, pady=2)
        
        # Detail pembayaran
        ttk.Label(scrollable_frame, text="RINCIAN PEMBAYARAN", font=("Arial", 11, "bold")).pack(pady=10)
        
        # Tabel pembayaran
        table_frame = ttk.Frame(scrollable_frame)
        table_frame.pack(fill=tk.X, padx=20, pady=5)
        
        # Header tabel
        ttk.Label(table_frame, text="Tahun", font=("Arial", 10, "bold"), borderwidth=1, relief="solid", width=10).grid(row=0, column=0, padx=1, pady=1)
        ttk.Label(table_frame, text="Iuran Wajib", font=("Arial", 10, "bold"), borderwidth=1, relief="solid", width=15).grid(row=0, column=1, padx=1, pady=1)
        ttk.Label(table_frame, text="Dibayar", font=("Arial", 10, "bold"), borderwidth=1, relief="solid", width=15).grid(row=0, column=2, padx=1, pady=1)
        ttk.Label(table_frame, text="Kekurangan", font=("Arial", 10, "bold"), borderwidth=1, relief="solid", width=15).grid(row=0, column=3, padx=1, pady=1)
        ttk.Label(table_frame, text="Status", font=("Arial", 10, "bold"), borderwidth=1, relief="solid", width=10).grid(row=0, column=4, padx=1, pady=1)
        
        # Data 2024
        ttk.Label(table_frame, text="2024", font=("Arial", 9), borderwidth=1, relief="solid", width=10).grid(row=1, column=0, padx=1, pady=1)
        ttk.Label(table_frame, text=self.format_rupiah(anggota['iuran_2024']), font=("Arial", 9), borderwidth=1, relief="solid", width=15).grid(row=1, column=1, padx=1, pady=1)
        ttk.Label(table_frame, text=self.format_rupiah(anggota['dibayar_2024']), font=("Arial", 9), borderwidth=1, relief="solid", width=15).grid(row=1, column=2, padx=1, pady=1)
        ttk.Label(table_frame, text=self.format_rupiah(anggota['kurang_2024']), font=("Arial", 9), borderwidth=1, relief="solid", width=15).grid(row=1, column=3, padx=1, pady=1)
        status_color = "green" if anggota['status_bayar_2024'] == "Lunas" else "red"
        ttk.Label(table_frame, text=anggota['status_bayar_2024'], font=("Arial", 9, "bold"), foreground=status_color, borderwidth=1, relief="solid", width=10).grid(row=1, column=4, padx=1, pady=1)
        
        # Data 2025
        ttk.Label(table_frame, text="2025", font=("Arial", 9), borderwidth=1, relief="solid", width=10).grid(row=2, column=0, padx=1, pady=1)
        ttk.Label(table_frame, text=self.format_rupiah(anggota['iuran_2025']), font=("Arial", 9), borderwidth=1, relief="solid", width=15).grid(row=2, column=1, padx=1, pady=1)
        ttk.Label(table_frame, text=self.format_rupiah(anggota['dibayar_2025']), font=("Arial", 9), borderwidth=1, relief="solid", width=15).grid(row=2, column=2, padx=1, pady=1)
        ttk.Label(table_frame, text=self.format_rupiah(anggota['kurang_2025']), font=("Arial", 9), borderwidth=1, relief="solid", width=15).grid(row=2, column=3, padx=1, pady=1)
        status_color = "green" if anggota['status_bayar_2025'] == "Lunas" else "red"
        ttk.Label(table_frame, text=anggota['status_bayar_2025'], font=("Arial", 9, "bold"), foreground=status_color, borderwidth=1, relief="solid", width=10).grid(row=2, column=4, padx=1, pady=1)
        
        # Catatan
        ttk.Label(scrollable_frame, text="Catatan:", font=("Arial", 10, "bold")).pack(anchor=tk.W, padx=20, pady=5)
        ttk.Label(scrollable_frame, text="1. Pembayaran dapat dilakukan melalui bendahara IPARI Mentawai", font=("Arial", 9)).pack(anchor=tk.W, padx=20)
        ttk.Label(scrollable_frame, text="2. Nota ini sah tanpa tanda tangan basah", font=("Arial", 9)).pack(anchor=tk.W, padx=20)
        ttk.Label(scrollable_frame, text="3. Simpan nota ini sebagai bukti pembayaran", font=("Arial", 9)).pack(anchor=tk.W, padx=20)
        
        # Tanda tangan (versi disederhanakan)
        tanda_tangan_frame = ttk.Frame(scrollable_frame)
        tanda_tangan_frame.pack(fill=tk.X, padx=20, pady=20)
        
        # Kolom kiri (Ketua)
        kadep_frame = ttk.Frame(tanda_tangan_frame)
        kadep_frame.pack(side=tk.LEFT, expand=True)
        
        ttk.Label(kadep_frame, text="Ketua Umum", font=("Arial", 10, "bold")).pack()
        ttk.Label(kadep_frame, text="IPARI Mentawai", font=("Arial", 10)).pack(pady=5)
        ttk.Label(kadep_frame, text="Harris Muda, S.Ag", font=("Arial", 10)).pack(pady=5)
        
        # Kolom kanan (Bendahara)
        bendahara_frame = ttk.Frame(tanda_tangan_frame)
        bendahara_frame.pack(side=tk.RIGHT, expand=True)
        
        ttk.Label(bendahara_frame, text="Bendahara", font=("Arial", 10, "bold")).pack()
        ttk.Label(bendahara_frame, text="IPARI Mentawai", font=("Arial", 10)).pack(pady=5)
        ttk.Label(bendahara_frame, text="Rolina Hutabarat, S.Th", font=("Arial", 10)).pack(pady=5)
        
        # Footer
        footer_frame = ttk.Frame(scrollable_frame)
        footer_frame.pack(fill=tk.X, padx=20, pady=10)
        
        ttk.Label(footer_frame, text="Developed by Oyabuya", font=("Arial", 8)).pack(side=tk.RIGHT)
        ttk.Label(footer_frame, text="© 2025 IPARI Mentawai", font=("Arial", 8)).pack(side=tk.LEFT)
        
        # Tombol cetak
        btn_frame = ttk.Frame(scrollable_frame)
        btn_frame.pack(pady=10)
        
        # Tambahkan flag untuk mencegah multiple click
        self.saving_in_progress = False
        
        def save_action():
            if not self.saving_in_progress:
                self.saving_in_progress = True
                self.simpan_nota(nota_window)
                self.saving_in_progress = False
        
        ttk.Button(btn_frame, text="Simpan sebagai Gambar", command=save_action).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Tutup", command=nota_window.destroy).pack(side=tk.LEFT, padx=5)
    
    def simpan_nota(self, window):
        try:
            # Dapatkan frame utama nota
            main_frame = window.winfo_children()[0]
            
            # Pastikan window sudah di-update untuk mendapatkan ukuran yang benar
            window.update_idletasks()
            
            # Dapatkan koordinat frame relatif terhadap jendela
            x = main_frame.winfo_x()
            y = main_frame.winfo_y()
            width = main_frame.winfo_width()
            height = main_frame.winfo_height()
            
            # Dapatkan koordinat absolut di layar
            abs_x = window.winfo_rootx() + x
            abs_y = window.winfo_rooty() + y
            
            # Buat nama file
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"nota_iuran_{timestamp}.png"
            
            # Periksa apakah file sudah ada sebelumnya
            if os.path.exists(filename):
                # Tambah angka di belakang jika file sudah ada
                base_name, ext = os.path.splitext(filename)
                counter = 1
                while os.path.exists(f"{base_name}_{counter}{ext}"):
                    counter += 1
                filename = f"{base_name}_{counter}{ext}"
            
            # Simpan hanya area nota
            from PIL import ImageGrab
            img = ImageGrab.grab(bbox=(abs_x, abs_y, abs_x + width, abs_y + height))
            img.save(filename)
            
            # Tampilkan konfirmasi hanya jika berhasil
            messagebox.showinfo("Sukses", f"Nota berhasil disimpan sebagai {filename}")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal menyimpan nota: {str(e)}")
    
    def clear_results(self):
        self.results_frame.pack_forget()
        self.current_anggota = None
        self.update_status("Silakan cari atau tambahkan anggota lain")
    
    def show_all_members(self):
        if not self.data_anggota:
            messagebox.showinfo("Anggota", "Belum ada data anggota")
            return
            
        # Buat daftar anggota
        anggota_list = "DAFTAR ANGGOTA:\n"
        anggota_list += "=" * 50 + "\n"
        
        for anggota in self.data_anggota:
            status = "Lunas" if anggota['kurang_2024'] == 0 and anggota['kurang_2025'] == 0 else "Belum Lunas"
            anggota_list += f"- {anggota['nama']} ({status})\n"
        
        messagebox.showinfo("Semua Anggota", anggota_list)
    
    def clear_all(self):
        if not self.data_anggota:
            return
            
        if messagebox.askyesno("Konfirmasi", "Apakah Anda yakin ingin menghapus semua data?"):
            self.data_anggota = []
            self.results_frame.pack_forget()
            self.current_anggota = None
            self.update_status("Semua data berhasil dihapus")
            self.save_data()
    
    def export_to_csv(self):
        if not self.data_anggota:
            messagebox.showwarning("Peringatan", "Tidak ada data untuk diekspor")
            return
            
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("File CSV", "*.csv"), ("Semua File", "*.*")],
            title="Simpan Rekap Iuran"
        )
        
        if not file_path:
            return
            
        try:
            with open(file_path, 'w', newline='', encoding='utf-8') as file:
                writer = csv.writer(file, delimiter=';')
                
                # Header
                writer.writerow(["REKAPITULASI IURAN ANGGOTA IPARI MENTAWAI"])
                writer.writerow([f"Tanggal Ekspor: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"])
                writer.writerow([])
                
                # Header kolom
                headers = [
                    "No", "Nama Anggota", "Status 2024", "Status 2025", "Kondisi", 
                    "Iuran 2024", "Iuran 2025", 
                    "Dibayar 2024", "Kurang 2024", "Status 2024",
                    "Dibayar 2025", "Kurang 2025", "Status 2025"
                ]
                writer.writerow(headers)
                
                # Data anggota
                for idx, anggota in enumerate(self.data_anggota, 1):
                    row = [
                        idx,
                        anggota['nama'],
                        anggota['status_2024'],
                        anggota['status_2025'],
                        anggota['kondisi'],
                        self.format_rupiah(anggota['iuran_2024']),
                        self.format_rupiah(anggota['iuran_2025']),
                        self.format_rupiah(anggota['dibayar_2024']),
                        self.format_rupiah(anggota['kurang_2024']),
                        anggota['status_bayar_2024'],
                        self.format_rupiah(anggota['dibayar_2025']),
                        self.format_rupiah(anggota['kurang_2025']),
                        anggota['status_bayar_2025']
                    ]
                    writer.writerow(row)
            
            self.update_status(f"Data berhasil diekspor ke: {file_path}")
            messagebox.showinfo("Sukses", "Data berhasil diekspor ke CSV")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal mengekspor data: {str(e)}")
    
    def show_statistics(self):
        """Menampilkan statistik ringkasan pembayaran"""
        if not self.data_anggota:
            messagebox.showinfo("Statistik", "Belum ada data anggota")
            return
            
        # Hitung statistik
        total_anggota = len(self.data_anggota)
        lunas_2024 = sum(1 for a in self.data_anggota if a['status_bayar_2024'] == 'Lunas')
        lunas_2025 = sum(1 for a in self.data_anggota if a['status_bayar_2025'] == 'Lunas')
        belum_2024 = total_anggota - lunas_2024
        belum_2025 = total_anggota - lunas_2025
        
        total_iuran_2024 = sum(a['dibayar_2024'] for a in self.data_anggota)
        total_iuran_2025 = sum(a['dibayar_2025'] for a in self.data_anggota)
        total_kurang_2024 = sum(a['kurang_2024'] for a in self.data_anggota)
        total_kurang_2025 = sum(a['kurang_2025'] for a in self.data_anggota)
        
        # Format pesan statistik
        stats = f"STATISTIK IPARI MENTAWAI\n"
        stats += "=" * 40 + "\n\n"
        stats += f"Total Anggota: {total_anggota}\n\n"
        
        stats += "TAHUN 2024:\n"
        stats += f"- Sudah Lunas: {lunas_2024} ({lunas_2024/total_anggota:.1%})\n"
        stats += f"- Belum Lunas: {belum_2024} ({belum_2024/total_anggota:.1%})\n"
        stats += f"Total Iuran Terkumpul: {self.format_rupiah(total_iuran_2024)}\n"
        stats += f"Total Kekurangan: {self.format_rupiah(total_kurang_2024)}\n\n"
        
        stats += "TAHUN 2025:\n"
        stats += f"- Sudah Lunas: {lunas_2025} ({lunas_2025/total_anggota:.1%})\n"
        stats += f"- Belum Lunas: {belum_2025} ({belum_2025/total_anggota:.1%})\n"
        stats += f"Total Iuran Terkumpul: {self.format_rupiah(total_iuran_2025)}\n"
        stats += f"Total Kekurangan: {self.format_rupiah(total_kurang_2025)}"
        
        messagebox.showinfo("Statistik Pembayaran Iuran", stats)
    
    def cetak_semua_nota(self):
        """Mencetak semua nota anggota dalam satu PDF"""
        if not self.data_anggota:
            messagebox.showwarning("Peringatan", "Tidak ada data anggota untuk dicetak")
            return
            
        # Buat folder untuk menyimpan nota
        folder_name = f"nota_iuran_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        os.makedirs(folder_name, exist_ok=True)
        
        # Cetak nota untuk setiap anggota
        for anggota in self.data_anggota:
            self.current_anggota = anggota
            
            # Buat jendela khusus untuk nota
            nota_window = tk.Toplevel(self.root)
            nota_window.title(f"Nota Pembayaran IPARI - {anggota['nama']}")
            nota_window.geometry("600x800")
            
            # Buat canvas dengan scrollbar
            canvas = Canvas(nota_window)
            scrollbar = ttk.Scrollbar(nota_window, orient="vertical", command=canvas.yview)
            scrollable_frame = ttk.Frame(canvas)
            
            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )
            
            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)
            
            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")
            
            # Header nota
            header_frame = ttk.Frame(scrollable_frame)
            header_frame.pack(fill=tk.X, padx=20, pady=10)
            
            # Logo IPARI - versi gambar
            try:
                # Muat gambar logo
                logo_img = Image.open(self.logo_path)
                
                # Resize dengan menjaga aspek rasio
                max_size = (180, 180)
                logo_img.thumbnail(max_size, Image.LANCZOS)
                
                # Konversi ke format yang kompatibel dengan Tkinter
                self.logo_photo = ImageTk.PhotoImage(logo_img)
                
                # Buat label untuk menampilkan logo
                logo_label = ttk.Label(header_frame, image=self.logo_photo)
                logo_label.image = self.logo_photo
                logo_label.pack(pady=5)
            except Exception as e:
                print(f"Gagal memuat logo: {e}")
                logo_label = ttk.Label(header_frame, text="LOGO IPARI MENTAWAI", 
                                      font=("Arial", 14, "bold"), foreground="blue")
                logo_label.pack(pady=5)
            
            # Judul
            ttk.Label(header_frame, text="IKATAN PENYULUH AGAMA REPUBLIK INDONESIA", 
                     font=("Arial", 12, "bold")).pack()
            ttk.Label(header_frame, text="Kabupaten Kepulauan Mentawai", 
                     font=("Arial", 11)).pack()
            ttk.Label(header_frame, text="NOTA PEMBAYARAN IURAN", 
                     font=("Arial", 12, "bold"), foreground="navy").pack(pady=10)
            
            # Garis pemisah
            ttk.Separator(scrollable_frame, orient='horizontal').pack(fill='x', padx=20, pady=5)
            
            # Informasi anggota
            info_frame = ttk.Frame(scrollable_frame)
            info_frame.pack(fill=tk.X, padx=20, pady=10)
            
            ttk.Label(info_frame, text="Nama Anggota:", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky=tk.W, pady=2)
            ttk.Label(info_frame, text=anggota['nama'], font=("Arial", 10)).grid(row=0, column=1, sticky=tk.W, pady=2)
            
            ttk.Label(info_frame, text="Status 2024:", font=("Arial", 10, "bold")).grid(row=1, column=0, sticky=tk.W, pady=2)
            ttk.Label(info_frame, text=anggota['status_2024'], font=("Arial", 10)).grid(row=1, column=1, sticky=tk.W, pady=2)
            
            ttk.Label(info_frame, text="Status 2025:", font=("Arial", 10, "bold")).grid(row=2, column=0, sticky=tk.W, pady=2)
            ttk.Label(info_frame, text=anggota['status_2025'], font=("Arial", 10)).grid(row=2, column=1, sticky=tk.W, pady=2)
            
            # Detail pembayaran
            ttk.Label(scrollable_frame, text="RINCIAN PEMBAYARAN", font=("Arial", 11, "bold")).pack(pady=10)
            
            # Tabel pembayaran
            table_frame = ttk.Frame(scrollable_frame)
            table_frame.pack(fill=tk.X, padx=20, pady=5)
            
            # Header tabel
            ttk.Label(table_frame, text="Tahun", font=("Arial", 10, "bold"), borderwidth=1, relief="solid", width=10).grid(row=0, column=0, padx=1, pady=1)
            ttk.Label(table_frame, text="Iuran Wajib", font=("Arial", 10, "bold"), borderwidth=1, relief="solid", width=15).grid(row=0, column=1, padx=1, pady=1)
            ttk.Label(table_frame, text="Dibayar", font=("Arial", 10, "bold"), borderwidth=1, relief="solid", width=15).grid(row=0, column=2, padx=1, pady=1)
            ttk.Label(table_frame, text="Kekurangan", font=("Arial", 10, "bold"), borderwidth=1, relief="solid", width=15).grid(row=0, column=3, padx=1, pady=1)
            ttk.Label(table_frame, text="Status", font=("Arial", 10, "bold"), borderwidth=1, relief="solid", width=10).grid(row=0, column=4, padx=1, pady=1)
            
            # Data 2024
            ttk.Label(table_frame, text="2024", font=("Arial", 9), borderwidth=1, relief="solid", width=10).grid(row=1, column=0, padx=1, pady=1)
            ttk.Label(table_frame, text=self.format_rupiah(anggota['iuran_2024']), font=("Arial", 9), borderwidth=1, relief="solid", width=15).grid(row=1, column=1, padx=1, pady=1)
            ttk.Label(table_frame, text=self.format_rupiah(anggota['dibayar_2024']), font=("Arial", 9), borderwidth=1, relief="solid", width=15).grid(row=1, column=2, padx=1, pady=1)
            ttk.Label(table_frame, text=self.format_rupiah(anggota['kurang_2024']), font=("Arial", 9), borderwidth=1, relief="solid", width=15).grid(row=1, column=3, padx=1, pady=1)
            status_color = "green" if anggota['status_bayar_2024'] == "Lunas" else "red"
            ttk.Label(table_frame, text=anggota['status_bayar_2024'], font=("Arial", 9, "bold"), foreground=status_color, borderwidth=1, relief="solid", width=10).grid(row=1, column=4, padx=1, pady=1)
            
            # Data 2025
            ttk.Label(table_frame, text="2025", font=("Arial", 9), borderwidth=1, relief="solid", width=10).grid(row=2, column=0, padx=1, pady=1)
            ttk.Label(table_frame, text=self.format_rupiah(anggota['iuran_2025']), font=("Arial", 9), borderwidth=1, relief="solid", width=15).grid(row=2, column=1, padx=1, pady=1)
            ttk.Label(table_frame, text=self.format_rupiah(anggota['dibayar_2025']), font=("Arial", 9), borderwidth=1, relief="solid", width=15).grid(row=2, column=2, padx=1, pady=1)
            ttk.Label(table_frame, text=self.format_rupiah(anggota['kurang_2025']), font=("Arial", 9), borderwidth=1, relief="solid", width=15).grid(row=2, column=3, padx=1, pady=1)
            status_color = "green" if anggota['status_bayar_2025'] == "Lunas" else "red"
            ttk.Label(table_frame, text=anggota['status_bayar_2025'], font=("Arial", 9, "bold"), foreground=status_color, borderwidth=1, relief="solid", width=10).grid(row=2, column=4, padx=1, pady=1)
            
            # Catatan
            ttk.Label(scrollable_frame, text="Catatan:", font=("Arial", 10, "bold")).pack(anchor=tk.W, padx=20, pady=5)
            ttk.Label(scrollable_frame, text="1. Pembayaran dapat dilakukan melalui bendahara IPARI Mentawai", font=("Arial", 9)).pack(anchor=tk.W, padx=20)
            ttk.Label(scrollable_frame, text="2. Nota ini sah tanpa tanda tangan basah", font=("Arial", 9)).pack(anchor=tk.W, padx=20)
            ttk.Label(scrollable_frame, text="3. Simpan nota ini sebagai bukti pembayaran", font=("Arial", 9)).pack(anchor=tk.W, padx=20)
            
            # Tanda tangan
            tanda_tangan_frame = ttk.Frame(scrollable_frame)
            tanda_tangan_frame.pack(fill=tk.X, padx=20, pady=20)
            
            # Kolom kiri (Ketua)
            kadep_frame = ttk.Frame(tanda_tangan_frame)
            kadep_frame.pack(side=tk.LEFT, expand=True)
            
            ttk.Label(kadep_frame, text="Ketua Umum", font=("Arial", 10, "bold")).pack()
            ttk.Label(kadep_frame, text="IPARI Mentawai", font=("Arial", 10)).pack(pady=5)
            ttk.Label(kadep_frame, text="Harris Muda, S.Ag", font=("Arial", 10)).pack(pady=5)
            
            # Kolom kanan (Bendahara)
            bendahara_frame = ttk.Frame(tanda_tangan_frame)
            bendahara_frame.pack(side=tk.RIGHT, expand=True)
            
            ttk.Label(bendahara_frame, text="Bendahara", font=("Arial", 10, "bold")).pack()
            ttk.Label(bendahara_frame, text="IPARI Mentawai", font=("Arial", 10)).pack(pady=5)
            ttk.Label(bendahara_frame, text="Rolina Hutabarat, S.Th", font=("Arial", 10)).pack(pady=5)
            
            # Footer
            footer_frame = ttk.Frame(scrollable_frame)
            footer_frame.pack(fill=tk.X, padx=20, pady=10)
            
            ttk.Label(footer_frame, text="Developed by Oyabuya", font=("Arial", 8)).pack(side=tk.RIGHT)
            ttk.Label(footer_frame, text="© 2025 IPARI Mentawai", font=("Arial", 8)).pack(side=tk.LEFT)
            
            # Pastikan window sudah di-update
            nota_window.update_idletasks()
            
            # Dapatkan frame utama nota
            main_frame = nota_window.winfo_children()[0]
            
            # Dapatkan koordinat frame relatif terhadap jendela
            x = main_frame.winfo_x()
            y = main_frame.winfo_y()
            width = main_frame.winfo_width()
            height = main_frame.winfo_height()
            
            # Dapatkan koordinat absolut di layar
            abs_x = nota_window.winfo_rootx() + x
            abs_y = nota_window.winfo_rooty() + y
            
            # Buat nama file
            safe_name = ''.join(c for c in anggota['nama'] if c.isalnum() or c in " _-")
            filename = os.path.join(folder_name, f"nota_{safe_name}.png")
            
            # Simpan hanya area nota
            from PIL import ImageGrab
            img = ImageGrab.grab(bbox=(abs_x, abs_y, abs_x + width, abs_y + height))
            img.save(filename)
            
            # Tutup jendela nota
            nota_window.destroy()
        
        # Tampilkan konfirmasi
        messagebox.showinfo("Sukses", f"Semua nota berhasil disimpan di folder:\n{folder_name}")
        # Buka folder
        os.startfile(folder_name)
    
    def update_status(self, message):
        self.status_var.set(message)
        self.root.update_idletasks()

if __name__ == "__main__":
    root = tk.Tk()
    app = IuranApp(root)
    root.mainloop()
