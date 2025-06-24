import tkinter as tk
from tkinter import ttk, messagebox, font
from PIL import Image, ImageTk
import pyodbc
from datetime import datetime, timedelta
import threading
import time
import os
from twilio.rest import Client

# Warna Modern
COLOR_PRIMARY = "#2b5876"  # Biru tua
COLOR_SECONDARY = "#4e4376"  # Ungu
COLOR_ACCENT = "#00c6fb"  # Biru muda
COLOR_BACKGROUND = "#f5f7fa"  # Abu-abu muda
COLOR_TEXT = "#333333"  # Hitam soft
COLOR_WHITE = "#ffffff"  # Putih

# Konfigurasi Twilio
TWILIO_ACCOUNT_SID = 'your_account_sid'
TWILIO_AUTH_TOKEN = 'your_auth_token'
TWILIO_PHONE_NUMBER = 'your_twilio_number'
client = Client(TWILIO_ACCOUNT_SID, TWILIO_AUTH_TOKEN)

def get_connection():
    try:
        conn = pyodbc.connect(
            'DRIVER={SQL Server};'
            'SERVER=DESKTOP-I3I1U61;'
            'DATABASE=inspecure;'
            'UID=sa;PWD=admin123;'
        )
        return conn
    except Exception as e:
        messagebox.showerror("Koneksi Error", f"Gagal terhubung ke database: {str(e)}")
        raise

def check_login(username, password):
    try:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM users WHERE username = ? AND password = ?", (username, password))
        result = cursor.fetchone()
        conn.close()
        return result is not None
    except Exception as e:
        messagebox.showerror("Database Error", f"Gagal terhubung ke database: {str(e)}")
        return False

def insert_data(seri_alat, nama_alat, perusahaan, tgl_sertifikasi, wa_perusahaan, wa_inspector):
    try:
        conn = get_connection()
        cursor = conn.cursor()

        tgl_awal_obj = datetime.strptime(tgl_sertifikasi, "%Y-%m-%d")
        tgl_akhir_obj = tgl_awal_obj + timedelta(days=3*365)

        sql = """INSERT INTO sertifikasi_alat 
                 (seri_alat, nama_alat, nama_perusahaan, tanggal_awal, tanggal_akhir, 
                  no_wa_perusahaan, no_wa_inspector)
                 VALUES (?, ?, ?, ?, ?, ?, ?)"""
        val = (seri_alat, nama_alat, perusahaan, tgl_awal_obj, tgl_akhir_obj, 
               wa_perusahaan, wa_inspector)

        cursor.execute(sql, val)
        conn.commit()
        conn.close()
        messagebox.showinfo("Sukses", "Data sertifikasi berhasil disimpan!")
        return True
    except Exception as e:
        messagebox.showerror("Error", str(e))
        return False

def delete_data(id_alat):
    try:
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute("DELETE FROM sertifikasi_alat WHERE id = ?", (id_alat,))
        conn.commit()
        conn.close()
        messagebox.showinfo("Sukses", "Data berhasil dihapus!")
        return True
    except Exception as e:
        messagebox.showerror("Error", str(e))
        return False

def check_notifications():
    while True:
        try:
            conn = get_connection()
            cursor = conn.cursor()
            
            tiga_bulan_lagi = datetime.now() + timedelta(days=90)
            cursor.execute("""
                SELECT nama_perusahaan, nama_alat, tanggal_akhir, no_wa_inspector, no_wa_perusahaan 
                FROM sertifikasi_alat 
                WHERE tanggal_akhir BETWEEN ? AND ? AND notifikasi_terkirim = 0
            """, (datetime.now(), tiga_bulan_lagi))
            
            records = cursor.fetchall()
            
            for record in records:
                perusahaan, alat, tgl_akhir, wa_inspector, wa_perusahaan = record
                pesan = (
                    f"🔔 *Pengingat Resertifikasi Alat* 🔔\n\n"
                    f"*Alat*: {alat}\n"
                    f"*Perusahaan*: {perusahaan}\n"
                    f"*Tanggal Resertifikasi*: {tgl_akhir.strftime('%d %B %Y')}\n\n"
                    f"Segera lakukan resertifikasi sebelum tanggal tersebut."
                )
                
                if wa_perusahaan:
                    send_whatsapp(wa_perusahaan, pesan)
                
                if wa_inspector:
                    send_whatsapp(wa_inspector, pesan)
                
                cursor.execute("""
                    UPDATE sertifikasi_alat 
                    SET notifikasi_terkirim = 1 
                    WHERE nama_perusahaan = ? AND nama_alat = ?
                """, (perusahaan, alat))
                conn.commit()
            
            conn.close()
        except Exception as e:
            print(f"Error dalam pengecekan notifikasi: {e}")
        
        time.sleep(86400)

def send_whatsapp(to_number, message):
    try:
        message = client.messages.create(
            body=message,
            from_='whatsapp:' + TWILIO_PHONE_NUMBER,
            to='whatsapp:' + to_number
        )
        print(f"Notifikasi terkirim ke {to_number}")
    except Exception as e:
        print(f"Gagal mengirim WA ke {to_number}: {e}")

def show_dashboard():
    win = tk.Tk()
    win.title("Inspecure - Sistem Manajemen Sertifikasi Alat")
    win.geometry("1200x700")
    win.configure(bg=COLOR_BACKGROUND)
    
    # Style configuration
    style = ttk.Style()
    style.theme_use('clam')
    
    # Configure styles
    style.configure('.', background=COLOR_BACKGROUND, foreground=COLOR_TEXT)
    style.configure('TFrame', background=COLOR_BACKGROUND)
    style.configure('TLabel', background=COLOR_BACKGROUND, foreground=COLOR_TEXT, font=('Helvetica', 10))
    style.configure('Header.TLabel', font=('Helvetica', 14, 'bold'), foreground=COLOR_PRIMARY)
    style.configure('TButton', font=('Helvetica', 10), 
                   background=COLOR_ACCENT, foreground=COLOR_WHITE,
                   borderwidth=0, focusthickness=3, focuscolor='none')
    style.map('TButton', background=[('active', COLOR_SECONDARY)])
    style.configure('TEntry', font=('Helvetica', 10), padding=5)
    style.configure('TNotebook', background=COLOR_BACKGROUND, borderwidth=0)
    style.configure('TNotebook.Tab', background=COLOR_BACKGROUND, foreground=COLOR_TEXT, 
                   padding=[10, 5], font=('Helvetica', 10, 'bold'))
    style.map('TNotebook.Tab', background=[('selected', COLOR_ACCENT)], foreground=[('selected', COLOR_WHITE)])
    style.configure('Treeview', background=COLOR_WHITE, fieldbackground=COLOR_WHITE, 
                   foreground=COLOR_TEXT, rowheight=25, borderwidth=0)
    style.configure('Treeview.Heading', background=COLOR_PRIMARY, foreground=COLOR_WHITE, 
                   font=('Helvetica', 10, 'bold'), borderwidth=0)
    style.configure('Treeview.Row', background=COLOR_WHITE, fieldbackground=COLOR_WHITE)
    
    # Header
    header_frame = ttk.Frame(win, style='TFrame')
    header_frame.pack(fill=tk.X, padx=20, pady=20)
    
    # Logo perusahaan
    if os.path.exists("Logo.png"):
        logo_img = Image.open("logo.png")
        logo_img = logo_img.resize((80, 40), Image.LANCZOS)
        logo = ImageTk.PhotoImage(logo_img)
        logo_label = tk.Label(header_frame, image=logo, bg=COLOR_BACKGROUND)
        logo_label.image = logo
        logo_label.pack(side=tk.LEFT, padx=10)
    
    title_frame = ttk.Frame(header_frame, style='TFrame')
    title_frame.pack(side=tk.LEFT)
    ttk.Label(title_frame, text="INSPECURE", style='Header.TLabel').pack(anchor='w')
    ttk.Label(title_frame, text="Inspector, Secure, and Ensure").pack(anchor='w')
    
    # Tab Control
    tab_control = ttk.Notebook(win)
    
    # Tab 1: Input Sertifikasi
    tab_input = ttk.Frame(tab_control)
    tab_control.add(tab_input, text='Input Sertifikasi')
    
    # Form container
    form_container = ttk.Frame(tab_input, style='TFrame')
    form_container.pack(padx=30, pady=30, fill=tk.BOTH, expand=True)
    
    # Form title
    ttk.Label(form_container, text="Input Data Sertifikasi", style='Header.TLabel').grid(row=0, column=0, columnspan=2, pady=(0, 20))
    
    # Form fields
    fields = [
        ("Nama Alat", "nama_alat"),
        ("Seri Alat", "seri_alat"),
        ("Nama Perusahaan", "perusahaan"),
        ("Tanggal Sertifikasi (YYYY-MM-DD)", "tgl_sertifikasi"),
        ("Kontak Perusahaan (WA)", "wa_perusahaan"),
        ("Kontak Inspektor (WA)", "wa_inspector")
    ]
    
    entries = {}
    for i, (label_text, field_name) in enumerate(fields, start=1):
        ttk.Label(form_container, text=label_text).grid(row=i, column=0, padx=10, pady=8, sticky='e')
        entry = ttk.Entry(form_container)
        entry.grid(row=i, column=1, padx=10, pady=8, sticky='ew', ipady=4)
        entries[field_name] = entry
    
    # Button frame
    btn_frame = ttk.Frame(form_container, style='TFrame')
    btn_frame.grid(row=len(fields)+1, column=0, columnspan=2, pady=20)
    
    def simpan():
        data = {
            'seri_alat': entries['seri_alat'].get(),
            'nama_alat': entries['nama_alat'].get(),
            'perusahaan': entries['perusahaan'].get(),
            'tgl_sertifikasi': entries['tgl_sertifikasi'].get(),
            'wa_perusahaan': entries['wa_perusahaan'].get(),
            'wa_inspector': entries['wa_inspector'].get()
        }
        
        if insert_data(**data):
            for entry in entries.values():
                entry.delete(0, tk.END)
            load_data()
    
    ttk.Button(btn_frame, text="Simpan Data", command=simpan).pack(side=tk.LEFT, padx=5)
    
    # Tab 2: Daftar Sertifikasi
    tab_daftar = ttk.Frame(tab_control)
    tab_control.add(tab_daftar, text='Daftar Sertifikasi')
    
    # Table container
    table_container = ttk.Frame(tab_daftar, style='TFrame')
    table_container.pack(padx=20, pady=20, fill=tk.BOTH, expand=True)
    
    # Treeview
    columns = ("no", "id", "alat", "seri", "perusahaan", "tgl_sertifikasi", "tgl_resertifikasi", 
               "kontak_perusahaan", "resertifikasi_ke", "status", "keterangan")
    
    tree = ttk.Treeview(table_container, columns=columns, show='headings', height=20, selectmode='browse')
    
    # Configure columns
    tree.heading("no", text="NO")
    tree.heading("id", text="ID")  # Hidden column for ID
    tree.heading("alat", text="Nama Alat")
    tree.heading("seri", text="Seri Alat")
    tree.heading("perusahaan", text="Perusahaan")
    tree.heading("tgl_sertifikasi", text="Tanggal Sertifikasi")
    tree.heading("tgl_resertifikasi", text="Tanggal Resertifikasi")
    tree.heading("kontak_perusahaan", text="Kontak Perusahaan")
    tree.heading("resertifikasi_ke", text="Resertifikasi Ke")
    tree.heading("status", text="Status")
    tree.heading("keterangan", text="Keterangan")
    
    tree.column("no", width=50, anchor='center')
    tree.column("id", width=0, stretch=tk.NO)  # Hide ID column
    tree.column("alat", width=150, anchor='w')
    tree.column("seri", width=100, anchor='w')
    tree.column("perusahaan", width=200, anchor='w')
    tree.column("tgl_sertifikasi", width=120, anchor='center')
    tree.column("tgl_resertifikasi", width=120, anchor='center')
    tree.column("kontak_perusahaan", width=120, anchor='center')
    tree.column("resertifikasi_ke", width=100, anchor='center')
    tree.column("status", width=150, anchor='center')
    tree.column("keterangan", width=150, anchor='center')
    
    # Scrollbar
    scrollbar = ttk.Scrollbar(table_container, orient=tk.VERTICAL, command=tree.yview)
    tree.configure(yscroll=scrollbar.set)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    tree.pack(fill=tk.BOTH, expand=True)
    
    # Button frame for table
    table_btn_frame = ttk.Frame(table_container, style='TFrame')
    table_btn_frame.pack(pady=10)
    
    def load_data():
        try:
            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute("""
                SELECT 
                    id, seri_alat, nama_alat, nama_perusahaan, 
                    CONVERT(varchar, tanggal_awal, 23) as tgl_awal, 
                    CONVERT(varchar, tanggal_akhir, 23) as tgl_akhir,
                    no_wa_perusahaan,
                    DATEDIFF(month, GETDATE(), tanggal_akhir) as bulan_akhir,
                    CASE 
                        WHEN DATEDIFF(month, GETDATE(), tanggal_akhir) <= 3 THEN 'Aktif - Segera Resertifikasi'
                        WHEN DATEDIFF(month, GETDATE(), tanggal_akhir) <= 6 THEN 'Aktif - Perlu Persiapan'
                        ELSE 'Aktif'
                    END as status,
                    CASE 
                        WHEN DATEDIFF(month, GETDATE(), tanggal_akhir) <= 0 THEN 'Sudah Kadaluarsa'
                        ELSE CAST(DATEDIFF(month, GETDATE(), tanggal_akhir) AS VARCHAR) + ' bulan lagi'
                    END as keterangan
                FROM sertifikasi_alat
                ORDER BY tanggal_akhir ASC
            """)
            
            for row in tree.get_children():
                tree.delete(row)
            
            for i, row in enumerate(cursor.fetchall(), 1):
                id_alat, seri, alat, perusahaan, tgl_awal, tgl_akhir, wa_perusahaan, bulan_akhir, status, keterangan = row
                resertifikasi_ke = (datetime.strptime(tgl_akhir, "%Y-%m-%d").year - datetime.strptime(tgl_awal, "%Y-%m-%d").year) // 3
                
                tree.insert("", tk.END, values=(
                    i, id_alat, alat, seri, perusahaan, tgl_awal, tgl_akhir, 
                    wa_perusahaan, resertifikasi_ke, status, keterangan
                ))
            
            conn.close()
        except Exception as e:
            messagebox.showerror("Error", f"Gagal memuat data: {str(e)}")
    
    def hapus_data():
        selected_item = tree.selection()
        if not selected_item:
            messagebox.showwarning("Peringatan", "Silakan pilih data yang akan dihapus")
            return
        
        item_data = tree.item(selected_item)
        id_alat = item_data['values'][1]  # Ambil ID dari kolom tersembunyi
        
        if messagebox.askyesno("Konfirmasi", "Apakah Anda yakin ingin menghapus data ini?"):
            if delete_data(id_alat):
                load_data()
    
    ttk.Button(table_btn_frame, text="Refresh Data", command=load_data).pack(side=tk.LEFT, padx=5)
    ttk.Button(table_btn_frame, text="Hapus Data", command=hapus_data).pack(side=tk.LEFT, padx=5)
    
    load_data()
    
    tab_control.pack(expand=1, fill="both", padx=10, pady=10)
    
    # Footer
    footer_frame = ttk.Frame(win, style='TFrame')
    footer_frame.pack(fill=tk.X, padx=20, pady=10)
    ttk.Label(footer_frame, text="© 2023 Inspecure - Sistem Manajemen Sertifikasi Alat").pack(side=tk.RIGHT)
    
    win.mainloop()

def show_login():
    root = tk.Tk()
    root.title("Inspecure - Login")
    root.geometry("500x450")
    root.configure(bg=COLOR_BACKGROUND)
    
    # Style configuration
    style = ttk.Style()
    style.theme_use('clam')
    
    # Configure styles
    style.configure('.', background=COLOR_BACKGROUND, foreground=COLOR_TEXT)
    style.configure('TFrame', background=COLOR_BACKGROUND)
    style.configure('TLabel', background=COLOR_BACKGROUND, foreground=COLOR_TEXT, font=('Helvetica', 12))
    style.configure('Header.TLabel', font=('Helvetica', 20, 'bold'), foreground=COLOR_PRIMARY)
    style.configure('Subheader.TLabel', font=('Helvetica', 14), foreground=COLOR_SECONDARY)
    style.configure('TButton', font=('Helvetica', 12), 
                   background=COLOR_ACCENT, foreground=COLOR_WHITE,
                   borderwidth=0, focusthickness=3, focuscolor='none')
    style.map('TButton', background=[('active', COLOR_SECONDARY)])
    style.configure('TEntry', font=('Helvetica', 12), padding=5)
    
    # Main container
    main_frame = ttk.Frame(root, style='TFrame')
    main_frame.pack(pady=40, padx=40, fill=tk.BOTH, expand=True)
    
    # Header with logo
    header_frame = ttk.Frame(main_frame, style='TFrame')
    header_frame.pack(pady=(0, 30))
    
    # Logo
    if os.path.exists("logo.png"):
        logo_img = Image.open("logo.png")
        logo_img = logo_img.resize((100, 100), Image.LANCZOS)
        logo = ImageTk.PhotoImage(logo_img)
        logo_label = tk.Label(header_frame, image=logo, bg=COLOR_BACKGROUND)
        logo_label.image = logo
        logo_label.pack()
    
    # App title
    ttk.Label(header_frame, text="SELAMAT DATANG DI INSPECURE", 
              style='Header.TLabel').pack(pady=10)
    ttk.Label(header_frame, text="Inspector, Secure, and Ensure", 
              style='Subheader.TLabel').pack()
    
    # Login form
    form_frame = ttk.Frame(main_frame, style='TFrame')
    form_frame.pack(pady=20)
    
    # Username field
    ttk.Label(form_frame, text="Username").grid(row=0, column=0, padx=10, pady=10, sticky='e')
    user_entry = ttk.Entry(form_frame)
    user_entry.grid(row=0, column=1, padx=10, pady=10, ipady=5)
    
    # Password field
    ttk.Label(form_frame, text="Password").grid(row=1, column=0, padx=10, pady=10, sticky='e')
    pass_entry = ttk.Entry(form_frame, show="*")
    pass_entry.grid(row=1, column=1, padx=10, pady=10, ipady=5)
    
    # Login button
    def login():
        user = user_entry.get()
        pw = pass_entry.get()
        if check_login(user, pw):
            root.destroy()
            notification_thread = threading.Thread(target=check_notifications, daemon=True)
            notification_thread.start()
            show_dashboard()
        else:
            messagebox.showerror("Gagal", "Username atau password salah!")
    
    btn_frame = ttk.Frame(main_frame, style='TFrame')
    btn_frame.pack(pady=20)
    
    login_btn = ttk.Button(btn_frame, text="Login", command=login, width=15)
    login_btn.pack()
    
    root.mainloop()

if __name__ == "__main__":
    show_login()