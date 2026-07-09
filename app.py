from flask import Flask, render_template, request, jsonify, redirect, url_for, send_file
import pandas as pd
import os

app = Flask(__name__)

# Tentukan lokasi file Excel master tersimpan
EXCEL_FILE = 'master_produk.xlsx'

# Variabel global untuk menyimpan data produk di memori Python
MASTER_PRODUK = {}

def muat_master_produk():
    """Fungsi untuk membaca data dari file Excel ke dalam aplikasi"""
    global MASTER_PRODUK
    MASTER_PRODUK = {}
    
    if os.path.exists(EXCEL_FILE):
        try:
            df = pd.read_excel(EXCEL_FILE)
            for _, row in df.iterrows():
                # Membersihkan format barcode dari spasi atau pecahan desimal
                bc = str(row['barcode']).strip().split('.')[0]
                
                MASTER_PRODUK[bc] = {
                    "sku": str(row['sku']).strip(),
                    "brand": str(row['brand']).strip(),
                    "name": str(row['product_name']).strip(),
                    "variant": str(row['variant']).strip()
                }
            print(f"\n==================================================")
            print(f" STATUS: BERHASIL MEMUAT {len(MASTER_PRODUK)} PRODUK DARI EXCEL!")
            print(f"==================================================\n")
        except Exception as e:
            print(f"\n[ERROR] Gagal membaca file Excel: {e}\n")
    else:
        print(f"\n[PERINGATAN] File '{EXCEL_FILE}' belum ada. Silakan upload via halaman /superuser\n")

# Jalankan muat produk otomatis saat aplikasi dinyalakan
muat_master_produk()

# Tempat menyimpan data hasil scan dari 17 user selama aplikasi berjalan
HASIL_OPNAME = []

# --- 1. ROUTE LOGIN ---
@app.route('/')
def index():
    zonas = ["PICKFACE", "STORAGE", "OFFLINE", "DAMAGED-STORE", "DAMAGED-WH"]
    return render_template('login.html', zonas=zonas)

# --- 2. ROUTE HALAMAN SCAN ---
@app.route('/scan')
def scan_page():
    username = request.args.get('username')
    zona = request.args.get('zona')
    if not username or not zona:
        return redirect(url_for('index'))
    return render_template('scan.html', username=username, zona=zona)

# --- 3. API: CEK BARCODE ---
@app.route('/api/cek_barcode', methods=['POST'])
def cek_barcode():
    data = request.json
    barcode_input = str(data.get('barcode')).strip().split('.')[0]
    produk = MASTER_PRODUK.get(barcode_input)
    
    if produk:
        return jsonify({"status": "success", "data": produk})
    return jsonify({"status": "error", "message": f"Barcode [{barcode_input}] tidak terdaftar di Master!"})

# --- 4. API: SIMPAN HASIL OPNAME ---
@app.route('/api/simpan_opname', methods=['POST'])
def simpan_opname():
    data = request.json
    record = {
        "id": len(HASIL_OPNAME) + 1,
        "username": data.get('username'),
        "zona": data.get('zona'),
        "sublokasi": data.get('sublokasi'),
        "barcode": data.get('barcode'),
        "sku": data.get('sku'),
        "brand": data.get('brand'),
        "product_name": data.get('product_name'),
        "variant": data.get('variant'),
        "qty": int(data.get('qty'))
    }
    HASIL_OPNAME.append(record)
    return jsonify({"status": "success", "message": "Data stock opname berhasil disimpan!"})

# --- 5. ROUTE: DASHBOARD SUPER USER & UPLOAD EXCEL MASTER ---
@app.route('/superuser', methods=['GET', 'POST'])
def superuser_page():
    if request.method == 'POST':
        if 'file_excel' not in request.files:
            return "Gagal: Tidak ada form file ditemukan.", 400
        
        file = request.files['file_excel']
        if file.filename == '':
            return "Gagal: Anda belum memilih file Excel.", 400
            
        if file and (file.filename.endswith('.xlsx') or file.filename.endswith('.xls')):
            try:
                file.save(EXCEL_FILE)
                muat_master_produk()
                return '''
                    <script>
                        alert("Master produk BHI berhasil diperbarui via Web!");
                        window.location.href = "/superuser";
                    </script>
                '''
            except Exception as e:
                return f"Gagal memproses file upload: {e}", 500
        else:
            return "Gagal: File harus berformat .xlsx atau .xls", 400
            
    return render_template('admin.html', data_so=HASIL_OPNAME)

# --- 6. API: EDIT MANUAL QTY (SUPER USER) ---
@app.route('/api/edit_qty', methods=['POST'])
def edit_qty():
    data = request.json
    record_id = int(data.get('id'))
    qty_baru = int(data.get('qty_baru'))
    
    for item in HASIL_OPNAME:
        if item['id'] == record_id:
            item['qty'] = qty_baru
            return jsonify({"status": "success", "message": f"Data ID {record_id} berhasil diubah menjadi {qty_baru}!"})
            
    return jsonify({"status": "error", "message": "Data tidak ditemukan!"})

# --- 7. ROUTE: DOWNLOAD EXCEL HASIL SCAN (SUPER USER) ---
@app.route('/superuser/download')
def download_excel():
    if not HASIL_OPNAME:
        return '''
            <script>
                alert("Belum ada data scan yang masuk, tidak ada data untuk didownload.");
                window.location.href = "/superuser";
            </script>
        '''
    try:
        df = pd.DataFrame(HASIL_OPNAME)
        kolom_rapi = ["id", "username", "zona", "sublokasi", "barcode", "sku", "brand", "product_name", "variant", "qty"]
        df = df[kolom_rapi]
        df.columns = ["ID", "Petugas", "Zona", "Sublokasi", "Barcode", "SKU", "Brand", "Nama Produk", "Varian", "Qty Fisik"]
        
        nama_file = "Hasil_Opname_BHI.xlsx"
        df.to_excel(nama_file, index=False)
        return send_file(nama_file, as_attachment=True)
    except Exception as e:
        return f"Gagal mengekspor data ke Excel: {e}", 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)