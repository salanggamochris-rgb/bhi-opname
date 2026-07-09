from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
import os

app = Flask(__name__)

# CONFIGURATION: Nama file master dan file simpan database lokal
MASTER_FILE = 'master_produk.xlsx'
OUTPUT_FILE = 'hasil_opname.xlsx'

# Counter untuk ID unik baris data hasil scan
scan_id_counter = 1

# Dummy data list untuk menampung hasil scan sementara di memory RAM
scan_results = []

# Load database master produk saat aplikasi pertama kali dijalankan
if os.path.exists(MASTER_FILE):
    df_master = pd.read_excel(MASTER_FILE)
    # Bersihkan nama kolom agar huruf kecil semua dan tidak ada spasi liar
    df_master.columns = df_master.columns.str.strip().str.lower()
    # Pastikan tipe data barcode di-force ke string biar match saat dicocokkan
    if 'barcode' in df_master.columns:
        df_master['barcode'] = df_master['barcode'].astype(str).str.strip()
else:
    # Buat dataframe template kosong jika file excel master belum ada
    df_master = pd.DataFrame(columns=['barcode', 'sku', 'brand', 'product_name', 'variant'])

@app.route('/')
def index():
    # Halaman login awal petugas memilih nama dan zona kerja
    return render_template('index.html')

@app.route('/scan')
def scan_page():
    # Mengambil parameter identitas petugas dari URL redirect halaman login
    petugas = request.args.get('petugas', 'Anonymous')
    zona = request.args.get('zona', '-')
    return render_template('scan.html', petugas=petugas, zona=zona)


# JALUR PADAT AMAN: Kita daftarkan kedua alamat alternatif ini agar tidak memicu 404 lagi!
@app.route('/super-user')
@app.route('/superuser')
def super_user_page():
    # Halaman monitoring live dashboard monitoring
    return render_template('super_user.html')


# ==================== API ENDPOINTS ====================

@app.route('/api/products', methods=['GET'])
def get_live_products():
    # Mengirimkan data tumpukan hasil scan real-time ke tabel dashboard super-user
    return jsonify(scan_results)

@app.route('/api/scan', methods=['POST'])
def process_scan():
    global scan_id_counter
    data = request.json
    barcode_input = str(data.get('barcode', '')).strip()
    petugas = data.get('petugas')
    zona = data.get('zona')
    sublokasi = data.get('sublokasi', '-').strip().upper()
    
    # Ambil nilai kuantitas inputan manual dari petugas (jika kosong, default ke 1)
    try:
        qty_input = int(data.get('qty', 1))
    except (ValueError, TypeError):
        qty_input = 1

    if not sublokasi or sublokasi == '-':
        return jsonify({'status': 'error', 'message': '⚠️ SILAKAN SCAN QR CODE RAK TERLEBIH DAHULU!'}), 400

    # Pencarian ke basis data master Excel
    product = df_master[df_master['barcode'] == barcode_input]

    if not product.empty:
        prod_data = product.iloc[0]
        
        # JALUR A: Jika request hanya untuk memverifikasi detail info produk saat kursor bergeser
        if data.get('check_only') is True:
            return jsonify({
                'status': 'success',
                'nama_produk': str(prod_data['product_name']),
                'variant': str(prod_data['variant']),
                'message': 'Produk terdaftar'
            }), 200

        # JALUR B: Proses simpan data final setelah input Qty manual dikirim
        existing_item = next((item for item in scan_results if item['barcode'] == barcode_input 
                              and item['petugas'] == petugas 
                              and item['zona'] == zona 
                              and item['sublokasi'] == sublokasi), None)
        
        if existing_item:
            # Akumulasikan kuantitas lama di lokasi rak yang sama dengan input baru
            existing_item['qty'] += qty_input
            current_qty = existing_item['qty']
        else:
            # Rekam record baris data baru ke database memori
            new_scan = {
                'id': scan_id_counter,
                'petugas': petugas,
                'zona': zona,
                'sublokasi': sublokasi,
                'barcode': barcode_input,
                'sku': str(prod_data['sku']),
                'brand': str(prod_data['brand']),
                'product_name': str(prod_data['product_name']),
                'variant': str(prod_data['variant']),
                'qty': qty_input
            }
            scan_results.append(new_scan)
            scan_id_counter += 1
            current_qty = qty_input

        return jsonify({
            'status': 'success',
            'nama_produk': str(prod_data['product_name']),
            'variant': str(prod_data['variant']),
            'qty': current_qty
        }), 200
    else:
        return jsonify({'status': 'error', 'message': f'Barcode [{barcode_input}] TIDAK TERDAFTAR di Master Excel!'}), 400

@app.route('/api/update-qty', methods=['POST'])
def update_qty_manual():
    # Endpoint koreksi instan nilai Qty dari tabel Live Monitoring Super User
    data = request.json
    row_id = int(data.get('id'))
    new_qty = int(data.get('qty', 0))
    
    for item in scan_results:
        if item['id'] == row_id:
            item['qty'] = new_qty
            return jsonify({'status': 'success', 'message': 'Kuantitas berhasil diperbarui!'})
    return jsonify({'status': 'error', 'message': 'Data tidak ditemukan!'}), 404

@app.route('/api/upload-master', methods=['POST'])
def upload_master():
    # Endpoint upload update Excel Master Produk terbaru dari Dashboard
    global df_master
    if 'file' not in request.files:
        return jsonify({'status': 'error', 'message': 'Tidak ada file diunggah!'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'status': 'error', 'message': 'Nama file kosong!'}), 400
    
    try:
        file.save(MASTER_FILE)
        df_master = pd.read_excel(MASTER_FILE)
        df_master.columns = df_master.columns.str.strip().str.lower()
        if 'barcode' in df_master.columns:
            df_master['barcode'] = df_master['barcode'].astype(str).str.strip()
        return jsonify({'status': 'success', 'message': 'Master Excel Sukses Diperbarui!'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': f'Gagal membaca file: {str(e)}'}), 500

@app.route('/api/download', methods=['GET'])
def download_excel():
    # Compile list RAM menjadi file Excel fisik saat tombol download diklik
    if not scan_results:
        # Jika kosong, buatkan dataframe dummy struktur kolomnya saja biar tidak error crash
        df_export = pd.DataFrame(columns=['ID', 'Petugas', 'Zona', 'Sublokasi/Rak', 'Barcode', 'SKU', 'Brand', 'Nama Produk', 'Varian', 'Qty Fisik'])
    else:
        df_export = pd.DataFrame(scan_results)
        df_export.columns = ['ID', 'Petugas', 'Zona', 'Sublokasi/Rak', 'Barcode', 'SKU', 'Brand', 'Nama Produk', 'Varian', 'Qty Fisik']
    
    df_export.to_excel(OUTPUT_FILE, index=False)
    return send_file(OUTPUT_FILE, as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
