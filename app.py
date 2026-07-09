from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
import os

app = Flask(__name__)

# CONFIGURATION: Lokasi file master
MASTER_FILE = 'master_produk.xlsx'
OUTPUT_FILE = '/tmp/hasil_opname.xlsx'  # Menggunakan folder /tmp agar writeable di serverless

# Counter ID & List penampung data live scan
scan_id_counter = 1
scan_results = []

def load_master_data():
    """Fungsi pembantu loading data excel secara aman"""
    if os.path.exists(MASTER_FILE):
        try:
            df = pd.read_excel(MASTER_FILE)
            df.columns = df.columns.str.strip().str.lower()
            if 'barcode' in df.columns:
                df['barcode'] = df['barcode'].astype(str).str.strip()
            return df
        except Exception:
            return pd.DataFrame(columns=['barcode', 'sku', 'brand', 'product_name', 'variant'])
    return pd.DataFrame(columns=['barcode', 'sku', 'brand', 'product_name', 'variant'])

# Inisialisasi awal database master produk
df_master = load_master_data()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/scan')
def scan_page():
    petugas = request.args.get('petugas', 'Anonymous')
    zona = request.args.get('zona', '-')
    return render_template('scan.html', petugas=petugas, zona=zona)

# Route ganda biar gak typo memicu 404/500
@app.route('/super-user')
@app.route('/superuser')
def super_user_page():
    # Menghindari crash jika template dicari serverless
    try:
        return render_template('super_user.html')
    except Exception as e:
        return f"Template super_user.html tidak ditemukan atau rusak: {str(e)}", 500

# ==================== API ENDPOINTS ====================

@app.route('/api/products', methods=['GET'])
def get_live_products():
    return jsonify(scan_results)

@app.route('/api/scan', methods=['POST'])
def process_scan():
    global scan_id_counter, df_master
    
    # Reload berkala jika berjalan di environment serverless stateless
    if df_master.empty:
        df_master = load_master_data()

    data = request.json or {}
    barcode_input = str(data.get('barcode', '')).strip()
    petugas = data.get('petugas', 'Anonymous')
    zona = data.get('zona', '-')
    sublokasi = data.get('sublokasi', '-').strip().upper()
    
    try:
        qty_input = int(data.get('qty', 1))
    except (ValueError, TypeError):
        qty_input = 1

    if not sublokasi or sublokasi == '-':
        return jsonify({'status': 'error', 'message': '⚠️ SILAKAN SCAN QR CODE RAK TERLEBIH DAHULU!'}), 400

    if df_master.empty:
        return jsonify({'status': 'error', 'message': '❌ File master_produk.xlsx kosong atau tidak terbaca di server!'}), 500

    product = df_master[df_master['barcode'] == barcode_input]

    if not product.empty:
        prod_data = product.iloc[0]
        
        # JALUR CHECKING (Selesai Scan Barcode, sebelum isi Qty)
        if data.get('check_only') is True:
            return jsonify({
                'status': 'success',
                'nama_produk': str(prod_data.get('product_name', 'Unknown')),
                'variant': str(prod_data.get('variant', '-')),
                'message': 'Produk terdaftar'
            }), 200

        # JALUR SIMPAN DATA FIX
        existing_item = next((item for item in scan_results if item['barcode'] == barcode_input 
                              and item['petugas'] == petugas 
                              and item['zona'] == zona 
                              and item['sublokasi'] == sublokasi), None)
        
        if existing_item:
            existing_item['qty'] += qty_input
            current_qty = existing_item['qty']
        else:
            new_scan = {
                'id': scan_id_counter,
                'petugas': petugas,
                'zona': zona,
                'sublokasi': sublokasi,
                'barcode': barcode_input,
                'sku': str(prod_data.get('sku', '-')),
                'brand': str(prod_data.get('brand', '-')),
                'product_name': str(prod_data.get('product_name', 'Unknown')),
                'variant': str(prod_data.get('variant', '-')),
                'qty': qty_input
            }
            scan_results.append(new_scan)
            scan_id_counter += 1
            current_qty = qty_input

        return jsonify({
            'status': 'success',
            'nama_produk': str(prod_data.get('product_name', 'Unknown')),
            'variant': str(prod_data.get('variant', '-')),
            'qty': current_qty
        }), 200
    else:
        return jsonify({'status': 'error', 'message': f'Barcode [{barcode_input}] TIDAK TERDAFTAR di Master Excel!'}), 400

@app.route('/api/update-qty', methods=['POST'])
def update_qty_manual():
    data = request.json or {}
    try:
        row_id = int(data.get('id'))
        new_qty = int(data.get('qty', 0))
    except (ValueError, TypeError):
        return jsonify({'status': 'error', 'message': 'Data ID atau Qty tidak valid!'}), 400
    
    for item in scan_results:
        if item['id'] == row_id:
            item['qty'] = new_qty
            return jsonify({'status': 'success', 'message': 'Kuantitas berhasil diperbarui!'})
    return jsonify({'status': 'error', 'message': 'Data tidak ditemukan!'}), 404

@app.route('/api/upload-master', methods=['POST'])
def upload_master():
    global df_master
    if 'file' not in request.files:
        return jsonify({'status': 'error', 'message': 'Tidak ada file diunggah!'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'status': 'error', 'message': 'Nama file kosong!'}), 400
    
    try:
        file.save(MASTER_FILE)
        df_master = load_master_data()
        return jsonify({'status': 'success', 'message': 'Master Excel Sukses Diperbarui!'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': f'Gagal membaca file: {str(e)}'}), 500

@app.route('/api/download', methods=['GET'])
def download_excel():
    if not scan_results:
        df_export = pd.DataFrame(columns=['ID', 'Petugas', 'Zona', 'Sublokasi/Rak', 'Barcode', 'SKU', 'Brand', 'Nama Produk', 'Varian', 'Qty Fisik'])
    else:
        df_export = pd.DataFrame(scan_results)
        df_export.columns = ['ID', 'Petugas', 'Zona', 'Sublokasi/Rak', 'Barcode', 'SKU', 'Brand', 'Nama Produk', 'Varian', 'Qty Fisik']
    
    df_export.to_excel(OUTPUT_FILE, index=False)
    return send_file(OUTPUT_FILE, as_attachment=True)

# Handler khusus Vercel Serverless WSGI
def handler(environ, start_response):
    return app(environ, start_response)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
