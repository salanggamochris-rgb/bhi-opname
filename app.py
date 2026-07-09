from flask import Flask, render_template, request, redirect, url_for, flash, send_file
import pandas as pd
import os

app = Flask(__name__)
app.secret_key = "bhi_opname_secret_key"

# Load Master Produk dari Excel
EXCEL_FILE = "master_produk.xlsx"
if os.path.exists(EXCEL_FILE):
    df_master = pd.read_excel(EXCEL_FILE)
    df_master['barcode'] = df_master['barcode'].astype(str).str.strip()
else:
    df_master = pd.DataFrame(columns=['barcode', 'sku', 'brand', 'product_name', 'variant'])

# Data Penyimpanan Utama (Di Memori Server)
scan_results = []
scan_id_counter = 1
# Default Zona Awal (Bisa ditambah/dihapus oleh Super User lewat Web)
master_zona = ["PICKFACE", "BULK", "RESERVE", "SHIPPING"]

@app.route('/')
def login():
    # Mengirim daftar master_zona ke halaman login user
    return render_template('login.html', master_zona=master_zona)

@app.route('/scan', methods=['GET', 'POST'])
def scan_page():
    if request.method == 'POST':
        petugas = request.form.get('petugas')
        zona = request.form.get('zona')
        sublokasi = request.form.get('sublokasi')
        return render_template('scan.html', petugas=petugas, zona=zona, sublokasi=sublokasi)
    return redirect(url_for('login'))

@app.route('/api/scan', methods=['POST'])
def process_scan():
    global scan_id_counter
    data = request.json
    barcode_input = str(data.get('barcode', '')).strip()
    petugas = data.get('petugas')
    zona = data.get('zona')
    sublokasi = data.get('sublokasi')

    product = df_master[df_master['barcode'] == barcode_input]

    if not product.empty:
        prod_data = product.iloc[0]
        existing_item = next((item for item in scan_results if item['barcode'] == barcode_input 
                              and item['petugas'] == petugas 
                              and item['zona'] == zona 
                              and item['sublokasi'] == sublokasi), None)
        
        if existing_item:
            existing_item['qty'] += 1
            current_qty = existing_item['qty']
        else:
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
                'qty': 1
            }
            scan_results.append(new_scan)
            scan_id_counter += 1
            current_qty = 1

        return {
            'status': 'success',
            'nama_produk': prod_data['product_name'],
            'variant': prod_data['variant'],
            'qty': current_qty
        }, 200
    else:
        return {'status': 'error', 'message': 'Barcode TIDAK TERDAFTAR di Master Produk!'}, 400

@app.route('/superuser')
def superuser_dashboard():
    selected_zona = request.args.get('filter_zona', '')
    
    if selected_zona:
        filtered_results = [item for item in scan_results if str(item.get('zona')) == selected_zona]
    else:
        filtered_results = scan_results
        
    return render_template('admin.html', 
                           results=filtered_results, 
                           master_zona=master_zona, 
                           selected_zona=selected_zona)

# Route untuk Menambah Zona Baru dari Web
@app.route('/superuser/add_zona', methods=['POST'])
def add_zona():
    new_zona_name = request.form.get('new_zona', '').strip().upper()
    if new_zona_name and new_zona_name not in master_zona:
        master_zona.append(new_zona_name)
    return redirect(url_for('superuser_dashboard'))

# Route untuk Menghapus Zona dari Web
@app.route('/superuser/delete_zona/<string:zona_name>', methods=['POST'])
def delete_zona(zona_name):
    if zona_name in master_zona:
        master_zona.remove(zona_name)
    return redirect(url_for('superuser_dashboard'))

@app.route('/superuser/edit/<int:scan_id>', methods=['POST'])
def edit_qty(scan_id):
    new_qty = request.form.get('qty')
    try:
        new_qty = int(new_qty)
        for item in scan_results:
            if item['id'] == scan_id:
                item['qty'] = new_qty
                break
    except ValueError:
        pass
    return redirect(url_for('superuser_dashboard', filter_zona=request.args.get('filter_zona', '')))

@app.route('/superuser/reset', methods=['POST'])
def reset_data():
    global scan_results, scan_id_counter
    scan_results = []
    scan_id_counter = 1
    return redirect(url_for('superuser_dashboard'))

@app.route('/superuser/download')
def download_excel():
    if not scan_results:
        return "Belum ada data scan.", 400
    df_download = pd.DataFrame(scan_results)
    if 'id' in df_download.columns:
        df_download = df_download.drop(columns=['id'])
    cols = ['petugas', 'zona', 'sublokasi', 'barcode', 'sku', 'brand', 'product_name', 'variant', 'qty']
    df_download = df_download[cols]
    output_file = "Hasil_Opname_BHI.xlsx"
    df_download.to_excel(output_file, index=False)
    return send_file(output_file, as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
