from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
import os

app = Flask(__name__)

MASTER_FILE = 'master_produk.xlsx'
OUTPUT_FILE = '/tmp/hasil_opname.xlsx'

scan_id_counter = 1
scan_results = []
# Penampung zona kustom cadangan jika admin menambahkan zona baru secara manual
custom_zones = []
# Penampung zona dari excel yang sengaja dihapus oleh admin
deleted_excel_zones = []

def load_master_data():
    """Membaca data Excel secara aman"""
    if os.path.exists(MASTER_FILE):
        try:
            df = pd.read_excel(MASTER_FILE)
            df.columns = df.columns.str.strip().str.lower()
            if 'barcode' in df.columns:
                df['barcode'] = df['barcode'].astype(str).str.strip()
            return df
        except Exception:
            return pd.DataFrame()
    return pd.DataFrame()

def get_active_zones():
    """Mengambil list unik zona dari file excel ditambah zona kustom minus yang dihapus"""
    zones = set()
    df = load_master_data()
    if not df.empty:
        # Cari nama kolom yang mirip 'zona' atau 'zone'
        zona_col = next((col for col in df.columns if col in ['zona', 'zone']), None)
        if zona_col:
            extracted = df[zona_col].dropna().astype(str).str.strip().str.upper().unique()
            zones.update(extracted)
    
    # Masukkan data zona tambahan dari dashboard admin
    zones.update(custom_zones)
    
    # Buang zona yang telah masuk daftar hapus (deleted_excel_zones)
    for d_zone in deleted_excel_zones:
        if d_zone in zones:
            zones.remove(d_zone)
    
    return sorted(list(zones))

# Inisialisasi data awal
df_master = load_master_data()

@app.route('/')
def login_page():
    # Kirim daftar zona aktif dari database ke dropdown login.html
    list_zona = get_active_zones()
    return render_template('login.html', list_zona=list_zona)

@app.route('/scan')
def scan_page():
    petugas = request.args.get('petugas', 'Anonymous')
    zona = request.args.get('zona', '-')
    return render_template('scan.html', petugas=petugas, zona=zona)


# =========================================================================
# 🛡️ ROUTE SUPER USER & ADMIN DENGAN MANAGEMENT ZONA PANEL + REMOVE OPTION
# =========================================================================
@app.route('/super-user')
@app.route('/superuser')
@app.route('/admin')
def super_user_page():
    return '''
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Super User Dashboard - Live Stock Opname</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body { background-color: #f4f6f9; font-family: 'Segoe UI', sans-serif; }
        .navbar-custom { background: #1e293b; box-shadow: 0 4px 12px rgba(0,0,0,0.1); }
        .card-stat { border: none; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); }
        .table-container { background: white; border-radius: 12px; padding: 20px; box-shadow: 0 4px 12px rgba(0,0,0,0.05); }
        .qty-input-edit { width: 80px; text-align: center; }
        .btn-delete-zone { padding: 2px 6px; font-size: 11px; border-radius: 4px; }
    </style>
</head>
<body>
    <nav class="navbar navbar-expand-lg navbar-dark navbar-custom py-3">
        <div class="container">
            <a class="navbar-brand fw-bold" href="#">🛡️ BHI Opname - Super User Panel</a>
            <div class="d-flex gap-2">
                <a href="/api/download" class="btn btn-success btn-sm fw-bold px-3">📥 Download Excel</a>
                <a href="/" class="btn btn-outline-light btn-sm">Halaman Login</a>
            </div>
        </div>
    </nav>

    <div class="container my-4">
        <div class="row g-3 mb-4">
            <div class="col-md-4">
                <div class="card card-stat bg-white p-3 d-flex flex-row align-items-center justify-content-between">
                    <div><h6 class="text-muted small text-uppercase mb-1">Total Unik Barcode</h6><h3 id="statTotalItems" class="fw-bold mb-0">0</h3></div>
                    <div class="fs-1 text-primary">📦</div>
                </div>
            </div>
            <div class="col-md-4">
                <div class="card card-stat bg-white p-3 d-flex flex-row align-items-center justify-content-between">
                    <div><h6 class="text-muted small text-uppercase mb-1">Total Qty Ter-scan</h6><h3 id="statTotalQty" class="fw-bold text-success mb-0">0</h3></div>
                    <div class="fs-1 text-success">🔢</div>
                </div>
            </div>
            <div class="col-md-4">
                <div class="card card-stat bg-white p-3 d-flex flex-row align-items-center justify-content-between">
                    <div><h6 class="text-muted small text-uppercase mb-1">Status Sinkronisasi</h6><h3 class="fw-bold text-warning mb-0 fs-5">⚡ LIVE AUTOMATIC</h3></div>
                </div>
            </div>
        </div>

        <div class="row g-4">
            <div class="col-lg-4">
                <div class="table-container mb-4">
                    <h5 class="fw-bold text-dark mb-3">⚙️ Update Master Produk</h5>
                    <form id="uploadMasterForm">
                        <div class="mb-3"><input class="form-control form-control-sm" type="file" id="masterFile" accept=".xlsx" required></div>
                        <button type="submit" class="btn btn-primary btn-sm w-100 fw-bold">📤 Upload Master</button>
                    </form>
                    <div id="uploadStatus" class="mt-2 small fw-bold text-center"></div>
                </div>

                <div class="table-container">
                    <h5 class="fw-bold text-dark mb-2">📍 Tambah / Kelola Zona Kerja</h5>
                    <p class="text-muted" style="font-size: 11px;">Gunakan form ini untuk menambah atau menghapus Zona operasional gess:</p>
                    <div class="input-group mb-3">
                        <input type="text" id="newZoneName" class="form-control form-control-sm" placeholder="Contoh: ZONA COLDROOM" style="text-transform: uppercase;">
                        <button onclick="addNewZone()" class="btn btn-dark btn-sm fw-bold">➕ Tambah</button>
                    </div>
                    <div class="fw-bold small text-secondary mb-1">Daftar Zona Kerja Aktif:</div>
                    <ul class="list-group list-group-flush border rounded" id="zoneListContainer" style="max-height: 200px; overflow-y: auto; font-size: 13px;">
                        <li class="list-group-item text-muted text-center py-2">Loading data zona...</li>
                    </ul>
                </div>
            </div>

            <div class="col-lg-8">
                <div class="table-container">
                    <div class="d-flex justify-content-between align-items-center mb-3">
                        <h5 class="fw-bold text-dark mb-0">📋 Data Hasil Scan Real-Time</h5>
                        <button onclick="fetchLiveRecords()" class="btn btn-sm btn-outline-secondary">🔄 Refresh</button>
                    </div>
                    <div class="table-responsive" style="max-height: 500px;">
                        <table class="table table-hover align-middle">
                            <thead>
                                <tr>
                                    <th>📍 Rak</th>
                                    <th>👤 Petugas</th>
                                    <th>📦 Detail Produk</th>
                                    <th class="text-center" style="width: 120px;">🔢 Qty Fisik</th>
                                    <th class="text-center">Aksi</th>
                                </tr>
                            </thead>
                            <tbody id="liveTableBody">
                                <tr><td colspan="5" class="text-center text-muted py-5">Menunggu data scan petugas...</td></tr>
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <script>
        setInterval(function() {
            fetchLiveRecords();
            renderActiveZones();
        }, 3000);

        function fetchLiveRecords() {
            fetch('/api/products')
            .then(res => res.json())
            .then(data => {
                const tbody = document.getElementById('liveTableBody');
                if (!data || data.length === 0) {
                    tbody.innerHTML = '<tr><td colspan="5" class="text-center text-muted py-5">Belum ada aktivitas scan dari petugas gudang.</td></tr>';
                    document.getElementById('statTotalItems').innerText = '0';
                    document.getElementById('statTotalQty').innerText = '0';
                    return;
                }
                let htmlContent = ''; let grandTotalQty = 0;
                for (let i = 0; i < data.length; i++) {
                    let item = data[i];
                    grandTotalQty += item.qty;
                    htmlContent += '<tr>' +
                        '<td><span class="badge bg-dark fw-bold fs-6">' + item.sublokasi + '</span><br><small class="text-muted">Zona: ' + item.zona + '</small></td>' +
                        '<td><strong class="text-secondary">' + item.petugas + '</strong></td>' +
                        '<td>' +
                            '<div class="fw-bold text-dark small mb-0">' + item.product_name + '</div>' +
                            '<div class="text-muted" style="font-size: 11px;">Barcode: <code>' + item.barcode + '</code> | SKU: ' + item.sku + ' | Brand: ' + item.brand + ' | Var: ' + item.variant + '</div>' +
                        '</td>' +
                        '<td class="text-center">' +
                            '<input type="number" id="qty_input_' + item.id + '" class="form-control form-control-sm qty-input-edit d-inline-block" value="' + item.qty + '">' +
                        '</td>' +
                        '<td class="text-center">' +
                            '<button onclick="saveQtyCorrection(' + item.id + ')" class="btn btn-sm btn-success py-1 px-2 fw-bold">Simpan</button>' +
                        '</td>' +
                    '</tr>';
                }
                tbody.innerHTML = htmlContent;
                document.getElementById('statTotalItems').innerText = data.length;
                document.getElementById('statTotalQty').innerText = grandTotalQty;
            }).catch(err => console.error(err));
        }

        function renderActiveZones() {
            fetch('/api/zones')
            .then(res => res.json())
            .then(zones => {
                const container = document.getElementById('zoneListContainer');
                let html = '';
                zones.forEach(z => {
                    html += `<li class="list-group-item py-2 d-flex justify-content-between align-items-center">` +
                            `<span>🔹 ${z}</span>` +
                            `<button onclick="removeZone('${z}')" class="btn btn-danger btn-sm btn-delete-zone fw-bold">❌ Hapus</button>` +
                            `</li>`;
                });
                container.innerHTML = html || `<li class="list-group-item text-muted text-center py-2">Tidak ada zona terdaftar</li>`;
            });
        }

        function addNewZone() {
            const el = document.getElementById('newZoneName');
            const name = el.value.trim().toUpperCase();
            if(!name) return;
            fetch('/api/zones', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ action: 'add', zona: name })
            }).then(res => res.json()).then(d => { el.value = ''; renderActiveZones(); });
        }

        function removeZone(zoneName) {
            if(confirm(`Apakah Anda yakin ingin menghapus "${zoneName}" dari sistem gess?`)) {
                fetch('/api/zones', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ action: 'remove', zona: zoneName })
                }).then(res => res.json()).then(d => { renderActiveZones(); });
            }
        }

        function saveQtyCorrection(rowId) {
            const inputField = document.getElementById('qty_input_' + rowId);
            const newQty = parseInt(inputField.value);
            if (isNaN(newQty) || newQty < 0) { alert("Qty harus angka valid!"); return; }
            fetch('/api/update-qty', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ id: rowId, qty: newQty })
            }).then(res => res.json()).then(d => { if(d.status === 'success') { alert("✅ Koreksi berhasil!"); fetchLiveRecords(); } });
        }

        document.getElementById('uploadMasterForm').addEventListener('submit', function(e) {
            e.preventDefault();
            const formData = new FormData();
            formData.append('file', document.getElementById('masterFile').files[0]);
            document.getElementById('uploadStatus').innerHTML = "Memproses...";
            fetch('/api/upload-master', { method: 'POST', body: formData })
            .then(res => res.json())
            .then(d => { document.getElementById('uploadStatus').innerHTML = d.status === 'success' ? "<span class='text-success'>✅ Sukses!</span>" : "<span class='text-danger'>❌ Gagal</span>"; });
        });

        fetchLiveRecords();
        renderActiveZones();
    </script>
</body>
</html>
    '''

# ==================== CONTROLLER BACKEND API ZONA ====================

@app.route('/api/zones', methods=['GET', 'POST'])
def handle_zones_api():
    global custom_zones, deleted_excel_zones
    if request.method == 'POST':
        data = request.json or {}
        action = data.get('action', 'add')
        target_z = str(data.get('zona', '')).strip().upper()
        
        if target_z:
            if action == 'add':
                if target_z in deleted_excel_zones:
                    deleted_excel_zones.remove(target_z)
                if target_z not in custom_zones:
                    custom_zones.append(target_z)
            elif action == 'remove':
                if target_z not in deleted_excel_zones:
                    deleted_excel_zones.append(target_z)
                if target_z in custom_zones:
                    custom_zones.remove(target_z)
                    
        return jsonify({'status': 'success'})
    return jsonify(get_active_zones())

@app.route('/api/products', methods=['GET'])
def get_live_products():
    return jsonify(scan_results)

@app.route('/api/scan', methods=['POST'])
def process_scan():
    global scan_id_counter, df_master
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
        return jsonify({'status': 'error', 'message': '❌ File master_produk.xlsx tidak terbaca atau kosong!'}), 500

    product = df_master[df_master['barcode'] == barcode_input]

    if not product.empty:
        prod_data = product.iloc[0]
        if data.get('check_only') is True:
            return jsonify({
                'status': 'success',
                'nama_produk': str(prod_data.get('product_name', 'Unknown')),
                'variant': str(prod_data.get('variant', '-')),
                'message': 'Produk terdaftar'
            }), 200

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
        return jsonify({'status': 'error', 'message': f'Barcode [{barcode_input}] TIDAK TERDAFTAR!'}), 400

@app.route('/api/update-qty', methods=['POST'])
def update_qty_manual():
    data = request.json or {}
    try:
        row_id = int(data.get('id'))
        new_qty = int(data.get('qty', 0))
    except (ValueError, TypeError):
        return jsonify({'status': 'error', 'message': 'Input tidak valid!'}), 400
    
    for item in scan_results:
        if item['id'] == row_id:
            item['qty'] = new_qty
            return jsonify({'status': 'success', 'message': 'Kuantitas diperbarui!'})
    return jsonify({'status': 'error', 'message': 'Data tidak ditemukan!'}), 404

@app.route('/api/upload-master', methods=['POST'])
def upload_master():
    global df_master
    if 'file' not in request.files:
        return jsonify({'status': 'error', 'message': 'Tidak ada file!'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'status': 'error', 'message': 'Nama file kosong!'}), 400
    try:
        file.save(MASTER_FILE)
        df_master = load_master_data()
        return jsonify({'status': 'success', 'message': 'Master Excel Diperbarui!'})
    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/api/download', methods=['GET'])
def download_excel():
    if not scan_results:
        df_export = pd.DataFrame(columns=['ID', 'Petugas', 'Zona', 'Sublokasi/Rak', 'Barcode', 'SKU', 'Brand', 'Nama Produk', 'Varian', 'Qty Fisik'])
    else:
        df_export = pd.DataFrame(scan_results)
        df_export.columns = ['ID', 'Petugas', 'Zona', 'Sublokasi/Rak', 'Barcode', 'SKU', 'Brand', 'Nama Produk', 'Varian', 'Qty Fisik']
    df_export.to_excel(OUTPUT_FILE, index=False)
    return send_file(OUTPUT_FILE, as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
