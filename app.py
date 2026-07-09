@app.route('/api/scan', methods=['POST'])
def process_scan():
    global scan_id_counter
    data = request.json
    barcode_input = str(data.get('barcode', '')).strip()
    petugas = data.get('petugas')
    zona = data.get('zona')
    sublokasi = data.get('sublokasi', '-').strip().upper()
    
    # Ambil qty inputan manual dari user (jika tidak ada/kosong, default ke 1)
    try:
        qty_input = int(data.get('qty', 1))
    except (ValueError, TypeError):
        qty_input = 1

    if not sublokasi or sublokasi == '-':
        return {'status': 'error', 'message': '⚠️ SILAKAN SCAN QR CODE RAK TERLEBIH DAHULU!'}, 400

    product = df_master[df_master['barcode'] == barcode_input]

    if not product.empty:
        prod_data = product.iloc[0]
        
        # Cari apakah petugas yang sama sudah pernah scan barang ini di rak yang sama
        existing_item = next((item for item in scan_results if item['barcode'] == barcode_input 
                              and item['petugas'] == petugas 
                              and item['zona'] == zona 
                              and item['sublokasi'] == sublokasi), None)
        
        if existing_item:
            # Jika sudah ada, tambahkan qty lama dengan qty inputan baru
            existing_item['qty'] += qty_input
            current_qty = existing_item['qty']
        else:
            # Jika barang baru di rak tersebut, simpan sesuai qty inputan baru
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

        return {
            'status': 'success',
            'nama_produk': prod_data['product_name'],
            'variant': prod_data['variant'],
            'qty': current_qty
        }, 200
    else:
        return {'status': 'error', 'message': f'Barcode [{barcode_input}] TIDAK TERDAFTAR di Master Excel!'}, 400
