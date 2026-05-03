from flask import Flask, render_template, request, redirect, url_for, flash, jsonify, send_file
import sqlite3
import pandas as pd
import io
from datetime import datetime
import csv
import os

app = Flask(__name__)
app.secret_key = "siber_master_key_2026"
DB_FILE = 'veritabani.db'

def get_db_connection():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    return conn

def veritabani_kontrol():
    conn = get_db_connection()
    conn.execute('''
        CREATE TABLE IF NOT EXISTS kayitlar (
            ID INTEGER PRIMARY KEY AUTOINCREMENT,
            Unvan TEXT,
            Ad TEXT,
            Soyad TEXT,
            Numara_TC TEXT
        )
    ''')
    conn.commit()
    conn.close()

@app.route('/')
def index():
    conn = get_db_connection()
    veriler = conn.execute('SELECT * FROM kayitlar ORDER BY Ad ASC').fetchall()
    
    stats = {
        'toplam': conn.execute('SELECT COUNT(*) FROM kayitlar').fetchone()[0],
        'ogrenci': conn.execute('SELECT COUNT(*) FROM kayitlar WHERE Unvan="Öğrenci"').fetchone()[0],
        'ogretmen': conn.execute('SELECT COUNT(*) FROM kayitlar WHERE Unvan="Öğretmen"').fetchone()[0],
        'veli': conn.execute('SELECT COUNT(*) FROM kayitlar WHERE Unvan="Veli"').fetchone()[0]
    }
    conn.close()
    return render_template('index.html', veriler=veriler, stats=stats)

@app.route('/ekle', methods=['GET', 'POST'])
def ekle():
    if request.method == 'POST':
        conn = get_db_connection()
        conn.execute('INSERT INTO kayitlar (Unvan, Ad, Soyad, Numara_TC) VALUES (?, ?, ?, ?)',
                     (request.form.get('unvan'), request.form.get('ad'), 
                      request.form.get('soyad'), request.form.get('numara')))
        conn.commit()
        conn.close()
        flash("✅ Başarıyla eklendi!", "success")
        return redirect(url_for('index'))
    return render_template('ekle.html')

@app.route('/guncelle/<int:id>', methods=['GET', 'POST'])
def guncelle(id):
    conn = get_db_connection()
    kisi = conn.execute('SELECT * FROM kayitlar WHERE ID = ?', (id,)).fetchone()
    
    if request.method == 'POST':
        conn.execute('UPDATE kayitlar SET Unvan=?, Ad=?, Soyad=?, Numara_TC=? WHERE ID=?',
                     (request.form.get('unvan'), request.form.get('ad'), 
                      request.form.get('soyad'), request.form.get('numara'), id))
        conn.commit()
        conn.close()
        flash("✅ Güncellendi!", "info")
        return redirect(url_for('index'))
    
    conn.close()
    return render_template('guncelleme.html', kisi=kisi)

@app.route('/sil/<int:id>', methods=['GET', 'POST'])
def sil(id):
    conn = get_db_connection()
    kisi = conn.execute('SELECT * FROM kayitlar WHERE ID = ?', (id,)).fetchone()
    
    if request.method == 'POST':
        conn.execute('DELETE FROM kayitlar WHERE ID = ?', (id,))
        conn.commit()
        conn.close()
        flash("🗑️ Kayıt silindi!", "danger")
        return redirect(url_for('index'))
    
    conn.close()
    return render_template('sil.html', kisi=kisi)

@app.route('/coklu_sil', methods=['POST'])
def coklu_sil():
    ids_to_delete = request.form.getlist('secili_id')
    if ids_to_delete:
        conn = get_db_connection()
        placeholders = ','.join('?' * len(ids_to_delete))
        conn.execute(f'DELETE FROM kayitlar WHERE ID IN ({placeholders})', ids_to_delete)
        conn.commit()
        conn.close()
        flash(f"🗑️ {len(ids_to_delete)} kayıt silindi!", "danger")
    return redirect(url_for('index'))

@app.route('/temizle', methods=['POST'])
def temizle():
    conn = get_db_connection()
    conn.execute('DELETE FROM kayitlar')
    conn.commit()
    conn.close()
    flash("💀 Tüm liste sıfırlandı!", "danger")
    return redirect(url_for('index'))

@app.route('/excel_yukle', methods=['POST'])
def excel_yukle():
    dosya = request.files.get('file')
    if not dosya or not dosya.filename.endswith(('.xlsx', '.xls')):
        return jsonify({'ok': False, 'mesaj': 'Geçerli bir Excel dosyası seçin!'})
    
    try:
        df = pd.read_excel(dosya, dtype=str).fillna("")
        df.columns = [c.strip() for c in df.columns]
        
        zorunlu = {'Unvan', 'Ad', 'Soyad', 'Numara_TC'}
        eksik = zorunlu - set(df.columns)
        if eksik:
            return jsonify({'ok': False, 'mesaj': f"Eksik sütunlar: {', '.join(eksik)}"})
        
        conn = get_db_connection()
        eklenen = 0
        for _, row in df.iterrows():
            if row.get('Unvan') in ['Öğrenci', 'Öğretmen', 'Veli'] and row.get('Ad') and row.get('Soyad') and row.get('Numara_TC'):
                conn.execute('INSERT INTO kayitlar (Unvan, Ad, Soyad, Numara_TC) VALUES (?, ?, ?, ?)',
                            (row['Unvan'], row['Ad'], row['Soyad'], row['Numara_TC']))
                eklenen += 1
        conn.commit()
        conn.close()
        
        return jsonify({'ok': True, 'eklenen': eklenen, 'mesaj': f"{eklenen} kayıt aktarıldı!"})
    except Exception as e:
        return jsonify({'ok': False, 'mesaj': f"Hata: {str(e)}"})

@app.route('/rapor_excel')
def rapor_excel():
    conn = get_db_connection()
    veriler = conn.execute('SELECT ID, Unvan, Ad, Soyad, Numara_TC FROM kayitlar').fetchall()
    conn.close()
    
    df = pd.DataFrame([dict(row) for row in veriler])
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='SiberPanelRaporu')
    
    output.seek(0)
    tarih = datetime.now().strftime('%Y%m%d_%H%M%S')
    return send_file(output, as_attachment=True, download_name=f'siber_rapor_{tarih}.xlsx', 
                    mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

@app.route('/rapor_csv')
def rapor_csv():
    conn = get_db_connection()
    veriler = conn.execute('SELECT ID, Unvan, Ad, Soyad, Numara_TC FROM kayitlar').fetchall()
    conn.close()
    
    output = io.BytesIO()
    output.write('\uFEFF'.encode('utf-8'))
    
    writer = csv.writer(output, delimiter=';')
    writer.writerow(['ID', 'Ünvan', 'Ad', 'Soyad', 'Numara/TC'])
    
    for row in veriler:
        writer.writerow([row['ID'], row['Unvan'], row['Ad'], row['Soyad'], row['Numara_TC']])
    
    output.seek(0)
    tarih = datetime.now().strftime('%Y%m%d_%H%M%S')
    return send_file(output, as_attachment=True, download_name=f'siber_rapor_{tarih}.csv', 
                    mimetype='text/csv')

@app.route('/rapor_pdf')
def rapor_pdf():
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    
    conn = get_db_connection()
    veriler = conn.execute('SELECT ID, Unvan, Ad, Soyad, Numara_TC FROM kayitlar').fetchall()
    conn.close()
    
    output = io.BytesIO()
    doc = SimpleDocTemplate(output, pagesize=A4, rightMargin=2*cm, leftMargin=2*cm, 
                           topMargin=2*cm, bottomMargin=2*cm)
    styles = getSampleStyleSheet()
    
    story = []
    
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], 
                                 fontSize=16, textColor=colors.HexColor('#7d5fff'), 
                                 alignment=1, spaceAfter=20)
    story.append(Paragraph("SİBER PANEL PRO - KAYIT RAPORU", title_style))
    story.append(Spacer(1, 0.5*cm))
    story.append(Paragraph(f"Oluşturulma: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}", styles['Normal']))
    story.append(Paragraph(f"Toplam Kayıt: {len(veriler)}", styles['Normal']))
    story.append(Spacer(1, 1*cm))
    
    data = [['ID', 'Ünvan', 'Ad', 'Soyad', 'Numara/TC']]
    for row in veriler:
        data.append([str(row['ID']), row['Unvan'], row['Ad'], row['Soyad'], row['Numara_TC']])
    
    table = Table(data)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#7d5fff')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
    ]))
    
    story.append(table)
    doc.build(story)
    output.seek(0)
    
    tarih = datetime.now().strftime('%Y%m%d_%H%M%S')
    return send_file(output, as_attachment=True, download_name=f'siber_rapor_{tarih}.pdf', 
                    mimetype='application/pdf')

if __name__ == '__main__':
    veritabani_kontrol()
    app.run(debug=True, port=5001)