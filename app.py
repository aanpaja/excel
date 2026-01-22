from flask import Flask, render_template, jsonify, request
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import re
import os
import urllib.request
import urllib.parse
from io import StringIO
import csv

app = Flask(__name__)

# Default spreadsheet URL - bisa diganti sesuai kebutuhan
DEFAULT_SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1HiQX-jn2pNfMO2-ZkBwaMbcolyKFFNb0/edit"

# Daftar bulan dalam bahasa Indonesia
MONTHS = ['JANUARI', 'FEBRUARI', 'MARET', 'APRIL', 'MEI', 'JUNI',
          'JULI', 'AGUSTUS', 'SEPTEMBER', 'OKTOBER', 'NOVEMBER', 'DESEMBER']

# Mapping nama kolom yang mungkin digunakan
COLUMN_MAPPINGS = {
    'durasi_respon': [
        'Durasi Respon', 'DURASI RESPON', 'durasi respon',
        'Durasi_Respon', 'Response Time', 'response_time',
        'Waktu Respon', 'WAKTU RESPON'
    ],
    'durasi_penanganan': [
        'Durasi Penanganan Gangguan', 'DURASI PENANGANAN GANGGUAN',
        'durasi penanganan gangguan', 'Durasi_Penanganan',
        'Durasi Penanganan', 'DURASI PENANGANAN',
        'Handling Time', 'handling_time', 'Waktu Penanganan'
    ],
    'lokasi': [
        'Lokasi Pelanggan', 'LOKASI PELANGGAN', 'lokasi pelanggan',
        'Lokasi', 'LOKASI', 'Location', 'location'
    ]
}


def parse_duration_to_minutes(duration_str):
    """Parse berbagai format durasi ke menit"""
    if pd.isna(duration_str):
        return 0

    if isinstance(duration_str, (int, float)):
        return float(duration_str)

    if isinstance(duration_str, timedelta):
        return duration_str.total_seconds() / 60

    duration_str = str(duration_str).strip()

    # Skip jika error atau kosong
    if not duration_str or duration_str in ['#DIV/0!', '#N/A', '#VALUE!', 'NaT', 'nan']:
        return 0

    try:
        # Format: "X days, HH:MM:SS" atau "HH:MM:SS"
        parts = duration_str.split(',')
        total_minutes = 0

        for part in parts:
            part = part.strip()

            if 'day' in part.lower():
                days = int(re.search(r'(\d+)', part).group(1))
                total_minutes += days * 24 * 60
            elif ':' in part:
                time_parts = part.split(':')
                if len(time_parts) == 3:
                    hours = int(time_parts[0])
                    minutes = int(time_parts[1])
                    seconds = float(time_parts[2])
                    total_minutes += hours * 60 + minutes + seconds / 60
                elif len(time_parts) == 2:
                    hours = int(time_parts[0])
                    minutes = int(time_parts[1])
                    total_minutes += hours * 60 + minutes

        return round(total_minutes, 2)
    except:
        return 0


def find_column(df, column_type):
    """Cari nama kolom yang sesuai dari mapping"""
    possible_names = COLUMN_MAPPINGS.get(column_type, [])
    for name in possible_names:
        if name in df.columns:
            return name
    return None


def extract_spreadsheet_id(url):
    """Extract spreadsheet ID dari URL"""
    patterns = [
        r'/spreadsheets/d/([a-zA-Z0-9-_]+)',
        r'key=([a-zA-Z0-9-_]+)',
        r'^([a-zA-Z0-9-_]+)$'
    ]

    for pattern in patterns:
        match = re.search(pattern, url)
        if match:
            return match.group(1)

    return url


def download_sheet_as_csv(sheet_id, sheet_name):
    """Download sheet tertentu sebagai CSV"""
    try:
        encoded_sheet = urllib.parse.quote(sheet_name)
        csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&sheet={encoded_sheet}"

        response = urllib.request.urlopen(csv_url, timeout=30)
        csv_data = response.read().decode('utf-8')
        return csv_data
    except Exception as e:
        print(f"Error downloading sheet {sheet_name}: {e}")
        return None


def parse_csv_to_dataframe(csv_data):
    """Parse CSV data ke DataFrame"""
    try:
        df = pd.read_csv(StringIO(csv_data))
        return df
    except:
        return None


def calculate_monthly_averages_from_raw(spreadsheet_url):
    """
    FUNGSI UTAMA: Menghitung average dari data mentah di setiap sheet bulanan
    """
    if 'docs.google.com/spreadsheets' not in spreadsheet_url:
        return None, None, []

    sheet_id = extract_spreadsheet_id(spreadsheet_url)
    print(f"Sheet ID: {sheet_id}")
    print("Menghitung average dari data mentah...")

    monthly_results = []
    all_location_data = {}
    total_records = 0

    for month in MONTHS:
        print(f"  Memproses sheet: {month}...")

        csv_data = download_sheet_as_csv(sheet_id, month)

        if csv_data is None:
            print(f"    - Sheet {month} tidak ditemukan atau tidak bisa diakses")
            monthly_results.append({
                'bulan': month,
                'avg_respon_minutes': 0,
                'avg_penanganan_minutes': 0,
                'total_tiket': 0,
                'status': 'not_found'
            })
            continue

        df = parse_csv_to_dataframe(csv_data)

        if df is None or df.empty:
            print(f"    - Sheet {month} kosong")
            monthly_results.append({
                'bulan': month,
                'avg_respon_minutes': 0,
                'avg_penanganan_minutes': 0,
                'total_tiket': 0,
                'status': 'empty'
            })
            continue

        # Cari kolom yang sesuai
        col_respon = find_column(df, 'durasi_respon')
        col_penanganan = find_column(df, 'durasi_penanganan')
        col_lokasi = find_column(df, 'lokasi')

        print(f"    - Kolom ditemukan: Respon={col_respon}, Penanganan={col_penanganan}, Lokasi={col_lokasi}")

        # Hitung average durasi respon
        avg_respon = 0
        if col_respon:
            respon_values = df[col_respon].dropna().apply(parse_duration_to_minutes)
            respon_values = respon_values[respon_values > 0]
            if len(respon_values) > 0:
                avg_respon = respon_values.mean()

        # Hitung average durasi penanganan
        avg_penanganan = 0
        if col_penanganan:
            penanganan_values = df[col_penanganan].dropna().apply(parse_duration_to_minutes)
            penanganan_values = penanganan_values[penanganan_values > 0]
            if len(penanganan_values) > 0:
                avg_penanganan = penanganan_values.mean()

        total_tiket = len(df)
        total_records += total_tiket

        print(f"    - Total tiket: {total_tiket}, Avg Respon: {avg_respon:.2f} min, Avg Penanganan: {avg_penanganan:.2f} min")

        monthly_results.append({
            'bulan': month,
            'avg_respon_minutes': round(avg_respon, 2),
            'avg_penanganan_minutes': round(avg_penanganan, 2),
            'total_tiket': total_tiket,
            'status': 'ok'
        })

        # Hitung per lokasi
        if col_lokasi and col_penanganan:
            for lokasi in df[col_lokasi].dropna().unique():
                if not lokasi or lokasi in ['GMEDIA', 'nan', '']:
                    continue

                lokasi_df = df[df[col_lokasi] == lokasi]
                penanganan_vals = lokasi_df[col_penanganan].dropna().apply(parse_duration_to_minutes)
                penanganan_vals = penanganan_vals[penanganan_vals > 0]

                if len(penanganan_vals) > 0:
                    if lokasi not in all_location_data:
                        all_location_data[lokasi] = {
                            'monthly_data': [],
                            'all_values': []
                        }

                    avg_lokasi = penanganan_vals.mean()
                    all_location_data[lokasi]['monthly_data'].append({
                        'bulan': month,
                        'value_minutes': round(avg_lokasi, 2),
                        'count': len(penanganan_vals)
                    })
                    all_location_data[lokasi]['all_values'].extend(penanganan_vals.tolist())

    # Proses location data
    location_results = []
    for lokasi, data in all_location_data.items():
        if data['all_values']:
            avg_total = sum(data['all_values']) / len(data['all_values'])
            category = get_location_category(lokasi)
            location_results.append({
                'location': lokasi,
                'category': category,
                'avg_minutes': round(avg_total, 2),
                'monthly_data': data['monthly_data'],
                'total_records': len(data['all_values'])
            })

    # Sort by category and name
    location_results = sorted(location_results, key=lambda x: (get_category_order(x['category']), x['location']))

    # Hitung quarterly (triwulan)
    quarterly_results = []
    quarters = [
        ('TRIWULAN 1', ['JANUARI', 'FEBRUARI', 'MARET']),
        ('TRIWULAN 2', ['APRIL', 'MEI', 'JUNI']),
        ('TRIWULAN 3', ['JULI', 'AGUSTUS', 'SEPTEMBER']),
        ('TRIWULAN 4', ['OKTOBER', 'NOVEMBER', 'DESEMBER'])
    ]

    for quarter_name, quarter_months in quarters:
        quarter_data = [m for m in monthly_results if m['bulan'] in quarter_months and m['status'] == 'ok']
        if quarter_data:
            avg_respon_q = sum(m['avg_respon_minutes'] for m in quarter_data) / len(quarter_data)
            avg_penanganan_q = sum(m['avg_penanganan_minutes'] for m in quarter_data) / len(quarter_data)
            quarterly_results.append({
                'period': quarter_name,
                'avg_respon_minutes': round(avg_respon_q, 2),
                'avg_penanganan_minutes': round(avg_penanganan_q, 2)
            })

    print(f"\nTotal records diproses: {total_records}")
    print(f"Total lokasi ditemukan: {len(location_results)}")

    return monthly_results, quarterly_results, location_results


def get_location_category(location_name):
    """Kategorikan lokasi berdasarkan nama"""
    location_lower = str(location_name).lower()

    if 'corporate' in location_lower:
        return 'Corporate'
    elif 'retail' in location_lower:
        return 'Retail'
    elif 'pop' in location_lower:
        return 'POP'
    elif 'pemerintahan' in location_lower:
        return 'Pemerintahan'
    elif 'disdik' in location_lower:
        return 'Disdik'
    elif 'tekkomdik' in location_lower:
        return 'TEKKOMDIK'
    else:
        return 'Lainnya'


def get_category_order(category):
    """Urutan kategori untuk sorting"""
    order = {
        'Corporate': 1,
        'Retail': 2,
        'POP': 3,
        'Pemerintahan': 4,
        'Disdik': 5,
        'TEKKOMDIK': 6,
        'Lainnya': 99
    }
    return order.get(category, 99)


def get_all_locations_from_spreadsheet(spreadsheet_url):
    """Ambil semua lokasi unik dari sheet JANUARI"""
    if 'docs.google.com/spreadsheets' not in spreadsheet_url:
        return []

    sheet_id = extract_spreadsheet_id(spreadsheet_url)
    csv_data = download_sheet_as_csv(sheet_id, 'JANUARI')

    if csv_data is None:
        return []

    df = parse_csv_to_dataframe(csv_data)
    if df is None:
        return []

    col_lokasi = find_column(df, 'lokasi')
    if not col_lokasi:
        return []

    locations = df[col_lokasi].dropna().unique().tolist()
    valid_locations = [loc for loc in locations if loc not in ['GMEDIA', 'nan', '', None]]

    return sorted(valid_locations)


def calculate_location_monthly_from_raw(spreadsheet_url, location_name):
    """Hitung data bulanan untuk lokasi tertentu dari data mentah"""
    if 'docs.google.com/spreadsheets' not in spreadsheet_url:
        return []

    sheet_id = extract_spreadsheet_id(spreadsheet_url)
    monthly_results = []

    for month in MONTHS:
        csv_data = download_sheet_as_csv(sheet_id, month)

        if csv_data is None:
            monthly_results.append({
                'bulan': month,
                'avg_minutes': 0,
                'count': 0
            })
            continue

        df = parse_csv_to_dataframe(csv_data)
        if df is None or df.empty:
            monthly_results.append({
                'bulan': month,
                'avg_minutes': 0,
                'count': 0
            })
            continue

        col_lokasi = find_column(df, 'lokasi')
        col_penanganan = find_column(df, 'durasi_penanganan')

        if not col_lokasi or not col_penanganan:
            monthly_results.append({
                'bulan': month,
                'avg_minutes': 0,
                'count': 0
            })
            continue

        # Filter by location
        location_df = df[df[col_lokasi] == location_name]

        if len(location_df) == 0:
            monthly_results.append({
                'bulan': month,
                'avg_minutes': 0,
                'count': 0
            })
            continue

        # Calculate average
        durations = location_df[col_penanganan].dropna().apply(parse_duration_to_minutes)
        durations = durations[durations > 0]

        if len(durations) > 0:
            avg_minutes = durations.mean()
            monthly_results.append({
                'bulan': month,
                'avg_minutes': round(avg_minutes, 2),
                'count': len(durations)
            })
        else:
            monthly_results.append({
                'bulan': month,
                'avg_minutes': 0,
                'count': 0
            })

    return monthly_results


# ==================== ROUTES ====================

@app.route('/')
def index():
    return render_template('dashboard.html', default_url=DEFAULT_SPREADSHEET_URL)


@app.route('/api/config')
def get_config():
    return jsonify({
        'default_spreadsheet_url': DEFAULT_SPREADSHEET_URL,
        'last_update': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    })


@app.route('/api/data', methods=['POST'])
def get_data():
    """
    API endpoint utama - menghitung average dari data mentah
    """
    try:
        data = request.json
        spreadsheet_url = data.get('spreadsheet_url', '')

        if not spreadsheet_url:
            spreadsheet_url = DEFAULT_SPREADSHEET_URL

        print(f"\n{'='*50}")
        print(f"Memproses spreadsheet: {spreadsheet_url}")
        print(f"{'='*50}")

        # Hitung dari data mentah
        monthly_data, quarterly_data, location_data = calculate_monthly_averages_from_raw(spreadsheet_url)

        if monthly_data is None:
            return jsonify({
                'success': False,
                'message': 'Gagal membaca data. Pastikan spreadsheet sudah di-share sebagai "Anyone with the link"'
            })

        # Filter hanya bulan dengan data valid
        valid_monthly = [m for m in monthly_data if m['status'] == 'ok' and (m['avg_respon_minutes'] > 0 or m['avg_penanganan_minutes'] > 0)]

        # Hitung yearly average
        yearly_avg_respon = 0
        yearly_avg_penanganan = 0

        if valid_monthly:
            respon_values = [m['avg_respon_minutes'] for m in valid_monthly if m['avg_respon_minutes'] > 0]
            penanganan_values = [m['avg_penanganan_minutes'] for m in valid_monthly if m['avg_penanganan_minutes'] > 0]

            if respon_values:
                yearly_avg_respon = sum(respon_values) / len(respon_values)
            if penanganan_values:
                yearly_avg_penanganan = sum(penanganan_values) / len(penanganan_values)

        return jsonify({
            'success': True,
            'calculated_from': 'raw_data',
            'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'data': {
                'monthly': [{
                    'bulan': m['bulan'],
                    'avg_respon_minutes': m['avg_respon_minutes'],
                    'avg_penanganan_minutes': m['avg_penanganan_minutes'],
                    'total_tiket': m.get('total_tiket', 0)
                } for m in monthly_data if m['status'] == 'ok'],
                'quarterly': quarterly_data,
                'yearly': {
                    'avg_respon_minutes': round(yearly_avg_respon, 2),
                    'avg_penanganan_minutes': round(yearly_avg_penanganan, 2)
                },
                'locations': location_data,
                'summary': {
                    'total_months_processed': len([m for m in monthly_data if m['status'] == 'ok']),
                    'total_locations': len(location_data),
                    'total_tiket': sum(m.get('total_tiket', 0) for m in monthly_data)
                }
            }
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'message': str(e)})


@app.route('/api/location-monthly', methods=['POST'])
def get_location_monthly():
    """API untuk mendapatkan data bulanan per lokasi"""
    try:
        data = request.json
        location = data.get('location', '')
        spreadsheet_url = data.get('spreadsheet_url', DEFAULT_SPREADSHEET_URL)

        if not location:
            return jsonify({'success': False, 'message': 'Lokasi tidak boleh kosong'})

        monthly_data = calculate_location_monthly_from_raw(spreadsheet_url, location)

        return jsonify({
            'success': True,
            'data': {
                'location': location,
                'monthly': monthly_data
            }
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'message': str(e)})


@app.route('/api/locations-list', methods=['POST'])
def get_locations_list():
    """API untuk mendapatkan daftar lokasi"""
    try:
        data = request.json
        spreadsheet_url = data.get('spreadsheet_url', DEFAULT_SPREADSHEET_URL)

        locations = get_all_locations_from_spreadsheet(spreadsheet_url)

        return jsonify({
            'success': True,
            'data': {
                'locations': locations
            }
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'message': str(e)})


@app.route('/api/raw-data', methods=['POST'])
def get_raw_data():
    """API untuk melihat data mentah dari sheet tertentu (untuk debugging)"""
    try:
        data = request.json
        spreadsheet_url = data.get('spreadsheet_url', DEFAULT_SPREADSHEET_URL)
        sheet_name = data.get('sheet_name', 'JANUARI')
        limit = data.get('limit', 10)

        sheet_id = extract_spreadsheet_id(spreadsheet_url)
        csv_data = download_sheet_as_csv(sheet_id, sheet_name)

        if csv_data is None:
            return jsonify({'success': False, 'message': f'Sheet {sheet_name} tidak ditemukan'})

        df = parse_csv_to_dataframe(csv_data)
        if df is None:
            return jsonify({'success': False, 'message': 'Gagal parse data'})

        return jsonify({
            'success': True,
            'data': {
                'sheet_name': sheet_name,
                'columns': df.columns.tolist(),
                'total_rows': len(df),
                'sample_data': df.head(limit).to_dict('records')
            }
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'message': str(e)})


@app.route('/api/month-detail', methods=['POST'])
def get_month_detail():
    """
    API untuk mendapatkan detail data untuk bulan tertentu
    Termasuk breakdown per lokasi
    """
    try:
        data = request.json
        spreadsheet_url = data.get('spreadsheet_url', DEFAULT_SPREADSHEET_URL)
        month = data.get('month', 'JANUARI').upper()

        if month not in MONTHS:
            return jsonify({'success': False, 'message': f'Bulan {month} tidak valid'})

        sheet_id = extract_spreadsheet_id(spreadsheet_url)
        csv_data = download_sheet_as_csv(sheet_id, month)

        if csv_data is None:
            return jsonify({'success': False, 'message': f'Sheet {month} tidak ditemukan'})

        df = parse_csv_to_dataframe(csv_data)
        if df is None or df.empty:
            return jsonify({'success': False, 'message': f'Sheet {month} kosong'})

        # Cari kolom yang sesuai
        col_respon = find_column(df, 'durasi_respon')
        col_penanganan = find_column(df, 'durasi_penanganan')
        col_lokasi = find_column(df, 'lokasi')

        # Hitung overall average untuk bulan ini
        avg_respon = 0
        avg_penanganan = 0
        total_tiket = len(df)

        if col_respon:
            respon_values = df[col_respon].dropna().apply(parse_duration_to_minutes)
            respon_values = respon_values[respon_values > 0]
            if len(respon_values) > 0:
                avg_respon = respon_values.mean()

        if col_penanganan:
            penanganan_values = df[col_penanganan].dropna().apply(parse_duration_to_minutes)
            penanganan_values = penanganan_values[penanganan_values > 0]
            if len(penanganan_values) > 0:
                avg_penanganan = penanganan_values.mean()

        # Hitung per lokasi
        location_breakdown = []
        if col_lokasi and col_penanganan:
            for lokasi in df[col_lokasi].dropna().unique():
                if not lokasi or lokasi in ['GMEDIA', 'nan', '', None]:
                    continue

                lokasi_df = df[df[col_lokasi] == lokasi]
                tiket_count = len(lokasi_df)

                # Avg respon per lokasi
                avg_respon_lokasi = 0
                if col_respon:
                    respon_vals = lokasi_df[col_respon].dropna().apply(parse_duration_to_minutes)
                    respon_vals = respon_vals[respon_vals > 0]
                    if len(respon_vals) > 0:
                        avg_respon_lokasi = respon_vals.mean()

                # Avg penanganan per lokasi
                avg_penanganan_lokasi = 0
                penanganan_vals = lokasi_df[col_penanganan].dropna().apply(parse_duration_to_minutes)
                penanganan_vals = penanganan_vals[penanganan_vals > 0]
                if len(penanganan_vals) > 0:
                    avg_penanganan_lokasi = penanganan_vals.mean()

                category = get_location_category(lokasi)
                location_breakdown.append({
                    'location': lokasi,
                    'category': category,
                    'total_tiket': tiket_count,
                    'avg_respon_minutes': round(avg_respon_lokasi, 2),
                    'avg_penanganan_minutes': round(avg_penanganan_lokasi, 2)
                })

        # Sort by category then by avg_penanganan descending
        location_breakdown = sorted(location_breakdown,
                                   key=lambda x: (get_category_order(x['category']), -x['avg_penanganan_minutes']))

        return jsonify({
            'success': True,
            'data': {
                'month': month,
                'summary': {
                    'total_tiket': total_tiket,
                    'avg_respon_minutes': round(avg_respon, 2),
                    'avg_penanganan_minutes': round(avg_penanganan, 2),
                    'total_lokasi': len(location_breakdown)
                },
                'location_breakdown': location_breakdown,
                'columns_found': {
                    'respon': col_respon,
                    'penanganan': col_penanganan,
                    'lokasi': col_lokasi
                }
            }
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'message': str(e)})


@app.route('/api/yearly-recap', methods=['POST'])
def get_yearly_recap():
    """
    API untuk mendapatkan rekap data 1 tahun dengan breakdown per triwulan
    """
    try:
        data = request.json
        spreadsheet_url = data.get('spreadsheet_url', DEFAULT_SPREADSHEET_URL)

        # Hitung dari data mentah
        monthly_data, quarterly_data, location_data = calculate_monthly_averages_from_raw(spreadsheet_url)

        if monthly_data is None:
            return jsonify({'success': False, 'message': 'Gagal membaca data'})

        # Organize data per quarter
        quarters_detail = {
            'Q1': {'months': ['JANUARI', 'FEBRUARI', 'MARET'], 'data': [], 'locations': {}},
            'Q2': {'months': ['APRIL', 'MEI', 'JUNI'], 'data': [], 'locations': {}},
            'Q3': {'months': ['JULI', 'AGUSTUS', 'SEPTEMBER'], 'data': [], 'locations': {}},
            'Q4': {'months': ['OKTOBER', 'NOVEMBER', 'DESEMBER'], 'data': [], 'locations': {}}
        }

        for m in monthly_data:
            if m['status'] != 'ok':
                continue
            for q_name, q_info in quarters_detail.items():
                if m['bulan'] in q_info['months']:
                    q_info['data'].append(m)
                    break

        # Calculate quarterly averages
        quarterly_summary = []
        for q_name in ['Q1', 'Q2', 'Q3', 'Q4']:
            q_data = quarters_detail[q_name]['data']
            if q_data:
                avg_respon = sum(m['avg_respon_minutes'] for m in q_data) / len(q_data)
                avg_penanganan = sum(m['avg_penanganan_minutes'] for m in q_data) / len(q_data)
                total_tiket = sum(m.get('total_tiket', 0) for m in q_data)
                quarterly_summary.append({
                    'quarter': q_name,
                    'quarter_name': f'Triwulan {q_name[1]}',
                    'months': quarters_detail[q_name]['months'],
                    'avg_respon_minutes': round(avg_respon, 2),
                    'avg_penanganan_minutes': round(avg_penanganan, 2),
                    'total_tiket': total_tiket,
                    'monthly_data': q_data
                })

        # Calculate yearly totals
        valid_months = [m for m in monthly_data if m['status'] == 'ok']
        yearly_avg_respon = 0
        yearly_avg_penanganan = 0
        yearly_total_tiket = 0

        if valid_months:
            respon_vals = [m['avg_respon_minutes'] for m in valid_months if m['avg_respon_minutes'] > 0]
            penanganan_vals = [m['avg_penanganan_minutes'] for m in valid_months if m['avg_penanganan_minutes'] > 0]

            if respon_vals:
                yearly_avg_respon = sum(respon_vals) / len(respon_vals)
            if penanganan_vals:
                yearly_avg_penanganan = sum(penanganan_vals) / len(penanganan_vals)

            yearly_total_tiket = sum(m.get('total_tiket', 0) for m in valid_months)

        # Location yearly summary
        location_yearly = []
        for loc in location_data:
            location_yearly.append({
                'location': loc['location'],
                'category': loc['category'],
                'avg_penanganan_minutes': loc['avg_minutes'],
                'total_records': loc.get('total_records', 0),
                'monthly_breakdown': loc.get('monthly_data', [])
            })

        return jsonify({
            'success': True,
            'data': {
                'yearly_summary': {
                    'avg_respon_minutes': round(yearly_avg_respon, 2),
                    'avg_penanganan_minutes': round(yearly_avg_penanganan, 2),
                    'total_tiket': yearly_total_tiket,
                    'total_months': len(valid_months),
                    'total_lokasi': len(location_data)
                },
                'quarterly': quarterly_summary,
                'monthly': [m for m in monthly_data if m['status'] == 'ok'],
                'locations': location_yearly
            }
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'message': str(e)})


if __name__ == '__main__':
    print("="*60)
    print("Dashboard Monitoring - Menghitung Average dari Data Mentah")
    print("="*60)
    print("URL: http://127.0.0.1:5000")
    print("")
    print("Fitur:")
    print("  - Membaca data dari setiap sheet bulanan (JANUARI-DESEMBER)")
    print("  - Menghitung sendiri average durasi respon & penanganan")
    print("  - Analisis per lokasi")
    print("  - Detail per bulan dengan breakdown lokasi")
    print("  - Rekap tahunan dengan breakdown triwulan")
    print("  - Real-time dari Google Spreadsheet")
    print("="*60)
    app.run(debug=True, host='0.0.0.0', port=5000)
