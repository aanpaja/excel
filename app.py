from flask import Flask, render_template, jsonify, request
import pandas as pd
import numpy as np
from datetime import timedelta
import re
import os

app = Flask(__name__)

def parse_duration_to_minutes(duration_str):
    if pd.isna(duration_str):
        return 0
    
    if isinstance(duration_str, (int, float)):
        return float(duration_str)
    
    if isinstance(duration_str, timedelta):
        return duration_str.total_seconds() / 60
    
    duration_str = str(duration_str).strip()
    
    try:
        parts = duration_str.split(',')
        total_minutes = 0
        
        for part in parts:
            part = part.strip()
            
            if 'day' in part:
                days = int(re.search(r'(\d+)', part).group(1))
                total_minutes += days * 24 * 60
            elif ':' in part:
                time_parts = part.split(':')
                if len(time_parts) == 3:
                    hours = int(time_parts[0])
                    minutes = int(time_parts[1])
                    seconds = float(time_parts[2])
                    total_minutes += hours * 60 + minutes + seconds / 60
        
        return round(total_minutes, 2)
    except:
        return 0

def get_spreadsheet_data(spreadsheet_url):
    try:
        if 'docs.google.com/spreadsheets' not in spreadsheet_url:
            return None
        
        sheet_id = spreadsheet_url.split('/d/')[1].split('/')[0]
        
        gid = None
        if 'gid=' in spreadsheet_url:
            gid = spreadsheet_url.split('gid=')[1].split('#')[0].split('&')[0]
        
        sheet_name = 'AVG'
        
        if gid:
            csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
        else:
            import urllib.parse
            encoded_sheet = urllib.parse.quote(sheet_name)
            csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&sheet={encoded_sheet}"
        
        print(f"Downloading from Google Sheets...")
        print(f"Sheet ID: {sheet_id}")
        
        try:
            import urllib.request
            response = urllib.request.urlopen(csv_url)
            csv_data = response.read().decode('utf-8')
            print("Download berhasil!")
        except Exception as e:
            if '403' in str(e):
                print("ERROR 403: Spreadsheet PRIVATE!")
                print("Solusi: Share spreadsheet ke 'Anyone with the link' (view access)")
            else:
                print(f"Error: {e}")
            return None
        
        from io import StringIO
        import csv
        
        lines = csv_data.split('\n')
        reader = csv.reader(lines)
        
        data_list = []
        for idx, row in enumerate(reader):
            if idx < 2:
                continue
            
            if len(row) < 4:
                continue
            
            bulan = row[1] if len(row) > 1 else ''
            respon = row[2] if len(row) > 2 else ''
            penanganan = row[3] if len(row) > 3 else ''
            
            if bulan and bulan.strip() and not bulan.startswith('TOTAL'):
                bulan_str = str(bulan).upper().strip()
                if '#DIV/0!' not in str(respon) and '#DIV/0!' not in str(penanganan):
                    data_list.append({
                        'BULAN': bulan_str,
                        'AVG_DURASI_RESPON': respon,
                        'AVG_PENANGANAN_GANGGUAN': penanganan
                    })
        
        df = pd.DataFrame(data_list)
        
        df['avg_respon_minutes'] = df['AVG_DURASI_RESPON'].apply(parse_duration_to_minutes)
        df['avg_penanganan_minutes'] = df['AVG_PENANGANAN_GANGGUAN'].apply(parse_duration_to_minutes)
        
        return df
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        return None

def extract_spreadsheet_id(url):
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

def read_local_excel(file_path):
    try:
        df = pd.read_excel(file_path, sheet_name='AVG', header=None)
        
        data_list = []
        for i in range(2, min(20, df.shape[0])):
            bulan = df.iloc[i, 1] if pd.notna(df.iloc[i, 1]) else df.iloc[i, 0]
            respon = df.iloc[i, 2]
            penanganan = df.iloc[i, 3]
            
            if pd.notna(bulan) and bulan != '' and not str(bulan).startswith('TOTAL'):
                bulan_str = str(bulan).upper().strip()
                if '#DIV/0!' not in str(respon) and '#DIV/0!' not in str(penanganan):
                    data_list.append({
                        'BULAN': bulan_str,
                        'AVG_DURASI_RESPON': respon,
                        'AVG_PENANGANAN_GANGGUAN': penanganan
                    })
        
        df = pd.DataFrame(data_list)
        
        df['avg_respon_minutes'] = df['AVG_DURASI_RESPON'].apply(parse_duration_to_minutes)
        df['avg_penanganan_minutes'] = df['AVG_PENANGANAN_GANGGUAN'].apply(parse_duration_to_minutes)
        
        return df
    except Exception as e:
        print(f"Error reading Excel: {e}")
        import traceback
        traceback.print_exc()
        return None

def read_location_data(file_path):
    try:
        df = pd.read_excel(file_path, sheet_name='AVG', header=None)
        
        locations = []
        for i in range(2, df.shape[1]):
            val = df.iloc[21, i]
            if pd.notna(val) and str(val) not in ['NO', 'BULAN', 'TOTAL AVG 1 TAHUN']:
                locations.append({'col_idx': i, 'name': str(val)})
        
        location_data = []
        for loc in locations:
            col_idx = loc['col_idx']
            loc_name = loc['name']
            
            monthly_values = []
            for row_idx in range(22, min(34, df.shape[0])):
                bulan = df.iloc[row_idx, 1]
                value = df.iloc[row_idx, col_idx]
                
                if pd.notna(bulan) and pd.notna(value) and '#DIV/0!' not in str(value):
                    bulan_str = str(bulan).upper().strip()
                    value_minutes = parse_duration_to_minutes(value)
                    monthly_values.append({
                        'bulan': bulan_str,
                        'value_minutes': value_minutes
                    })
            
            if monthly_values:
                avg_value = sum([m['value_minutes'] for m in monthly_values]) / len(monthly_values)
                category = get_location_category(loc_name)
                location_data.append({
                    'location': loc_name,
                    'category': category,
                    'avg_minutes': avg_value,
                    'monthly_data': monthly_values
                })
        
        return sorted(location_data, key=lambda x: (get_category_order(x['category']), x['location']))
    except Exception as e:
        print(f"Error reading location data: {e}")
        import traceback
        traceback.print_exc()
        return []

def get_location_category(location_name):
    location_lower = location_name.lower()
    
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

def read_monthly_data_per_location(file_path, location_name):
    months = ['JANUARI', 'FEBRUARI', 'MARET', 'APRIL', 'MEI', 'JUNI', 
              'JULI', 'AGUSTUS', 'SEPTEMBER', 'OKTOBER', 'NOVEMBER', 'DESEMBER']
    
    monthly_averages = []
    
    for month in months:
        try:
            df = pd.read_excel(file_path, sheet_name=month, header=0)
            
            if 'Lokasi Pelanggan' not in df.columns or 'Durasi Penanganan Gangguan' not in df.columns:
                monthly_averages.append({
                    'bulan': month,
                    'avg_minutes': 0,
                    'count': 0
                })
                continue
            
            location_data = df[df['Lokasi Pelanggan'] == location_name]
            
            if len(location_data) == 0:
                monthly_averages.append({
                    'bulan': month,
                    'avg_minutes': 0,
                    'count': 0
                })
                continue
            
            durations = location_data['Durasi Penanganan Gangguan'].dropna()
            
            if len(durations) == 0:
                monthly_averages.append({
                    'bulan': month,
                    'avg_minutes': 0,
                    'count': 0
                })
                continue
            
            total_minutes = 0
            for duration in durations:
                minutes = parse_duration_to_minutes(duration)
                total_minutes += minutes
            
            avg_minutes = total_minutes / len(durations)
            
            monthly_averages.append({
                'bulan': month,
                'avg_minutes': avg_minutes,
                'count': len(durations)
            })
            
        except Exception as e:
            print(f"Error reading {month} for {location_name}: {e}")
            monthly_averages.append({
                'bulan': month,
                'avg_minutes': 0,
                'count': 0
            })
    
    return monthly_averages

def get_all_locations_from_monthly_sheets(file_path):
    try:
        df = pd.read_excel(file_path, sheet_name='JANUARI', header=0)
        
        if 'Lokasi Pelanggan' not in df.columns:
            return []
        
        locations = df['Lokasi Pelanggan'].dropna().unique().tolist()
        
        valid_locations = [loc for loc in locations if loc not in ['GMEDIA', 'nan', '']]
        
        return sorted(valid_locations)
    except Exception as e:
        print(f"Error getting locations: {e}")
        return []

def read_location_data_from_csv(csv_data):
    try:
        from io import StringIO
        import csv
        
        lines = csv_data.split('\n')
        reader = csv.reader(lines)
        
        rows = list(reader)
        
        if len(rows) < 22:
            return []
        
        header_row = rows[21]
        locations = []
        for col_idx, val in enumerate(header_row):
            if col_idx >= 2 and val and val not in ['NO', 'BULAN', 'TOTAL AVG 1 TAHUN']:
                locations.append({'col_idx': col_idx, 'name': val})
        
        location_data = []
        for loc in locations:
            col_idx = loc['col_idx']
            loc_name = loc['name']
            
            monthly_values = []
            for row_idx in range(22, min(34, len(rows))):
                if row_idx >= len(rows):
                    break
                
                row = rows[row_idx]
                if len(row) <= col_idx or len(row) <= 1:
                    continue
                
                bulan = row[1] if len(row) > 1 else ''
                value = row[col_idx] if len(row) > col_idx else ''
                
                if bulan and value and '#DIV/0!' not in str(value):
                    bulan_str = str(bulan).upper().strip()
                    value_minutes = parse_duration_to_minutes(value)
                    if value_minutes > 0:
                        monthly_values.append({
                            'bulan': bulan_str,
                            'value_minutes': value_minutes
                        })
            
            if monthly_values:
                avg_value = sum([m['value_minutes'] for m in monthly_values]) / len(monthly_values)
                category = get_location_category(loc_name)
                location_data.append({
                    'location': loc_name,
                    'category': category,
                    'avg_minutes': avg_value,
                    'monthly_data': monthly_values
                })
        
        return sorted(location_data, key=lambda x: (get_category_order(x['category']), x['location']))
    except Exception as e:
        print(f"Error reading location data from CSV: {e}")
        import traceback
        traceback.print_exc()
        return []

@app.route('/')
def index():
    return render_template('dashboard.html')

@app.route('/api/data', methods=['POST'])
def get_data():
    try:
        data = request.json
        spreadsheet_url = data.get('spreadsheet_url', '')
        
        location_data = []
        
        if spreadsheet_url:
            df = get_spreadsheet_data(spreadsheet_url)
            
            if df is not None:
                sheet_id = spreadsheet_url.split('/d/')[1].split('/')[0]
                gid = None
                if 'gid=' in spreadsheet_url:
                    gid = spreadsheet_url.split('gid=')[1].split('#')[0].split('&')[0]
                
                sheet_name = 'AVG'
                if gid:
                    csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
                else:
                    import urllib.parse
                    encoded_sheet = urllib.parse.quote(sheet_name)
                    csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&sheet={encoded_sheet}"
                
                try:
                    import urllib.request
                    response = urllib.request.urlopen(csv_url)
                    csv_data = response.read().decode('utf-8')
                    location_data = read_location_data_from_csv(csv_data)
                except:
                    pass
        else:
            df = read_local_excel('/mnt/user-data/uploads/LAPORAN_HARIAN_HELPDESK_2025_BARU.xlsx')
            location_data = read_location_data('/mnt/user-data/uploads/LAPORAN_HARIAN_HELPDESK_2025_BARU.xlsx')
        
        if df is None:
            return jsonify({'success': False, 'message': 'Gagal membaca data. Pastikan spreadsheet sudah di-share sebagai "Anyone with the link"'})
        
        monthly_data = []
        quarterly_data = []
        
        for idx, row in df.iterrows():
            bulan = str(row['BULAN']).upper()
            
            if 'TRIWULAN' in bulan or 'KUARTAL' in bulan:
                quarterly_data.append({
                    'period': bulan,
                    'avg_respon_minutes': float(row['avg_respon_minutes']),
                    'avg_penanganan_minutes': float(row['avg_penanganan_minutes'])
                })
            elif bulan and bulan != 'NAN':
                monthly_data.append({
                    'bulan': bulan,
                    'avg_respon_minutes': float(row['avg_respon_minutes']),
                    'avg_penanganan_minutes': float(row['avg_penanganan_minutes'])
                })
        
        yearly_avg_respon = df[~df['BULAN'].str.contains('TRIWULAN|KUARTAL', na=False)]['avg_respon_minutes'].mean()
        yearly_avg_penanganan = df[~df['BULAN'].str.contains('TRIWULAN|KUARTAL', na=False)]['avg_penanganan_minutes'].mean()
        
        return jsonify({
            'success': True,
            'data': {
                'monthly': monthly_data,
                'quarterly': quarterly_data,
                'yearly': {
                    'avg_respon_minutes': float(yearly_avg_respon),
                    'avg_penanganan_minutes': float(yearly_avg_penanganan)
                },
                'locations': location_data
            }
        })
    
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'message': str(e)})

@app.route('/api/location-monthly', methods=['POST'])
def get_location_monthly():
    try:
        data = request.json
        location = data.get('location', '')
        
        if not location:
            return jsonify({'success': False, 'message': 'Lokasi tidak boleh kosong'})
        
        monthly_data = read_monthly_data_per_location(
            '/mnt/user-data/uploads/LAPORAN_HARIAN_HELPDESK_2025_BARU.xlsx',
            location
        )
        
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

@app.route('/api/locations-list', methods=['GET'])
def get_locations_list():
    try:
        locations = get_all_locations_from_monthly_sheets(
            '/mnt/user-data/uploads/LAPORAN_HARIAN_HELPDESK_2025_BARU.xlsx'
        )
        
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

if __name__ == '__main__':
    print("Dashboard Average Durasi Respon dan Penanganan Gangguan 2025")
    print("URL: http://127.0.0.1:5000")
    app.run(debug=True, host='0.0.0.0', port=5000)