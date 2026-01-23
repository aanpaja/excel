from flask import Flask, render_template, jsonify, request
import pandas as pd
import numpy as np
from datetime import timedelta
import re
import os

app = Flask(__name__)

# Configuration - Can read from Google Sheets OR local Excel file
# If SPREADSHEET_URL is empty or None, will use local Excel file
SPREADSHEET_URL = ''  # Set to Google Sheets URL if needed, e.g., 'https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID/edit'
EXCEL_FILE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'LAPORAN HARIAN HELPDESK 2025 BARU.xlsx')

# Indonesian month names
MONTH_NAMES = ['JANUARI', 'FEBRUARI', 'MARET', 'APRIL', 'MEI', 'JUNI',
               'JULI', 'AGUSTUS', 'SEPTEMBER', 'OKTOBER', 'NOVEMBER', 'DESEMBER']

def parse_duration_to_minutes(duration_str):
    """Parse various duration formats to minutes"""
    if pd.isna(duration_str):
        return 0

    if isinstance(duration_str, (int, float)):
        return float(duration_str)

    if isinstance(duration_str, timedelta):
        return duration_str.total_seconds() / 60

    duration_str = str(duration_str).strip()

    # Handle Excel date-time format like "1900-01-01 09:04:04.394000"
    if '1900-01-01' in duration_str or '1899-12-31' in duration_str:
        try:
            time_part = duration_str.split(' ')[1] if ' ' in duration_str else duration_str
            parts = time_part.split(':')
            if len(parts) >= 2:
                hours = int(parts[0])
                minutes = int(parts[1])
                seconds = float(parts[2].split('.')[0]) if len(parts) > 2 else 0
                return round(hours * 60 + minutes + seconds / 60, 2)
        except:
            pass

    # Handle HH:MM:SS format (without days)
    if ':' in duration_str and 'day' not in duration_str.lower():
        try:
            parts = duration_str.split(':')
            if len(parts) == 3:
                hours = int(parts[0])
                minutes = int(parts[1])
                seconds = float(parts[2].split('.')[0]) if '.' in parts[2] else float(parts[2])
                total_minutes = hours * 60 + minutes + seconds / 60
                return round(total_minutes, 2)
            elif len(parts) == 2:
                hours = int(parts[0])
                minutes = int(parts[1])
                return round(hours * 60 + minutes, 2)
        except:
            pass

    # Handle "X days, HH:MM:SS" format
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
                    seconds = float(time_parts[2].split('.')[0]) if '.' in time_parts[2] else float(time_parts[2])
                    total_minutes += hours * 60 + minutes + seconds / 60
                elif len(time_parts) == 2:
                    hours = int(time_parts[0])
                    minutes = int(time_parts[1])
                    total_minutes += hours * 60 + minutes

        return round(total_minutes, 2)
    except:
        return 0


def get_location_category(location_name):
    """Categorize location based on name"""
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
    """Get sort order for categories"""
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


# ============================================
# LOCAL EXCEL FILE READING FUNCTIONS
# ============================================

def read_local_excel_avg(file_path):
    """Read AVG sheet from local Excel file"""
    try:
        print(f"\n{'='*60}")
        print(f"Reading from LOCAL EXCEL: {file_path}")
        print(f"{'='*60}")

        if not os.path.exists(file_path):
            print(f"ERROR: File not found: {file_path}")
            return None

        df = pd.read_excel(file_path, sheet_name='AVG', header=None)

        print(f"AVG sheet loaded: {df.shape[0]} rows, {df.shape[1]} columns")

        data_list = []
        # Row 2-13 contain monthly data (JANUARI to DESEMBER)
        for i in range(2, min(14, df.shape[0])):
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

        # Row 14-17 contain quarterly data (TRIWULAN 1-4)
        for i in range(14, min(18, df.shape[0])):
            bulan = df.iloc[i, 1] if pd.notna(df.iloc[i, 1]) else df.iloc[i, 0]
            respon = df.iloc[i, 2]
            penanganan = df.iloc[i, 3]

            if pd.notna(bulan) and 'TRIWULAN' in str(bulan).upper():
                bulan_str = str(bulan).upper().strip()
                if '#DIV/0!' not in str(respon) and '#DIV/0!' not in str(penanganan):
                    data_list.append({
                        'BULAN': bulan_str,
                        'AVG_DURASI_RESPON': respon,
                        'AVG_PENANGANAN_GANGGUAN': penanganan
                    })

        df_result = pd.DataFrame(data_list)

        # Parse duration columns to minutes
        df_result['avg_respon_minutes'] = df_result['AVG_DURASI_RESPON'].apply(parse_duration_to_minutes)
        df_result['avg_penanganan_minutes'] = df_result['AVG_PENANGANAN_GANGGUAN'].apply(parse_duration_to_minutes)

        print(f"✓ Parsed {len(data_list)} rows from AVG sheet")
        for _, row in df_result.iterrows():
            print(f"  - {row['BULAN']}: Respon={row['avg_respon_minutes']:.2f}min, Penanganan={row['avg_penanganan_minutes']:.2f}min")

        return df_result
    except Exception as e:
        print(f"ERROR reading local Excel AVG: {e}")
        import traceback
        traceback.print_exc()
        return None


def read_local_excel_location_data(file_path):
    """Read per-location data from AVG sheet (rows 21+)"""
    try:
        df = pd.read_excel(file_path, sheet_name='AVG', header=None)

        # Row 21 contains location headers (Corporate, Retail BTG - BB, etc.)
        header_row = 21

        # Find all location columns (starting from column 2)
        locations = []
        for col_idx in range(2, df.shape[1]):
            val = df.iloc[header_row, col_idx]
            if pd.notna(val) and str(val).strip() not in ['NO', 'BULAN', 'TOTAL AVG 1 TAHUN', '']:
                locations.append({'col_idx': col_idx, 'name': str(val).strip()})

        print(f"\nFound {len(locations)} locations in AVG sheet")

        location_data = []
        for loc in locations:
            col_idx = loc['col_idx']
            loc_name = loc['name']

            monthly_values = []
            # Row 22-33 contain monthly data per location
            for row_idx in range(22, min(34, df.shape[0])):
                bulan = df.iloc[row_idx, 1]
                value = df.iloc[row_idx, col_idx]

                if pd.notna(bulan) and pd.notna(value) and '#DIV/0!' not in str(value):
                    bulan_str = str(bulan).upper().strip()
                    if bulan_str in MONTH_NAMES:
                        value_minutes = parse_duration_to_minutes(value)
                        monthly_values.append({
                            'bulan': bulan_str,
                            'avg_penanganan_minutes': value_minutes,
                            'avg_respon_minutes': 0  # AVG sheet doesn't have per-location respon data
                        })

            if monthly_values:
                avg_value = sum([m['avg_penanganan_minutes'] for m in monthly_values]) / len(monthly_values)
                category = get_location_category(loc_name)
                location_data.append({
                    'location': loc_name,
                    'category': category,
                    'avg_minutes': avg_value,
                    'monthly_data': monthly_values
                })

        # Sort by category and name
        location_data = sorted(location_data, key=lambda x: (get_category_order(x['category']), x['location']))

        print(f"✓ Processed {len(location_data)} locations")

        return location_data
    except Exception as e:
        print(f"ERROR reading location data from Excel: {e}")
        import traceback
        traceback.print_exc()
        return []


def read_monthly_sheets_from_excel(file_path):
    """Read all monthly sheets and aggregate data per location with both respon and penanganan"""
    try:
        print(f"\n{'='*60}")
        print("Reading monthly sheets from Excel...")
        print(f"{'='*60}")

        all_monthly_data = []
        location_monthly_data = {}

        xlsx = pd.ExcelFile(file_path)
        available_sheets = xlsx.sheet_names

        for month in MONTH_NAMES:
            if month not in available_sheets:
                print(f"  ⚠ Sheet {month} not found, skipping...")
                continue

            try:
                df_month = pd.read_excel(file_path, sheet_name=month, header=0)
                df_month.columns = df_month.columns.str.strip()

                # Find required columns
                lokasi_col = None
                respon_col = None
                penanganan_col = None

                for col in df_month.columns:
                    col_lower = col.lower().strip()
                    if 'lokasi' in col_lower and 'pelanggan' in col_lower:
                        lokasi_col = col
                    elif 'durasi' in col_lower and 'respon' in col_lower:
                        respon_col = col
                    elif 'durasi' in col_lower and 'penanganan' in col_lower:
                        penanganan_col = col

                if not lokasi_col:
                    print(f"  ⚠ Sheet {month}: 'Lokasi Pelanggan' column not found")
                    continue

                # Clean data - remove rows with empty location
                df_month_clean = df_month[df_month[lokasi_col].notna()].copy()

                # Parse duration columns
                if respon_col and respon_col in df_month_clean.columns:
                    df_month_clean['durasi_respon_minutes'] = df_month_clean[respon_col].apply(parse_duration_to_minutes)
                else:
                    df_month_clean['durasi_respon_minutes'] = 0

                if penanganan_col and penanganan_col in df_month_clean.columns:
                    df_month_clean['durasi_penanganan_minutes'] = df_month_clean[penanganan_col].apply(parse_duration_to_minutes)
                else:
                    df_month_clean['durasi_penanganan_minutes'] = 0

                # Calculate monthly averages
                avg_respon = df_month_clean['durasi_respon_minutes'].mean() if len(df_month_clean) > 0 else 0
                avg_penanganan = df_month_clean['durasi_penanganan_minutes'].mean() if len(df_month_clean) > 0 else 0

                all_monthly_data.append({
                    'bulan': month,
                    'avg_respon_minutes': avg_respon,
                    'avg_penanganan_minutes': avg_penanganan
                })

                # Group by location
                location_grouped = df_month_clean.groupby(lokasi_col).agg({
                    'durasi_respon_minutes': 'mean',
                    'durasi_penanganan_minutes': 'mean'
                }).reset_index()

                for _, row in location_grouped.iterrows():
                    loc_name = str(row[lokasi_col]).strip()

                    if loc_name not in location_monthly_data:
                        location_monthly_data[loc_name] = []

                    location_monthly_data[loc_name].append({
                        'bulan': month,
                        'avg_respon_minutes': row['durasi_respon_minutes'],
                        'avg_penanganan_minutes': row['durasi_penanganan_minutes']
                    })

                print(f"  ✓ {month}: {len(df_month_clean)} records, {len(location_grouped)} locations")

            except Exception as e:
                print(f"  ⚠ Error reading sheet {month}: {e}")
                continue

        # Convert location_monthly_data to list format
        locations_data = []
        for loc_name, monthly_data in location_monthly_data.items():
            category = get_location_category(loc_name)
            avg_penanganan = sum([m['avg_penanganan_minutes'] for m in monthly_data]) / len(monthly_data) if monthly_data else 0
            locations_data.append({
                'location': loc_name,
                'category': category,
                'avg_minutes': avg_penanganan,
                'monthly_data': monthly_data
            })

        # Sort by category and name
        locations_data = sorted(locations_data, key=lambda x: (get_category_order(x['category']), x['location']))

        print(f"\n✓ Total months processed: {len(all_monthly_data)}")
        print(f"✓ Total locations: {len(locations_data)}")

        return all_monthly_data, locations_data

    except Exception as e:
        print(f"ERROR in read_monthly_sheets_from_excel: {e}")
        import traceback
        traceback.print_exc()
        return [], []


# ============================================
# GOOGLE SHEETS READING FUNCTIONS
# ============================================

def list_available_google_sheets(spreadsheet_url):
    """Try to detect available sheets from Google Sheets"""
    try:
        import urllib.parse
        import urllib.request
        import urllib.error

        spreadsheet_url = spreadsheet_url.split('?')[0].split('#')[0]

        if '/d/' not in spreadsheet_url:
            print("ERROR: Invalid spreadsheet URL format!")
            return []

        sheet_id = spreadsheet_url.split('/d/')[1].split('/')[0]

        print(f"\nGoogle Spreadsheet ID: {sheet_id}")
        print("Detecting available sheets...")

        # Test common month name formats
        test_names = MONTH_NAMES.copy()

        available_sheets = []

        for sheet_name in test_names:
            try:
                encoded_sheet = urllib.parse.quote(sheet_name)
                csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&sheet={encoded_sheet}"

                response = urllib.request.urlopen(csv_url, timeout=10)
                csv_data = response.read().decode('utf-8')

                # Check if it's a valid monthly sheet
                if 'Lokasi Pelanggan' in csv_data or 'LOKASI PELANGGAN' in csv_data.upper():
                    available_sheets.append(sheet_name)
                    print(f"  ✓ Found: {sheet_name}")
            except urllib.error.HTTPError as e:
                if e.code == 403:
                    print(f"  ⚠ ERROR 403: Spreadsheet might be PRIVATE!")
                    print(f"     Please share spreadsheet as 'Anyone with the link' → Viewer")
                    break
                # 400 means sheet not found, which is expected
            except Exception:
                pass

        return available_sheets

    except Exception as e:
        print(f"ERROR listing Google sheets: {e}")
        return []


def read_google_sheets_data(spreadsheet_url):
    """Read data from Google Sheets"""
    try:
        import urllib.parse
        import urllib.request
        import urllib.error
        from io import StringIO

        spreadsheet_url = spreadsheet_url.split('?')[0].split('#')[0]
        sheet_id = spreadsheet_url.split('/d/')[1].split('/')[0]

        # First try to read AVG sheet
        print(f"\n{'='*60}")
        print("READING FROM GOOGLE SHEETS...")
        print(f"{'='*60}")

        # Try to get AVG sheet
        encoded_sheet = urllib.parse.quote('AVG')
        csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&sheet={encoded_sheet}"

        try:
            response = urllib.request.urlopen(csv_url, timeout=30)
            csv_data = response.read().decode('utf-8')
            print("✓ Downloaded AVG sheet from Google Sheets")

            # Parse AVG sheet data
            df = pd.read_csv(StringIO(csv_data), header=None)

            data_list = []
            for i in range(2, min(14, df.shape[0])):
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

            for i in range(14, min(18, df.shape[0])):
                bulan = df.iloc[i, 1] if pd.notna(df.iloc[i, 1]) else df.iloc[i, 0]
                respon = df.iloc[i, 2]
                penanganan = df.iloc[i, 3]

                if pd.notna(bulan) and 'TRIWULAN' in str(bulan).upper():
                    bulan_str = str(bulan).upper().strip()
                    if '#DIV/0!' not in str(respon) and '#DIV/0!' not in str(penanganan):
                        data_list.append({
                            'BULAN': bulan_str,
                            'AVG_DURASI_RESPON': respon,
                            'AVG_PENANGANAN_GANGGUAN': penanganan
                        })

            df_result = pd.DataFrame(data_list)
            df_result['avg_respon_minutes'] = df_result['AVG_DURASI_RESPON'].apply(parse_duration_to_minutes)
            df_result['avg_penanganan_minutes'] = df_result['AVG_PENANGANAN_GANGGUAN'].apply(parse_duration_to_minutes)

            return df_result

        except urllib.error.HTTPError as e:
            if e.code == 403:
                print("ERROR 403: Spreadsheet is PRIVATE!")
                print("Solusi: Share spreadsheet ke 'Anyone with the link' (view access)")
            else:
                print(f"ERROR {e.code}: Could not download AVG sheet")
            return None

    except Exception as e:
        print(f"ERROR reading Google Sheets: {e}")
        import traceback
        traceback.print_exc()
        return None


def read_google_sheets_location_data(spreadsheet_url):
    """Read per-location data from Google Sheets AVG sheet"""
    try:
        import urllib.parse
        import urllib.request
        from io import StringIO

        spreadsheet_url = spreadsheet_url.split('?')[0].split('#')[0]
        sheet_id = spreadsheet_url.split('/d/')[1].split('/')[0]

        encoded_sheet = urllib.parse.quote('AVG')
        csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&sheet={encoded_sheet}"

        response = urllib.request.urlopen(csv_url, timeout=30)
        csv_data = response.read().decode('utf-8')

        import csv
        lines = csv_data.split('\n')
        reader = csv.reader(lines)
        rows = list(reader)

        if len(rows) < 22:
            return []

        # Row 21 contains location headers
        header_row = rows[21]
        locations = []
        for col_idx, val in enumerate(header_row):
            if col_idx >= 2 and val and val.strip() not in ['NO', 'BULAN', 'TOTAL AVG 1 TAHUN', '']:
                locations.append({'col_idx': col_idx, 'name': val.strip()})

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
                    if bulan_str in MONTH_NAMES:
                        value_minutes = parse_duration_to_minutes(value)
                        if value_minutes > 0:
                            monthly_values.append({
                                'bulan': bulan_str,
                                'avg_penanganan_minutes': value_minutes,
                                'avg_respon_minutes': 0
                            })

            if monthly_values:
                avg_value = sum([m['avg_penanganan_minutes'] for m in monthly_values]) / len(monthly_values)
                category = get_location_category(loc_name)
                location_data.append({
                    'location': loc_name,
                    'category': category,
                    'avg_minutes': avg_value,
                    'monthly_data': monthly_values
                })

        return sorted(location_data, key=lambda x: (get_category_order(x['category']), x['location']))

    except Exception as e:
        print(f"ERROR reading Google Sheets location data: {e}")
        return []


# ============================================
# API ROUTES
# ============================================

@app.route('/')
def index():
    return render_template('dashboard.html')


@app.route('/api/data', methods=['POST', 'GET'])
def get_data():
    try:
        # Check if request has JSON data with spreadsheet_url
        spreadsheet_url = None
        if request.method == 'POST' and request.json:
            spreadsheet_url = request.json.get('spreadsheet_url', '')

        # If no URL provided, use configured default
        if not spreadsheet_url:
            spreadsheet_url = SPREADSHEET_URL

        location_data = []

        # Decide whether to use Google Sheets or local Excel
        if spreadsheet_url and 'docs.google.com/spreadsheets' in spreadsheet_url:
            print("\n" + "="*80)
            print("MODE: GOOGLE SHEETS")
            print("="*80)

            df = read_google_sheets_data(spreadsheet_url)
            location_data = read_google_sheets_location_data(spreadsheet_url) if df is not None else []
        else:
            print("\n" + "="*80)
            print("MODE: LOCAL EXCEL FILE")
            print("="*80)

            df = read_local_excel_avg(EXCEL_FILE_PATH)
            location_data = read_local_excel_location_data(EXCEL_FILE_PATH)

        if df is None or len(df) == 0:
            return jsonify({
                'success': False,
                'message': 'Gagal membaca data. Pastikan file Excel ada atau Google Sheets sudah di-share sebagai "Anyone with the link"'
            })

        # Separate monthly and quarterly data
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
            elif bulan and bulan != 'NAN' and bulan in MONTH_NAMES:
                monthly_data.append({
                    'bulan': bulan,
                    'avg_respon_minutes': float(row['avg_respon_minutes']),
                    'avg_penanganan_minutes': float(row['avg_penanganan_minutes'])
                })

        # Calculate yearly average (excluding quarterly data)
        monthly_only = df[~df['BULAN'].str.contains('TRIWULAN|KUARTAL', na=False)]
        yearly_avg_respon = monthly_only['avg_respon_minutes'].mean() if len(monthly_only) > 0 else 0
        yearly_avg_penanganan = monthly_only['avg_penanganan_minutes'].mean() if len(monthly_only) > 0 else 0

        print(f"\n✓ Yearly average respon: {yearly_avg_respon:.2f} minutes")
        print(f"✓ Yearly average penanganan: {yearly_avg_penanganan:.2f} minutes")
        print("="*80 + "\n")

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
    """Get detailed monthly data for a specific location from monthly sheets"""
    try:
        data = request.json
        location = data.get('location', '')

        if not location:
            return jsonify({'success': False, 'message': 'Lokasi tidak boleh kosong'})

        monthly_data = []

        for month in MONTH_NAMES:
            try:
                df = pd.read_excel(EXCEL_FILE_PATH, sheet_name=month, header=0)
                df.columns = df.columns.str.strip()

                # Find columns
                lokasi_col = None
                penanganan_col = None
                respon_col = None

                for col in df.columns:
                    col_lower = col.lower()
                    if 'lokasi' in col_lower and 'pelanggan' in col_lower:
                        lokasi_col = col
                    elif 'durasi' in col_lower and 'penanganan' in col_lower:
                        penanganan_col = col
                    elif 'durasi' in col_lower and 'respon' in col_lower:
                        respon_col = col

                if not lokasi_col or not penanganan_col:
                    monthly_data.append({
                        'bulan': month,
                        'avg_minutes': 0,
                        'avg_respon_minutes': 0,
                        'count': 0
                    })
                    continue

                # Filter by location
                location_data = df[df[lokasi_col] == location]

                if len(location_data) == 0:
                    monthly_data.append({
                        'bulan': month,
                        'avg_minutes': 0,
                        'avg_respon_minutes': 0,
                        'count': 0
                    })
                    continue

                # Calculate averages
                durations = location_data[penanganan_col].dropna()
                total_penanganan = sum([parse_duration_to_minutes(d) for d in durations])
                avg_penanganan = total_penanganan / len(durations) if len(durations) > 0 else 0

                avg_respon = 0
                if respon_col:
                    respon_durations = location_data[respon_col].dropna()
                    total_respon = sum([parse_duration_to_minutes(d) for d in respon_durations])
                    avg_respon = total_respon / len(respon_durations) if len(respon_durations) > 0 else 0

                monthly_data.append({
                    'bulan': month,
                    'avg_minutes': avg_penanganan,
                    'avg_respon_minutes': avg_respon,
                    'count': len(durations)
                })

            except Exception as e:
                print(f"Error reading {month} for {location}: {e}")
                monthly_data.append({
                    'bulan': month,
                    'avg_minutes': 0,
                    'avg_respon_minutes': 0,
                    'count': 0
                })

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
    """Get list of all unique locations from monthly sheets"""
    try:
        # Try to get locations from JANUARI sheet first
        try:
            df = pd.read_excel(EXCEL_FILE_PATH, sheet_name='JANUARI', header=0)
            df.columns = df.columns.str.strip()

            lokasi_col = None
            for col in df.columns:
                if 'lokasi' in col.lower() and 'pelanggan' in col.lower():
                    lokasi_col = col
                    break

            if lokasi_col:
                locations = df[lokasi_col].dropna().unique().tolist()
                valid_locations = [loc for loc in locations if loc and str(loc).strip() not in ['GMEDIA', 'nan', '', 'NaN']]
                return jsonify({
                    'success': True,
                    'data': {
                        'locations': sorted(valid_locations)
                    }
                })
        except Exception as e:
            print(f"Error getting locations from JANUARI: {e}")

        return jsonify({
            'success': False,
            'message': 'Tidak dapat membaca daftar lokasi'
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'message': str(e)})


if __name__ == '__main__':
    print("\n" + "="*60)
    print("Dashboard Average Durasi Respon dan Penanganan Gangguan 2025")
    print("="*60)
    print(f"URL: http://127.0.0.1:5000")
    print(f"Excel file: {EXCEL_FILE_PATH}")

    if os.path.exists(EXCEL_FILE_PATH):
        print(f"✓ Excel file found!")
    else:
        print(f"⚠ WARNING: Excel file not found at {EXCEL_FILE_PATH}")

    if SPREADSHEET_URL:
        print(f"Google Sheets URL: {SPREADSHEET_URL}")
    else:
        print("Mode: LOCAL EXCEL FILE (no Google Sheets URL configured)")

    print("="*60 + "\n")
    app.run(debug=True, host='0.0.0.0', port=5000)
