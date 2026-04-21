from flask import Flask, render_template, jsonify, request, send_file
from flask_socketio import SocketIO, emit
import sqlite3
import pandas as pd
import io
import os
import json
from datetime import datetime, timedelta
from collections import defaultdict
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)

DATABASE = os.getenv('DB_PATH', 'software_log.db')

UPLOAD_FOLDER = os.getenv('UPLOAD_FOLDER', r'C:\log')
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

UPDATE = os.getenv('UPDATE_FOLDER', r'C:\log')
if not os.path.exists(UPDATE):
    os.makedirs(UPDATE)

def get_data(query='', column='', start_time=None, end_time=None, limit=None):
    conn = sqlite3.connect(DATABASE)
    cursor = conn.cursor()

    column = column if column in ["USERNAME", "ProgramName", "PID", "StartTime", "EndTime", "USTime", "COMPUTERNAME"] else "USERNAME"
    limit_clause = f"LIMIT {limit}" if limit else ""

    if column == 'USERNAME' and ',' in query:
        computer_names = [name.strip() for name in query.split(',')]
        like_clauses = [f"USERNAME LIKE ?" for _ in computer_names]
        where_clause = " OR ".join(like_clauses)
        params = [f"%{name}%" for name in computer_names]
    else:
        where_clause = f"{column} LIKE ?"
        params = [f"%{query}%"]


    if start_time and end_time:
        where_clause += """
        AND (
            (StartTime <= strftime('%Y-%m-%d %H:%M:%S', ?) AND EndTime >= strftime('%Y-%m-%d %H:%M:%S', ?))
            OR (StartTime BETWEEN strftime('%Y-%m-%d %H:%M:%S', ?) AND strftime('%Y-%m-%d %H:%M:%S', ?))
            OR (EndTime BETWEEN strftime('%Y-%m-%d %H:%M:%S', ?) AND strftime('%Y-%m-%d %H:%M:%S', ?))
        )
        """
        params.extend([start_time, end_time, start_time, end_time, start_time, end_time])


    cursor.execute(f'''
        SELECT USERNAME, ProgramName, PID, StartTime, EndTime, USTime, COMPUTERNAME
        FROM program_usage
        WHERE {where_clause}
        ORDER BY id DESC
        {limit_clause}
    ''', tuple(params))

    data = cursor.fetchall()
    conn.close()

    formatted_data = []
    for row in data:
        username = row[0]
        formatted_username = username.split("\\")[-1]
        formatted_data.append((formatted_username, *row[1:]))

    return formatted_data



@app.route('/')
def index():
    data = get_data(limit=15)

    sorted_data = sorted(data, key=lambda x: x[4], reverse=True)
    latest_end_times = [row[4] for row in sorted_data[:5]]

    return render_template('index.html', data=data, latest_end_times=latest_end_times)


@app.route('/update')
def update():
    data = get_data(limit=15)
    return jsonify(data)

@app.route('/RDP_Status')
def rdp_status():
    try:
        if not os.path.exists(RDP_LOG_PATH):
            return jsonify({'error': 'File not found'}), 404

        with open(RDP_LOG_PATH, 'r') as file:
            content = file.read()
        return jsonify({'content': content}), 200
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/ordinary_search')
def ordinary_search():
    return render_template('ordinary_search.html')

@app.route('/advanced_search')
def advanced_search():
    return render_template('advanced_search.html')


@app.route('/search_results')
def search_results():
    query = request.args.get('query', '')
    column = request.args.get('column', 'USERNAME')
    start_time = request.args.get('start_time')
    end_time = request.args.get('end_time')



    if column == 'department':
        column = 'USERNAME'
        query = f"%{query}%"


        if ',' in query:
            query = query.strip()
        else:
            query = f"{query}"

    data = get_data(query, column, start_time, end_time)
    return render_template('search_results.html', data=data)

@app.route('/graph')
def graph():
    return render_template('graph.html')



@app.route('/export')
def export():
    query = request.args.get('query', '')
    column = request.args.get('column', 'USERNAME')
    start_time = request.args.get('start_date', '')
    end_time = request.args.get('end_date', '')
    include_pages = request.args.get('include-checkbox', 'no')


    data = get_data(query, column, start_time=start_time, end_time=end_time)


    df = pd.DataFrame(data, columns=["使用者", "軟體名稱", "PID", "開始使用時間", "結束使用時間", "軟體使用時間", "電腦名稱"])

    df["使用者"] = df["使用者"].str.replace(r"^Hy-", "", regex=True)


    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        if include_pages == 'yes':
            for user, group in df.groupby("使用者"):
                group.to_excel(writer, index=False, sheet_name=user[:30])
        else:
            df.to_excel(writer, index=False, sheet_name='查詢結果')

    output.seek(0)
    timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
    download_name = f'{timestamp}.xlsx'

    return send_file(output, as_attachment=True, download_name=download_name, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'message': 'No file part'}), 400

    file = request.files['file']

    if file.filename == '':
        return jsonify({'message': 'No selected file'}), 400

    if file:
        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(file_path)
        return jsonify({'message': f'File {file.filename} uploaded successfully'}), 200

@app.route('/uploadnew', methods=['POST'])
def upload_json():
    data = request.get_json()

    if not data:
        return jsonify({'message': 'No JSON data received'}), 400

    try:
        conn = sqlite3.connect(DATABASE)
        cursor = conn.cursor()

        cursor.execute('''
            SELECT * FROM program_usage
            WHERE ProgramName=? AND PID=? AND StartTime=?
        ''', (data['ProgramName'], data['PID'], data['StartTime']))
        existing = cursor.fetchone()

        if existing:
            cursor.execute('''
                UPDATE program_usage
                SET EndTime=?, USTime=?, CPU_AVG=?, MEMORY_AVG=?
                WHERE ProgramName=? AND PID=? AND StartTime=?
            ''', (
                data['EndTime'], data['USTime'], data['CPU_AVG'], data['MEMORY_AVG'],
                data['ProgramName'], data['PID'], data['StartTime']
            ))
            msg = 'Updated existing record'
        else:
            cursor.execute('''
                INSERT INTO program_usage (
                    USERNAME, ProgramName, PID, StartTime, EndTime, USTime, FilePath,
                    COMPUTERNAME, CPU_AVG, MEMORY_AVG
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                data['USERNAME'], data['ProgramName'], data['PID'], data['StartTime'],
                data['EndTime'], data['USTime'], str(data['FilePath']),
                data['COMPUTERNAME'], data['CPU_AVG'], data['MEMORY_AVG']
            ))
            msg = 'Inserted new record'

        conn.commit()
        conn.close()
        return jsonify({'message': msg}), 200

    except Exception as e:
        return jsonify({'message': f'Error processing data: {str(e)}'}), 500

if __name__ == '__main__':
    init_db()
    app.run(host='0.0.0.0', port=5000)


@app.route('/json')
def get_json():
    try:

        json_path = os.path.join(UPLOAD_FOLDER, 'target_process.json')

        if not os.path.exists(json_path):
            return jsonify({'error': 'File not found'}), 404


        with open(json_path, 'r') as file:
            data = json.load(file)

        return jsonify(data)
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/update/version.ini', methods=['GET'])
def get_version():
    version_file_path = os.path.join(UPDATE, 'version.ini')

    if os.path.exists(version_file_path):
        return send_file(version_file_path, as_attachment=False)
    else:
        return jsonify({'error': 'Version file not found'}), 404


@app.route('/update/copyright.ini', methods=['GET'])
def get_copyright():
    copyright_file_path = os.path.join(UPDATE, 'copyright.ini')

    if os.path.exists(copyright_file_path):
        return send_file(copyright_file_path, as_attachment=False)
    else:
        return jsonify({'error': 'copyright file not found'}), 404

@app.route('/update/SWUL.exe', methods=['GET'])
def get_program():
    program_file_path = os.path.join(UPDATE, 'SWUL.exe')

    if os.path.exists(program_file_path):
        return send_file(program_file_path, as_attachment=True)
    else:
        return jsonify({'error': 'Program file not found'}), 404

@app.route('/update_version', methods=['POST'])
def update_version():
    data = request.json

    version = data.get('version')
    computer_name = data.get('computer_name')

    if not computer_name or not version:
        return jsonify({"error": "Missing computer_name or version"}), 400

    try:
        conn = sqlite3.connect(DATABASE)
        cursor = conn.cursor()


        cursor.execute('''
            SELECT id, version, timestamp
            FROM version_info
            WHERE computer_name = ?
        ''', (computer_name,))
        existing_record = cursor.fetchone()

        if existing_record:

            record_id, existing_version, existing_timestamp = existing_record

            cursor.execute('''
                UPDATE version_info
                SET version = ?, timestamp = datetime('now', '+8 hours')
                WHERE id = ?
            ''', (version, record_id))
            message = "Version and timestamp updated successfully!"
        else:
            cursor.execute('''
                INSERT INTO version_info (computer_name, version, timestamp)
                VALUES (?, ?, datetime('now', '+8 hours'))
            ''', (computer_name, version))
            message = "New version record added successfully!"

        conn.commit()
        conn.close()

        return jsonify({"message": message}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/graph_data')
def graph_data():
    chart_type = request.args.get('chart_type', 'line')
    days = int(request.args.get('days', 30))
    selected_date = request.args.get('date')

    if selected_date:
        selected_date = datetime.strptime(selected_date, "%Y-%m-%d").date()
    else:
        selected_date = datetime.now().date()

    data = get_data('', 'USERNAME')
    df = pd.DataFrame(data, columns=["使用者", "軟體名稱", "PID", "開始使用時間", "結束使用時間", "軟體使用時間", "電腦名稱"])

    expanded_data = []
    for _, row in df.iterrows():
        time_data = split_time_by_day(row['開始使用時間'], row['結束使用時間'])
        for day_data in time_data:
            expanded_data.append({
                "使用者": row['使用者'],
                "軟體名稱": row['軟體名稱'],
                "電腦名稱": row['電腦名稱'],
                "日期": day_data['date'],
                "開始使用時間": day_data['start'],
                "結束使用時間": day_data['end']
            })

    expanded_df = pd.DataFrame(expanded_data)

    if expanded_df.empty:
        return jsonify({"error": "No data available."})


    def merge_overlapping_time_ranges(group):
        sorted_group = group.sort_values(by="開始使用時間")
        merged_ranges = []
        current_start = sorted_group.iloc[0]["開始使用時間"]
        current_end = sorted_group.iloc[0]["結束使用時間"]

        for i in range(1, len(sorted_group)):
            row = sorted_group.iloc[i]
            if row["開始使用時間"] <= current_end:
                current_end = max(current_end, row["結束使用時間"])
            else:

                merged_ranges.append((current_start, current_end))
                current_start = row["開始使用時間"]
                current_end = row["結束使用時間"]

        merged_ranges.append((current_start, current_end))
        return sum((end - start).total_seconds() / 60 for start, end in merged_ranges)


    expanded_df["開始使用時間"] = pd.to_datetime(expanded_df["開始使用時間"])
    expanded_df["結束使用時間"] = pd.to_datetime(expanded_df["結束使用時間"])


    merged_times = expanded_df.groupby(["使用者", "軟體名稱", "日期"]).apply(merge_overlapping_time_ranges).reset_index(name="軟體使用時間")
    merged_times['日期'] = pd.to_datetime(merged_times['日期'])


    end_date = datetime.now().date()
    start_date = end_date - timedelta(days=days)


    filtered_df = merged_times[(merged_times['日期'].dt.date >= start_date) & (merged_times['日期'].dt.date <= end_date)]


    today_df = merged_times[merged_times['日期'].dt.date == selected_date]


    user_software_usage = today_df.groupby(["使用者", "軟體名稱"])["軟體使用時間"].sum().unstack(fill_value=0)

    bar_chart_data = {
        "labels": user_software_usage.index.tolist(),
        "datasets": []
    }

    for software in user_software_usage.columns:
        bar_chart_data["datasets"].append({
            "label": software,
            "data": user_software_usage[software].tolist(),
            "backgroundColor": "rgba(54, 162, 235, 0.2)",
            "borderColor": "rgba(54, 162, 235, 1)",
            "borderWidth": 1
        })


    if not filtered_df.empty:
        usage_per_day = filtered_df.groupby(['日期', '軟體名稱'])['軟體使用時間'].sum().unstack(fill_value=0)
        line_chart_data = {
            "labels": usage_per_day.index.strftime('%Y-%m-%d').tolist(),
            "datasets": [
                {"label": software, "data": [round(value / 60, 1) for value in usage_per_day[software].tolist()]}
                for software in usage_per_day.columns
            ]
        }
    else:
        line_chart_data = {"labels": [], "datasets": []}

    print("長條圖:", bar_chart_data)
    print("折線圖:", line_chart_data)

    if chart_type == 'line':
        return jsonify({"line_chart": line_chart_data})
    elif chart_type == 'bar':
        return jsonify({"bar_chart": bar_chart_data})
    else:
        return jsonify({"error": "Invalid chart type."})

@app.route('/get_software_list', methods=['GET'])
def get_software_list():
    conn = sqlite3.connect(DATABASE)
    cursor = conn.cursor()

    cursor.execute("""
        SELECT DISTINCT
            CASE
                WHEN LOWER(ProgramName) LIKE 'solidworks%' THEN 'SolidWorks'
                ELSE LOWER(REPLACE(ProgramName, '.exe', ''))
            END AS software_name
        FROM program_usage
    """)

    software_list = [row[0] for row in cursor.fetchall()]
    conn.close()

    return jsonify(software_list)




def merge_overlapping_times(time_data):
    if not time_data:
        return []


    time_data.sort(key=lambda x: (x['date'], x['start']))

    merged = []
    current_date = time_data[0]['date']
    current_start = time_data[0]['start']
    current_end = time_data[0]['end']

    for i in range(1, len(time_data)):
        entry = time_data[i]


        if entry['date'] == current_date:
            if entry['start'] <= current_end:
                current_end = max(current_end, entry['end'])
            else:

                merged.append({'date': current_date, 'start': current_start, 'end': current_end})
                current_start = entry['start']
                current_end = entry['end']
        else:

            merged.append({'date': current_date, 'start': current_start, 'end': current_end})
            current_date = entry['date']
            current_start = entry['start']
            current_end = entry['end']


    merged.append({'date': current_date, 'start': current_start, 'end': current_end})

    return merged


def split_time_by_day(start_time_str, end_time_str):
    try:
        start_time = datetime.strptime(start_time_str, '%Y-%m-%d %H:%M:%S')
        end_time = datetime.strptime(end_time_str, '%Y-%m-%d %H:%M:%S')

        time_data = []
        current_day_start = start_time

        while current_day_start.date() <= end_time.date():
            if current_day_start.date() == start_time.date():
                current_day_end = current_day_start.replace(hour=23, minute=59, second=59)
            elif current_day_start.date() == end_time.date():
                current_day_end = end_time
            else:
                current_day_end = current_day_start.replace(hour=23, minute=59, second=59)

            time_data.append({
                'date': current_day_start.date(),
                'start': max(current_day_start, start_time),
                'end': min(current_day_end, end_time)
            })


            current_day_start += timedelta(days=1)
            current_day_start = current_day_start.replace(hour=0, minute=0, second=0)


        return merge_overlapping_times(time_data)

    except Exception as e:
        print(f"Time conversion error: {e} for values {start_time_str}, {end_time_str}")
        return []


if __name__ == '__main__':
    host = os.getenv('FLASK_HOST', '127.0.0.1')
    port = int(os.getenv('FLASK_PORT', 5000))
    app.run(host=host, port=port)