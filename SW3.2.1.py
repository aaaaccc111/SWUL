#V.3.2.1
#2024/12/11 新增icon圖示
import psutil
import time
from datetime import datetime, timedelta
import pandas as pd
import io
import getpass
import platform
import win32api
import win32security
import win32con
import requests
from pystray import Icon
from PIL import Image
import os
from dotenv import load_dotenv

load_dotenv()

boot_time_threshold = 30
start_times = {}
end_times = {}
last_update_times = {}
file_paths = {}
cpu_usages = {}
memory_usages = {}
process_users = {}

user_name = getpass.getuser()
computer_name = platform.node()
current_date = datetime.now().strftime("%Y%m")
url = os.getenv('SERVER_URL')
json_url = os.getenv('JSON_CONFIG_URL')
icon_image_path = os.getenv('ICON_PATH')


def create_excel_file():
    return pd.DataFrame(columns=["USERNAME", "ProgramName", "PID", "StartTime", "EndTime", "USTime", "FilePath", "COMPUTERNAME", "CPU_AVG", "MEMORY_AVG"])

# 工具列LOGO
def create_image():
    return Image.open(icon_image_path)


def start_tray_icon():
    icon = Icon("TestIcon", create_image(), title="Your Application Name")
    icon.title = "軟體使用紀錄 - {}".format(user_name)
    icon.run()

def load_target_processes(url):
    try:
        response = requests.get(url)
        response.raise_for_status()
        data = response.json()
        print(f"Response JSON: {data}")
        return data.get('target_process_names', [])
    except Exception as e:
        print(f"Error loading target processes from URL {url}: {e}")
        return []

# 取得指定 PID 的使用者名稱
def get_process_user(pid):
    try:
        process_handle = win32api.OpenProcess(win32con.PROCESS_QUERY_INFORMATION, False, pid)
        token_handle = win32security.OpenProcessToken(process_handle, win32security.TOKEN_QUERY)
        user_info = win32security.GetTokenInformation(token_handle, win32security.TokenUser)
        user_sid = user_info[0]
        user_name, domain, _ = win32security.LookupAccountSid(None, user_sid)
        return f"{domain}\\{user_name}"
    except Exception as e:
        print(f"Error retrieving user for PID {pid}: {e}")
        return "Unknown"

# 取得特定程式的所有進程
def get_process_info(name):
    processes = []
    for proc in psutil.process_iter(['pid', 'name']):
        if name.lower() in proc.info['name'].lower():
            user_name = get_process_user(proc.info['pid'])
            processes.append((proc.info['pid'], proc.info['name'], user_name))
    #print(f"Processes found for {name}: {processes}")  # 打印找到的進程
    return processes

# 檢查進程是否為使用者啟動
def is_user_initiated(process):
    try:
        boot_time = psutil.boot_time()
        process_start_time = process.create_time()
        return (process_start_time - boot_time) > boot_time_threshold
    except psutil.NoSuchProcess:
        print(f"Process no longer exists: {process.pid}")
        return False

# 當無法取得進程資訊時，進行錯誤處理
def get_file_path(pid):
    try:
        proc = psutil.Process(pid)
        return proc.cmdline()
    except (psutil.NoSuchProcess, psutil.AccessDenied) as e:
        print(f"Error retrieving command line for PID {pid}: {e}")
        return "None"

# 記錄程式使用狀況
def log_usage(user_name, pname, pid, start, end, usage_time, path, computer_name, avg_cpu, avg_memory):
    global df
    existing_row = df[(df["ProgramName"] == pname) & (df["PID"] == str(pid)) & (df["StartTime"] == start.strftime('%Y-%m-%d %H:%M:%S'))]
    if not existing_row.empty:
        idx = existing_row.index[0]
        df.at[idx, "EndTime"] = end.strftime('%Y-%m-%d %H:%M:%S')
        df.at[idx, "USTime"] = str(usage_time).split('.')[0]
        df.at[idx, "CPU_AVG"] = avg_cpu
        df.at[idx, "MEMORY_AVG"] = avg_memory
        print(f"Updated log for {pname} (PID: {pid}) at index {idx}")
    else:
        new_log = pd.DataFrame([{
            "USERNAME": user_name,
            "ProgramName": pname,
            "PID": str(pid),
            "StartTime": start.strftime('%Y-%m-%d %H:%M:%S'),
            "EndTime": end.strftime('%Y-%m-%d %H:%M:%S'),
            "USTime": str(usage_time).split('.')[0],
            "FilePath": path,
            "COMPUTERNAME": computer_name,
            "CPU_AVG": avg_cpu,
            "MEMORY_AVG": avg_memory
        }])
        df = pd.concat([df, new_log], ignore_index=True)
        print(f"Added new log for {pname} (PID: {pid})")
    save_logs()

# 儲存紀錄到 Excel 並上傳
def save_logs():
    global df, current_date, user_name
    this_date = datetime.now().strftime("%Y%m")
    if this_date != current_date:
        current_date = this_date
        df = create_excel_file()  # Reset DataFrame for new month

    try:
        # 使用 BytesIO 處理 Excel 檔案
        with io.BytesIO() as output:
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Log')
                workbook = writer.book
                worksheet = writer.sheets['Log']
                date_format = workbook.add_format({'num_format': 'yyyy-mm-dd hh:mm:ss'})
                worksheet.set_column('C:D', 20, date_format)
                worksheet.set_column('C:C', None)
            output.seek(0)

            # 上傳檔案
            files = {'file': (f"{current_date}_{user_name}.xlsx", output, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}
            response = requests.post(url, files=files)

            if response.status_code == 200:
                print(f"File uploaded successfully: {response.json()}")
            else:
                print(f"Failed to upload file. Status code: {response.status_code}, Response: {response.text}")

    except Exception as e:
        print(f"Error saving or uploading logs: {e}")

# 更新程式結束時間
def update_end_time(pname, pid):
    if start_times.get(pid) and last_update_times.get(pid):
        end_times[pid] = datetime.now()
        usage_time = end_times[pid] - start_times[pid]
        avg_cpu = sum(cpu_usages[pid]) / len(cpu_usages[pid]) if cpu_usages[pid] else 0
        avg_memory = sum(memory_usages[pid]) / len(memory_usages[pid]) if memory_usages[pid] else 0
        log_usage(process_users.get(pid, "Unknown"), pname, pid, start_times[pid], end_times[pid], usage_time, file_paths[pid], computer_name, avg_cpu, avg_memory)
        last_update_times[pid] = end_times[pid]
        print(f"Ended log for {pname} (PID: {pid})")


def main():
    global start_times, end_times, last_update_times, file_paths, cpu_usages, memory_usages, process_users, target_process_names, last_json_check_time
    update_interval = timedelta(minutes=10)
    json_check_interval = timedelta(minutes=5)

    target_process_names = load_target_processes(json_url)
    last_json_check_time = datetime.now()

    while True:
        current_time = datetime.now()

        if current_time - last_json_check_time >= json_check_interval:
            new_target_process_names = load_target_processes(json_url)
            if new_target_process_names != target_process_names:
                print("Updated target process names from JSON.")
                target_process_names = new_target_process_names
            last_json_check_time = current_time

        for pid in list(start_times.keys()):
            try:
                pname = psutil.Process(pid).name()
                if pid in start_times and (current_time - last_update_times[pid] >= update_interval):
                    update_end_time(pname, pid)
                    last_update_times[pid] = current_time
            except psutil.NoSuchProcess:
                update_end_time(pname, pid)
                del start_times[pid]
                del end_times[pid]
                del last_update_times[pid]
                del file_paths[pid]
                del cpu_usages[pid]
                del memory_usages[pid]
                del process_users[pid]

        for name in target_process_names:
            processes = get_process_info(name)
            for pid, pname, user_name in processes:
                try:
                    # 如果 PID 尚未記錄，並且進程是由使用者啟動
                    if pid not in start_times and is_user_initiated(psutil.Process(pid)):
                        start_times[pid] = datetime.now()
                        last_update_times[pid] = start_times[pid]
                        cmdline = get_file_path(pid)
                        file_paths[pid] = cmdline if cmdline else "None"
                        process_users[pid] = user_name
                        cpu_usages[pid] = []
                        memory_usages[pid] = []
                        print(f"{pname} (PID: {pid}) started at {start_times[pid]}, file: {file_paths[pid]}, user: {process_users[pid]}")

                    # 如果 PID 已記錄，更新其 CPU 和記憶體使用率
                    if pid in start_times:
                        p = psutil.Process(pid)
                        cpu_usages[pid].append(p.cpu_percent(interval=1))
                        memory_usages[pid].append(p.memory_percent())

                except psutil.NoSuchProcess:
                    # 如果進程不存在，移除該 PID 的所有記錄
                    print(f"Process {pname} (PID: {pid}) no longer exists. Cleaning up.")
                    if pid in start_times:
                        update_end_time(pname, pid)  # 記錄結束時間
                        del start_times[pid]
                        del end_times[pid]
                        del last_update_times[pid]
                        del file_paths[pid]
                        del cpu_usages[pid]
                        del memory_usages[pid]
                        del process_users[pid]

                except Exception as e:
                    # 捕捉其他可能的異常
                    print(f"Error processing PID {pid}: {e}")

        time.sleep(10)

if __name__ == "__main__":
    df = create_excel_file()
    import threading
    tray_thread = threading.Thread(target=start_tray_icon)
    tray_thread.daemon = True
    tray_thread.start()
    main()
