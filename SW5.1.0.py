#V.5.1.0
#修改紀錄直接上傳資料庫不做成excel檔、多個程式icon訊息變更
#2025/04/21開始修正
#2025/4/28上線
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
from pystray import Icon, MenuItem, Menu
from PIL import Image
from win10toast import ToastNotifier
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
pname_map = {}


user_name = getpass.getuser()
computer_name = platform.node()
current_date = datetime.now().strftime("%Y%m")
excel_file_name = f"{current_date}_{user_name}.xlsx"
base_url = os.getenv('API_BASE_URL', 'http://localhost')
url = f"{base_url}/upload"
urlnew = f"{base_url}/uploadnew"
json_url = f"{base_url}/json"
icon_image_path = os.getenv('ICON_PATH')
program_path = os.getenv('EXE_PATH')

tray_icon = None  # 初始化tray_icon變數
active_pid_map = {}

def create_image():
    return Image.open(icon_image_path)

def on_update_clicked(icon, item):
    program_path = r"C:\Program Files\AutoSWUL\AutoSWUL"
    os.startfile(program_path)


def start_tray_icon():
    global tray_icon
    if tray_icon is None:
        menu = Menu(
            MenuItem("更新", on_update_clicked),
        )
        tray_icon = Icon("TestIcon", create_image(), title=f"軟體使用紀錄 - {user_name}",menu=menu)
        tray_icon.run()

def update_tray_title_from_map():
    global tray_icon, active_pid_map
    if tray_icon:
        if active_pid_map:
            pname_list = sorted(set(active_pid_map.values()))
            tray_icon.title = f"軟體使用紀錄 - {user_name} ({', '.join(pname_list)})"
        else:
            tray_icon.title = f"軟體使用紀錄 - {user_name}"

def handle_detected_software(pname, pid):
    global tray_icon, active_pid_map

    active_pid_map[pid] = pname

    if tray_icon:
        pname_str = ", ".join(sorted(set(active_pid_map.values())))
        tray_icon.title = f"軟體使用紀錄 - {user_name} ({pname_str})"

def handle_software_closed(pid):
    global tray_icon, active_pid_map

    if pid in active_pid_map:
        del active_pid_map[pid]

    if tray_icon:
        if active_pid_map:
            pname_str = ", ".join(sorted(set(active_pid_map.values())))
            tray_icon.title = f"軟體使用紀錄 - {user_name} ({pname_str})"
        else:
            tray_icon.title = f"軟體使用紀錄 - {user_name}"


def create_excel_file():
    return pd.DataFrame(columns=["USERNAME", "ProgramName", "PID", "StartTime", "EndTime", "USTime", "FilePath", "COMPUTERNAME", "CPU_AVG", "MEMORY_AVG"])


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

# 取得指定PID的使用者名稱
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

# 取得指定程式的所有程序
def get_process_info(name):
    processes = []
    for proc in psutil.process_iter(['pid', 'name']):
        if name.lower() in proc.info['name'].lower():
            user_name = get_process_user(proc.info['pid'])
            processes.append((proc.info['pid'], proc.info['name'], user_name))
    #print(f"Processes found for {name}: {processes}")  # 印出找到的行程
    return processes

# 檢查程序是否為使用者開啟
def is_user_initiated(process):
    try:
        boot_time = psutil.boot_time()
        process_start_time = process.create_time()
        return (process_start_time - boot_time) > boot_time_threshold
    except psutil.NoSuchProcess:
        print(f"Process no longer exists: {process.pid}")
        return False

# 當無法取得進程序資料時，進行錯誤處理
def get_file_path(pid):
    try:
        proc = psutil.Process(pid)
        return proc.cmdline()
    except (psutil.NoSuchProcess, psutil.AccessDenied) as e:
        print(f"Error retrieving command line for PID {pid}: {e}")
        return "None"

# 記錄指定程式使用狀況
def log_usage(user_name, pname, pid, start, end, usage_time, path, computer_name, avg_cpu, avg_memory):
    if user_name == "Unknown":
        return  # 如果使用者名稱無法取得，則不記錄

    data = {
        "USERNAME": user_name,
        "ProgramName": pname,
        "PID": str(pid),
        "StartTime": start.strftime('%Y-%m-%d %H:%M:%S'),
        "EndTime": end.strftime('%Y-%m-%d %H:%M:%S'),
        "USTime": str(usage_time).split('.')[0],
        "FilePath": str(path),
        "COMPUTERNAME": computer_name,
        "CPU_AVG": float(avg_cpu),
        "MEMORY_AVG": float(avg_memory)
    }

    try:
        response = requests.post(urlnew, json=data)
        if response.status_code == 200:
            print(f"Data uploaded successfully: {response.json()}")
        else:
            print(f"Failed to upload data. Status code: {response.status_code}, Response: {response.text}")
    except Exception as e:
        print(f"Error sending data to server: {e}")


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
        handle_software_closed(pid)


# 主程式
def main():
    global start_times, end_times, last_update_times, file_paths, cpu_usages, memory_usages, process_users, target_process_names, last_json_check_time
    update_interval = timedelta(minutes=10)
    json_check_interval = timedelta(minutes=5)

    #讀取JSON中指定軟體名稱
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
                # 從 start_times 或 active_pid_map 中獲取 pname
                pname = active_pid_map.get(pid, None)
                if pname is None:
                    # print(f"找不到 PID {pid} 的軟體名稱")
                    continue  # 可跳過處理(繼續)
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
                    # 如果該PID尚未被記錄，並且程序是由使用者開啟
                    if pid not in start_times and is_user_initiated(psutil.Process(pid)):
                        start_times[pid] = datetime.now()
                        last_update_times[pid] = start_times[pid]
                        cmdline = get_file_path(pid)
                        file_paths[pid] = cmdline if cmdline else "None"
                        process_users[pid] = user_name
                        cpu_usages[pid] = []
                        memory_usages[pid] = []
                        print(f"{pname} (PID: {pid}) started at {start_times[pid]}, file: {file_paths[pid]}, user: {process_users[pid]}")
                        handle_detected_software(pname,pid)
                except psutil.NoSuchProcess:
                    # 如果程序不存在，移除該PID的所有記錄
                    print(f"Process {pname} (PID: {pid}) no longer exists. Cleaning up.")
                    if pid in start_times:
                        update_end_time(pname, pid)  # 更新結束時間
                        del start_times[pid]
                        del end_times[pid]
                        del last_update_times[pid]
                        del file_paths[pid]
                        del cpu_usages[pid]
                        del memory_usages[pid]
                        del process_users[pid]

                except Exception as e:
                    print(f"Error processing PID {pid}: {e}")

        time.sleep(10)

if __name__ == "__main__":
    df = create_excel_file()
    import threading
    tray_thread = threading.Thread(target=start_tray_icon)
    tray_thread.daemon = True
    tray_thread.start()
    main()
