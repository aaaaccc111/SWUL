#V.4.0.2
#開發中
#修正把參數移到伺服器設定，減少修改使用者端程式碼問題
import asyncio
import psutil
import getpass
import platform
from datetime import datetime, timedelta
import requests
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

if not url or not json_url:
    print("錯誤：找不到環境變數 SERVER_URL 或 JSON_CONFIG_URL")

async def load_target_processes(url):
    try:
        response = requests.get(url)
        response.raise_for_status()
        data = response.json()
        return data.get('target_process_names', [])
    except Exception as e:
        print(f"Error loading target processes from URL {url}: {e}")
        return []


def get_process_user(pid):
    try:
        process = psutil.Process(pid)
        user_info = process.username()
        formatted_user_name = user_info.rsplit("\\", 1)[-1]  # 格式化取 \ 後方的使用者名稱
        return formatted_user_name
    except Exception as e:
        print(f"Error retrieving user for PID {pid}: {e}")
        return "Unknown"


def get_process_info(name):
    processes = []
    for proc in psutil.process_iter(['pid', 'name']):
        if name.lower() in proc.info['name'].lower():
            user_name = get_process_user(proc.info['pid'])
            processes.append((proc.info['pid'], proc.info['name'], user_name))
    return processes


def is_user_initiated(process):
    boot_time = psutil.boot_time()
    process_start_time = process.create_time()
    return (process_start_time - boot_time) > boot_time_threshold


def get_file_path(pid):
    try:
        proc = psutil.Process(pid)
        return proc.cmdline()
    except (psutil.NoSuchProcess, psutil.AccessDenied) as e:
        print(f"Error retrieving command line for PID {pid}: {e}")
        return "None"


def log_usage(user_name, pname, pid, start, end, usage_time, path, computer_name, avg_cpu, avg_memory):
    data = {
        "USERNAME": user_name,
        "ProgramName": pname,
        "PID": pid,
        "StartTime": start.strftime('%Y-%m-%d %H:%M:%S'),
        "EndTime": end.strftime('%Y-%m-%d %H:%M:%S'),
        "USTime": str(usage_time).split('.')[0],
        "FilePath": path,
        "ComputerName": computer_name,
        "CPU_AVG": avg_cpu,
        "MEMORY_AVG": avg_memory
    }

    try:
        response = requests.post(url, json=data)
        if response.status_code == 200:
            print(f"Data uploaded successfully: {response.json()}")
        else:
            print(f"Failed to upload data. Status code: {response.status_code}, Response: {response.text}")
    except Exception as e:
        print(f"Error sending data to server: {e}")


def update_end_time(pname, pid):
    if start_times.get(pid) and last_update_times.get(pid):
        end_times[pid] = datetime.now()
        usage_time = end_times[pid] - start_times[pid]
        avg_cpu = sum(cpu_usages[pid]) / len(cpu_usages[pid]) if cpu_usages[pid] else 0
        avg_memory = sum(memory_usages[pid]) / len(memory_usages[pid]) if memory_usages[pid] else 0
        log_usage(process_users.get(pid, "Unknown"), pname, pid, start_times[pid], end_times[pid], usage_time, file_paths[pid], computer_name, avg_cpu, avg_memory)
        last_update_times[pid] = end_times[pid]
        print(f"Ended log for {pname} (PID: {pid})")


async def main():
    global start_times, end_times, last_update_times, file_paths, cpu_usages, memory_usages, process_users, target_process_names, last_json_check_time
    update_interval = timedelta(minutes=10)
    json_check_interval = timedelta(minutes=5)

    target_process_names = await load_target_processes(json_url)
    last_json_check_time = datetime.now()

    while True:
        current_time = datetime.now()

        # 每五分鐘檢查一次 JSON 檔案，更新進程名稱列表
        if current_time - last_json_check_time >= json_check_interval:
            new_target_process_names = await load_target_processes(json_url)
            if new_target_process_names != target_process_names:
                print("Updated target process names from JSON.")
                target_process_names = new_target_process_names
            last_json_check_time = current_time

        # 更新進程的結束時間
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

        # 檢查所有目標程式的進程
        for name in target_process_names:
            processes = get_process_info(name)
            for pid, pname, user_name in processes:
                if pid not in start_times and is_user_initiated(psutil.Process(pid)):
                    start_times[pid] = datetime.now()
                    last_update_times[pid] = start_times[pid]
                    cmdline = get_file_path(pid)
                    file_paths[pid] = cmdline if cmdline else "None"
                    process_users[pid] = user_name
                    cpu_usages[pid] = []
                    memory_usages[pid] = []
                    print(f"Started monitoring {pname} (PID: {pid})")

        await asyncio.sleep(5)  # 設定每 5 秒檢查一次

if __name__ == "__main__":
    asyncio.run(main())

