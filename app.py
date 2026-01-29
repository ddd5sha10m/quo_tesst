import os
import csv
import time
import shutil
import threading
import io
import tkinter as tk
from tkinter import messagebox
from datetime import datetime
import subprocess
import re

# 導入處理圖片與條碼的套件
from barcode import Code128
from barcode.writer import ImageWriter
from PIL import Image, ImageTk

# 導入檔案系統監控套件 (跨平台支援)
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# ================= 設定區 =================
# 1. 儀器磁碟機的根目錄 (用來執行掛載/卸載)
VOLUME_PATH = '/Volumes/QTEST1A9166' 

# 2. 實際存放 RES 檔案的資料夾 (如果是根目錄下的 Log 資料夾)
SOURCE_FOLDER = os.path.join(VOLUME_PATH, 'Log')

# 存放在你電腦本機的路徑
OUTPUT_CSV = './instrument_results.csv'
LOG_FILE = './processed_history.txt'

# 檢查間隔 (秒) - 建議設長一點，因為掛載需要時間
CHECK_INTERVAL = 300
# =========================================

class ResFileHandler(FileSystemEventHandler):
    """檔案系統事件處理器：監控 .res 檔案的建立"""
    def __init__(self, app):
        self.app = app
        self.processing_lock = threading.Lock()
    
    def on_created(self, event):
        """當有新檔案建立時觸發"""
        if event.is_directory:
            return
        
        filename = os.path.basename(event.src_path)
        
        # 只處理 .res 檔案，且排除幽靈檔案
        if filename.lower().endswith('.res') and not filename.startswith('._'):
            # 等待一下確保檔案寫入完成
            time.sleep(1)
            
            with self.processing_lock:
                processed_files = self.app.get_processed_files()
                if filename not in processed_files:
                    self.app.log_message(f"🔔 偵測到新檔案: {filename}")
                    self.app.process_files([filename])
    
    def on_modified(self, event):
        """當檔案被修改時觸發（某些系統會先建立空檔再寫入）"""
        if event.is_directory:
            return
        
        filename = os.path.basename(event.src_path)
        
        # 只處理 .res 檔案，且排除幽靈檔案
        if filename.lower().endswith('.res') and not filename.startswith('._'):
            # 檢查檔案大小，確保不是空檔
            try:
                if os.path.getsize(event.src_path) > 0:
                    time.sleep(0.5)  # 等待寫入完成
                    
                    with self.processing_lock:
                        processed_files = self.app.get_processed_files()
                        if filename not in processed_files:
                            self.app.log_message(f"📝 檔案已更新: {filename}")
                            self.app.process_files([filename])
            except:
                pass

class InstrumentApp:
    def __init__(self, root):
        self.root = root
        self.root.title("儀器資料監控與條碼助手")
        self.root.geometry("500x750") # 拉長視窗高度以容納按鈕
        
        # 檔案系統監控器
        self.observer = None

        # --- 上半部：條碼產生器 ---
        self.frame_top = tk.LabelFrame(root, text="條碼產生器", padx=10, pady=10)
        self.frame_top.pack(fill="both", expand="yes", padx=10, pady=5)

        tk.Label(self.frame_top, text="輸入 ID (如 19990510):").pack(pady=5)
        
        self.entry_code = tk.Entry(self.frame_top, font=("Arial", 14))
        self.entry_code.pack(pady=5)
        self.entry_code.bind('<Return>', lambda event: self.generate_barcode())
        
        btn_gen = tk.Button(self.frame_top, text="產生條碼", command=self.generate_barcode, bg="#ddd")
        btn_gen.pack(pady=5)

        self.lbl_image = tk.Label(self.frame_top, text="(條碼將顯示於此)")
        self.lbl_image.pack(pady=10)

        # --- 中間：儀器控制區 (這段是你之前漏掉的) ---
        self.frame_mid = tk.LabelFrame(root, text="儀器連線控制", padx=10, pady=10)
        self.frame_mid.pack(fill="x", padx=10, pady=5)
        
        # 顯示目前設定路徑
        tk.Label(self.frame_mid, text=f"監控路徑: ...{SOURCE_FOLDER[-20:]}", fg="gray").pack()

        # 手動刷新按鈕
        btn_refresh = tk.Button(self.frame_mid, text="🔄 強制刷新儀器 (Remount)", 
                                command=self.manual_refresh, bg="#ffdddd")
        btn_refresh.pack(fill="x", pady=5)

        # --- 下半部：系統監控紀錄 ---
        self.frame_bottom = tk.LabelFrame(root, text="系統監控紀錄 (System Log)", padx=10, pady=10)
        self.frame_bottom.pack(fill="both", expand="yes", padx=10, pady=5)

        self.txt_log = tk.Text(self.frame_bottom, height=12, state='disabled', bg="#f0f0f0")
        self.scrollbar = tk.Scrollbar(self.frame_bottom, command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=self.scrollbar.set)
        
        self.scrollbar.pack(side="right", fill="y")
        self.txt_log.pack(side="left", fill="both", expand=True)

        # --- 啟動背景監控 ---
        self.log_message("程式介面已載入，準備啟動背景監控...")
        self.start_monitoring_thread()

    def log_message(self, msg):
        """將訊息顯示在視窗下方的文字框"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        final_msg = f"[{timestamp}] {msg}\n"
        self.txt_log.config(state='normal')
        self.txt_log.insert('end', final_msg)
        self.txt_log.see('end')
        self.txt_log.config(state='disabled')

    def generate_barcode(self):
        """產生並顯示條碼"""
        content = self.entry_code.get().strip()
        if not content:
            messagebox.showwarning("提示", "請輸入 ID 內容！")
            return

        try:
            fp = io.BytesIO()
            Code128(content, writer=ImageWriter()).write(fp)
            fp.seek(0)
            img = Image.open(fp)
            base_width = 300
            w_percent = (base_width / float(img.size[0]))
            h_size = int((float(img.size[1]) * float(w_percent)))
            img = img.resize((base_width, h_size), Image.Resampling.LANCZOS)
            img_tk = ImageTk.PhotoImage(img)
            self.lbl_image.config(image=img_tk, text="")
            self.lbl_image.image = img_tk 
            self.log_message(f"已生成條碼: {content}")
        except Exception as e:
            self.log_message(f"條碼錯誤: {e}")
    
    # === 重新掛載功能 ===
    def get_disk_identifier(self, target_path):
        """找出儀器對應的硬體代號 (例如 disk2s1)"""
        try:
            # 使用 diskutil info 指令查詢
            cmd = ['diskutil', 'info', target_path]
            result = subprocess.check_output(cmd).decode('utf-8')
            match = re.search(r'Device Identifier:\s+(\w+)', result)
            if match:
                return match.group(1)
        except Exception as e:
            print(f"查詢磁碟代號失敗: {e}")
        return None

    def remount_drive(self):
        """執行卸載再掛載"""
        # 1. 檢查路徑是否存在，如果不存在，嘗試檢查父目錄或忽略
        target_volume = VOLUME_PATH 
        if not os.path.exists(target_volume):
            self.log_message("⚠️ 無法刷新：找不到儀器路徑")
            return False

        # 2. 取得 Device ID (例如 disk4s1)
        device_id = self.get_disk_identifier(target_volume)
        if not device_id:
            self.log_message("⚠️ 無法刷新：找不到裝置代號")
            return False

        try:
            self.log_message(f"🔄 正在刷新連線 ({device_id})...")
            
            # 3. 卸載 (Unmount) - 使用 device id 最穩
            subprocess.run(['diskutil', 'unmount', f'/dev/{device_id}'], check=True)
            
            # 等待 3 秒讓系統反應
            time.sleep(3)
            
            # 4. 掛載 (Mount)
            subprocess.run(['diskutil', 'mount', f'/dev/{device_id}'], check=True)
            
            # 再等一下確保檔案系統準備好
            time.sleep(2)
            
            self.log_message("✅ 刷新完成，檔案列表已更新")
            return True
        except Exception as e:
            self.log_message(f"❌ 刷新失敗: {e}")
            return False

    def manual_refresh(self):
        """按鈕觸發的刷新"""
        threading.Thread(target=self.remount_drive).start()

    def start_monitoring_thread(self):
        """啟動雙重監控機制：即時監控 + 定期掃描"""
        # 1. 啟動即時檔案系統監控
        if os.path.exists(SOURCE_FOLDER):
            try:
                event_handler = ResFileHandler(self)
                self.observer = Observer()
                self.observer.schedule(event_handler, SOURCE_FOLDER, recursive=False)
                self.observer.start()
                self.log_message("✅ 即時檔案監控已啟動")
            except Exception as e:
                self.log_message(f"⚠️ 即時監控啟動失敗: {e}")
                self.log_message("將使用定期掃描模式")
        
        # 2. 啟動定期掃描（作為備援機制）
        thread = threading.Thread(target=self.monitor_logic, daemon=True)
        thread.start()

    def monitor_logic(self):
        """定期掃描邏輯（作為即時監控的備援機制）"""
        
        # 確保本地檔案存在
        if not os.path.exists(OUTPUT_CSV): open(OUTPUT_CSV, 'a').close()
        if not os.path.exists(LOG_FILE): open(LOG_FILE, 'a').close()
        
        self.log_message(f"🔄 定期掃描已啟動 (間隔: {CHECK_INTERVAL}秒)")

        while True:
            try:
                time.sleep(CHECK_INTERVAL)
                
                if os.path.exists(SOURCE_FOLDER):
                    # 抓取所有 .res 檔案（不分大小寫）
                    all_files = [f for f in os.listdir(SOURCE_FOLDER) if f.lower().endswith('.res')]
                    
                    # 過濾掉 ._ 開頭的幽靈檔案
                    valid_files = [f for f in all_files if not f.startswith('._')]

                    processed_files = self.get_processed_files()
                    files_to_process = [f for f in valid_files if f not in processed_files]

                    if files_to_process:
                        self.log_message(f"🔎 定期掃描發現 {len(files_to_process)} 個新檔案")
                        self.process_files(files_to_process)
                else:
                    self.log_message(f"⚠️ 找不到資料夾: {SOURCE_FOLDER}")

            except Exception as e:
                self.log_message(f"定期掃描錯誤: {e}")
                time.sleep(CHECK_INTERVAL)

    def get_processed_files(self):
        if not os.path.exists(LOG_FILE): return set()
        with open(LOG_FILE, 'r', encoding='utf-8') as f:
            return set(line.strip() for line in f)

    def mark_as_processed(self, filename):
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(f"{filename}\n")

    def process_files(self, files):
        file_exists = os.path.isfile(OUTPUT_CSV)
        try:
            with open(OUTPUT_CSV, mode='a', newline='', encoding='utf-8-sig') as csvfile:
                fieldnames = ['Patient_ID', 'Sample_Seq', 'Timestamp', 'Test_Name', 'Result_Value', 'Unit', 'Source_File']
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                
                if not file_exists:
                    writer.writeheader()

                for filename in files:
                    source_path = os.path.join(SOURCE_FOLDER, filename)
                    result_data = self.parse_res_file(source_path)
                    
                    if result_data:
                        writer.writerow(result_data)
                        self.log_message(f"  ➜ 匯入: ID {result_data['Patient_ID']} ({filename})")
                        self.mark_as_processed(filename)
                    else:
                        self.log_message(f"  ❌ 跳過(格式不符): {filename}")
                        self.mark_as_processed(filename) 

        except Exception as e:
            self.log_message(f"寫入 CSV 失敗: {e}")
    

    def parse_res_file(self, file_path):
        """
        雙重解析模式：
        1. 優先讀取檔案內容 (標準格式)
        2. 若失敗，則嘗試讀取檔名 (救援模式)
        檔名範例: 0080p_A1C_5.5.res
        """
        data = {}
        filename = os.path.basename(file_path)
        
        # --- 策略 A: 嘗試讀取檔案內容 ---
        try:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read().strip()
                
            # 檢查內容是否正常 (要有 | 分隔符號)
            if '|' in content:
                parts = content.split('|')
                if len(parts) >= 4:
                    # === 既有的解析邏輯 ===
                    data['Patient_ID'] = parts[0].strip()
                    data['Sample_Seq'] = parts[1].strip()
                    meta = parts[2].split('^')
                    data['Test_Name'] = meta[0] if len(meta) > 0 else ""
                    raw_time = meta[2] if len(meta) > 2 else ""
                    data['Timestamp'] = raw_time[:19] if len(raw_time) >= 19 else datetime.now().strftime("%Y/%m/%d %H:%M:%S")
                    
                    res_block = parts[3].split('^')[0]
                    if '%' in res_block:
                        v, u = res_block.split('%', 1)
                        data['Result_Value'] = v
                        data['Unit'] = '%' + u
                    else:
                        data['Result_Value'] = res_block
                        data['Unit'] = ''
                    
                    data['Source_File'] = filename
                    return data # 成功回傳，結束函式

        except Exception as e:
            # 讀取失敗沒關係，我們還有 Plan B
            pass 

        # --- 策略 B: 檔名救援模式 ---
        # 如果上面失敗了 (內容是空的，或沒有 | )，我們來解析檔名
        # 假設檔名格式: 0080p_A1C_5.5.res
        try:
            # 去除副檔名 -> 0080p_A1C_5.5
            name_body = os.path.splitext(filename)[0]
            
            # 用底線 _ 切割
            parts = name_body.split('_')
            
            # 確保至少切出 3 塊 (序號, 項目, 結果)
            if len(parts) >= 3:
                self.log_message(f"⚠️ 啟動檔名解析模式: {filename}")
                
                # 0080p -> 去掉 p 當作序號或ID
                raw_id = parts[0].replace('p', '').replace('P', '')
                data['Patient_ID'] = "Unknown" # 檔名沒給病人ID，先填未知
                data['Sample_Seq'] = raw_id
                
                data['Test_Name'] = parts[1] # A1C
                data['Result_Value'] = parts[2] # 5.5
                data['Unit'] = "" # 檔名通常沒單位
                data['Timestamp'] = datetime.now().strftime("%Y/%m/%d %H:%M:%S") # 用現在時間
                data['Source_File'] = filename
                
                return data
            else:
                self.log_message(f"❌ 檔名格式也不符: {filename}")
                return None

        except Exception as e:
            self.log_message(f"解析全失敗 {filename}: {e}")
            return None

if __name__ == "__main__":
    root = tk.Tk()
    app = InstrumentApp(root)
    root.mainloop()