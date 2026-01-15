這是一份專業且易於閱讀的 `README.md` 文件草稿。我特別針對你程式中的「跨平台路徑設定」與「Mac 特有的刷新機制」做了詳細說明，讓使用 Windows 的同事也能順利上手。

你可以將以下內容複製並存檔為 `README.md`（或 `說明文件.txt`）。

---

# 儀器資料監控與條碼助手 (Instrument Data Monitor & Barcode Helper)

這是一個專為實驗室儀器（如 QTEST1A9166）設計的自動化工具。它能自動偵測儀器產出的檢驗報告（.RES 檔），將其轉換為易讀的 Excel/CSV 報表，並提供即時的條碼生成功能。

## ✨ 主要功能

* **自動監控**：背景程式自動掃描儀器資料夾，發現新檔案立即處理。
* **格式轉換**：自動解析 `.RES` 原始碼，提取病人 ID、檢體序號、檢測結果與時間，匯出為 `.csv` 檔。
* **條碼產生器**：內建 Code128 條碼生成器，輸入 ID 即可顯示條碼（方便補印標籤或測試）。
* **自動刷新 (Mac 專用)**：解決 macOS 對外接裝置的快取問題，自動重新掛載儀器以讀取最新檔案。
* **斷點續傳**：程式具備記憶功能，重啟後會自動比對並補齊漏掉的舊檔案。

---

## 🛠️ 安裝教學 (Installation)

### 1. 安裝 Python

請確保電腦已安裝 Python 3.8 或以上版本。

* [下載 Python](https://www.python.org/downloads/)

### 2. 安裝必要套件

開啟終端機 (Mac Terminal) 或 命令提示字元 (Windows CMD)，執行以下指令來安裝影像與條碼處理套件：

```bash
pip install python-barcode pillow

```

* `python-barcode`: 用於產生條碼。
* `pillow`: 用於在視窗介面中顯示圖片。
* *(註：tkinter, os, csv, threading 等為內建套件，無需額外安裝)*

---

## ⚙️ 設定說明 (Configuration)

在執行程式前，請使用文字編輯器（如 Notepad, VS Code）開啟 `app.py`，並根據你的作業系統修改最上方的 **設定區**。

### 🍎 Mac 使用者設定

Mac 需要指定掛載路徑 (Volume) 以便執行自動刷新。

```python
# ================= 設定區 =================
# 1. 儀器磁碟機的根目錄 (用來執行掛載/卸載)
VOLUME_PATH = '/Volumes/QTEST1A9166' 

# 2. 實際存放 RES 檔案的資料夾
SOURCE_FOLDER = os.path.join(VOLUME_PATH, 'Log')

# 3. 輸出檔案位置 (建議放在桌面或本機文件夾)
OUTPUT_CSV = './instrument_results.csv'
LOG_FILE = './processed_history.txt'

```

### 🪟 Windows 使用者設定

Windows 使用磁碟機代號（如 `E:` 或 `F:`），且**不需要**自動刷新功能（Windows 會自動抓到新檔）。

```python
# ================= 設定區 =================
# 1. Windows 不使用 VOLUME_PATH，可留空或忽略
VOLUME_PATH = 'E:\\' 

# 2. 指定儀器的路徑 (注意：Windows 路徑建議用雙斜線 \\ 或前綴 r)
SOURCE_FOLDER = r'E:\Log' 

# ...其他設定相同...

```

> **⚠️ Windows 特別注意**：
> 本程式的「自動刷新 (Remount)」功能使用了 Mac 專屬指令 (`diskutil`)。在 Windows 上執行時，該功能會自動失敗並顯示錯誤訊息，但不影響檔案讀取與轉檔功能。Windows 使用者可忽略刷新相關的錯誤提示。

---

## 🚀 如何使用 (Usage)

1. **連接儀器**：將儀器透過 USB 連接至電腦。
2. **執行程式**：
在終端機輸入：
```bash
python app.py

```


3. **操作介面**：
* **上方 (條碼區)**：輸入 ID 按 Enter，即時產生條碼圖片。
* **中間 (控制區)**：顯示目前監控路徑。Mac 用戶可點擊「強制刷新」手動更新連線。
* **下方 (紀錄區)**：顯示即時監控狀態，包含讀取到的新檔案 ID 與結果。


4. **查看結果**：
打開資料夾中的 `instrument_results.csv`，即可看到整理好的表格數據。

---

## ❓ 常見問題 (FAQ)

**Q1: 程式顯示「找不到路徑」？**

* 請確認儀器已正確連接電腦。
* 請確認 `app.py` 中的 `SOURCE_FOLDER` 路徑名稱是否與你的儀器名稱完全一致。

**Q2: Mac 上跳出 Finder 視窗很干擾？**

* 這是正常的。因為程式執行了「重新掛載 (Remount)」以強迫系統更新檔案列表，Mac 預設會開啟新掛載的資料夾。
* 若覺得困擾，可在程式碼中將 `CHECK_INTERVAL` 秒數調大（例如 30 秒刷新一次）。

**Q3: 如何重新處理所有檔案？**

* 若想讓程式重新讀取所有舊檔案，請刪除資料夾內的 `processed_history.txt` 紀錄檔，重新啟動程式即可。

---

### 📝 版本資訊

* **Version**: 1.0.0
* **Last Updated**: 2026-01-02
* **Author**: 程式夥伴 (AI Partner)