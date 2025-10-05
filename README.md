# KatFile 增強版上傳工具 v3.3

完整功能的KatFile.com檔案上傳工具，支援檔案壓縮、密碼保護和Word文件記錄。

## 🚀 主要功能

### 📤 檔案上傳
- 支援單檔案和批次檔案上傳
- 支援整個資料夾上傳
- 自動獲取真實直接下載連結
- 支援上傳到指定資料夾
- 即時進度顯示和狀態追蹤

### 🗜️ 檔案壓縮
- 支援ZIP和7Z格式壓縮
- 可設定解壓密碼保護
- 上傳前自動壓縮檔案
- 壓縮測試功能

### 📄 Word文件記錄
- 每個檔案自動生成Word記錄文件
- 基於自訂範本格式
- 自動填入檔案資訊和下載連結
- 解壓密碼自動填入

### 📁 資料夾管理
- 瀏覽和管理KatFile資料夾
- 建立新資料夾
- 選擇上傳目標資料夾

### 🔐 帳戶管理
- API金鑰安全儲存
- 即時顯示帳戶資訊
- 網路連線診斷

## 📋 系統需求

- **Python 3.7+** (建議使用Python 3.8或更高版本)
- **作業系統**: Windows 10+, macOS 10.14+, Linux (Ubuntu 18.04+)
- **網路連線**: 需要穩定的網路連線
- **儲存空間**: 至少100MB可用空間

## 🛠️ 安裝說明

### 方法一：自動安裝（推薦）

1. **下載並解壓縮**程式檔案
2. **雙擊執行** `啟動KatFile上傳工具.bat` (Windows)
3. 程式會自動檢查並安裝所需依賴

### 方法二：手動安裝

1. **安裝Python依賴套件**：
   ```bash
   pip install requests python-docx py7zr
   ```

2. **檢查tkinter是否可用**：
   ```bash
   python -c "import tkinter; print('tkinter可用')"
   ```

3. **如果tkinter缺失**：
   - **Windows**: 重新安裝Python，確保勾選tkinter
   - **Linux**: `sudo apt-get install python3-tk`
   - **macOS**: 使用官方Python安裝包

4. **執行依賴檢查**：
   ```bash
   python install_dependencies.py
   ```

5. **啟動主程式**：
   ```bash
   python katfile_uploader_enhanced.py
   ```

## 🎯 使用指南

### 首次設定

1. **啟動程式**
2. **設定API金鑰**：
   - 在「檔案上傳」頁面輸入您的KatFile API金鑰
   - 點擊「測試」驗證金鑰有效性
   - 點擊「儲存」保存設定

3. **配置壓縮設定**（可選）：
   - 切換到「壓縮設定」頁面
   - 啟用壓縮功能
   - 選擇壓縮格式（ZIP/7Z）
   - 設定解壓密碼

4. **配置Word文件**（可選）：
   - 切換到「文件記錄」頁面
   - 啟用Word文件生成
   - 選擇使用內建範本或自訂範本

### 日常使用

1. **選擇檔案**：
   - 回到「檔案上傳」頁面
   - 點擊「選擇檔案」或「選擇資料夾」
   - 選擇要上傳的檔案

2. **選擇目標資料夾**：
   - 在左側資料夾列表中雙擊選擇目標資料夾
   - 或保持「根目錄」作為上傳位置

3. **開始上傳**：
   - 點擊「開始上傳」
   - 程式會自動處理壓縮、上傳、移動檔案等步驟
   - 查看日誌了解上傳進度

4. **檢查結果**：
   - 上傳完成後，檢查生成的Word記錄文件
   - 複製下載連結進行分享

## 🔧 故障排除

### 常見問題

**Q: 程式無法啟動，顯示「No module named 'tkinter'」**
A: tkinter缺失，請：
- Windows: 重新安裝Python，確保勾選tkinter
- Linux: 執行 `sudo apt-get install python3-tk`
- macOS: 使用官方Python安裝包

**Q: API金鑰測試失敗**
A: 請檢查：
- API金鑰是否正確（無多餘空格）
- 網路連線是否正常
- 點擊「診斷」按鈕檢查網路狀態

**Q: 上傳失敗或連線中斷**
A: 請嘗試：
- 檢查網路連線穩定性
- 使用VPN（如果網路受限）
- 重新測試API金鑰
- 查看詳細錯誤日誌

## 📝 檔案說明

- `katfile_uploader_enhanced.py` - 主程式
- `install_dependencies.py` - 依賴檢查和安裝腳本
- `start_katfile_uploader.py` - 啟動腳本（含錯誤處理）
- `啟動KatFile上傳工具.bat` - Windows一鍵啟動腳本
- `README_完整版.md` - 完整使用說明
