#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
KatFile 上傳工具啟動腳本
包含完整的錯誤處理和診斷功能
"""

import sys
import os
import traceback

def check_dependencies():
    """檢查依賴套件"""
    missing_packages = []
    
    try:
        import tkinter
    except ImportError:
        missing_packages.append("tkinter (GUI支援)")
    
    try:
        import requests
    except ImportError:
        missing_packages.append("requests")
    
    try:
        import docx
    except ImportError:
        missing_packages.append("python-docx")
    
    try:
        import py7zr
    except ImportError:
        missing_packages.append("py7zr")
    
    return missing_packages

def show_error_dialog(title, message):
    """顯示錯誤對話框"""
    try:
        import tkinter as tk
        from tkinter import messagebox
        
        root = tk.Tk()
        root.withdraw()  # 隱藏主視窗
        messagebox.showerror(title, message)
        root.destroy()
    except:
        # 如果無法顯示GUI對話框，使用控制台輸出
        print(f"\n❌ {title}")
        print(f"   {message}")
        input("\n按Enter鍵退出...")

def main():
    """主要啟動流程"""
    try:
        print("🚀 啟動 KatFile 上傳工具...")
        
        # 檢查依賴
        missing = check_dependencies()
        if missing:
            error_msg = f"""缺少必要的套件：
{chr(10).join(f'• {pkg}' for pkg in missing)}

請執行以下命令安裝：
pip install requests python-docx py7zr

如果tkinter缺失：
• Windows: 重新安裝Python，確保勾選tkinter
• Linux: sudo apt-get install python3-tk
• macOS: 使用官方Python安裝包

或者執行 install_dependencies.py 自動安裝"""
            
            show_error_dialog("依賴套件缺失", error_msg)
            return False
        
        # 檢查主程式檔案
        main_script = "katfile_uploader_enhanced.py"
        if not os.path.exists(main_script):
            error_msg = f"找不到主程式檔案: {main_script}\n請確保檔案在同一目錄中"
            show_error_dialog("檔案缺失", error_msg)
            return False
        
        # 啟動主程式
        print("✅ 依賴檢查通過，啟動主程式...")
        
        # 導入並執行主程式
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        
        try:
            from katfile_uploader_enhanced import main as main_app
            main_app()
            return True
            
        except ImportError as e:
            error_msg = f"導入主程式失敗: {str(e)}\n請檢查 katfile_uploader_enhanced.py 檔案是否完整"
            show_error_dialog("導入錯誤", error_msg)
            return False
            
    except Exception as e:
        error_msg = f"啟動失敗: {str(e)}\n\n詳細錯誤:\n{traceback.format_exc()}"
        show_error_dialog("啟動錯誤", error_msg)
        return False

if __name__ == "__main__":
    try:
        success = main()
        if not success:
            print("\n❌ 啟動失敗")
            print("請執行 install_dependencies.py 檢查依賴")
            
    except KeyboardInterrupt:
        print("\n\n👋 程式已取消")
    except Exception as e:
        print(f"\n❌ 發生未預期的錯誤: {e}")
        print("\n詳細錯誤資訊:")
        traceback.print_exc()
        input("\n按Enter鍵退出...")
