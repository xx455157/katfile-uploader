#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
KatFile 上傳工具依賴安裝腳本
檢查並安裝所需的Python套件
"""

import sys
import subprocess
import importlib

def check_python_version():
    """檢查Python版本"""
    print("🐍 檢查Python版本...")
    version = sys.version_info
    print(f"   Python版本: {version.major}.{version.minor}.{version.micro}")
    
    if version.major < 3 or (version.major == 3 and version.minor < 7):
        print("❌ 需要Python 3.7或更高版本")
        return False
    else:
        print("✅ Python版本符合要求")
        return True

def install_package(package_name, import_name=None):
    """安裝Python套件"""
    if import_name is None:
        import_name = package_name
    
    try:
        importlib.import_module(import_name)
        print(f"✅ {package_name} 已安裝")
        return True
    except ImportError:
        print(f"📦 正在安裝 {package_name}...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package_name])
            print(f"✅ {package_name} 安裝成功")
            return True
        except subprocess.CalledProcessError:
            print(f"❌ {package_name} 安裝失敗")
            return False

def check_tkinter():
    """檢查tkinter是否可用"""
    print("🖼️ 檢查GUI支援...")
    try:
        import tkinter
        print("✅ tkinter 可用")
        return True
    except ImportError:
        print("❌ tkinter 不可用")
        print("   在Windows上，tkinter通常隨Python一起安裝")
        print("   在Linux上，請執行: sudo apt-get install python3-tk")
        print("   在macOS上，請確保使用官方Python安裝包")
        return False

def main():
    """主要檢查流程"""
    print("🔧 KatFile 上傳工具依賴檢查")
    print("=" * 50)
    
    # 檢查Python版本
    if not check_python_version():
        return False
    
    print()
    
    # 檢查tkinter
    if not check_tkinter():
        return False
    
    print()
    
    # 需要安裝的套件
    packages = [
        ("requests", "requests"),
        ("python-docx", "docx"),
        ("py7zr", "py7zr"),
    ]
    
    print("📦 檢查並安裝Python套件...")
    all_success = True
    
    for package_name, import_name in packages:
        if not install_package(package_name, import_name):
            all_success = False
    
    print()
    
    if all_success:
        print("🎉 所有依賴檢查完成！")
        print("✅ 您現在可以執行 katfile_uploader_enhanced.py")
        
        # 測試導入所有模組
        print("\n🧪 測試模組導入...")
        try:
            import tkinter
            import requests
            import docx
            import py7zr
            print("✅ 所有模組導入成功")
            
            # 嘗試啟動GUI測試
            print("\n🖼️ 測試GUI啟動...")
            root = tkinter.Tk()
            root.title("依賴測試成功")
            root.geometry("300x100")
            label = tkinter.Label(root, text="✅ GUI測試成功！\n可以關閉此視窗")
            label.pack(expand=True)
            
            # 自動關閉視窗
            root.after(3000, root.destroy)
            root.mainloop()
            
            print("✅ GUI測試成功")
            
        except Exception as e:
            print(f"❌ 模組測試失敗: {e}")
            all_success = False
    else:
        print("❌ 部分依賴安裝失敗")
        print("請手動安裝失敗的套件，或檢查網路連線")
    
    return all_success

if __name__ == "__main__":
    try:
        success = main()
        if success:
            print("\n🚀 準備就緒！請執行主程式：")
            print("   python katfile_uploader_enhanced.py")
        else:
            print("\n⚠️ 請解決上述問題後再次執行此腳本")
        
        input("\n按Enter鍵退出...")
        
    except KeyboardInterrupt:
        print("\n\n👋 安裝已取消")
    except Exception as e:
        print(f"\n❌ 發生錯誤: {e}")
        input("\n按Enter鍵退出...")
