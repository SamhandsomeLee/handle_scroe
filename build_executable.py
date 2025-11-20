"""
使用 PyInstaller 打包成可执行文件
支持 Windows 和 macOS
"""
import PyInstaller.__main__
import sys
import os

def build_windows():
    """打包 Windows 版本"""
    print("=" * 80)
    print("正在打包 Windows 版本...")
    print("=" * 80)
    
    PyInstaller.__main__.run([
        'score_import_gui.py',
        '--name=成绩导入转换工具',
        '--onefile',
        '--windowed',
        '--icon=icon.ico',  # 需要提供图标文件
        '--add-data=excel_handler.py:.',
        '--collect-all=PySide6',
        '--hidden-import=xlrd',
        '--hidden-import=xlwt',
        '--hidden-import=openpyxl',
        '--distpath=./dist/windows',
        '--buildpath=./build/windows',
        '--specpath=./build/windows',
    ])
    
    print("\n✓ Windows 版本打包完成！")
    print("  输出目录: ./dist/windows/")


def build_macos():
    """打包 macOS 版本"""
    print("=" * 80)
    print("正在打包 macOS 版本...")
    print("=" * 80)
    
    PyInstaller.__main__.run([
        'score_import_gui.py',
        '--name=成绩导入转换工具',
        '--onefile',
        '--windowed',
        '--icon=icon.icns',  # macOS 需要 .icns 格式
        '--add-data=excel_handler.py:.',
        '--collect-all=PySide6',
        '--hidden-import=xlrd',
        '--hidden-import=xlwt',
        '--hidden-import=openpyxl',
        '--distpath=./dist/macos',
        '--buildpath=./build/macos',
        '--specpath=./build/macos',
        '--osx-bundle-identifier=com.example.score-import-tool',
    ])
    
    print("\n✓ macOS 版本打包完成！")
    print("  输出目录: ./dist/macos/")


def build_linux():
    """打包 Linux 版本"""
    print("=" * 80)
    print("正在打包 Linux 版本...")
    print("=" * 80)
    
    PyInstaller.__main__.run([
        'score_import_gui.py',
        '--name=score-import-tool',
        '--onefile',
        '--add-data=excel_handler.py:.',
        '--collect-all=PySide6',
        '--hidden-import=xlrd',
        '--hidden-import=xlwt',
        '--hidden-import=openpyxl',
        '--distpath=./dist/linux',
        '--buildpath=./build/linux',
        '--specpath=./build/linux',
    ])
    
    print("\n✓ Linux 版本打包完成！")
    print("  输出目录: ./dist/linux/")


if __name__ == "__main__":
    if len(sys.argv) > 1:
        platform = sys.argv[1].lower()
        if platform == "windows":
            build_windows()
        elif platform == "macos":
            build_macos()
        elif platform == "linux":
            build_linux()
        else:
            print(f"未知平台: {platform}")
            print("支持的平台: windows, macos, linux")
    else:
        # 检测当前系统并打包
        if sys.platform == "win32":
            build_windows()
        elif sys.platform == "darwin":
            build_macos()
        elif sys.platform == "linux":
            build_linux()
        else:
            print(f"不支持的系统: {sys.platform}")
