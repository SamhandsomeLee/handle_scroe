#!/bin/bash
# 快速打包脚本 - 支持 macOS 和 Linux

set -e

echo "======================================================================"
echo "成绩导入转换工具 - 快速打包脚本"
echo "======================================================================"

# 检测系统
if [[ "$OSTYPE" == "darwin"* ]]; then
    PLATFORM="macos"
    EXECUTABLE_NAME="成绩导入转换工具"
elif [[ "$OSTYPE" == "linux-gnu"* ]]; then
    PLATFORM="linux"
    EXECUTABLE_NAME="score-import-tool"
else
    echo "不支持的系统: $OSTYPE"
    exit 1
fi

echo "检测到系统: $PLATFORM"
echo ""

# 检查 Python
if ! command -v python3 &> /dev/null; then
    echo "错误: 未找到 Python 3"
    exit 1
fi

PYTHON_VERSION=$(python3 --version | cut -d' ' -f2)
echo "Python 版本: $PYTHON_VERSION"

# 检查依赖
echo ""
echo "检查依赖..."
pip3 install --upgrade pip setuptools wheel

# 安装项目依赖
echo "安装项目依赖..."
pip3 install -r requirements.txt

# 安装打包工具
echo "安装打包工具..."
pip3 install PyInstaller

# 清理旧的构建
echo ""
echo "清理旧的构建文件..."
rm -rf build dist *.spec

# 打包
echo ""
echo "开始打包 $PLATFORM 版本..."
echo ""

if [ "$PLATFORM" = "macos" ]; then
    pyinstaller \
        score_import_gui.py \
        --name="$EXECUTABLE_NAME" \
        --onefile \
        --windowed \
        --add-data "excel_handler.py:." \
        --collect-all=PySide6 \
        --hidden-import=xlrd \
        --hidden-import=xlwt \
        --hidden-import=openpyxl \
        --osx-bundle-identifier=com.example.score-import-tool \
        --distpath=./dist/macos \
        --buildpath=./build/macos \
        --specpath=./build/macos
    
    echo ""
    echo "======================================================================"
    echo "✓ macOS 版本打包完成！"
    echo "======================================================================"
    echo ""
    echo "应用位置: dist/macos/$EXECUTABLE_NAME.app"
    echo ""
    echo "使用方法:"
    echo "  1. 双击打开应用"
    echo "  2. 或在终端运行: open dist/macos/$EXECUTABLE_NAME.app"
    echo ""
    echo "如果提示'无法验证开发者'，请运行:"
    echo "  sudo xattr -rd com.apple.quarantine dist/macos/$EXECUTABLE_NAME.app"
    echo ""
    
elif [ "$PLATFORM" = "linux" ]; then
    pyinstaller \
        score_import_gui.py \
        --name="$EXECUTABLE_NAME" \
        --onefile \
        --add-data "excel_handler.py:." \
        --collect-all=PySide6 \
        --hidden-import=xlrd \
        --hidden-import=xlwt \
        --hidden-import=openpyxl \
        --distpath=./dist/linux \
        --buildpath=./build/linux \
        --specpath=./build/linux
    
    # 添加执行权限
    chmod +x dist/linux/$EXECUTABLE_NAME
    
    echo ""
    echo "======================================================================"
    echo "✓ Linux 版本打包完成！"
    echo "======================================================================"
    echo ""
    echo "应用位置: dist/linux/$EXECUTABLE_NAME"
    echo ""
    echo "使用方法:"
    echo "  1. 在终端运行: ./dist/linux/$EXECUTABLE_NAME"
    echo "  2. 或双击打开（如果文件管理器支持）"
    echo ""
fi

echo "打包文件大小:"
du -sh dist/*/

echo ""
echo "完成！"
