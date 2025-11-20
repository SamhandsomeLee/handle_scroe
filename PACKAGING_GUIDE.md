# 打包指南 - 跨平台发布

本指南说明如何将成绩导入转换工具打包成可在 Windows、macOS 和 Linux 上运行的独立应用程序。

## 目录

1. [环境准备](#环境准备)
2. [打包方案](#打包方案)
3. [Windows 打包](#windows-打包)
4. [macOS 打包](#macos-打包)
5. [Linux 打包](#linux-打包)
6. [分发方式](#分发方式)

---

## 环境准备

### 通用要求

```bash
# 1. 确保 Python 版本 >= 3.8
python --version

# 2. 升级 pip
python -m pip install --upgrade pip

# 3. 安装打包工具
pip install PyInstaller
pip install setuptools wheel

# 4. 安装项目依赖
pip install -r requirements.txt
```

### 平台特定要求

**Windows**
- Visual C++ Build Tools（可选，某些库需要编译）
- 图标文件：`icon.ico`

**macOS**
- Xcode Command Line Tools
- 图标文件：`icon.icns`
- 代码签名证书（可选，用于发布）

**Linux**
- GCC 编译工具
- 开发库：`libgl1-mesa-glx`, `libxkbcommon-x11-0`

---

## 打包方案

### 方案对比

| 方案 | 优点 | 缺点 | 适用场景 |
|------|------|------|---------|
| **PyInstaller** | 简单易用，支持多平台 | 文件较大 | 推荐用于独立应用 |
| **cx_Freeze** | 跨平台支持好 | 配置复杂 | 复杂应用 |
| **py2exe/py2app** | 平台优化好 | 仅支持单一平台 | 特定平台发布 |
| **Nuitka** | 性能最优 | 编译时间长 | 性能关键应用 |

**推荐**: 使用 **PyInstaller**（已配置）

---

## Windows 打包

### 方法1：使用自动脚本（推荐）

```bash
# 打包当前系统版本
python build_executable.py

# 或指定平台
python build_executable.py windows
```

### 方法2：手动 PyInstaller 命令

```bash
pyinstaller ^
  score_import_gui.py ^
  --name="成绩导入转换工具" ^
  --onefile ^
  --windowed ^
  --icon=icon.ico ^
  --add-data "excel_handler.py:." ^
  --collect-all=PySide6 ^
  --hidden-import=xlrd ^
  --hidden-import=xlwt ^
  --hidden-import=openpyxl
```

### 输出

```
dist/
└── 成绩导入转换工具.exe  (单个可执行文件)
```

### 创建安装程序（可选）

使用 NSIS 或 Inno Setup 创建 `.msi` 安装程序：

```bash
# 安装 NSIS
# 下载: https://nsis.sourceforge.io/

# 创建 installer.nsi 脚本
# 编译生成 .exe 安装程序
makensis installer.nsi
```

---

## macOS 打包

### 前置条件

```bash
# 1. 安装 Xcode Command Line Tools
xcode-select --install

# 2. 安装依赖
pip install -r requirements.txt
pip install PyInstaller
```

### 方法1：使用自动脚本（推荐）

```bash
python build_executable.py macos
```

### 方法2：手动 PyInstaller 命令

```bash
pyinstaller \
  score_import_gui.py \
  --name="成绩导入转换工具" \
  --onefile \
  --windowed \
  --icon=icon.icns \
  --add-data "excel_handler.py:." \
  --collect-all=PySide6 \
  --hidden-import=xlrd \
  --hidden-import=xlwt \
  --hidden-import=openpyxl \
  --osx-bundle-identifier=com.example.score-import-tool
```

### 输出

```
dist/
└── 成绩导入转换工具.app/  (macOS 应用包)
    └── Contents/
        ├── MacOS/
        │   └── 成绩导入转换工具  (可执行文件)
        ├── Resources/
        ├── Info.plist
        └── ...
```

### 代码签名（可选，用于发布）

```bash
# 签名应用
codesign --deep --force --verify --verbose \
  --sign "Developer ID Application" \
  dist/成绩导入转换工具.app

# 验证签名
codesign --verify --verbose=4 \
  dist/成绩导入转换工具.app

# 公证（Apple 要求）
xcrun altool --notarize-app \
  -f dist/成绩导入转换工具.dmg \
  -t osx \
  --primary-bundle-id com.example.score-import-tool \
  -u your-apple-id@example.com \
  -p your-app-specific-password
```

### 创建 DMG 安装程序

```bash
# 使用 create-dmg 工具
npm install -g create-dmg

create-dmg \
  'dist/成绩导入转换工具.app' \
  dist/ \
  --overwrite
```

---

## Linux 打包

### 前置条件

```bash
# Ubuntu/Debian
sudo apt-get install -y \
  python3-dev \
  libgl1-mesa-glx \
  libxkbcommon-x11-0 \
  libdbus-1-3

# Fedora/RHEL
sudo dnf install -y \
  python3-devel \
  mesa-libGL \
  libxkbcommon-x11

# 安装依赖
pip install -r requirements.txt
pip install PyInstaller
```

### 方法1：使用自动脚本

```bash
python build_executable.py linux
```

### 方法2：手动 PyInstaller 命令

```bash
pyinstaller \
  score_import_gui.py \
  --name=score-import-tool \
  --onefile \
  --add-data "excel_handler.py:." \
  --collect-all=PySide6 \
  --hidden-import=xlrd \
  --hidden-import=xlwt \
  --hidden-import=openpyxl
```

### 输出

```
dist/
└── score-import-tool  (可执行文件)
```

### 创建 AppImage（可选）

```bash
# 安装 appimagetool
wget https://github.com/AppImage/AppImageKit/releases/download/continuous/appimagetool-x86_64.AppImage
chmod +x appimagetool-x86_64.AppImage

# 创建 AppImage
./appimagetool-x86_64.AppImage dist/score-import-tool dist/score-import-tool.AppImage
```

---

## 分发方式

### 1. GitHub Releases

```bash
# 创建 GitHub 仓库
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/yourusername/score-import-tool.git
git push -u origin main

# 创建 Release
# 在 GitHub 网页上：
# 1. 点击 "Releases"
# 2. 点击 "Create a new release"
# 3. 上传打包好的文件
# 4. 发布
```

### 2. 直接下载链接

创建 `DOWNLOAD.md` 文件：

```markdown
# 下载

## Windows
- [成绩导入转换工具.exe](https://example.com/downloads/windows/score-import-tool.exe)

## macOS
- [成绩导入转换工具.dmg](https://example.com/downloads/macos/score-import-tool.dmg)

## Linux
- [score-import-tool](https://example.com/downloads/linux/score-import-tool)
- [score-import-tool.AppImage](https://example.com/downloads/linux/score-import-tool.AppImage)
```

### 3. 包管理器

**Windows - Chocolatey**
```bash
# 创建 chocolatey 包
choco new score-import-tool
# 编辑 tools/chocolateyinstall.ps1
# 提交到 Chocolatey Community Repository
```

**macOS - Homebrew**
```bash
# 创建 Homebrew Formula
brew create https://github.com/yourusername/score-import-tool/releases/download/v1.0.0/score-import-tool.tar.gz
# 编辑 formula 文件
# 提交 Pull Request 到 homebrew-core
```

**Linux - Snap**
```bash
# 安装 snapcraft
sudo apt install snapcraft

# 创建 snapcraft.yaml
# 构建 snap
snapcraft

# 发布到 Snap Store
snapcraft upload score-import-tool_1.0.0_amd64.snap --release=stable
```

---

## 常见问题

### Q1: 打包后文件太大？

**解决方案**：
```bash
# 使用 UPX 压缩
pip install upx
pyinstaller ... --upx-dir=/path/to/upx

# 或删除不必要的文件
# 在 spec 文件中配置
```

### Q2: macOS 提示"无法验证开发者"？

**解决方案**：
```bash
# 方法1：右键打开，选择"打开"
# 方法2：执行以下命令
sudo xattr -rd com.apple.quarantine /Applications/成绩导入转换工具.app

# 方法3：进行代码签名和公证
```

### Q3: Windows 打包后无法运行？

**解决方案**：
- 检查是否缺少 Visual C++ Runtime
- 下载并安装：https://support.microsoft.com/en-us/help/2977003
- 或在打包时包含运行时库

### Q4: Linux 上缺少依赖库？

**解决方案**：
```bash
# 检查缺少的库
ldd ./dist/score-import-tool

# 安装缺少的库
sudo apt-get install <library-name>
```

### Q5: 如何自动更新？

**推荐方案**：
- 使用 `PyUpdater` 库
- 或集成 `Sparkle`（macOS）
- 或 `WinSparkle`（Windows）

---

## 完整工作流

### 开发阶段
```bash
# 1. 开发和测试
python score_import_gui.py

# 2. 运行测试
python test_conversion.py
```

### 发布阶段
```bash
# 1. 更新版本号
# 编辑 setup.py 和 score_import_gui.py 中的版本

# 2. 打包
python build_executable.py windows
python build_executable.py macos
python build_executable.py linux

# 3. 测试打包结果
./dist/windows/成绩导入转换工具.exe
./dist/macos/成绩导入转换工具.app
./dist/linux/score-import-tool

# 4. 创建 GitHub Release
git tag v1.0.0
git push origin v1.0.0

# 5. 上传文件到 Release
# 在 GitHub 网页上操作
```

---

## 文件清单

打包完成后的文件结构：

```
dist/
├── windows/
│   └── 成绩导入转换工具.exe
├── macos/
│   └── 成绩导入转换工具.app/
└── linux/
    └── score-import-tool

build/
├── windows/
├── macos/
└── linux/

# 源代码保持不变
score_import_gui.py
excel_handler.py
requirements.txt
...
```

---

## 性能优化建议

### 启动时间优化
```python
# 在 score_import_gui.py 中
import sys
import os

# 延迟导入
def main():
    from PySide6.QtWidgets import QApplication
    from score_import_gui import MainWindow
    
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
```

### 文件大小优化
- 使用 `--onefile` 创建单个可执行文件
- 移除不必要的库
- 使用 UPX 压缩

---

## 安全建议

1. **代码签名**
   - Windows: 使用代码签名证书
   - macOS: 进行公证
   - Linux: 提供 GPG 签名

2. **病毒扫描**
   - 使用 VirusTotal 扫描发布前的文件

3. **依赖安全**
   - 定期更新依赖库
   - 检查已知漏洞

---

## 参考资源

- [PyInstaller 官方文档](https://pyinstaller.readthedocs.io/)
- [PySide6 部署指南](https://doc.qt.io/qtforpython/deployment.html)
- [GitHub Releases 文档](https://docs.github.com/en/repositories/releasing-projects-on-github/about-releases)
- [Homebrew 贡献指南](https://docs.brew.sh/Formula-Cookbook)

---

**最后更新**: 2025-11-20  
**版本**: v1.0
