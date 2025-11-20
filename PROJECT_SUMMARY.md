# 成绩导入转换工具 - 项目总结

## 项目概述

基于 PySide6 的现代化 GUI 工具，用于将包含两次考试数据的 Excel 文件转换为标准的8工作表格式。

**开发日期**: 2025-11-20

## 核心功能

### 输入处理
- 读取包含 `高二一调` 和 `高二期中` 两个工作表的 Excel 文件
- 自动识别不同的表头格式（单层/双层）
- 保持原始数据完整性

### 输出生成
生成包含8个工作表的标准格式文件：

1. **历次成绩原始** - 合并两次考试的所有数据
2. **历次成绩打印** - 合并两次考试的所有数据
3. **尖子生成绩** - 排名前20的优秀学生
4. **高一期末** - 空表（仅表头）
5. **高二一调** - 原始数据直接复制
6. **高二期中** - 原始数据直接复制
7. **均值 姓名排序** - 两次排名均值，按姓名排序
8. **两次 均值 升序** - 两次排名均值，按均值升序

## 技术架构

### 文件结构

```
handle_score/
├── score_import_gui.py          # 主程序（GUI界面）
├── excel_handler.py             # 核心处理模块
├── requirements.txt             # 依赖列表
├── README.md                    # 使用文档
├── PROJECT_SUMMARY.md           # 本文件
├── test_conversion.py           # 测试脚本
└── data/                        # 测试数据目录
    ├── 22历次成绩.xls           # 目标格式参考
    ├── 22高二一调.xls           # 测试输入
    └── 22高二期中.xls           # 测试输入
```

### 核心模块

**excel_handler.py**
- `ExamSheet` - 工作表数据封装
- `ExcelHandler` - 文件处理主类
  - `read_excel_file()` - 读取Excel文件
  - `write_excel_file()` - 生成8工作表
  - `get_file_info()` - 获取文件信息
  - 私有方法处理各类工作表生成

**score_import_gui.py**
- `MainWindow` - 主窗口类
- `ProcessThread` - 后台处理线程
- 现代化UI设计（Fusion风格）

## 数据处理流程

### 1. 读取阶段
```
输入文件 (2个工作表)
    ↓
xlrd 读取所有行列数据
    ↓
ExamSheet 对象封装
    ↓
学生数据解析
```

### 2. 转换阶段
```
高二一调 + 高二期中
    ↓
├─→ 合并 → 历次成绩原始/打印
├─→ 筛选 → 尖子生成绩
├─→ 复制 → 高二一调/期中
├─→ 计算 → 均值排序
└─→ 排序 → 均值升序
```

### 3. 输出阶段
```
8个工作表
    ↓
xlwt 写入Excel
    ↓
输出文件 (22历次成绩.xls格式)
```

## 关键设计决策

### 1. 表头识别
- 自动检测第2行是否包含"得分"或"校次"标签
- 支持混合格式（某些工作表有标签行，某些没有）

### 2. 数据合并
- 保持原始表头和标签行
- 先写高二一调，再写高二期中
- 避免重复的表头行

### 3. 尖子生筛选
- 基于总校次排名（第16列）
- 取排名最小的前20名
- 保持原始行格式

### 4. 均值计算
- 使用排名值而非分数
- 公式：(高二一调排名 + 高二期中排名) / 2
- 支持两种排序方式

## 测试结果

### 转换测试
```
输入: 22高二期中.xls (2个工作表, 56行)
输出: test_output.xls (8个工作表)

工作表验证:
✓ 历次成绩原始: 111行 × 16列
✓ 历次成绩打印: 111行 × 16列
✓ 尖子生成绩: 22行 × 16列
✓ 高一期末: 2行 × 16列
✓ 高二一调: 56行 × 16列
✓ 高二期中: 56行 × 16列
✓ 均值 姓名排序: 56行 × 6列
✓ 两次 均值 升序: 56行 × 6列
```

## 已知限制

1. **格式限制**
   - 仅支持 `.xls` 格式（xlwt库限制）
   - 不支持 `.xlsx` 格式的直接写入

2. **功能限制**
   - 尖子生筛选固定为前20名
   - 高一期末工作表为空表
   - 不支持自定义表头

3. **性能限制**
   - 大文件处理速度待优化
   - 单线程处理

## 未来改进方向

- [ ] 支持 `.xlsx` 格式
- [ ] 可配置的尖子生数量
- [ ] 数据验证和错误检查
- [ ] 批量文件处理
- [ ] 导出报表功能
- [ ] 数据统计分析
- [ ] 多线程处理优化

## 使用指南

### 快速开始

```bash
# 1. 安装依赖
pip install -r requirements.txt

# 2. 启动程序
python score_import_gui.py

# 3. 选择输入文件
# 4. 设置输出路径
# 5. 点击"开始转换"
```

### 命令行使用

```python
from excel_handler import ExcelHandler

# 读取文件
sheets = ExcelHandler.read_excel_file("input.xls")

# 生成输出
ExcelHandler.write_excel_file(sheets, "output.xls")

# 获取文件信息
info = ExcelHandler.get_file_info("input.xls")
```

## 代码质量

### 代码规范
- 遵循 PEP 8 风格指南
- 完整的类型注解
- 详细的文档字符串
- 异常处理完善

### 测试覆盖
- 单元测试脚本 (`test_conversion.py`)
- 实际数据验证
- 输出格式验证

## 依赖管理

```
PySide6>=6.5.0      # GUI框架
xlrd>=2.0.1         # Excel读取
xlwt>=1.3.0         # Excel写入
openpyxl>=3.1.0     # 备用库
```

## 许可证

本项目仅供学习和内部使用。

## 开发者备注

### 关键代码片段

**表头检测**
```python
def detect_header_structure(sheet):
    row2_has_labels = any(
        str(sheet.cell(1, col_idx).value) in ['得分', '校次']
        for col_idx in range(min(sheet.ncols, 10))
    )
    return 2 if row2_has_labels else 1
```

**数据合并**
```python
# 先写高二一调
for row in sheet1.data[2:]:
    write_row(row)

# 再写高二期中（跳过表头）
for row in sheet2.data[1:]:
    write_row(row)
```

**均值计算**
```python
avg = (float(rank1) + float(rank2)) / 2
```

## 联系方式

如有问题或建议，请反馈。

---

**项目状态**: ✅ 完成  
**最后更新**: 2025-11-20  
**版本**: v1.0
