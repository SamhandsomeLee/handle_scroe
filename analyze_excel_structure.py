"""
分析 Excel 表格结构工具
"""
import pandas as pd
import xlrd
import sys
import io

# 设置标准输出为 UTF-8 编码
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def analyze_excel_structure(file_path):
    """分析 Excel 文件的结构"""
    print(f"正在分析文件: {file_path}\n")
    print("=" * 80)
    
    # 使用 xlrd 读取 .xls 文件
    workbook = xlrd.open_workbook(file_path, formatting_info=True)
    
    print(f"工作簿包含 {workbook.nsheets} 个工作表\n")
    
    for sheet_idx in range(workbook.nsheets):
        sheet = workbook.sheet_by_index(sheet_idx)
        print(f"\n{'=' * 80}")
        print(f"工作表 {sheet_idx + 1}: {sheet.name}")
        print(f"{'=' * 80}")
        print(f"行数: {sheet.nrows}")
        print(f"列数: {sheet.ncols}")
        print()
        
        # 显示前10行数据
        print("前10行数据预览:")
        print("-" * 80)
        
        max_rows = min(10, sheet.nrows)
        for row_idx in range(max_rows):
            row_data = []
            for col_idx in range(sheet.ncols):
                cell = sheet.cell(row_idx, col_idx)
                cell_value = cell.value
                
                # 处理不同类型的单元格
                if cell.ctype == xlrd.XL_CELL_DATE:
                    cell_value = xlrd.xldate_as_datetime(cell.value, workbook.datemode)
                elif cell.ctype == xlrd.XL_CELL_NUMBER:
                    # 如果是整数，显示为整数
                    if cell_value == int(cell_value):
                        cell_value = int(cell_value)
                
                row_data.append(str(cell_value))
            
            print(f"行 {row_idx + 1}: {' | '.join(row_data)}")
        
        # 分析列结构
        if sheet.nrows > 0:
            print(f"\n{'-' * 80}")
            print("列结构分析:")
            print(f"{'-' * 80}")
            
            # 假设第一行是表头
            headers = []
            for col_idx in range(sheet.ncols):
                header = sheet.cell(0, col_idx).value
                headers.append(header)
                
                # 分析该列的数据类型
                col_types = set()
                sample_values = []
                for row_idx in range(1, min(sheet.nrows, 20)):  # 取前20行样本
                    cell = sheet.cell(row_idx, col_idx)
                    if cell.ctype == xlrd.XL_CELL_EMPTY:
                        col_types.add("空值")
                    elif cell.ctype == xlrd.XL_CELL_TEXT:
                        col_types.add("文本")
                        if len(sample_values) < 3:
                            sample_values.append(str(cell.value))
                    elif cell.ctype == xlrd.XL_CELL_NUMBER:
                        col_types.add("数字")
                        if len(sample_values) < 3:
                            sample_values.append(str(cell.value))
                    elif cell.ctype == xlrd.XL_CELL_DATE:
                        col_types.add("日期")
                    elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
                        col_types.add("布尔值")
                
                print(f"列 {col_idx + 1}: {header}")
                print(f"  数据类型: {', '.join(col_types)}")
                if sample_values:
                    print(f"  样本值: {', '.join(sample_values[:3])}")
                print()

if __name__ == "__main__":
    file_path = r"d:\handle_score\data\22历次成绩.xls"
    try:
        analyze_excel_structure(file_path)
    except Exception as e:
        print(f"错误: {e}")
        import traceback
        traceback.print_exc()
