"""
分析单独的两个 Excel 文件
"""
import xlrd
import sys
import io

# 设置标准输出为 UTF-8 编码
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def analyze_single_excel(file_path):
    """分析单个 Excel 文件"""
    print(f"\n{'#' * 80}")
    print(f"文件: {file_path}")
    print(f"{'#' * 80}\n")
    
    workbook = xlrd.open_workbook(file_path, formatting_info=True)
    print(f"工作表数量: {workbook.nsheets}\n")
    
    for sheet_idx in range(workbook.nsheets):
        sheet = workbook.sheet_by_index(sheet_idx)
        print(f"\n{'=' * 80}")
        print(f"工作表 {sheet_idx + 1}: {sheet.name}")
        print(f"{'=' * 80}")
        print(f"行数: {sheet.nrows}")
        print(f"列数: {sheet.ncols}")
        print()
        
        # 显示前15行数据
        print("前15行数据预览:")
        print("-" * 80)
        
        max_rows = min(15, sheet.nrows)
        for row_idx in range(max_rows):
            row_data = []
            for col_idx in range(sheet.ncols):
                cell = sheet.cell(row_idx, col_idx)
                cell_value = cell.value
                
                if cell.ctype == xlrd.XL_CELL_DATE:
                    cell_value = xlrd.xldate_as_datetime(cell.value, workbook.datemode)
                elif cell.ctype == xlrd.XL_CELL_NUMBER:
                    if cell_value == int(cell_value):
                        cell_value = int(cell_value)
                elif cell.ctype == xlrd.XL_CELL_EMPTY:
                    cell_value = ""
                
                row_data.append(str(cell_value))
            
            print(f"行 {row_idx + 1}: {' | '.join(row_data)}")
        
        # 列结构分析
        if sheet.nrows > 0:
            print(f"\n{'-' * 80}")
            print("列结构详细分析:")
            print(f"{'-' * 80}")
            
            for col_idx in range(sheet.ncols):
                # 收集该列前3行的值（通常是表头）
                header_values = []
                for row_idx in range(min(3, sheet.nrows)):
                    cell = sheet.cell(row_idx, col_idx)
                    if cell.ctype != xlrd.XL_CELL_EMPTY:
                        header_values.append(str(cell.value))
                    else:
                        header_values.append("(空)")
                
                # 分析数据类型和样本
                col_types = set()
                sample_values = []
                empty_count = 0
                
                for row_idx in range(3, min(sheet.nrows, 10)):  # 从第4行开始取样本
                    cell = sheet.cell(row_idx, col_idx)
                    if cell.ctype == xlrd.XL_CELL_EMPTY:
                        empty_count += 1
                    elif cell.ctype == xlrd.XL_CELL_TEXT:
                        col_types.add("文本")
                        if len(sample_values) < 3:
                            sample_values.append(str(cell.value))
                    elif cell.ctype == xlrd.XL_CELL_NUMBER:
                        col_types.add("数字")
                        if len(sample_values) < 3:
                            val = cell.value
                            if val == int(val):
                                sample_values.append(str(int(val)))
                            else:
                                sample_values.append(str(val))
                    elif cell.ctype == xlrd.XL_CELL_DATE:
                        col_types.add("日期")
                    elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
                        col_types.add("布尔值")
                
                if empty_count > 0:
                    col_types.add(f"空值({empty_count})")
                
                print(f"\n列 {col_idx + 1}:")
                print(f"  前3行: {' → '.join(header_values)}")
                print(f"  数据类型: {', '.join(col_types) if col_types else '全空'}")
                if sample_values:
                    print(f"  样本值(第4-10行): {', '.join(sample_values)}")

if __name__ == "__main__":
    files = [
        r"d:\handle_score\data\22高二一调.xls",
        r"d:\handle_score\data\22高二期中.xls"
    ]
    
    for file_path in files:
        try:
            analyze_single_excel(file_path)
        except Exception as e:
            print(f"\n错误处理文件 {file_path}: {e}")
            import traceback
            traceback.print_exc()
    
    print("\n\n" + "=" * 80)
    print("分析完成")
    print("=" * 80)
