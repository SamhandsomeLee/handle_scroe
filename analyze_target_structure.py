"""
分析目标输出格式（22历次成绩.xls中的高二期中工作表）
"""
import xlrd
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def main():
    file_path = r"d:\handle_score\data\22历次成绩.xls"
    
    workbook = xlrd.open_workbook(file_path)
    sheet = workbook.sheet_by_name("高二期中")
    
    print("=" * 120)
    print(f"目标格式 - 工作表: {sheet.name}")
    print(f"行数: {sheet.nrows}, 列数: {sheet.ncols}")
    print("=" * 120)
    
    # 显示前5行的所有列
    print("\n前5行完整数据:")
    print("-" * 120)
    
    for row_idx in range(min(5, sheet.nrows)):
        row_data = []
        for col_idx in range(sheet.ncols):
            cell = sheet.cell(row_idx, col_idx)
            if cell.ctype == xlrd.XL_CELL_NUMBER:
                val = cell.value
                if val == int(val):
                    row_data.append(str(int(val)))
                else:
                    row_data.append(str(val))
            elif cell.ctype == xlrd.XL_CELL_EMPTY:
                row_data.append("[空]")
            else:
                row_data.append(str(cell.value))
        
        print(f"行{row_idx+1}: {' | '.join(row_data)}")
    
    # 分析表头
    print("\n" + "=" * 120)
    print("表头分析（前3行）:")
    print("=" * 120)
    
    for row_idx in range(min(3, sheet.nrows)):
        print(f"\n行{row_idx+1}:")
        for col_idx in range(sheet.ncols):
            cell = sheet.cell(row_idx, col_idx)
            if cell.ctype == xlrd.XL_CELL_NUMBER:
                val = cell.value
                if val == int(val):
                    val = int(val)
            elif cell.ctype == xlrd.XL_CELL_EMPTY:
                val = "[空]"
            else:
                val = str(cell.value)
            
            print(f"  列{col_idx+1}: {val}")

if __name__ == "__main__":
    main()
