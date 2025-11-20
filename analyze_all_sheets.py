"""
分析 22历次成绩.xls 的所有8个工作表结构
"""
import xlrd
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def main():
    file_path = r"d:\handle_score\data\22历次成绩.xls"
    
    workbook = xlrd.open_workbook(file_path)
    
    print("=" * 100)
    print(f"文件: 22历次成绩.xls")
    print(f"总工作表数: {workbook.nsheets}")
    print("=" * 100)
    
    for sheet_idx in range(workbook.nsheets):
        sheet = workbook.sheet_by_index(sheet_idx)
        
        print(f"\n工作表 {sheet_idx + 1}: {sheet.name}")
        print(f"  行数: {sheet.nrows}, 列数: {sheet.ncols}")
        
        # 显示前3行
        print(f"  前3行数据:")
        for row_idx in range(min(3, sheet.nrows)):
            row_data = []
            for col_idx in range(min(8, sheet.ncols)):  # 只显示前8列
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
            
            print(f"    行{row_idx+1}: {' | '.join(row_data)}")

if __name__ == "__main__":
    main()
