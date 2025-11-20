"""
详细对比高二期中数据的差异
"""
import xlrd
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def main():
    summary_file = r"d:\handle_score\data\22历次成绩.xls"
    qizhong_file = r"d:\handle_score\data\22高二期中.xls"
    
    wb_summary = xlrd.open_workbook(summary_file)
    wb_qizhong = xlrd.open_workbook(qizhong_file)
    
    sheet_summary = wb_summary.sheet_by_name("高二期中")
    sheet_qizhong = wb_qizhong.sheet_by_index(0)  # 第一个工作表
    
    print("=" * 80)
    print("高二期中数据详细对比")
    print("=" * 80)
    
    print(f"\n汇总文件工作表: {sheet_summary.name}")
    print(f"独立文件工作表: {sheet_qizhong.name}")
    print()
    
    # 显示汇总文件的前5行
    print("汇总文件 - 前5行:")
    print("-" * 80)
    for row_idx in range(min(5, sheet_summary.nrows)):
        row = []
        for col_idx in range(min(8, sheet_summary.ncols)):  # 只显示前8列
            cell = sheet_summary.cell(row_idx, col_idx)
            if cell.ctype == xlrd.XL_CELL_NUMBER:
                val = cell.value
                if val == int(val):
                    row.append(str(int(val)))
                else:
                    row.append(str(val))
            elif cell.ctype == xlrd.XL_CELL_EMPTY:
                row.append("(空)")
            else:
                row.append(str(cell.value))
        print(f"行{row_idx+1}: {' | '.join(row)}")
    
    print("\n独立文件 - 前5行:")
    print("-" * 80)
    for row_idx in range(min(5, sheet_qizhong.nrows)):
        row = []
        for col_idx in range(min(8, sheet_qizhong.ncols)):
            cell = sheet_qizhong.cell(row_idx, col_idx)
            if cell.ctype == xlrd.XL_CELL_NUMBER:
                val = cell.value
                if val == int(val):
                    row.append(str(int(val)))
                else:
                    row.append(str(val))
            elif cell.ctype == xlrd.XL_CELL_EMPTY:
                row.append("(空)")
            else:
                row.append(str(cell.value))
        print(f"行{row_idx+1}: {' | '.join(row)}")
    
    # 分析差异
    print("\n" + "=" * 80)
    print("差异分析")
    print("=" * 80)
    
    # 检查是否第二个工作表才是高二期中
    if wb_qizhong.nsheets > 1:
        print(f"\n注意: 独立文件包含 {wb_qizhong.nsheets} 个工作表:")
        for i in range(wb_qizhong.nsheets):
            sheet = wb_qizhong.sheet_by_index(i)
            print(f"  工作表{i+1}: {sheet.name} ({sheet.nrows}行 × {sheet.ncols}列)")
        
        # 尝试对比第二个工作表
        if wb_qizhong.nsheets >= 2:
            print("\n尝试对比第2个工作表...")
            sheet_qizhong2 = wb_qizhong.sheet_by_index(1)
            print(f"\n独立文件 - 第2个工作表'{sheet_qizhong2.name}' - 前5行:")
            print("-" * 80)
            for row_idx in range(min(5, sheet_qizhong2.nrows)):
                row = []
                for col_idx in range(min(8, sheet_qizhong2.ncols)):
                    cell = sheet_qizhong2.cell(row_idx, col_idx)
                    if cell.ctype == xlrd.XL_CELL_NUMBER:
                        val = cell.value
                        if val == int(val):
                            row.append(str(int(val)))
                        else:
                            row.append(str(val))
                    elif cell.ctype == xlrd.XL_CELL_EMPTY:
                        row.append("(空)")
                    else:
                        row.append(str(cell.value))
                print(f"行{row_idx+1}: {' | '.join(row)}")
            
            # 对比数据
            print("\n" + "-" * 80)
            print("检查第2个工作表的数据是否匹配...")
            print("-" * 80)
            
            differences = 0
            for row_idx in range(min(sheet_summary.nrows, sheet_qizhong2.nrows)):
                for col_idx in range(min(sheet_summary.ncols, sheet_qizhong2.ncols)):
                    val1 = sheet_summary.cell(row_idx, col_idx).value
                    val2 = sheet_qizhong2.cell(row_idx, col_idx).value
                    
                    if isinstance(val1, (int, float)) and isinstance(val2, (int, float)):
                        if abs(val1 - val2) > 0.001:
                            differences += 1
                    elif str(val1).strip() != str(val2).strip():
                        differences += 1
            
            if differences == 0:
                print("✅ 第2个工作表的数据与汇总文件完全匹配！")
            else:
                print(f"⚠️  第2个工作表仍有 {differences} 处差异")

if __name__ == "__main__":
    main()
