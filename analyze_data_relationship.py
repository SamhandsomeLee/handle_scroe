"""
分析三个文件之间的数据关系
- 22历次成绩.xls (汇总文件，8个工作表)
- 22高二一调.xls (独立文件)
- 22高二期中.xls (独立文件)
"""
import xlrd
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def read_sheet_data(workbook, sheet_index_or_name):
    """读取工作表数据，返回二维列表"""
    if isinstance(sheet_index_or_name, int):
        sheet = workbook.sheet_by_index(sheet_index_or_name)
    else:
        sheet = workbook.sheet_by_name(sheet_index_or_name)
    
    data = []
    for row_idx in range(sheet.nrows):
        row = []
        for col_idx in range(sheet.ncols):
            cell = sheet.cell(row_idx, col_idx)
            if cell.ctype == xlrd.XL_CELL_NUMBER:
                val = cell.value
                if val == int(val):
                    row.append(int(val))
                else:
                    row.append(val)
            elif cell.ctype == xlrd.XL_CELL_EMPTY:
                row.append("")
            else:
                row.append(cell.value)
        data.append(row)
    return data, sheet.name

def compare_data(data1, data2, name1, name2):
    """对比两个数据集"""
    print(f"\n{'=' * 80}")
    print(f"对比: {name1} vs {name2}")
    print(f"{'=' * 80}")
    
    print(f"{name1}: {len(data1)} 行 × {len(data1[0]) if data1 else 0} 列")
    print(f"{name2}: {len(data2)} 行 × {len(data2[0]) if data2 else 0} 列")
    
    # 检查行列数
    if len(data1) != len(data2):
        print(f"⚠️  行数不同！差异: {abs(len(data1) - len(data2))} 行")
        return False
    
    if data1 and data2 and len(data1[0]) != len(data2[0]):
        print(f"⚠️  列数不同！{name1}: {len(data1[0])}列, {name2}: {len(data2[0])}列")
        return False
    
    # 逐行对比
    differences = []
    for row_idx in range(len(data1)):
        for col_idx in range(min(len(data1[row_idx]), len(data2[row_idx]))):
            val1 = data1[row_idx][col_idx]
            val2 = data2[row_idx][col_idx]
            
            # 处理数值比较
            if isinstance(val1, (int, float)) and isinstance(val2, (int, float)):
                if abs(val1 - val2) > 0.001:  # 浮点数容差
                    differences.append((row_idx, col_idx, val1, val2))
            elif str(val1).strip() != str(val2).strip():
                differences.append((row_idx, col_idx, val1, val2))
    
    if differences:
        print(f"\n❌ 发现 {len(differences)} 处差异：")
        for i, (r, c, v1, v2) in enumerate(differences[:10]):  # 只显示前10个
            print(f"  位置[{r+1},{c+1}]: '{v1}' ≠ '{v2}'")
        if len(differences) > 10:
            print(f"  ... 还有 {len(differences) - 10} 处差异")
        return False
    else:
        print(f"\n✅ 数据完全一致！")
        return True

def analyze_structure(data, name):
    """分析数据结构"""
    print(f"\n{'-' * 80}")
    print(f"{name} 的结构分析")
    print(f"{'-' * 80}")
    
    if len(data) < 3:
        print("数据行数不足，跳过分析")
        return
    
    # 分析前3行
    print("前3行内容:")
    for i in range(min(3, len(data))):
        print(f"  行{i+1}: {data[i][:5]}...")  # 只显示前5列
    
    # 判断表头结构
    if len(data) >= 2:
        row2_has_labels = any(str(cell) in ['得分', '校次'] for cell in data[1])
        if row2_has_labels:
            print("\n📋 表头结构: 双层表头（第2行为'得分/校次'标签）")
            print(f"   数据起始行: 第3行")
            data_start = 2
        else:
            print("\n📋 表头结构: 单层表头")
            print(f"   数据起始行: 第2行")
            data_start = 1
        
        # 检查第1列
        if data_start < len(data):
            first_col_values = [data[i][0] for i in range(data_start, min(data_start+5, len(data)))]
            unique_values = set(str(v) for v in first_col_values if v)
            if len(unique_values) == 1:
                print(f"   第1列: 冗余（所有行都是'{list(unique_values)[0]}'）")
            else:
                print(f"   第1列: 有意义数据")

def main():
    print("=" * 80)
    print("Excel 文件数据关系分析")
    print("=" * 80)
    
    # 读取三个文件
    summary_file = r"d:\handle_score\data\22历次成绩.xls"
    yidiao_file = r"d:\handle_score\data\22高二一调.xls"
    qizhong_file = r"d:\handle_score\data\22高二期中.xls"
    
    print(f"\n读取汇总文件: {summary_file}")
    wb_summary = xlrd.open_workbook(summary_file)
    print(f"  包含 {wb_summary.nsheets} 个工作表:")
    for i in range(wb_summary.nsheets):
        print(f"    {i+1}. {wb_summary.sheet_by_index(i).name}")
    
    print(f"\n读取独立文件1: {yidiao_file}")
    wb_yidiao = xlrd.open_workbook(yidiao_file)
    print(f"  包含 {wb_yidiao.nsheets} 个工作表")
    
    print(f"\n读取独立文件2: {qizhong_file}")
    wb_qizhong = xlrd.open_workbook(qizhong_file)
    print(f"  包含 {wb_qizhong.nsheets} 个工作表")
    
    # 读取数据
    print("\n" + "=" * 80)
    print("数据读取")
    print("=" * 80)
    
    # 汇总文件中的"高二一调"工作表
    data_summary_yidiao, name_summary_yidiao = read_sheet_data(wb_summary, "高二一调")
    print(f"✓ 汇总文件 - 工作表'{name_summary_yidiao}': {len(data_summary_yidiao)}行")
    
    # 汇总文件中的"高二期中"工作表
    data_summary_qizhong, name_summary_qizhong = read_sheet_data(wb_summary, "高二期中")
    print(f"✓ 汇总文件 - 工作表'{name_summary_qizhong}': {len(data_summary_qizhong)}行")
    
    # 独立文件
    data_yidiao, name_yidiao = read_sheet_data(wb_yidiao, 0)
    print(f"✓ 独立文件 - {name_yidiao}: {len(data_yidiao)}行")
    
    data_qizhong, name_qizhong = read_sheet_data(wb_qizhong, 0)
    print(f"✓ 独立文件 - {name_qizhong}: {len(data_qizhong)}行")
    
    # 结构分析
    print("\n" + "=" * 80)
    print("结构分析")
    print("=" * 80)
    
    analyze_structure(data_summary_yidiao, "汇总文件-高二一调")
    analyze_structure(data_yidiao, "独立文件-高二一调")
    analyze_structure(data_summary_qizhong, "汇总文件-高二期中")
    analyze_structure(data_qizhong, "独立文件-高二期中")
    
    # 数据对比
    print("\n" + "=" * 80)
    print("数据一致性检查")
    print("=" * 80)
    
    result1 = compare_data(
        data_summary_yidiao, data_yidiao,
        "汇总文件[高二一调]", "独立文件[高二一调.xls]"
    )
    
    result2 = compare_data(
        data_summary_qizhong, data_qizhong,
        "汇总文件[高二期中]", "独立文件[高二期中.xls]"
    )
    
    # 总结
    print("\n" + "=" * 80)
    print("关系总结")
    print("=" * 80)
    
    if result1 and result2:
        print("\n✅ 结论: 独立文件中的数据与汇总文件中对应工作表的数据完全一致")
        print("\n数据流向推测:")
        print("  【独立文件】 → 【汇总文件】")
        print("  说明: 汇总文件是从各个独立文件中提取数据整合而成")
    else:
        print("\n⚠️  数据存在差异，需要进一步检查数据来源")
    
    print("\n汇总文件的其他工作表:")
    for i in range(wb_summary.nsheets):
        sheet_name = wb_summary.sheet_by_index(i).name
        if sheet_name not in ["高二一调", "高二期中"]:
            print(f"  - {sheet_name}")
    print("\n推测: 这些工作表可能来自其他独立的Excel文件")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n错误: {e}")
        import traceback
        traceback.print_exc()
