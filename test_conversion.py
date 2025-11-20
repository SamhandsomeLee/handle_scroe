"""
测试完整的转换流程
"""
import sys
import io
from excel_handler import ExcelHandler

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

def main():
    input_file = r"d:\handle_score\data\22高二期中.xls"
    output_file = r"d:\handle_score\data\test_output.xls"
    
    print("=" * 80)
    print("开始转换测试")
    print("=" * 80)
    
    try:
        print(f"\n1. 读取输入文件: {input_file}")
        sheets = ExcelHandler.read_excel_file(input_file)
        print(f"   ✓ 读取成功，共 {len(sheets)} 个工作表")
        for sheet in sheets:
            print(f"     - {sheet.name}: {len(sheet.data)} 行, {len(sheet.students)} 个学生")
        
        print(f"\n2. 生成输出文件: {output_file}")
        ExcelHandler.write_excel_file(sheets, output_file)
        print(f"   ✓ 生成成功")
        
        # 验证输出文件
        print(f"\n3. 验证输出文件")
        import xlrd
        wb = xlrd.open_workbook(output_file)
        print(f"   ✓ 输出文件包含 {wb.nsheets} 个工作表:")
        for i in range(wb.nsheets):
            sheet = wb.sheet_by_index(i)
            print(f"     {i+1}. {sheet.name}: {sheet.nrows} 行 × {sheet.ncols} 列")
        
        print("\n" + "=" * 80)
        print("✓ 转换测试完成！")
        print("=" * 80)
        
    except Exception as e:
        print(f"\n✗ 错误: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
