"""
Excel 数据处理模块
基于分析结果处理不同格式的成绩表

输入格式（22高二期中.xls）：
- 2个工作表：高二一调、高二期中
- 行1：表头（科目名称）
- 行2+：学生数据

输出格式（22历次成绩.xls）：
- 8个工作表：
  1. 历次成绩原始（合并两次考试）
  2. 历次成绩打印（合并两次考试）
  3. 尖子生成绩（筛选）
  4. 高一期末（空）
  5. 高二一调（复制）
  6. 高二期中（复制）
  7. 均值 姓名排序（计算）
  8. 两次 均值 升序（计算）
"""
import xlrd
import xlwt
from typing import List, Dict, Tuple, Optional


class ExamSheet:
    """考试成绩工作表数据"""
    
    def __init__(self, name: str, data: List[List]):
        self.name = name
        self.data = data  # 保持原始数据结构
        self.students = []
        
    def parse_data(self):
        """解析学生数据"""
        self.students = []
        # 数据从第2行(索引1)开始
        for row_idx in range(1, len(self.data)):
            row = self.data[row_idx]
            if len(row) >= 2 and row[1]:  # 确保有姓名
                student_data = {
                    'exam_name': str(row[0]).strip() if row[0] else '',
                    'name': str(row[1]).strip(),
                    'scores': {}
                }
                
                # 解析各科成绩和排名
                # 列3-16：语文(3,4), 数学(5,6), 英语(7,8), 物理(9,10), 化学(11,12), 生物(13,14), 总分(15,16)
                subjects = ['语文', '数学', '英语', '物理', '化学', '生物', '总分']
                col_idx = 2
                
                for subject in subjects:
                    if col_idx < len(row):
                        score = row[col_idx]
                        rank = row[col_idx + 1] if col_idx + 1 < len(row) else None
                        
                        student_data['scores'][subject] = {
                            'score': score if score != '' else None,
                            'rank': rank if rank != '' else None
                        }
                        col_idx += 2
                
                self.students.append(student_data)


class ExcelHandler:
    """Excel 文件处理器"""
    
    @staticmethod
    def read_excel_file(file_path: str) -> List[ExamSheet]:
        """
        读取 Excel 文件
        返回所有工作表的数据
        
        保持原始格式不变，直接复制所有数据
        """
        workbook = xlrd.open_workbook(file_path, formatting_info=False)
        sheets = []
        
        for sheet_idx in range(workbook.nsheets):
            sheet = workbook.sheet_by_index(sheet_idx)
            
            # 读取所有数据
            data = []
            for row_idx in range(sheet.nrows):
                row = []
                for col_idx in range(sheet.ncols):
                    cell = sheet.cell(row_idx, col_idx)
                    if cell.ctype == xlrd.XL_CELL_NUMBER:
                        val = cell.value
                        # 保持数值精度
                        row.append(val)
                    elif cell.ctype == xlrd.XL_CELL_EMPTY:
                        row.append('')
                    else:
                        row.append(str(cell.value))
                data.append(row)
            
            # 创建工作表对象
            exam_sheet = ExamSheet(sheet.name, data)
            exam_sheet.parse_data()
            sheets.append(exam_sheet)
        
        return sheets
    
    @staticmethod
    def write_excel_file(sheets: List[ExamSheet], output_path: str):
        """
        写入 Excel 文件
        生成8个工作表的标准格式
        
        输入: 2个工作表（高二一调、高二期中）
        输出: 8个工作表
        """
        workbook = xlwt.Workbook(encoding='utf-8')
        
        # 找到两个输入工作表
        sheet_yidiao = None
        sheet_qizhong = None
        
        for sheet in sheets:
            if '一调' in sheet.name:
                sheet_yidiao = sheet
            elif '期中' in sheet.name:
                sheet_qizhong = sheet
        
        if not sheet_yidiao or not sheet_qizhong:
            raise ValueError("输入文件必须包含'高二一调'和'高二期中'两个工作表")
        
        # 工作表1: 历次成绩原始 - 合并两次考试
        ExcelHandler._write_merged_sheet(workbook, "历次成绩原始", sheet_yidiao, sheet_qizhong)
        
        # 工作表2: 历次成绩打印 - 合并两次考试
        ExcelHandler._write_merged_sheet(workbook, "历次成绩打印", sheet_yidiao, sheet_qizhong)
        
        # 工作表3: 尖子生成绩 - 筛选排名前20的学生
        ExcelHandler._write_top_students_sheet(workbook, "尖子生成绩", sheet_yidiao, sheet_qizhong)
        
        # 工作表4: 高一期末 - 空表
        ExcelHandler._write_empty_sheet(workbook, "高一期末")
        
        # 工作表5: 高二一调 - 直接复制
        ExcelHandler._write_single_sheet(workbook, sheet_yidiao)
        
        # 工作表6: 高二期中 - 直接复制
        ExcelHandler._write_single_sheet(workbook, sheet_qizhong)
        
        # 工作表7: 均值 姓名排序 - 计算两次排名均值
        ExcelHandler._write_average_sheet(workbook, "均值 姓名排序", sheet_yidiao, sheet_qizhong, sort_by_name=True)
        
        # 工作表8: 两次 均值 升序 - 按均值排序
        ExcelHandler._write_average_sheet(workbook, "两次 均值 升序", sheet_yidiao, sheet_qizhong, sort_by_name=False)
        
        workbook.save(output_path)
    
    @staticmethod
    def _write_single_sheet(workbook, exam_sheet: ExamSheet):
        """写入单个工作表（直接复制）"""
        ws = workbook.add_sheet(exam_sheet.name)
        
        for row_idx, row_data in enumerate(exam_sheet.data):
            for col_idx, cell_value in enumerate(row_data):
                if isinstance(cell_value, (int, float)):
                    ws.write(row_idx, col_idx, cell_value)
                else:
                    ws.write(row_idx, col_idx, str(cell_value))
    
    @staticmethod
    def _write_merged_sheet(workbook, sheet_name: str, sheet1: ExamSheet, sheet2: ExamSheet):
        """写入合并两次考试的工作表"""
        ws = workbook.add_sheet(sheet_name)
        
        # 写入表头（从sheet1复制）
        for col_idx, cell_value in enumerate(sheet1.data[0]):
            ws.write(0, col_idx, str(cell_value))
        
        # 写入标签行（从sheet1复制）
        for col_idx, cell_value in enumerate(sheet1.data[1]):
            ws.write(1, col_idx, str(cell_value))
        
        # 写入sheet1的数据
        row_idx = 2
        for data_row in sheet1.data[2:]:
            for col_idx, cell_value in enumerate(data_row):
                if isinstance(cell_value, (int, float)):
                    ws.write(row_idx, col_idx, cell_value)
                else:
                    ws.write(row_idx, col_idx, str(cell_value))
            row_idx += 1
        
        # 写入sheet2的数据
        for data_row in sheet2.data[1:]:  # 跳过sheet2的表头
            for col_idx, cell_value in enumerate(data_row):
                if isinstance(cell_value, (int, float)):
                    ws.write(row_idx, col_idx, cell_value)
                else:
                    ws.write(row_idx, col_idx, str(cell_value))
            row_idx += 1
    
    @staticmethod
    def _write_top_students_sheet(workbook, sheet_name: str, sheet1: ExamSheet, sheet2: ExamSheet):
        """写入尖子生工作表（筛选排名前20的学生）"""
        ws = workbook.add_sheet(sheet_name)
        
        # 写入表头
        for col_idx, cell_value in enumerate(sheet1.data[0]):
            ws.write(0, col_idx, str(cell_value))
        
        # 写入标签行
        for col_idx, cell_value in enumerate(sheet1.data[1]):
            ws.write(1, col_idx, str(cell_value))
        
        # 收集所有学生的总排名（从总分校次列）
        students_with_rank = []
        for data_row in sheet1.data[2:]:
            if len(data_row) >= 2 and data_row[1]:
                total_rank = data_row[15] if len(data_row) > 15 else 9999
                try:
                    total_rank = int(total_rank) if isinstance(total_rank, (int, float)) else 9999
                except:
                    total_rank = 9999
                students_with_rank.append((total_rank, data_row))
        
        # 按排名排序，取前20
        students_with_rank.sort(key=lambda x: x[0])
        top_students = [row for _, row in students_with_rank[:20]]
        
        # 写入尖子生数据
        row_idx = 2
        for data_row in top_students:
            for col_idx, cell_value in enumerate(data_row):
                if isinstance(cell_value, (int, float)):
                    ws.write(row_idx, col_idx, cell_value)
                else:
                    ws.write(row_idx, col_idx, str(cell_value))
            row_idx += 1
    
    @staticmethod
    def _write_empty_sheet(workbook, sheet_name: str):
        """写入空表（高一期末）"""
        ws = workbook.add_sheet(sheet_name)
        # 只写表头
        headers = ['班级', '姓名', '语文', '', '数学', '', '英语', '', '物理', '', '化学', '', '生物', '', '总分', '校次']
        for col_idx, header in enumerate(headers):
            ws.write(0, col_idx, header)
        
        # 写标签行
        labels = ['', '', '得分', '校次', '得分', '校次', '得分', '校次', '得分', '校次', '得分', '校次', '得分', '校次', '得分', '校次']
        for col_idx, label in enumerate(labels):
            ws.write(1, col_idx, label)
    
    @staticmethod
    def _write_average_sheet(workbook, sheet_name: str, sheet1: ExamSheet, sheet2: ExamSheet, sort_by_name: bool = True):
        """写入均值工作表"""
        ws = workbook.add_sheet(sheet_name)
        
        # 构建学生数据字典
        students_data = {}
        
        # 从sheet1（高二一调）读取数据
        for data_row in sheet1.data[2:]:
            if len(data_row) >= 2 and data_row[1]:
                name = str(data_row[1]).strip()
                total_score1 = data_row[14] if len(data_row) > 14 else None
                rank1 = data_row[15] if len(data_row) > 15 else None
                students_data[name] = {
                    'score1': total_score1,
                    'rank1': rank1,
                    'score2': None,
                    'rank2': None
                }
        
        # 从sheet2（高二期中）读取数据
        for data_row in sheet2.data[1:]:
            if len(data_row) >= 2 and data_row[1]:
                name = str(data_row[1]).strip()
                total_score2 = data_row[14] if len(data_row) > 14 else None
                rank2 = data_row[15] if len(data_row) > 15 else None
                
                if name in students_data:
                    students_data[name]['score2'] = total_score2
                    students_data[name]['rank2'] = rank2
        
        # 计算均值
        for name in students_data:
            rank1 = students_data[name]['rank1']
            rank2 = students_data[name]['rank2']
            
            if rank1 and rank2:
                try:
                    rank1_val = float(rank1)
                    rank2_val = float(rank2)
                    avg = (rank1_val + rank2_val) / 2
                    students_data[name]['avg'] = avg
                except:
                    students_data[name]['avg'] = 9999
            else:
                students_data[name]['avg'] = 9999
        
        # 排序
        if sort_by_name:
            sorted_students = sorted(students_data.items(), key=lambda x: x[0])
        else:
            sorted_students = sorted(students_data.items(), key=lambda x: x[1]['avg'])
        
        # 写表头
        headers = ['', '高二一调', '校次', '高二期中', '', '']
        for col_idx, header in enumerate(headers):
            ws.write(0, col_idx, header)
        
        # 写标签行
        labels = ['姓名', '得分', '校次', '总分', '校次', '均值']
        for col_idx, label in enumerate(labels):
            ws.write(1, col_idx, label)
        
        # 写数据
        row_idx = 2
        for name, data in sorted_students:
            ws.write(row_idx, 0, name)
            ws.write(row_idx, 1, data['score1'] if data['score1'] else '')
            ws.write(row_idx, 2, data['rank1'] if data['rank1'] else '')
            ws.write(row_idx, 3, data['score2'] if data['score2'] else '')
            ws.write(row_idx, 4, data['rank2'] if data['rank2'] else '')
            ws.write(row_idx, 5, data['avg'] if data['avg'] != 9999 else '')
            row_idx += 1
    
    @staticmethod
    def get_file_info(file_path: str) -> Dict:
        """
        获取文件信息
        """
        try:
            workbook = xlrd.open_workbook(file_path)
            sheets_info = []
            
            for sheet_idx in range(workbook.nsheets):
                sheet = workbook.sheet_by_index(sheet_idx)
                
                # 统计学生数量（数据从第2行开始）
                student_count = 0
                for row_idx in range(1, sheet.nrows):
                    if sheet.cell(row_idx, 1).value:  # 有姓名
                        student_count += 1
                
                sheets_info.append({
                    'name': sheet.name,
                    'rows': sheet.nrows,
                    'cols': sheet.ncols,
                    'student_count': student_count
                })
            
            return {
                'success': True,
                'sheets': sheets_info,
                'total_sheets': workbook.nsheets
            }
        except Exception as e:
            return {
                'success': False,
                'error': str(e)
            }
