"""
成绩导入转换工具 - 现代化 GUI
使用 PySide6 实现
"""
import sys
import os
from pathlib import Path
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QTextEdit, QGroupBox,
    QProgressBar, QMessageBox, QTableWidget, QTableWidgetItem,
    QHeaderView, QSplitter
)
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QFont, QIcon, QColor
from excel_handler import ExcelHandler


class ProcessThread(QThread):
    """后台处理线程"""
    progress = Signal(int)
    status = Signal(str)
    finished = Signal(bool, str)
    
    def __init__(self, input_file: str, output_file: str):
        super().__init__()
        self.input_file = input_file
        self.output_file = output_file
    
    def run(self):
        try:
            self.status.emit("正在读取文件...")
            self.progress.emit(30)
            
            # 读取 Excel 文件
            sheets = ExcelHandler.read_excel_file(self.input_file)
            
            self.status.emit(f"已读取 {len(sheets)} 个工作表")
            self.progress.emit(60)
            
            # 写入输出文件
            self.status.emit("正在生成输出文件...")
            ExcelHandler.write_excel_file(sheets, self.output_file)
            
            self.progress.emit(100)
            self.status.emit("处理完成！")
            self.finished.emit(True, f"成功生成文件：{self.output_file}")
            
        except Exception as e:
            self.status.emit(f"错误：{str(e)}")
            self.finished.emit(False, str(e))


class MainWindow(QMainWindow):
    """主窗口"""
    
    def __init__(self):
        super().__init__()
        self.input_file = None
        self.output_file = None
        self.process_thread = None
        
        self.init_ui()
    
    def init_ui(self):
        """初始化 UI"""
        self.setWindowTitle("成绩导入转换工具 v1.0")
        self.setMinimumSize(900, 650)
        
        # 创建中央 widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 主布局
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)
        
        # 标题
        title_label = QLabel("成绩 Excel 文件导入转换")
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)
        
        # 说明文本
        info_label = QLabel(
            "导入包含'高二一调'和'高二期中'两个工作表的 Excel 文件\n"
            "输出包含8个工作表的标准格式文件（历次成绩原始、历次成绩打印、尖子生成绩、高一期末、高二一调、高二期中、均值排序、均值升序）"
        )
        info_label.setAlignment(Qt.AlignCenter)
        info_label.setStyleSheet("color: #666; padding: 10px;")
        main_layout.addWidget(info_label)
        
        # 文件选择区域
        file_group = QGroupBox("文件选择")
        file_layout = QVBoxLayout()
        
        # 输入文件
        input_layout = QHBoxLayout()
        input_layout.addWidget(QLabel("输入文件:"))
        self.input_label = QLabel("未选择文件")
        self.input_label.setStyleSheet(
            "QLabel { background-color: #f5f5f5; padding: 8px; "
            "border: 1px solid #ddd; border-radius: 4px; }"
        )
        input_layout.addWidget(self.input_label, 1)
        self.btn_select_input = QPushButton("选择文件")
        self.btn_select_input.clicked.connect(self.select_input_file)
        self.btn_select_input.setStyleSheet(
            "QPushButton { background-color: #2196F3; color: white; "
            "padding: 8px 20px; border: none; border-radius: 4px; }"
            "QPushButton:hover { background-color: #1976D2; }"
        )
        input_layout.addWidget(self.btn_select_input)
        file_layout.addLayout(input_layout)
        
        # 输出文件
        output_layout = QHBoxLayout()
        output_layout.addWidget(QLabel("输出文件:"))
        self.output_label = QLabel("未设置")
        self.output_label.setStyleSheet(
            "QLabel { background-color: #f5f5f5; padding: 8px; "
            "border: 1px solid #ddd; border-radius: 4px; }"
        )
        output_layout.addWidget(self.output_label, 1)
        self.btn_select_output = QPushButton("设置路径")
        self.btn_select_output.clicked.connect(self.select_output_file)
        self.btn_select_output.setStyleSheet(
            "QPushButton { background-color: #4CAF50; color: white; "
            "padding: 8px 20px; border: none; border-radius: 4px; }"
            "QPushButton:hover { background-color: #45a049; }"
        )
        output_layout.addWidget(self.btn_select_output)
        file_layout.addLayout(output_layout)
        
        file_group.setLayout(file_layout)
        main_layout.addWidget(file_group)
        
        # 文件信息区域
        info_group = QGroupBox("文件信息")
        info_layout = QVBoxLayout()
        self.info_table = QTableWidget()
        self.info_table.setColumnCount(4)
        self.info_table.setHorizontalHeaderLabels(
            ["工作表名称", "总行数", "总列数", "学生数"]
        )
        self.info_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.info_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.info_table.setAlternatingRowColors(True)
        info_layout.addWidget(self.info_table)
        info_group.setLayout(info_layout)
        main_layout.addWidget(info_group)
        
        # 处理按钮
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        
        self.btn_process = QPushButton("开始转换")
        self.btn_process.setEnabled(False)
        self.btn_process.clicked.connect(self.start_process)
        self.btn_process.setStyleSheet(
            "QPushButton { background-color: #FF9800; color: white; "
            "padding: 12px 40px; border: none; border-radius: 4px; "
            "font-size: 14px; font-weight: bold; }"
            "QPushButton:hover { background-color: #F57C00; }"
            "QPushButton:disabled { background-color: #ccc; }"
        )
        btn_layout.addWidget(self.btn_process)
        btn_layout.addStretch()
        main_layout.addLayout(btn_layout)
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setStyleSheet(
            "QProgressBar { border: 2px solid #ddd; border-radius: 5px; "
            "text-align: center; }"
            "QProgressBar::chunk { background-color: #4CAF50; }"
        )
        main_layout.addWidget(self.progress_bar)
        
        # 日志区域
        log_group = QGroupBox("处理日志")
        log_layout = QVBoxLayout()
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(150)
        self.log_text.setStyleSheet(
            "QTextEdit { background-color: #f9f9f9; "
            "border: 1px solid #ddd; font-family: Consolas, monospace; }"
        )
        log_layout.addWidget(self.log_text)
        log_group.setLayout(log_layout)
        main_layout.addWidget(log_group)
        
        # 设置样式
        self.setStyleSheet("""
            QMainWindow {
                background-color: #fafafa;
            }
            QGroupBox {
                font-weight: bold;
                border: 2px solid #ddd;
                border-radius: 6px;
                margin-top: 10px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }
        """)
        
        self.log("程序已启动，等待导入文件...")
    
    def log(self, message: str):
        """添加日志"""
        self.log_text.append(f"[INFO] {message}")
    
    def select_input_file(self):
        """选择输入文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择输入 Excel 文件",
            "",
            "Excel Files (*.xls *.xlsx);;All Files (*)"
        )
        
        if file_path:
            self.input_file = file_path
            self.input_label.setText(os.path.basename(file_path))
            self.log(f"已选择输入文件: {file_path}")
            
            # 获取文件信息
            info = ExcelHandler.get_file_info(file_path)
            if info['success']:
                self.display_file_info(info)
                
                # 自动设置输出文件名
                if not self.output_file:
                    base_dir = os.path.dirname(file_path)
                    base_name = os.path.splitext(os.path.basename(file_path))[0]
                    default_output = os.path.join(base_dir, f"{base_name}_转换.xls")
                    self.output_file = default_output
                    self.output_label.setText(os.path.basename(default_output))
                
                self.update_process_button()
            else:
                QMessageBox.critical(self, "错误", f"读取文件失败: {info['error']}")
    
    def select_output_file(self):
        """选择输出文件"""
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "设置输出 Excel 文件",
            self.output_file if self.output_file else "",
            "Excel Files (*.xls);;All Files (*)"
        )
        
        if file_path:
            # 确保文件扩展名是 .xls
            if not file_path.endswith('.xls'):
                file_path += '.xls'
            
            self.output_file = file_path
            self.output_label.setText(os.path.basename(file_path))
            self.log(f"输出路径设置为: {file_path}")
            self.update_process_button()
    
    def display_file_info(self, info: dict):
        """显示文件信息"""
        sheets = info['sheets']
        self.info_table.setRowCount(len(sheets))
        
        for row, sheet_info in enumerate(sheets):
            self.info_table.setItem(row, 0, QTableWidgetItem(sheet_info['name']))
            self.info_table.setItem(row, 1, QTableWidgetItem(str(sheet_info['rows'])))
            self.info_table.setItem(row, 2, QTableWidgetItem(str(sheet_info['cols'])))
            self.info_table.setItem(row, 3, QTableWidgetItem(str(sheet_info['student_count'])))
        
        self.log(f"文件包含 {len(sheets)} 个工作表")
    
    def update_process_button(self):
        """更新处理按钮状态"""
        self.btn_process.setEnabled(self.input_file is not None and self.output_file is not None)
    
    def start_process(self):
        """开始处理"""
        if not self.input_file or not self.output_file:
            QMessageBox.warning(self, "警告", "请先选择输入和输出文件")
            return
        
        self.log("=" * 50)
        self.log("开始转换处理...")
        
        # 禁用按钮
        self.btn_process.setEnabled(False)
        self.btn_select_input.setEnabled(False)
        self.btn_select_output.setEnabled(False)
        
        # 显示进度条
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        
        # 启动处理线程
        self.process_thread = ProcessThread(self.input_file, self.output_file)
        self.process_thread.progress.connect(self.update_progress)
        self.process_thread.status.connect(self.log)
        self.process_thread.finished.connect(self.process_finished)
        self.process_thread.start()
    
    def update_progress(self, value: int):
        """更新进度条"""
        self.progress_bar.setValue(value)
    
    def process_finished(self, success: bool, message: str):
        """处理完成"""
        if success:
            self.log(f"✓ {message}")
            QMessageBox.information(self, "成功", message)
        else:
            self.log(f"✗ 错误: {message}")
            QMessageBox.critical(self, "错误", f"处理失败:\n{message}")
        
        # 恢复按钮
        self.btn_process.setEnabled(True)
        self.btn_select_input.setEnabled(True)
        self.btn_select_output.setEnabled(True)
        
        # 重置进度条
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(False)


def main():
    app = QApplication(sys.argv)
    
    # 设置应用样式
    app.setStyle('Fusion')
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
