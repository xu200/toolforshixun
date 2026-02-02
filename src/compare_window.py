"""
成绩对比窗口
独立的 GUI 界面，用于对比人工 Excel 和系统 Excel 的实训成绩
"""
import os
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QMessageBox,
    QGroupBox, QStatusBar
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont

from .modules.comparator import ScoreComparator
from .modules.highlighter import ScoreHighlighter


class CompareWindow(QDialog):
    """成绩对比窗口"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("成绩对比检查")
        self.setMinimumSize(550, 400)
        
        self.comparator = ScoreComparator()
        self.highlighter = ScoreHighlighter()
        
        self.manual_file = ""
        self.auto_files = []
        self.compared = False
        
        self.setup_ui()
        self.update_button_states()
    
    def setup_ui(self):
        """设置界面"""
        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        
        title_label = QLabel("成绩对比检查")
        title_label.setFont(QFont("微软雅黑", 14, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)
        
        desc_label = QLabel(
            "说明：对比人工整理的班级成绩 Excel 与系统导出的汇总 Excel，\n"
            "找出实训成绩不一致的单元格并标红。平时测验成绩不参与对比。"
        )
        desc_label.setWordWrap(True)
        desc_label.setStyleSheet("color: #666; padding: 5px;")
        layout.addWidget(desc_label)
        
        import_group = QGroupBox("导入文件")
        import_layout = QVBoxLayout(import_group)
        import_layout.setSpacing(10)
        
        manual_row = QHBoxLayout()
        self.manual_btn = QPushButton("导入人工成绩 Excel")
        self.manual_btn.setMinimumWidth(180)
        self.manual_btn.clicked.connect(self.import_manual)
        self.manual_label = QLabel("未导入")
        self.manual_label.setStyleSheet("color: #999;")
        manual_row.addWidget(self.manual_btn)
        manual_row.addWidget(self.manual_label, 1)
        import_layout.addLayout(manual_row)
        
        auto_row = QHBoxLayout()
        self.auto_btn = QPushButton("导入系统成绩 Excel")
        self.auto_btn.setMinimumWidth(180)
        self.auto_btn.clicked.connect(self.import_auto)
        self.auto_label = QLabel("未导入（可导入多份）")
        self.auto_label.setStyleSheet("color: #999;")
        auto_row.addWidget(self.auto_btn)
        auto_row.addWidget(self.auto_label, 1)
        import_layout.addLayout(auto_row)
        
        # 清空系统成绩按钮
        clear_row = QHBoxLayout()
        self.clear_auto_btn = QPushButton("清空系统成绩")
        self.clear_auto_btn.setMinimumWidth(180)
        self.clear_auto_btn.clicked.connect(self.clear_auto)
        clear_row.addWidget(self.clear_auto_btn)
        clear_row.addStretch()
        import_layout.addLayout(clear_row)
        
        layout.addWidget(import_group)
        
        result_group = QGroupBox("对比结果")
        result_layout = QVBoxLayout(result_group)
        
        self.result_label = QLabel("请先导入两个 Excel 文件")
        self.result_label.setAlignment(Qt.AlignCenter)
        self.result_label.setStyleSheet("font-size: 14px; padding: 20px;")
        result_layout.addWidget(self.result_label)
        
        layout.addWidget(result_group)
        
        action_layout = QHBoxLayout()
        action_layout.setSpacing(20)
        
        self.compare_btn = QPushButton("开始对比")
        self.compare_btn.setMinimumWidth(120)
        self.compare_btn.clicked.connect(self.do_compare)
        
        self.export_btn = QPushButton("导出对比结果")
        self.export_btn.setMinimumWidth(150)
        self.export_btn.clicked.connect(self.export_result)
        self.export_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:disabled {
                background-color: #cccccc;
            }
        """)
        
        action_layout.addStretch()
        action_layout.addWidget(self.compare_btn)
        action_layout.addWidget(self.export_btn)
        action_layout.addStretch()
        
        layout.addLayout(action_layout)
        
        self.status_bar = QStatusBar()
        layout.addWidget(self.status_bar)
        self.status_bar.showMessage("就绪")
    
    def import_manual(self):
        """导入人工成绩 Excel"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择人工成绩 Excel",
            "",
            "Excel 文件 (*.xlsx *.xls)"
        )
        
        if not file_path:
            return
        
        self.status_bar.showMessage("正在加载人工成绩 Excel...")
        
        success, message = self.comparator.load_manual_excel(file_path)
        
        if success:
            self.manual_file = file_path
            self.manual_label.setText(f"✓ {os.path.basename(file_path)}")
            self.manual_label.setStyleSheet("color: #4CAF50;")
            self.compared = False
            self.result_label.setText("请点击「开始对比」")
        else:
            QMessageBox.warning(self, "导入失败", message)
        
        self.status_bar.showMessage(message)
        self.update_button_states()
    
    def import_auto(self):
        """导入系统成绩 Excel（支持多次导入）"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择系统成绩 Excel（可多次导入不同班级）",
            "",
            "Excel 文件 (*.xlsx *.xls)"
        )
        
        if not file_path:
            return
        
        self.status_bar.showMessage("正在加载系统成绩 Excel...")
        
        success, message = self.comparator.load_auto_excel(file_path)
        
        if success:
            self.auto_files.append(file_path)
            file_count = len(self.auto_files)
            student_count = len(self.comparator.auto_map)
            if file_count == 1:
                self.auto_label.setText(f"✓ {os.path.basename(file_path)} ({student_count}人)")
            else:
                self.auto_label.setText(f"✓ 已导入 {file_count} 份，共 {student_count} 人")
            self.auto_label.setStyleSheet("color: #4CAF50;")
            self.compared = False
            self.result_label.setText("请点击「开始对比」")
        else:
            QMessageBox.warning(self, "导入失败", message)
        
        self.status_bar.showMessage(message)
        self.update_button_states()
    
    def clear_auto(self):
        """清空系统成绩数据"""
        self.comparator.clear_auto_data()
        self.auto_files = []
        self.auto_label.setText("未导入（可导入多份）")
        self.auto_label.setStyleSheet("color: #999;")
        self.compared = False
        self.result_label.setText("请先导入两个 Excel 文件")
        self.result_label.setStyleSheet("font-size: 14px; padding: 20px;")
        self.status_bar.showMessage("已清空系统成绩数据")
        self.update_button_states()
    
    def do_compare(self):
        """执行对比"""
        self.status_bar.showMessage("正在对比...")
        
        success, message, mismatch_count = self.comparator.compare()
        
        if success:
            self.compared = True
            if mismatch_count > 0:
                self.result_label.setText(f"发现 {mismatch_count} 处成绩不一致（将标红）")
                self.result_label.setStyleSheet("font-size: 14px; padding: 20px; color: #FF0000;")
            else:
                self.result_label.setText("所有实训成绩完全一致 ✓")
                self.result_label.setStyleSheet("font-size: 14px; padding: 20px; color: #4CAF50;")
        else:
            QMessageBox.warning(self, "对比失败", message)
        
        self.status_bar.showMessage(message)
        self.update_button_states()
    
    def export_result(self):
        """导出对比结果"""
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "保存对比结果",
            "成绩对比检查结果.xlsx",
            "Excel 文件 (*.xlsx)"
        )
        
        if not file_path:
            return
        
        self.status_bar.showMessage("正在导出...")
        
        success, message = self.highlighter.highlight_and_save(
            self.comparator.get_manual_df(),
            self.comparator.get_mismatches(),
            file_path
        )
        
        if success:
            QMessageBox.information(self, "导出成功", message)
        else:
            QMessageBox.warning(self, "导出失败", message)
        
        self.status_bar.showMessage(message)
    
    def update_button_states(self):
        """更新按钮状态"""
        has_manual = bool(self.manual_file)
        has_auto = len(self.auto_files) > 0
        
        self.compare_btn.setEnabled(has_manual and has_auto)
        self.export_btn.setEnabled(self.compared)
        self.clear_auto_btn.setEnabled(has_auto)
