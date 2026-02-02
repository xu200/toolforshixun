"""
多次实训成绩汇总窗口
独立的 GUI 界面，用于导入多次实训成绩并生成汇总表
"""
import os
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QGridLayout,
    QPushButton, QLabel, QFileDialog, QMessageBox,
    QGroupBox, QStatusBar
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont

from .modules.excel_loader import ExcelLoader
from .modules.merger import ScoreMerger
from .modules.exporter import SummaryExporter


class SummaryWindow(QDialog):
    """多次实训成绩汇总窗口"""
    
    EXPERIMENT_NAMES = [
        "第一次实训",
        "第二次实训",
        "第三次实训",
        "第四次实训",
        "第五次实训",
        "第六次实训",
        "第七次实训",
        "第八次实训",
    ]
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("多次实训成绩汇总")
        self.setMinimumSize(600, 500)
        
        self.loader = ExcelLoader()
        self.merger = ScoreMerger()
        self.exporter = SummaryExporter()
        
        self.import_buttons = []
        self.status_labels = []
        
        self.setup_ui()
        self.update_button_states()
    
    def setup_ui(self):
        """设置界面"""
        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        
        title_label = QLabel("多次实训成绩汇总")
        title_label.setFont(QFont("微软雅黑", 14, QFont.Bold))
        title_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(title_label)
        
        desc_label = QLabel(
            "说明：按顺序导入各次实训的成绩 Excel 文件，系统将按学号自动合并。\n"
            "导入顺序即为实训编号，第1次导入对应「第一次实训分数」列。"
        )
        desc_label.setWordWrap(True)
        desc_label.setStyleSheet("color: #666; padding: 5px;")
        layout.addWidget(desc_label)
        
        import_group = QGroupBox("导入实训成绩")
        import_layout = QGridLayout(import_group)
        import_layout.setSpacing(10)
        
        for i in range(8):
            exp_index = i + 1
            
            btn = QPushButton(f"导入{self.EXPERIMENT_NAMES[i]}")
            btn.setMinimumWidth(150)
            btn.clicked.connect(lambda checked, idx=exp_index: self.import_experiment(idx))
            self.import_buttons.append(btn)
            
            status_label = QLabel("未导入")
            status_label.setStyleSheet("color: #999;")
            status_label.setMinimumWidth(250)
            self.status_labels.append(status_label)
            
            row = i // 2
            col = (i % 2) * 2
            import_layout.addWidget(btn, row, col)
            import_layout.addWidget(status_label, row, col + 1)
        
        layout.addWidget(import_group)
        
        stats_group = QGroupBox("统计信息")
        stats_layout = QHBoxLayout(stats_group)
        
        self.student_count_label = QLabel("学生总数: 0")
        self.imported_count_label = QLabel("已导入: 0/8 次实训")
        
        stats_layout.addWidget(self.student_count_label)
        stats_layout.addStretch()
        stats_layout.addWidget(self.imported_count_label)
        
        layout.addWidget(stats_group)
        
        action_layout = QHBoxLayout()
        action_layout.setSpacing(20)
        
        self.reset_btn = QPushButton("重置全部")
        self.reset_btn.setMinimumWidth(120)
        self.reset_btn.clicked.connect(self.reset_all)
        
        self.export_btn = QPushButton("导出汇总 Excel")
        self.export_btn.setMinimumWidth(150)
        self.export_btn.clicked.connect(self.export_summary)
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
        action_layout.addWidget(self.reset_btn)
        action_layout.addWidget(self.export_btn)
        action_layout.addStretch()
        
        layout.addLayout(action_layout)
        
        self.status_bar = QStatusBar()
        layout.addWidget(self.status_bar)
        self.status_bar.showMessage("就绪")
    
    def import_experiment(self, experiment_index: int):
        """导入某次实训的成绩"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            f"选择{self.EXPERIMENT_NAMES[experiment_index - 1]}成绩 Excel",
            "",
            "Excel 文件 (*.xlsx *.xls)"
        )
        
        if not file_path:
            return
        
        self.status_bar.showMessage(f"正在导入 {os.path.basename(file_path)}...")
        
        success, message, records = self.loader.load(file_path)
        
        if not success:
            QMessageBox.warning(self, "导入失败", message)
            self.status_bar.showMessage("导入失败")
            return
        
        merged_count = self.merger.merge(experiment_index, records, os.path.basename(file_path))
        
        self.status_labels[experiment_index - 1].setText(
            f"✓ {os.path.basename(file_path)} ({merged_count}条)"
        )
        self.status_labels[experiment_index - 1].setStyleSheet("color: #4CAF50;")
        
        self.update_button_states()
        self.status_bar.showMessage(f"成功导入 {merged_count} 条记录")
    
    def reset_all(self):
        """重置所有数据"""
        reply = QMessageBox.question(
            self, "确认重置",
            "确定要清空所有已导入的数据吗？",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.No:
            return
        
        self.merger.reset()
        
        for i in range(8):
            self.status_labels[i].setText("未导入")
            self.status_labels[i].setStyleSheet("color: #999;")
        
        self.update_button_states()
        self.status_bar.showMessage("已重置所有数据")
    
    def export_summary(self):
        """导出汇总 Excel"""
        if self.merger.get_imported_count() == 0:
            QMessageBox.warning(self, "无数据", "请先导入至少一次实训的成绩")
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "保存汇总 Excel",
            "实训成绩汇总.xlsx",
            "Excel 文件 (*.xlsx)"
        )
        
        if not file_path:
            return
        
        self.status_bar.showMessage("正在导出...")
        
        success, message = self.exporter.export(
            self.merger.get_all_students(),
            file_path
        )
        
        if success:
            QMessageBox.information(self, "导出成功", message)
            self.status_bar.showMessage(message)
        else:
            QMessageBox.warning(self, "导出失败", message)
            self.status_bar.showMessage("导出失败")
    
    def update_button_states(self):
        """更新按钮状态"""
        for i in range(8):
            exp_index = i + 1
            if self.merger.is_imported(exp_index):
                self.import_buttons[i].setEnabled(False)
                self.import_buttons[i].setText(f"{self.EXPERIMENT_NAMES[i]} ✓")
            else:
                self.import_buttons[i].setEnabled(True)
                self.import_buttons[i].setText(f"导入{self.EXPERIMENT_NAMES[i]}")
        
        self.student_count_label.setText(f"学生总数: {self.merger.get_student_count()}")
        self.imported_count_label.setText(f"已导入: {self.merger.get_imported_count()}/8 次实训")
        
        self.export_btn.setEnabled(self.merger.get_imported_count() > 0)
