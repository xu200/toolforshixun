"""
GUI 主界面模块
基于 PySide6 实现教师实训报告批量处理系统界面
"""
import os
from typing import Optional
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTableWidget, QTableWidgetItem, QPushButton, QLabel,
    QFileDialog, QMessageBox, QHeaderView, QTextEdit,
    QDialog, QDialogButtonBox, QStatusBar, QGroupBox,
    QLineEdit, QAbstractItemView, QStyledItemDelegate, QApplication,
    QInputDialog
)
from PySide6.QtCore import Qt, Signal, QTimer, QObject, QThread
from PySide6.QtGui import QColor, QFont, QIntValidator

from .doc_parser import DocParser, StudentInfo
from .file_renamer import FileRenamer
from .comment_generator import CommentGenerator, CommentWriter
from .excel_exporter import ExcelExporter
from .licensing import LicenseManager, LicenseCheckResult, get_machine_code, get_machine_code_display
from .lab_name_fixer import LabNameFixer


class ScoreDelegate(QStyledItemDelegate):
    def createEditor(self, parent, option, index):
        editor = QLineEdit(parent)
        editor.setValidator(QIntValidator(0, 100, editor))
        editor.setAlignment(Qt.AlignCenter)
        editor.setFont(QFont("微软雅黑", 10))
        editor.setStyleSheet("QLineEdit{border:0;padding:0 6px;background:transparent;}")
        editor.setMinimumWidth(90)
        return editor

    def setEditorData(self, editor, index):
        data = index.data()
        editor.setText("" if data is None else str(data))

    def setModelData(self, editor, model, index):
        model.setData(index, editor.text().strip())

    def updateEditorGeometry(self, editor, option, index):
        editor.setGeometry(option.rect)


class CommentEditDialog(QDialog):
    """评语编辑对话框"""
    
    def __init__(self, student: StudentInfo, parent=None):
        super().__init__(parent)
        self.student = student
        self.setWindowTitle(f"编辑评语 - {student.student_name}")
        self.setMinimumSize(500, 300)
        self.setup_ui()
    
    def setup_ui(self):
        layout = QVBoxLayout(self)
        
        # 学生信息
        info_label = QLabel(f"学号: {self.student.student_id}  姓名: {self.student.student_name}  分数: {self.student.score}")
        info_label.setFont(QFont("微软雅黑", 10))
        layout.addWidget(info_label)
        
        # 评语编辑区
        self.text_edit = QTextEdit()
        self.text_edit.setText(self.student.comment)
        self.text_edit.setFont(QFont("宋体", 10))
        layout.addWidget(self.text_edit)
        
        # 按钮
        button_box = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel
        )
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
    
    def get_comment(self) -> str:
        return self.text_edit.toPlainText()


class ActivationDialog(QDialog):
    def __init__(self, license_manager: LicenseManager, parent=None):
        super().__init__(parent)
        self.license_manager = license_manager
        self.setWindowTitle("软件激活")
        self.setMinimumSize(520, 320)

        self._machine_code_raw: str = ""
        self._mc_thread: Optional[QThread] = None
        self._mc_worker: Optional[QObject] = None

        layout = QVBoxLayout(self)

        code_row = QHBoxLayout()
        code_row.addWidget(QLabel("机器码:"))
        self.machine_code_edit = QLineEdit()
        self.machine_code_edit.setReadOnly(True)
        self.machine_code_edit.setText("计算中...")
        code_row.addWidget(self.machine_code_edit, stretch=1)
        copy_btn = QPushButton("复制")
        copy_btn.clicked.connect(self.copy_machine_code)
        code_row.addWidget(copy_btn)
        layout.addLayout(code_row)

        self._start_machine_code_fetch()

        layout.addWidget(QLabel("激活码:"))
        self.license_edit = QTextEdit()
        self.license_edit.setFont(QFont("Consolas", 9))
        self.license_edit.setPlaceholderText("粘贴激活码（payload.signature）")
        layout.addWidget(self.license_edit, stretch=1)

        self.status_label = QLabel("")
        self.status_label.setStyleSheet("color: gray;")
        layout.addWidget(self.status_label)

        action_row = QHBoxLayout()
        self.activate_btn = QPushButton("校验并激活")
        self.activate_btn.clicked.connect(self.activate)
        action_row.addWidget(self.activate_btn)

        self.deactivate_btn = QPushButton("注销授权")
        self.deactivate_btn.clicked.connect(self.deactivate)
        action_row.addWidget(self.deactivate_btn)

        action_row.addStretch()
        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(self.reject)
        action_row.addWidget(close_btn)
        layout.addLayout(action_row)

    def _start_machine_code_fetch(self):
        if self._mc_thread is not None:
            try:
                if self._mc_thread.isRunning():
                    return
            except RuntimeError:
                self._mc_thread = None
                self._mc_worker = None

        class _MachineCodeWorker(QObject):
            finished = Signal(object)

            def run(self):
                try:
                    raw = get_machine_code()
                    display = get_machine_code_display()
                    self.finished.emit((raw, display))
                except Exception:
                    self.finished.emit(("", ""))

        self._mc_thread = QThread(self)
        self._mc_worker = _MachineCodeWorker()
        self._mc_worker.moveToThread(self._mc_thread)
        self._mc_thread.started.connect(self._mc_worker.run)
        self._mc_worker.finished.connect(self._on_machine_code_ready)
        self._mc_worker.finished.connect(self._mc_thread.quit)
        self._mc_worker.finished.connect(self._mc_worker.deleteLater)
        self._mc_thread.finished.connect(self._mc_thread.deleteLater)
        self._mc_thread.finished.connect(self._clear_machine_code_thread)
        self._mc_thread.start()

    def _clear_machine_code_thread(self):
        self._mc_thread = None
        self._mc_worker = None

    def _on_machine_code_ready(self, result):
        raw, display = result
        self._machine_code_raw = raw or ""
        if display:
            self.machine_code_edit.setText(display)
            self.machine_code_edit.setCursorPosition(0)
        else:
            self.machine_code_edit.setText("获取失败")

    def copy_machine_code(self):
        if not self._machine_code_raw:
            self.status_label.setText("机器码计算中，请稍候")
            return
        QApplication.clipboard().setText(self._machine_code_raw)
        self.status_label.setText("机器码已复制")

    def activate(self):
        license_key = self.license_edit.toPlainText().strip()
        result = self.license_manager.activate(license_key)
        self.status_label.setText(result.message)
        if result.ok:
            self.accept()

    def deactivate(self):
        reply = QMessageBox.question(
            self,
            "确认",
            "将清除本机授权文件，软件会恢复为未激活状态。是否继续？",
            QMessageBox.Yes | QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return
        ok = self.license_manager.deactivate()
        self.license_edit.clear()
        if ok:
            self.status_label.setText("已注销，本机已恢复未激活状态")
            parent = self.parent()
            if parent is not None and hasattr(parent, "license_result") and hasattr(parent, "update_button_states"):
                try:
                    parent.license_result = LicenseCheckResult(False, "未激活")
                    parent.update_button_states()
                    if hasattr(parent, "_lic_check_seq"):
                        parent._lic_check_seq += 1
                    if hasattr(parent, "start_license_check"):
                        QTimer.singleShot(0, parent.start_license_check)
                except Exception:
                    pass
        else:
            self.status_label.setText("注销失败：可能没有权限或文件被占用")


class MainWindow(QMainWindow):
    """主窗口"""
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("教师实训报告批量处理系统")
        self.setMinimumSize(1200, 700)
        
        # 初始化组件
        self.doc_parser = DocParser()
        self.file_renamer = FileRenamer()
        self.comment_generator = CommentGenerator()
        self.comment_writer = CommentWriter()
        self.excel_exporter = ExcelExporter()
        self.lab_name_fixer = LabNameFixer()

        self.license_manager = LicenseManager()
        self.license_result: Optional[LicenseCheckResult] = None
        self._activation_prompted = False
        self._lic_check_seq = 0
        self._lic_check_pending = False
        self._lic_thread: Optional[QThread] = None
        self._lic_worker: Optional[QObject] = None
        
        # 数据
        self.students: list[StudentInfo] = []
        self.folder_path: str = ""
        self.course_name: str = ""
        self.training_name: str = ""
        
        self.setup_ui()
        self.update_button_states()

        QTimer.singleShot(0, self.start_license_check)

    def start_license_check(self):
        if self._lic_thread is not None:
            try:
                if self._lic_thread.isRunning():
                    self._lic_check_pending = True
                    return
            except RuntimeError:
                self._lic_thread = None
                self._lic_worker = None

        self._lic_check_seq += 1
        seq = self._lic_check_seq

        class _LicenseCheckWorker(QObject):
            finished = Signal(object)

            def __init__(self, license_manager: LicenseManager):
                super().__init__()
                self.license_manager = license_manager

            def run(self):
                try:
                    self.finished.emit((seq, self.license_manager.check()))
                except Exception:
                    self.finished.emit((seq, LicenseCheckResult(False, "授权校验失败")))

        if hasattr(self, "status_bar"):
            msg = self.status_bar.currentMessage().strip() if self.status_bar.currentMessage() else ""
            if (not msg) or (msg == "就绪"):
                self.status_bar.showMessage("正在校验授权...")

        self._lic_thread = QThread(self)
        self._lic_worker = _LicenseCheckWorker(self.license_manager)
        self._lic_worker.moveToThread(self._lic_thread)
        self._lic_thread.started.connect(self._lic_worker.run)
        self._lic_worker.finished.connect(self._on_license_checked)
        self._lic_worker.finished.connect(self._lic_thread.quit)
        self._lic_worker.finished.connect(self._lic_worker.deleteLater)
        self._lic_thread.finished.connect(self._lic_thread.deleteLater)
        self._lic_thread.finished.connect(self._clear_license_thread)
        self._lic_thread.start()

    def _clear_license_thread(self):
        self._lic_thread = None
        self._lic_worker = None
        if self._lic_check_pending:
            self._lic_check_pending = False
            QTimer.singleShot(0, self.start_license_check)

    def _on_license_checked(self, result):
        try:
            r_seq, r_value = result
        except Exception:
            r_seq, r_value = self._lic_check_seq, result

        if int(r_seq) != int(self._lic_check_seq):
            return

        self.license_result = r_value
        self.update_button_states()

        if getattr(self.license_result, "ok", False):
            if hasattr(self, "status_bar"):
                msg = self.status_bar.currentMessage().strip() if self.status_bar.currentMessage() else ""
                if msg == "正在校验授权...":
                    self.status_bar.showMessage("就绪")
            return

        if not self._activation_prompted:
            self._activation_prompted = True
            dialog = ActivationDialog(self.license_manager, self)
            dialog.exec()
            QTimer.singleShot(0, self.start_license_check)
    
    def setup_ui(self):
        """设置界面"""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        main_layout = QHBoxLayout(central_widget)
        
        # 左侧：主操作区
        left_layout = QVBoxLayout()
        
        # 顶部任务区
        self.setup_task_area(left_layout)
        
        # 中央学生列表
        self.setup_student_table(left_layout)
        
        main_layout.addLayout(left_layout, stretch=4)
        
        # 右侧：功能按钮区
        self.setup_button_area(main_layout)
        
        # 底部状态栏
        self.setup_status_bar()
    
    def setup_task_area(self, parent_layout: QVBoxLayout):
        """设置顶部任务区"""
        task_group = QGroupBox("当前任务")
        task_layout = QVBoxLayout(task_group)
        
        # 文件夹选择行
        folder_layout = QHBoxLayout()
        folder_layout.addWidget(QLabel("实训文件夹:"))
        self.folder_label = QLabel("未选择")
        self.folder_label.setStyleSheet("color: gray;")
        folder_layout.addWidget(self.folder_label, stretch=1)
        self.select_folder_btn = QPushButton("选择文件夹")
        self.select_folder_btn.clicked.connect(self.select_folder)
        folder_layout.addWidget(self.select_folder_btn)
        task_layout.addLayout(folder_layout)
        
        # 课程信息行
        info_layout = QHBoxLayout()
        info_layout.addWidget(QLabel("课程名称:"))
        self.course_label = QLabel("-")
        info_layout.addWidget(self.course_label, stretch=1)
        info_layout.addWidget(QLabel("实训名称:"))
        self.training_label = QLabel("-")
        info_layout.addWidget(self.training_label, stretch=1)
        task_layout.addLayout(info_layout)
        
        parent_layout.addWidget(task_group)
    
    def setup_student_table(self, parent_layout: QVBoxLayout):
        """设置学生列表表格"""
        table_group = QGroupBox("学生列表")
        table_layout = QVBoxLayout(table_group)
        
        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(["学号", "姓名", "实验室名称", "文件名", "分数", "评语预览"])
        
        # 设置列宽
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.Stretch)
        header.setSectionResizeMode(4, QHeaderView.Fixed)
        header.resizeSection(4, 130)
        
        self.table.verticalHeader().setDefaultSectionSize(50)
        header.setSectionResizeMode(5, QHeaderView.Stretch)

        self.table.setItemDelegateForColumn(4, ScoreDelegate(self.table))
        
        # 设置选择模式
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        
        # 单击分数列即可编辑
        self.table.setEditTriggers(QAbstractItemView.AllEditTriggers)
        
        # 双击编辑评语
        self.table.cellDoubleClicked.connect(self.on_cell_double_clicked)
        
        # 连接分数编辑信号
        self.table.cellChanged.connect(self.on_score_changed)
        
        table_layout.addWidget(self.table)
        parent_layout.addWidget(table_group, stretch=1)
    
    def setup_button_area(self, parent_layout: QHBoxLayout):
        """设置右侧功能按钮区"""
        button_group = QGroupBox("操作")
        button_layout = QVBoxLayout(button_group)
        button_layout.setSpacing(10)

        self.activate_view_btn = QPushButton("激活/查看机器码")
        self.activate_view_btn.clicked.connect(self.open_activation_dialog)
        button_layout.addWidget(self.activate_view_btn)
        button_layout.addSpacing(10)
        
        # 解析文档按钮
        self.parse_btn = QPushButton("解析实训文档")
        self.parse_btn.clicked.connect(self.parse_documents)
        button_layout.addWidget(self.parse_btn)

        self.auto_score_btn = QPushButton("执行自动化评分")
        self.auto_score_btn.clicked.connect(self.apply_auto_scoring)
        button_layout.addWidget(self.auto_score_btn)
        
        # 重命名按钮
        self.rename_btn = QPushButton("统一重命名文件")
        self.rename_btn.clicked.connect(self.rename_files)
        button_layout.addWidget(self.rename_btn)
        
        # 批量修正实验室名称按钮
        self.fix_lab_name_btn = QPushButton("一键修正实验（训）室名称")
        self.fix_lab_name_btn.clicked.connect(self.fix_lab_name)
        button_layout.addWidget(self.fix_lab_name_btn)
        
        button_layout.addSpacing(20)
        
        # 生成评语按钮
        self.generate_btn = QPushButton("生成评语")
        self.generate_btn.clicked.connect(self.generate_comments)
        button_layout.addWidget(self.generate_btn)
        
        # 重新生成评语按钮
        self.regenerate_btn = QPushButton("重新生成评语")
        self.regenerate_btn.clicked.connect(self.regenerate_comments)
        button_layout.addWidget(self.regenerate_btn)
        
        button_layout.addSpacing(20)
        
        # 写入 Word 按钮
        self.write_btn = QPushButton("写入 Word 文档")
        self.write_btn.clicked.connect(self.write_to_word)
        button_layout.addWidget(self.write_btn)
        
        # 导出 Excel 按钮
        self.export_btn = QPushButton("导出成绩汇总 Excel")
        self.export_btn.clicked.connect(self.export_excel)
        button_layout.addWidget(self.export_btn)
        
        button_layout.addStretch()
        
        parent_layout.addWidget(button_group)
    
    def setup_status_bar(self):
        """设置状态栏"""
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("就绪")
    
    def update_button_states(self):
        """更新按钮状态"""
        has_folder = bool(self.folder_path)
        has_students = len(self.students) > 0
        has_scores = any(s.score is not None for s in self.students)
        has_comments = any(s.comment for s in self.students)

        can_auto_score = any(
            s.is_valid and (getattr(s, 'score_source', "") not in ("teacher", "manual"))
            for s in self.students
        )

        licensed = bool(getattr(self, "license_result", None) and self.license_result.ok)
        
        self.parse_btn.setEnabled(licensed and has_folder)
        self.auto_score_btn.setEnabled(licensed and has_students and can_auto_score)
        self.rename_btn.setEnabled(licensed and has_students)
        self.fix_lab_name_btn.setEnabled(licensed and has_folder)
        self.generate_btn.setEnabled(licensed and has_scores)
        self.regenerate_btn.setEnabled(licensed and has_comments)
        self.write_btn.setEnabled(licensed and has_comments)
        self.export_btn.setEnabled(licensed and has_students)

        if hasattr(self, "status_bar") and not licensed:
            msg = self.status_bar.currentMessage().strip() if self.status_bar.currentMessage() else ""
            if (not msg) or (msg == "就绪"):
                self.status_bar.showMessage("未激活：核心功能已禁用，请先激活")

    def open_activation_dialog(self):
        dialog = ActivationDialog(self.license_manager, self)
        dialog.exec()
        self.start_license_check()
    
    def select_folder(self):
        """选择文件夹"""
        folder = QFileDialog.getExistingDirectory(
            self, "选择实训报告文件夹", ""
        )
        
        if folder:
            if self.students:
                reply = QMessageBox.question(
                    self, "确认",
                    "选择新文件夹将清空当前数据，是否继续？",
                    QMessageBox.Yes | QMessageBox.No
                )
                if reply == QMessageBox.No:
                    return
            
            self.folder_path = folder
            self.folder_label.setText(folder)
            self.folder_label.setStyleSheet("color: black;")
            self.students = []
            self.table.setRowCount(0)
            self.update_button_states()
            self.status_bar.showMessage(f"已选择文件夹: {folder}")
            
            # 自动解析（仅在已激活时）
            if getattr(self, "license_result", None) and self.license_result.ok:
                self.parse_documents()
            else:
                self.status_bar.showMessage("未激活：请先激活后再解析实训文档")
    
    def parse_documents(self):
        """解析文档"""
        if not (getattr(self, "license_result", None) and self.license_result.ok):
            dialog = ActivationDialog(self.license_manager, self)
            dialog.exec()
            try:
                self.license_result = self.license_manager.check()
            except Exception:
                self.license_result = LicenseCheckResult(False, "未激活")
            self.update_button_states()
            if not (getattr(self, "license_result", None) and self.license_result.ok):
                return

        if not self.folder_path:
            QMessageBox.warning(self, "警告", "请先选择文件夹")
            return
        
        # 如果已有数据，提示用户
        if self.students:
            reply = QMessageBox.question(
                self, "确认",
                "重新解析将更新文档解析结果，但会保留已输入的分数和评语，确定要继续吗？",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply == QMessageBox.No:
                return
        
        self.status_bar.showMessage("正在解析文档...")
        
        # 保存旧数据（按文件路径索引）
        old_data = {}
        for student in self.students:
            old_data[student.file_path] = {
                'score': student.score,
                'score_source': getattr(student, 'score_source', ""),
                'comment': student.comment
            }
        
        # 重新解析
        self.students = self.doc_parser.parse_folder(self.folder_path)
        self.course_name, self.training_name = self.doc_parser.get_course_info()
        
        # 恢复旧数据
        for student in self.students:
            if student.file_path in old_data:
                old_score = old_data[student.file_path].get('score')
                old_source = old_data[student.file_path].get('score_source', "")
                if (student.score is None) and (getattr(student, 'score_source', "") != "teacher"):
                    if old_score is not None:
                        if old_source == "auto":
                            student.score = old_score
                            student.score_source = "auto"
                        else:
                            student.score = old_score
                            student.score_source = "manual"
                student.comment = old_data[student.file_path].get('comment', "")
        
        # 更新课程信息显示
        self.course_label.setText(self.course_name if self.course_name else "-")
        self.training_label.setText(self.training_name if self.training_name else "-")
        
        # 更新表格
        self.refresh_table()
        
        valid_count = sum(1 for s in self.students if s.is_valid)
        self.status_bar.showMessage(
            f"解析完成: 共 {len(self.students)} 份文档，{valid_count} 份有效"
        )
        self.update_button_states()

    def _reparse_documents_silent(self):
        if not self.folder_path:
            return

        self.status_bar.showMessage("正在重新解析文档刷新列表...")
        QApplication.processEvents()

        old_data = {}
        for student in self.students:
            old_data[student.file_path] = {
                'score': student.score,
                'score_source': getattr(student, 'score_source', ""),
                'comment': student.comment
            }

        self.students = self.doc_parser.parse_folder(self.folder_path)
        self.course_name, self.training_name = self.doc_parser.get_course_info()

        for student in self.students:
            if student.file_path in old_data:
                old_score = old_data[student.file_path].get('score')
                old_source = old_data[student.file_path].get('score_source', "")
                if (student.score is None) and (getattr(student, 'score_source', "") != "teacher"):
                    if old_score is not None:
                        if old_source == "auto":
                            student.score = old_score
                            student.score_source = "auto"
                        else:
                            student.score = old_score
                            student.score_source = "manual"
                student.comment = old_data[student.file_path].get('comment', "")

        self.course_label.setText(self.course_name if self.course_name else "-")
        self.training_label.setText(self.training_name if self.training_name else "-")
        self.refresh_table()
        self.update_button_states()

    def apply_auto_scoring(self):
        if not (getattr(self, "license_result", None) and self.license_result.ok):
            dialog = ActivationDialog(self.license_manager, self)
            dialog.exec()
            try:
                self.license_result = self.license_manager.check()
            except Exception:
                self.license_result = LicenseCheckResult(False, "未激活")
            self.update_button_states()
            if not (getattr(self, "license_result", None) and self.license_result.ok):
                return

        if not self.students:
            QMessageBox.warning(self, "警告", "请先解析文档")
            return

        self.auto_score_btn.setEnabled(False)
        self.status_bar.showMessage("正在执行自动化评分...")
        QApplication.processEvents()

        changed = self.doc_parser.apply_auto_scores(self.students)
        self.refresh_table()
        self.status_bar.showMessage(f"自动化评分完成: 共更新 {changed} 条分数")
        self.update_button_states()
    
    def refresh_table(self):
        """刷新表格数据"""
        # 阻止信号触发避免循环
        self.table.blockSignals(True)
        self.table.setRowCount(len(self.students))
        
        for row, student in enumerate(self.students):
            # 学号
            id_item = QTableWidgetItem(student.student_id)
            id_item.setFlags(id_item.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(row, 0, id_item)
            
            # 姓名
            name_item = QTableWidgetItem(student.student_name)
            name_item.setFlags(name_item.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(row, 1, name_item)
            
            # 实验室名称
            lab_name = getattr(student, 'lab_name', '')
            lab_item = QTableWidgetItem(lab_name)
            lab_item.setFlags(lab_item.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(row, 2, lab_item)
            
            # 文件名
            file_item = QTableWidgetItem(student.file_name)
            file_item.setFlags(file_item.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(row, 3, file_item)
            
            # 分数
            score_text = str(int(student.score)) if student.score is not None else ""
            score_item = QTableWidgetItem(score_text)
            score_item.setFont(QFont("微软雅黑", 10))
            score_item.setTextAlignment(Qt.AlignCenter)
            score_item.setFlags(score_item.flags() | Qt.ItemIsEditable)
            self.table.setItem(row, 4, score_item)

            score_source = getattr(student, 'score_source', "")
            if score_source == "teacher":
                score_item.setBackground(QColor(225, 255, 225))
            elif score_source == "auto":
                score_item.setBackground(QColor(235, 240, 255))
            elif score_source == "manual":
                score_item.setBackground(QColor(255, 245, 225))
            
            # 评语预览
            comment_preview = student.comment[:50] + "..." if len(student.comment) > 50 else student.comment
            comment_item = QTableWidgetItem(comment_preview)
            comment_item.setFlags(comment_item.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(row, 5, comment_item)
            
            # 标记无效行
            if not student.is_valid:
                for col in range(6):
                    item = self.table.item(row, col)
                    if item:
                        item.setBackground(QColor(255, 230, 230))
        
        # 恢复信号
        self.table.blockSignals(False)
    
    def on_score_changed(self, row: int, col: int):
        """分数变更处理"""
        if col != 4:
            return
        
        item = self.table.item(row, col)
        if not item:
            return
        
        text = item.text().strip()
        if not text:
            self.students[row].score = None
            if hasattr(self.students[row], 'score_source'):
                self.students[row].score_source = ""
        else:
            try:
                score = float(text)
                if 0 <= score <= 100:
                    self.students[row].score = score
                    if hasattr(self.students[row], 'score_source'):
                        self.students[row].score_source = "manual"
                    normalized = str(int(score))
                    if item.text() != normalized:
                        self.table.blockSignals(True)
                        item.setText(normalized)
                        self.table.blockSignals(False)
                else:
                    self.table.blockSignals(True)
                    QMessageBox.warning(self, "警告", "分数应在 0-100 之间")
                    item.setText("")
                    self.students[row].score = None
                    if hasattr(self.students[row], 'score_source'):
                        self.students[row].score_source = ""
                    self.table.blockSignals(False)
            except ValueError:
                self.table.blockSignals(True)
                QMessageBox.warning(self, "警告", "请输入有效的数字")
                item.setText("")
                self.students[row].score = None
                if hasattr(self.students[row], 'score_source'):
                    self.students[row].score_source = ""
                self.table.blockSignals(False)
        
        self.update_button_states()
    
    def on_cell_double_clicked(self, row: int, col: int):
        """双击单元格处理"""
        if col == 5:  # 评语列
            self.edit_comment(row)
    
    def edit_comment(self, row: int):
        """编辑评语"""
        student = self.students[row]
        dialog = CommentEditDialog(student, self)
        
        if dialog.exec() == QDialog.Accepted:
            student.comment = dialog.get_comment()
            self.refresh_table_row(row)
            self.update_button_states()
    
    def refresh_table_row(self, row: int):
        """刷新单行数据"""
        student = self.students[row]
        
        # 更新评语预览
        comment_preview = student.comment[:50] + "..." if len(student.comment) > 50 else student.comment
        comment_item = QTableWidgetItem(comment_preview)
        comment_item.setFlags(comment_item.flags() & ~Qt.ItemIsEditable)
        self.table.setItem(row, 5, comment_item)
    
    def rename_files(self):
        """重命名文件"""
        reply = QMessageBox.question(
            self, "确认",
            "确定要按规范格式重命名所有文件吗？",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.No:
            return
        
        results = self.file_renamer.rename_files(self.students)
        
        success_count = sum(1 for r in results if r[2])
        self.status_bar.showMessage(f"重命名完成: {success_count}/{len(results)} 成功")
        
        # 刷新表格
        self.refresh_table()
    
    def generate_comments(self):
        """生成评语"""
        count = self.comment_generator.generate_for_all(self.students)
        self.refresh_table()
        self.status_bar.showMessage(f"已为 {count} 名学生生成评语")
        self.update_button_states()
    
    def regenerate_comments(self):
        """重新生成评语"""
        selected_rows = self.table.selectionModel().selectedRows()
        
        if selected_rows:
            # 只重新生成选中的
            for index in selected_rows:
                row = index.row()
                self.comment_generator.regenerate_for_student(self.students[row])
                self.refresh_table_row(row)
            self.status_bar.showMessage(f"已重新生成 {len(selected_rows)} 条评语")
        else:
            # 全部重新生成
            count = self.comment_generator.generate_for_all(self.students)
            self.refresh_table()
            self.status_bar.showMessage(f"已重新生成 {count} 条评语")
    
    def write_to_word(self):
        """写入 Word 文档"""
        reply = QMessageBox.question(
            self, "确认",
            "确定要将评语写入所有 Word 文档吗？",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.No:
            return
        
        self.status_bar.showMessage("正在写入 Word 文档...")
        
        results = []
        total = len(self.students)
        for idx, student in enumerate(self.students, start=1):
            if student.comment:
                success, message = self.comment_writer.write_to_document(student)
                results.append((student.file_name, success, message))
            else:
                results.append((student.file_name, True, ""))

            if idx % 5 == 0 or idx == total:
                self.status_bar.showMessage(f"正在写入 Word 文档... {idx}/{total}")
                from PySide6.QtWidgets import QApplication
                QApplication.processEvents()
        
        success_count = sum(1 for r in results if r[1])
        self.status_bar.showMessage(f"写入完成: {success_count}/{len(results)} 成功")
        
        if success_count < len(results):
            # 显示详细的失败信息
            failed_info = []
            for r in results:
                if not r[1]:
                    failed_info.append(f"{r[0]}: {r[2]}")
            
            QMessageBox.warning(
                self, "部分失败",
                f"以下文件写入失败:\n" + "\n".join(failed_info[:10])
            )
    
    def export_excel(self):
        """导出 Excel"""
        default_name = self.excel_exporter.generate_filename(
            self.course_name, self.training_name
        )
        default_path = os.path.join(self.folder_path, default_name)
        
        file_path, _ = QFileDialog.getSaveFileName(
            self, "保存 Excel 文件",
            default_path,
            "Excel 文件 (*.xlsx)"
        )
        
        if not file_path:
            return
        
        success, message = self.excel_exporter.export(
            self.students,
            file_path,
            self.course_name,
            self.training_name
        )
        
        if success:
            self.status_bar.showMessage(message)
            QMessageBox.information(self, "成功", message)
        else:
            self.status_bar.showMessage(message)
            QMessageBox.warning(self, "失败", message)
    
    def fix_lab_name(self):
        """批量修正实验（训）室名称"""
        if not self.folder_path:
            QMessageBox.warning(self, "警告", "请先选择文件夹")
            return
        
        new_lab_name, ok = QInputDialog.getText(
            self,
            "修正实验（训）室名称",
            "请输入新的实验（训）室名称：\n\n"
            "注意：\n"
            "1. 将批量修改所有 Word 文档中的实验（训）室名称\n"
            "2. 仅修改字段内容，不改变其他任何内容和格式\n"
            "3. 修改后将原地保存文件\n\n"
            "新名称："
        )
        
        if not ok:
            return
        
        new_lab_name = new_lab_name.strip()
        if not new_lab_name:
            QMessageBox.warning(self, "警告", "实验室名称不能为空")
            return
        
        reply = QMessageBox.question(
            self, "确认",
            f"确定要将所有文档的实验（训）室名称统一修改为：\n\n{new_lab_name}\n\n"
            "此操作将直接修改文件，无法撤销！",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.No:
            return
        
        self.status_bar.showMessage("正在批量修正实验室名称...")
        
        success_files, failed_files = self.lab_name_fixer.batch_fix_lab_name(
            self.folder_path, new_lab_name
        )
        
        total = len(success_files) + len(failed_files)
        success_count = len(success_files)
        
        self.status_bar.showMessage(
            f"修正完成: {success_count}/{total} 成功"
        )
        
        if failed_files:
            failed_info = "\n".join(failed_files[:20])
            if len(failed_files) > 20:
                failed_info += f"\n... 还有 {len(failed_files) - 20} 个文件"
            
            QMessageBox.warning(
                self, "部分失败",
                f"成功: {success_count} 个文件\n失败: {len(failed_files)} 个文件\n\n"
                f"失败文件:\n{failed_info}"
            )
        else:
            QMessageBox.information(
                self, "成功",
                f"已成功修正 {success_count} 个文件的实验（训）室名称"
            )

        if success_count > 0:
            self._reparse_documents_silent()
