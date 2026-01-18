"""
现代化 GUI 样式表
采用蓝绿色系专业配色方案
"""

MODERN_STYLE = """
/* 全局样式 - 浅灰蓝背景 */
QMainWindow {
    background-color: #f0f4f8;
}

QWidget {
    font-family: "Microsoft YaHei", "微软雅黑", sans-serif;
}

/* 分组框样式 */
QGroupBox {
    font-size: 14px;
    font-weight: bold;
    color: #1a365d;
    border: 1px solid #cbd5e0;
    border-radius: 10px;
    margin-top: 14px;
    padding-top: 12px;
    background-color: #ffffff;
}

QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 15px;
    padding: 0 10px;
    background-color: #ffffff;
    color: #2b6cb0;
}

/* 按钮样式 - 主色调蓝色 */
QPushButton {
    background-color: #4299e1;
    color: white;
    border: none;
    border-radius: 8px;
    padding: 10px 20px;
    font-size: 13px;
    font-weight: 500;
    min-height: 38px;
}

QPushButton:hover {
    background-color: #3182ce;
}

QPushButton:pressed {
    background-color: #2b6cb0;
}

QPushButton:disabled {
    background-color: #a0aec0;
    color: #e2e8f0;
}

/* 成功操作按钮 - 绿色 */
QPushButton#primaryBtn {
    background-color: #48bb78;
}

QPushButton#primaryBtn:hover {
    background-color: #38a169;
}

/* 危险操作按钮 - 红色 */
QPushButton#dangerBtn {
    background-color: #fc8181;
}

QPushButton#dangerBtn:hover {
    background-color: #f56565;
}

/* 表格样式 */
QTableWidget {
    background-color: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 8px;
    gridline-color: #edf2f7;
    selection-background-color: #bee3f8;
    selection-color: #1a365d;
    alternate-background-color: #f7fafc;
}

QTableWidget::item {
    padding: 8px;
    border-bottom: 1px solid #edf2f7;
    color: #2d3748;
    font-size: 12px;
}

QDoubleSpinBox {
    font-size: 12px;
}

QTableWidget::item:selected {
    background-color: #bee3f8;
    color: #1a365d;
}

QTableWidget::item:hover {
    background-color: #e6fffa;
}

/* 表头样式 - 深蓝色 */
QHeaderView::section {
    background-color: #2b6cb0;
    color: white;
    padding: 12px;
    border: none;
    font-weight: bold;
    font-size: 13px;
}

QHeaderView::section:first {
    border-top-left-radius: 8px;
}

QHeaderView::section:last {
    border-top-right-radius: 8px;
}

/* 标签样式 */
QLabel {
    color: #2d3748;
    font-size: 13px;
}

QLabel#titleLabel {
    font-size: 18px;
    font-weight: bold;
    color: #1a365d;
}

QLabel#pathLabel {
    color: #718096;
    font-size: 12px;
}

/* 状态栏样式 - 深蓝灰色 */
QStatusBar {
    background-color: #2d3748;
    color: #e2e8f0;
    font-size: 12px;
    padding: 6px;
}

/* 文本编辑框 */
QTextEdit {
    border: 2px solid #e2e8f0;
    border-radius: 8px;
    padding: 12px;
    background-color: #ffffff;
    font-size: 13px;
    color: #2d3748;
}

QTextEdit:focus {
    border: 2px solid #4299e1;
}

/* 对话框 */
QDialog {
    background-color: #f7fafc;
}

QDialogButtonBox QPushButton {
    min-width: 90px;
}

/* 滚动条样式 */
QScrollBar:vertical {
    background-color: #edf2f7;
    width: 10px;
    border-radius: 5px;
    margin: 2px;
}

QScrollBar::handle:vertical {
    background-color: #a0aec0;
    border-radius: 5px;
    min-height: 30px;
}

QScrollBar::handle:vertical:hover {
    background-color: #718096;
}

QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}

QScrollBar:horizontal {
    background-color: #edf2f7;
    height: 10px;
    border-radius: 5px;
    margin: 2px;
}

QScrollBar::handle:horizontal {
    background-color: #a0aec0;
    border-radius: 5px;
    min-width: 30px;
}

QScrollBar::handle:horizontal:hover {
    background-color: #718096;
}

QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
    width: 0px;
}

/* 消息框样式 */
QMessageBox {
    background-color: #ffffff;
}

QMessageBox QLabel {
    font-size: 13px;
    color: #2d3748;
}

/* 输入框样式 */
QLineEdit {
    border: 2px solid #e2e8f0;
    border-radius: 6px;
    padding: 8px 12px;
    background-color: #ffffff;
    font-size: 13px;
    color: #2d3748;
}

QLineEdit:focus {
    border: 2px solid #4299e1;
}

/* 工具提示 */
QToolTip {
    background-color: #2d3748;
    color: #ffffff;
    border: none;
    border-radius: 4px;
    padding: 6px 10px;
    font-size: 12px;
}
"""
