"""
教师实训报告批量处理系统
程序入口
"""
import sys
from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QFont, QIcon

from src.main_window import MainWindow
from src.styles import MODERN_STYLE


def main():
    app = QApplication(sys.argv)
    
    # 设置默认字体
    font = QFont("微软雅黑", 9)
    app.setFont(font)
    
    # 设置应用样式
    app.setStyle("Fusion")
    
    # 应用现代化样式表
    app.setStyleSheet(MODERN_STYLE)
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
