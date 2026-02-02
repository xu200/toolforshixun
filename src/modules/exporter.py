"""
汇总导出模块
负责将合并后的成绩数据导出为 Excel 文件
"""
import pandas as pd
from typing import Dict, Any, Tuple
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side


class SummaryExporter:
    """汇总导出器"""
    
    COLUMN_NAMES = [
        "学号",
        "姓名",
        "第一次实训分数",
        "第二次实训分数",
        "第三次实训分数",
        "第四次实训分数",
        "第五次实训分数",
        "第六次实训分数",
        "第七次实训分数",
        "第八次实训分数",
    ]
    
    EXPERIMENT_INDEX_MAP = {
        1: "第一次实训分数",
        2: "第二次实训分数",
        3: "第三次实训分数",
        4: "第四次实训分数",
        5: "第五次实训分数",
        6: "第六次实训分数",
        7: "第七次实训分数",
        8: "第八次实训分数",
    }
    
    def __init__(self):
        pass
    
    def export(self, students: Dict[str, Dict[str, Any]], output_path: str) -> Tuple[bool, str]:
        """
        导出汇总 Excel
        
        Args:
            students: 学生数据字典
            output_path: 输出文件路径
            
        Returns:
            (是否成功, 消息)
        """
        if not students:
            return False, "没有数据可导出"
        
        try:
            rows = []
            
            for student_id in sorted(students.keys()):
                student = students[student_id]
                row = {
                    "学号": student_id,
                    "姓名": student["name"],
                }
                
                for exp_idx, col_name in self.EXPERIMENT_INDEX_MAP.items():
                    score = student["scores"].get(exp_idx)
                    row[col_name] = score if score is not None else ""
                
                rows.append(row)
            
            df = pd.DataFrame(rows, columns=self.COLUMN_NAMES)
            
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='成绩汇总')
                
                workbook = writer.book
                worksheet = writer.sheets['成绩汇总']
                
                self._apply_styles(worksheet, len(rows))
            
            return True, f"成功导出 {len(rows)} 名学生的成绩汇总"
            
        except PermissionError:
            return False, "文件被占用，请关闭后重试"
        except Exception as e:
            return False, f"导出失败: {str(e)}"
    
    def _apply_styles(self, worksheet, data_rows: int):
        """应用样式"""
        column_widths = {
            'A': 15,
            'B': 12,
            'C': 18,
            'D': 18,
            'E': 18,
            'F': 18,
            'G': 18,
            'H': 18,
            'I': 18,
            'J': 18,
        }
        
        for col, width in column_widths.items():
            worksheet.column_dimensions[col].width = width
        
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        for row in range(1, data_rows + 2):
            for col in range(1, len(self.COLUMN_NAMES) + 1):
                cell = worksheet.cell(row=row, column=col)
                cell.border = thin_border
                cell.font = Font(name="宋体", size=10)
                cell.alignment = Alignment(horizontal="center", vertical="center")
        
        for col in range(1, len(self.COLUMN_NAMES) + 1):
            cell = worksheet.cell(row=1, column=col)
            cell.font = Font(name="宋体", size=11, bold=True)
