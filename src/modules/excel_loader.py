"""
Excel 导入模块
负责读取单次实训成绩 Excel 文件
"""
import pandas as pd
from typing import Dict, List, Tuple, Optional, Any


class ExcelLoader:
    """Excel 文件加载器"""
    
    REQUIRED_COLUMNS = ["学号", "姓名", "分数"]
    
    def __init__(self):
        pass
    
    def load(self, file_path: str) -> Tuple[bool, str, List[Dict[str, Any]]]:
        """
        加载 Excel 文件并提取学生成绩数据
        
        Args:
            file_path: Excel 文件路径
            
        Returns:
            (是否成功, 消息, 数据列表)
            数据列表格式: [{"student_id": "xxx", "name": "xxx", "score": 85}, ...]
        """
        try:
            df = pd.read_excel(file_path)
        except FileNotFoundError:
            return False, "文件不存在", []
        except PermissionError:
            return False, "文件被占用，请关闭 Excel 后重试", []
        except Exception as e:
            return False, f"读取文件失败: {str(e)}", []
        
        missing_cols = [col for col in self.REQUIRED_COLUMNS if col not in df.columns]
        if missing_cols:
            return False, f"Excel 缺少必需列: {', '.join(missing_cols)}", []
        
        records: List[Dict[str, Any]] = []
        
        for _, row in df.iterrows():
            raw_id = row["学号"]
            if pd.notna(raw_id):
                if isinstance(raw_id, float):
                    student_id = str(int(raw_id))
                else:
                    student_id = str(raw_id).strip()
                    if student_id.endswith(".0"):
                        student_id = student_id[:-2]
            else:
                student_id = ""
            
            if not student_id or student_id.lower() == "nan":
                continue
            
            name = str(row["姓名"]).strip() if pd.notna(row["姓名"]) else ""
            
            score: Optional[int] = None
            if pd.notna(row["分数"]):
                try:
                    score = int(float(row["分数"]))
                except (ValueError, TypeError):
                    score = None
            
            records.append({
                "student_id": student_id,
                "name": name,
                "score": score
            })
        
        if not records:
            return False, "Excel 中没有有效的学生数据", []
        
        return True, f"成功读取 {len(records)} 条记录", records
