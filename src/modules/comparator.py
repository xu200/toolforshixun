"""
成绩对比模块
负责对比人工 Excel 和系统 Excel 的实训成绩
"""
import pandas as pd
from typing import Dict, List, Tuple, Optional, Any


class ScoreComparator:
    """成绩对比器"""
    
    COLUMN_MAPPING = {
        "第一次实训成绩": "第一次实训分数",
        "第二次实训成绩": "第二次实训分数",
        "第三次实训成绩": "第三次实训分数",
        "第四次实训成绩": "第四次实训分数",
        "第五次实训成绩": "第五次实训分数",
        "第六次实训成绩": "第六次实训分数",
        "第七次实训成绩": "第七次实训分数",
        "第八次实训成绩": "第八次实训分数",
    }
    
    def __init__(self):
        self.manual_df: Optional[pd.DataFrame] = None
        self.auto_files: List[str] = []
        self.auto_map: Dict[int, Dict[str, Optional[int]]] = {}
        self.mismatches: List[Tuple[int, str, Any, Any]] = []
    
    def load_manual_excel(self, file_path: str) -> Tuple[bool, str]:
        """
        加载人工成绩 Excel
        
        Returns:
            (是否成功, 消息)
        """
        try:
            self.manual_df = pd.read_excel(file_path)
        except FileNotFoundError:
            return False, "文件不存在"
        except PermissionError:
            return False, "文件被占用，请关闭 Excel 后重试"
        except Exception as e:
            return False, f"读取文件失败: {str(e)}"
        
        if "学号" not in self.manual_df.columns:
            return False, "人工 Excel 缺少「学号」列"
        
        manual_cols = [col for col in self.COLUMN_MAPPING.keys() if col in self.manual_df.columns]
        if not manual_cols:
            return False, "人工 Excel 中未找到实训成绩列"
        
        return True, f"成功加载人工 Excel，共 {len(self.manual_df)} 行，{len(manual_cols)} 列实训成绩"
    
    def load_auto_excel(self, file_path: str) -> Tuple[bool, str]:
        """
        加载系统成绩 Excel（支持多次导入，数据会合并）
        
        Returns:
            (是否成功, 消息)
        """
        try:
            auto_df = pd.read_excel(file_path)
        except FileNotFoundError:
            return False, "文件不存在"
        except PermissionError:
            return False, "文件被占用，请关闭 Excel 后重试"
        except Exception as e:
            return False, f"读取文件失败: {str(e)}"
        
        if "学号" not in auto_df.columns:
            return False, "系统 Excel 缺少「学号」列"
        
        new_count = 0
        for _, row in auto_df.iterrows():
            try:
                sid = self._normalize_sid(row["学号"])
                if sid is None:
                    continue
            except (ValueError, TypeError):
                continue
            
            if sid not in self.auto_map:
                self.auto_map[sid] = {}
                new_count += 1
            
            for auto_col in self.COLUMN_MAPPING.values():
                if auto_col in auto_df.columns:
                    score = self._parse_score(row.get(auto_col))
                    if score is not None:
                        self.auto_map[sid][auto_col] = score
        
        self.auto_files.append(file_path)
        
        return True, f"成功加载系统 Excel，新增 {new_count} 名学生，当前共 {len(self.auto_map)} 名"
    
    def clear_auto_data(self):
        """清空系统成绩数据"""
        self.auto_map = {}
        self.auto_files = []
    
    def get_auto_file_count(self) -> int:
        """获取已导入的系统 Excel 数量"""
        return len(self.auto_files)
    
    def compare(self) -> Tuple[bool, str, int]:
        """
        执行对比
        
        Returns:
            (是否成功, 消息, 不一致数量)
        """
        if self.manual_df is None:
            return False, "请先导入人工成绩 Excel", 0
        if not self.auto_map:
            return False, "请先导入系统成绩 Excel", 0
        
        self.mismatches = []
        
        for row_idx, row in self.manual_df.iterrows():
            try:
                sid = self._normalize_sid(row["学号"])
                if sid is None:
                    continue
            except (ValueError, TypeError):
                continue
            
            if sid not in self.auto_map:
                continue
            
            auto_scores = self.auto_map[sid]
            
            for manual_col, auto_col in self.COLUMN_MAPPING.items():
                if manual_col not in self.manual_df.columns:
                    continue
                if auto_col not in auto_scores:
                    continue
                
                manual_score = self._parse_score(row.get(manual_col))
                auto_score = auto_scores.get(auto_col)
                
                if manual_score is None or auto_score is None:
                    continue
                
                if manual_score != auto_score:
                    self.mismatches.append((row_idx, manual_col, manual_score, auto_score))
        
        return True, f"对比完成，发现 {len(self.mismatches)} 处成绩不一致", len(self.mismatches)
    
    def get_mismatches(self) -> List[Tuple[int, str, Any, Any]]:
        """获取不一致列表: [(行索引, 列名, 人工分数, 系统分数), ...]"""
        return self.mismatches
    
    def get_manual_df(self) -> Optional[pd.DataFrame]:
        """获取人工 Excel DataFrame"""
        return self.manual_df
    
    def _normalize_sid(self, sid) -> Optional[int]:
        """标准化学号为整数"""
        if pd.isna(sid):
            return None
        try:
            return int(float(sid))
        except (ValueError, TypeError):
            return None
    
    def _parse_score(self, value) -> Optional[int]:
        """解析分数为整数"""
        if pd.isna(value):
            return None
        try:
            return int(float(value))
        except (ValueError, TypeError):
            return None
