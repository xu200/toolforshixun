"""
数据合并模块
负责将多次实训成绩合并到统一的数据结构中
"""
from typing import Dict, List, Any, Optional


class ScoreMerger:
    """成绩合并器"""
    
    MAX_EXPERIMENTS = 8
    
    def __init__(self):
        self.students: Dict[str, Dict[str, Any]] = {}
        self.imported_files: Dict[int, str] = {}
    
    def reset(self):
        """重置所有数据"""
        self.students = {}
        self.imported_files = {}
    
    def merge(self, experiment_index: int, records: List[Dict[str, Any]], file_name: str) -> int:
        """
        合并一次实训的成绩数据
        
        Args:
            experiment_index: 实训编号 (1-8)
            records: 学生成绩记录列表
            file_name: 导入的文件名
            
        Returns:
            合并的记录数
        """
        if experiment_index < 1 or experiment_index > self.MAX_EXPERIMENTS:
            raise ValueError(f"实训编号必须在 1-{self.MAX_EXPERIMENTS} 之间")
        
        merged_count = 0
        
        for record in records:
            student_id = record["student_id"]
            name = record["name"]
            score = record["score"]
            
            if student_id not in self.students:
                self.students[student_id] = {
                    "name": name,
                    "scores": {}
                }
            
            self.students[student_id]["scores"][experiment_index] = score
            merged_count += 1
        
        self.imported_files[experiment_index] = file_name
        
        return merged_count
    
    def is_imported(self, experiment_index: int) -> bool:
        """检查某次实训是否已导入"""
        return experiment_index in self.imported_files
    
    def get_imported_file(self, experiment_index: int) -> Optional[str]:
        """获取某次实训导入的文件名"""
        return self.imported_files.get(experiment_index)
    
    def get_all_students(self) -> Dict[str, Dict[str, Any]]:
        """获取所有学生数据"""
        return self.students
    
    def get_student_count(self) -> int:
        """获取学生总数"""
        return len(self.students)
    
    def get_imported_count(self) -> int:
        """获取已导入的实训次数"""
        return len(self.imported_files)
