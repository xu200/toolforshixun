"""
文件重命名模块
将实训报告文件按规范格式重命名
"""
import os
import shutil
from typing import List, Tuple
from .doc_parser import StudentInfo


class FileRenamer:
    """文件重命名器"""
    
    def __init__(self, create_copy: bool = False):
        """
        初始化重命名器
        
        Args:
            create_copy: True 则创建新文件，False 则覆盖原文件名
        """
        self.create_copy = create_copy
    
    def rename_files(self, students: List[StudentInfo]) -> List[Tuple[str, str, bool, str]]:
        """
        批量重命名文件
        
        Args:
            students: 学生信息列表
            
        Returns:
            重命名结果列表: [(原文件名, 新文件名, 是否成功, 错误信息)]
        """
        results = []
        
        for student in students:
            if not student.is_valid:
                results.append((
                    student.file_name,
                    "",
                    False,
                    f"跳过无效文件: {student.error_message}"
                ))
                continue
            
            new_filename = student.get_new_filename()
            
            # 如果文件名相同，跳过
            if new_filename == student.file_name:
                results.append((
                    student.file_name,
                    new_filename,
                    True,
                    "文件名已规范"
                ))
                continue
            
            result = self._rename_single_file(student, new_filename)
            results.append(result)
        
        return results
    
    def _rename_single_file(
        self, 
        student: StudentInfo, 
        new_filename: str
    ) -> Tuple[str, str, bool, str]:
        """
        重命名单个文件
        
        Args:
            student: 学生信息
            new_filename: 新文件名
            
        Returns:
            (原文件名, 新文件名, 是否成功, 消息)
        """
        try:
            old_path = student.file_path
            dir_path = os.path.dirname(old_path)
            new_path = os.path.join(dir_path, new_filename)
            
            # 检查目标文件是否已存在
            if os.path.exists(new_path) and old_path != new_path:
                return (
                    student.file_name,
                    new_filename,
                    False,
                    "目标文件已存在"
                )
            
            if self.create_copy:
                shutil.copy2(old_path, new_path)
            else:
                os.rename(old_path, new_path)
                # 更新学生信息中的路径
                student.file_path = new_path
                student.file_name = new_filename
            
            return (
                student.file_name if self.create_copy else os.path.basename(old_path),
                new_filename,
                True,
                "重命名成功"
            )
            
        except Exception as e:
            return (
                student.file_name,
                new_filename,
                False,
                f"重命名失败: {str(e)}"
            )
    
    def preview_rename(self, students: List[StudentInfo]) -> List[Tuple[str, str]]:
        """
        预览重命名结果（不实际执行）
        
        Args:
            students: 学生信息列表
            
        Returns:
            预览列表: [(原文件名, 新文件名)]
        """
        previews = []
        
        for student in students:
            if student.is_valid:
                new_filename = student.get_new_filename()
                previews.append((student.file_name, new_filename))
            else:
                previews.append((student.file_name, f"[无法重命名: {student.error_message}]"))
        
        return previews
