"""
课程名称批量修正模块
仿照实验室名称修正模块实现
"""
import os
import io
import re
import stat
import time
from typing import Tuple, List
from docx import Document


class CourseNameFixer:
    """课程名称批量修正器"""
    
    FIELD_PATTERN = re.compile(r"课\s*程\s*名\s*称\s*[：:]")
    
    def __init__(self):
        pass
    
    def extract_course_name(self, file_path: str) -> str:
        """
        提取文档中的课程名称
        
        Args:
            file_path: 文件路径
        
        Returns:
            课程名称，如果未找到则返回空字符串
        """
        try:
            with open(file_path, "rb") as f:
                data = f.read()
            doc = Document(io.BytesIO(data))
            return self.extract_course_name_from_doc(doc)
        except Exception:
            return ""
    
    def extract_course_name_from_doc(self, doc) -> str:
        """
        从已打开的 Document 对象中提取课程名称
        
        Args:
            doc: Document 对象
        
        Returns:
            课程名称，如果未找到则返回空字符串
        """
        for paragraph in self._iter_all_paragraphs(doc):
            course_name = self._extract_from_paragraph(paragraph)
            if course_name:
                return course_name
        
        return ""
    
    def _extract_from_paragraph(self, paragraph) -> str:
        """
        从段落中提取课程名称
        
        Args:
            paragraph: 段落对象
        
        Returns:
            课程名称，如果未找到则返回空字符串
        """
        full_text = paragraph.text

        m = self.FIELD_PATTERN.search(full_text)
        if not m:
            return ""

        rest = full_text[m.end():]
        cut = len(rest)

        m_nl = re.search(r"[\r\n]", rest)
        if m_nl and m_nl.start() < cut:
            cut = m_nl.start()

        m_space = re.search(r"\s{2,}", rest)
        if m_space and m_space.start() < cut:
            cut = m_space.start()

        return rest[:cut].strip()
    
    def batch_fix_course_name(self, folder_path: str, new_course_name: str) -> Tuple[List[str], List[str]]:
        """
        批量修正课程名称
        
        Args:
            folder_path: 文件夹路径
            new_course_name: 新的课程名称（已去除首尾空格）
        
        Returns:
            (success_files, failed_files): 成功和失败的文件列表
        """
        success_files = []
        failed_files = []
        
        if not os.path.isdir(folder_path):
            return success_files, failed_files
        
        docx_files = []
        for filename in os.listdir(folder_path):
            if filename.endswith('.docx') and not filename.startswith('~$'):
                docx_files.append(os.path.join(folder_path, filename))
        
        for file_path in docx_files:
            try:
                success = self._fix_single_file(file_path, new_course_name)
                if success:
                    success_files.append(os.path.basename(file_path))
                else:
                    failed_files.append(os.path.basename(file_path))
            except Exception as e:
                failed_files.append(f"{os.path.basename(file_path)} ({type(e).__name__})")
        
        return success_files, failed_files
    
    def _fix_single_file(self, file_path: str, new_course_name: str) -> bool:
        """
        修正单个文件的课程名称
        
        Args:
            file_path: 文件路径
            new_course_name: 新的课程名称
        
        Returns:
            是否成功
        """
        file_path = os.path.normpath(file_path)
        
        try:
            with open(file_path, "rb") as f:
                data = f.read()
        except Exception as e:
            print(f"读取文件失败 {file_path}: {e}")
            return False
        
        try:
            doc = Document(io.BytesIO(data))
        except Exception as e:
            print(f"解析文档失败 {file_path}: {e}")
            return False
        
        modified = False

        for paragraph in self._iter_all_paragraphs(doc):
            if self._contains_field(paragraph.text):
                self._replace_text_after_label(paragraph, new_course_name)
                modified = True
        
        if not modified:
            print(f"未找到字段 {file_path}")
            return False
        
        tmp_path = f"{file_path}.tmp"
        
        for attempt in range(3):
            try:
                doc.save(tmp_path)
                try:
                    os.chmod(file_path, stat.S_IWRITE)
                except Exception:
                    pass
                os.replace(tmp_path, file_path)
                return True
            except Exception as e:
                print(f"保存失败 (尝试 {attempt + 1}/3) {file_path}: {type(e).__name__}: {e}")
                try:
                    if os.path.exists(tmp_path):
                        os.remove(tmp_path)
                except Exception:
                    pass
                if attempt < 2:
                    time.sleep(0.4 * (attempt + 1))
                    continue
        
        print(f"保存最终失败 {file_path}")
        return False
    
    def _contains_field(self, text: str) -> bool:
        """
        检查文本是否包含字段名
        
        Args:
            text: 文本
        
        Returns:
            是否包含字段名
        """
        return self.FIELD_PATTERN.search(text or "") is not None
    
    def _replace_text_after_label(self, paragraph, new_course_name: str):
        """
        替换字段名后面的内容，保留原格式
        
        Args:
            paragraph: 段落对象
            new_course_name: 新的课程名称
        """
        full_text = paragraph.text
        m = self.FIELD_PATTERN.search(full_text)
        if not m:
            return

        rest = full_text[m.end():]
        cut = len(rest)

        m_nl = re.search(r"[\r\n]", rest)
        if m_nl and m_nl.start() < cut:
            cut = m_nl.start()

        m_space = re.search(r"\s{2,}", rest)
        if m_space and m_space.start() < cut:
            cut = m_space.start()

        new_text = full_text[:m.end()] + new_course_name + rest[cut:]

        if not paragraph.runs:
            paragraph.add_run(new_text)
            return

        first_run = paragraph.runs[0]
        for i in range(len(paragraph.runs)):
            paragraph.runs[i].text = ""
        first_run.text = new_text

    def _iter_all_paragraphs(self, doc):
        def iter_table_paragraphs(table):
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        yield p
                    for t in getattr(cell, "tables", []):
                        yield from iter_table_paragraphs(t)

        for p in doc.paragraphs:
            yield p
        for t in doc.tables:
            yield from iter_table_paragraphs(t)

        for section in getattr(doc, "sections", []):
            for attr in (
                "header",
                "footer",
                "first_page_header",
                "first_page_footer",
                "even_page_header",
                "even_page_footer",
            ):
                part = getattr(section, attr, None)
                if part is None:
                    continue
                for p in getattr(part, "paragraphs", []):
                    yield p
                for t in getattr(part, "tables", []):
                    yield from iter_table_paragraphs(t)
