"""
评语生成与写回模块
生成评语并写入 Word 文档
"""
import io
import os
import stat
import time
import re
from typing import List, Dict, Set, Tuple, Optional
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

from .comment_rules import generate_comment, get_level_by_score
from .doc_parser import StudentInfo


class CommentGenerator:
    """评语生成器"""
    
    def __init__(self):
        # 按等级记录已使用的评语组合，避免重复
        self.used_combinations: Dict[str, Set[Tuple[int, int, int]]] = {}
    
    def reset_combinations(self):
        """重置已使用的组合记录"""
        self.used_combinations = {}
    
    def generate_for_student(self, student: StudentInfo) -> str:
        """
        为单个学生生成评语
        
        Args:
            student: 学生信息
            
        Returns:
            生成的评语
        """
        if student.score is None:
            return ""
        
        level = get_level_by_score(student.score)
        
        # 获取该等级已使用的组合
        if level not in self.used_combinations:
            self.used_combinations[level] = set()
        
        comment, combination = generate_comment(
            student.score, 
            self.used_combinations[level]
        )
        
        # 记录使用的组合
        self.used_combinations[level].add(combination)
        
        student.comment = comment
        return comment
    
    def generate_for_all(self, students: List[StudentInfo]) -> int:
        """
        为所有学生批量生成评语
        
        Args:
            students: 学生信息列表
            
        Returns:
            成功生成评语的数量
        """
        self.reset_combinations()
        count = 0
        
        for student in students:
            if student.score is not None:
                self.generate_for_student(student)
                count += 1
        
        return count
    
    def regenerate_for_student(self, student: StudentInfo) -> str:
        """
        为单个学生重新生成评语
        
        Args:
            student: 学生信息
            
        Returns:
            新生成的评语
        """
        if student.score is None:
            return ""
        
        level = get_level_by_score(student.score)
        
        # 不使用已记录的组合，强制生成新的
        comment, _ = generate_comment(student.score)
        student.comment = comment
        return comment


class CommentWriter:
    """评语写入器"""
    
    # 写入格式配置
    FONT_NAME = "宋体"
    FONT_SIZE = Pt(10.5)  # 五号字

    def _clear_paragraph(self, para) -> None:
        try:
            p = para._element
            for child in list(p):
                p.remove(child)
        except Exception:
            try:
                for r in list(para.runs):
                    r.text = ""
            except Exception:
                pass
    
    def write_to_document(self, student: StudentInfo) -> Tuple[bool, str]:
        """
        将评语写入 Word 文档
        
        Args:
            student: 学生信息
            
        Returns:
            (是否成功, 消息)
        """
        if not student.comment:
            return False, "评语为空"
        
        file_path = student.file_path
        if not file_path:
            return False, "文件路径为空"

        file_path = os.path.normpath(file_path)
        student.file_path = file_path

        if not os.path.exists(file_path):
            return False, f"文件不存在: {file_path}"
        
        try:
            with open(file_path, "rb") as f:
                data = f.read()
            doc = Document(io.BytesIO(data))
            
            # 查找"教师评阅"位置并写入
            review_text = self._build_review_text(student)
            success = self._find_and_write(doc, review_text)
            
            if not success:
                return False, f"未找到教师评阅位置 | 路径: {file_path}"
            
            tmp_path = f"{file_path}.tmp"
            last_error: Optional[Exception] = None
            for attempt in range(3):
                try:
                    doc.save(tmp_path)
                    try:
                        os.chmod(file_path, stat.S_IWRITE)
                    except Exception:
                        pass
                    os.replace(tmp_path, file_path)
                    return True, "写入成功"
                except Exception as e:
                    last_error = e
                    winerror = getattr(e, 'winerror', None)
                    if attempt >= 2 and winerror == 32:
                        try:
                            root, ext = os.path.splitext(file_path)
                            fallback_path = f"{root}_待替换{ext}"
                            idx = 1
                            while os.path.exists(fallback_path):
                                fallback_path = f"{root}_待替换{idx}{ext}"
                                idx += 1
                            os.replace(tmp_path, fallback_path)
                            return False, (
                                f"原文件被占用(WinError 32)，无法覆盖保存: {file_path}"
                                f"\n已生成待替换文件(评语已写入): {fallback_path}"
                                f"\n请关闭占用程序后再次写入，或手动用待替换文件替换原文件"
                            )
                        except Exception:
                            pass
                    try:
                        if os.path.exists(tmp_path):
                            os.remove(tmp_path)
                    except Exception:
                        pass
                    if attempt < 2:
                        time.sleep(0.4 * (attempt + 1))
                        continue

            winerror = getattr(last_error, 'winerror', None) if last_error else None
            if winerror == 32:
                return False, (
                    f"文件被占用(WinError 32): {file_path}"
                    f"\n可能原因: Word/资源管理器预览窗格/OneDrive同步/杀毒软件正在占用"
                )
            if last_error:
                return False, f"写入失败: {type(last_error).__name__}: {last_error} | 路径: {file_path}"
            return False, f"写入失败: 未知错误 | 路径: {file_path}"
        except Exception as e:
            import traceback
            err_type = type(e).__name__
            winerror = getattr(e, 'winerror', None)
            if winerror is not None:
                brief = f"{err_type}(WinError {winerror}): {e}"
            else:
                brief = f"{err_type}: {e}"
            error_detail = traceback.format_exc()
            print(f"写入错误详情(路径={file_path}):\n{error_detail}")
            return False, f"{brief} | 路径: {file_path}"

    def _build_review_text(self, student: StudentInfo) -> str:
        parts: List[str] = []
        if getattr(student, "score", None) is not None:
            try:
                parts.append(f"成绩：{int(float(student.score))}")
            except Exception:
                parts.append(f"成绩：{student.score}")
        if student.comment:
            parts.append(f"评语：{student.comment.strip()}")
        return "\n".join([p for p in parts if p]).strip()

    def _looks_like_date(self, text: str) -> bool:
        t = (text or "").strip()
        if not t:
            return False
        if "年" in t and "月" in t and ("日" in t or t.endswith("月")):
            return True
        if re.search(r"(19|20)\d{2}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日?", t):
            return True
        if re.search(r"\d{4}\s*[\-/\.]\s*\d{1,2}\s*[\-/\.]\s*\d{1,2}", t):
            return True
        return False

    def _overwrite_review_in_paragraphs(self, paragraphs, review_text: str, aggressive_clear: bool) -> bool:
        hit_tokens = ("教师评阅", "教师评语")
        sig_tokens = ("教师签名", "签名", "日期")

        for i, para in enumerate(paragraphs):
            t = para.text or ""
            if not any(tok in t for tok in hit_tokens):
                continue

            colon_pos = t.find("：")
            if colon_pos < 0:
                colon_pos = t.find(":")
            prefix = t[: colon_pos + 1] if colon_pos >= 0 else "教师评阅："

            sig_positions = [t.find(tok) for tok in sig_tokens if t.find(tok) >= 0]
            sig_pos = min(sig_positions) if sig_positions else -1
            review_body = ("\n" + (review_text or "").strip()) if (review_text or "").strip() else ""

            if sig_pos >= 0:
                self._rewrite_para_keep_suffix_runs(para, sig_pos, prefix, review_body)
                return True

            self._clear_paragraph(para)
            title_run = para.add_run(prefix)
            self._apply_format(title_run)
            body_run = para.add_run(review_body)
            self._apply_format(body_run)

            end_idx: Optional[int] = None
            scan_limit = len(paragraphs) if aggressive_clear else min(len(paragraphs), i + 9)
            for j in range(i + 1, scan_limit):
                tj = (paragraphs[j].text or "").strip()
                if not tj:
                    continue
                if any(tok in tj for tok in sig_tokens) or self._looks_like_date(tj):
                    end_idx = j
                    break

            if aggressive_clear and end_idx is None:
                end_idx = len(paragraphs)

            if end_idx is not None:
                for k in range(i + 1, end_idx):
                    tk = (paragraphs[k].text or "").strip()
                    if not tk:
                        continue
                    self._clear_paragraph(paragraphs[k])
            return True
        return False

    def _rewrite_para_keep_suffix_runs(self, para, sig_char_index: int, prefix: str, review_body: str):
        runs = list(para.runs)
        pos = 0
        sig_run_idx: Optional[int] = None
        sig_offset = 0
        for idx, r in enumerate(runs):
            rt = r.text or ""
            if pos + len(rt) > sig_char_index:
                sig_run_idx = idx
                sig_offset = sig_char_index - pos
                break
            pos += len(rt)

        if sig_run_idx is None:
            self._clear_paragraph(para)
            title_run = para.add_run(prefix)
            self._apply_format(title_run)
            body_run = para.add_run(review_body)
            self._apply_format(body_run)
            return

        to_remove = runs[:sig_run_idx]
        for r in to_remove:
            try:
                para._element.remove(r._element)
            except Exception:
                pass

        sig_run = runs[sig_run_idx]
        if sig_offset > 0:
            try:
                sig_run.text = (sig_run.text or "")[sig_offset:]
            except Exception:
                pass

        prefix_run = para.add_run(prefix)
        self._apply_format(prefix_run)
        body_run = para.add_run(review_body)
        self._apply_format(body_run)
        try:
            body_run.add_break()
        except Exception:
            try:
                body_run.text = (body_run.text or "") + "\n"
            except Exception:
                pass

        p = para._element
        try:
            sig_el = sig_run._element
            sig_pos = p.index(sig_el)

            p.remove(body_run._element)
            p.remove(prefix_run._element)

            p.insert(sig_pos, body_run._element)
            p.insert(sig_pos, prefix_run._element)
        except Exception:
            pass
    
    def _find_and_write(self, doc: Document, comment: str) -> bool:
        """
        查找教师评阅位置并写入评语（支持段落和表格）
        
        Args:
            doc: Word 文档对象
            comment: 评语内容
            
        Returns:
            是否找到并写入
        """
        if self._overwrite_review_in_paragraphs(doc.paragraphs, comment, aggressive_clear=False):
            return True

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if self._overwrite_review_in_paragraphs(cell.paragraphs, comment, aggressive_clear=True):
                        return True

        return False
    
    def _append_comment(self, doc: Document, comment: str):
        """
        在文档末尾添加评语
        
        Args:
            doc: Word 文档对象
            comment: 评语内容
        """
        # 添加教师评阅标题
        title_para = doc.add_paragraph()
        title_run = title_para.add_run("教师评阅：")
        self._apply_format(title_run)
        
        # 添加评语内容
        comment_para = doc.add_paragraph()
        comment_run = comment_para.add_run(comment)
        self._apply_format(comment_run)
    
    def _apply_format(self, run):
        """
        应用格式设置
        
        Args:
            run: Word run 对象
        """
        run.font.name = self.FONT_NAME
        run.font.size = self.FONT_SIZE
        run.font.bold = False
        run.font.italic = False
        
        # 安全设置中文字体
        r = run._element
        rPr = r.get_or_add_rPr()
        rFonts = rPr.get_or_add_rFonts()
        rFonts.set(qn('w:eastAsia'), self.FONT_NAME)
    
    def write_all(self, students: List[StudentInfo]) -> List[Tuple[str, bool, str]]:
        """
        批量写入评语到所有文档
        
        Args:
            students: 学生信息列表
            
        Returns:
            写入结果列表: [(文件名, 是否成功, 消息)]
        """
        results = []
        
        for student in students:
            if student.comment:
                success, message = self.write_to_document(student)
                results.append((student.file_name, success, message))
            else:
                results.append((student.file_name, False, "无评语"))
        
        return results
