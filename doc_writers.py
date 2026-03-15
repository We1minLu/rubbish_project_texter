"""文档修改模块：docx 段落修改 + xlsx 单元格修改 + 自动备份。"""
from __future__ import annotations
import io
import json
import os
import shutil
import stat
from datetime import datetime
from pathlib import Path

from config import PROJECTS_DIR


def _resolve(project_name: str, filename: str) -> Path:
    path = PROJECTS_DIR / project_name / filename
    path = path.resolve()
    base = PROJECTS_DIR.resolve()
    if not str(path).startswith(str(base)):
        raise ValueError(f"路径不合法: {path}")
    return path


def _ensure_writable(path: Path) -> None:
    """若文件为只读，先去掉只读属性。"""
    current = os.stat(str(path)).st_mode
    if not (current & stat.S_IWRITE):
        os.chmod(str(path), current | stat.S_IWRITE)


def _backup(path: Path) -> Path:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = path.with_name(f"{path.stem}_backup_{ts}{path.suffix}")
    shutil.copy2(str(path), str(backup_path))
    return backup_path


# ---------------------------------------------------------------------------
# modify_docx_paragraph
# ---------------------------------------------------------------------------

def modify_docx_paragraph(
    project_name: str,
    filename: str,
    paragraph_index: int,
    new_text: str,
) -> str:
    try:
        from docx import Document
    except ImportError:
        return json.dumps({"error": "python-docx 未安装"}, ensure_ascii=False)

    path = _resolve(project_name, filename)
    if not path.exists():
        return json.dumps({"error": f"文件不存在: {filename}"}, ensure_ascii=False)

    try:
        doc = Document(str(path))
    except Exception as e:
        return json.dumps({"error": f"无法打开文档: {e}"}, ensure_ascii=False)

    paragraphs = doc.paragraphs
    if paragraph_index < 0 or paragraph_index >= len(paragraphs):
        return json.dumps(
            {"error": f"段落索引 {paragraph_index} 超出范围 [0, {len(paragraphs) - 1}]"},
            ensure_ascii=False,
        )

    para = paragraphs[paragraph_index]
    old_text = para.text

    # 保留第一个 run 的格式，清空其他 runs
    if para.runs:
        first_run = para.runs[0]
        first_run.text = new_text
        for run in para.runs[1:]:
            run.text = ""
    else:
        para.add_run(new_text)

    _ensure_writable(path)
    backup_path = _backup(path)
    doc.save(str(path))

    return json.dumps(
        {
            "status": "success",
            "filename": filename,
            "paragraph_index": paragraph_index,
            "old_value": old_text,
            "new_value": new_text,
            "backup": backup_path.name,
        },
        ensure_ascii=False,
    )


# ---------------------------------------------------------------------------
# modify_excel_cell
# ---------------------------------------------------------------------------

def modify_excel_cell(
    project_name: str,
    filename: str,
    sheet_name: str,
    cell_address: str,
    new_value: str,
) -> str:
    path = _resolve(project_name, filename)
    if not path.exists():
        return json.dumps({"error": f"文件不存在: {filename}"}, ensure_ascii=False)

    suffix = path.suffix.lower()
    if suffix == ".xls":
        return json.dumps(
            {"error": ".xls 格式为只读，请将文件另存为 .xlsx 后再修改"},
            ensure_ascii=False,
        )
    if suffix != ".xlsx":
        return json.dumps({"error": f"不支持的格式: {suffix}"}, ensure_ascii=False)

    try:
        import openpyxl
    except ImportError:
        return json.dumps({"error": "openpyxl 未安装"}, ensure_ascii=False)

    try:
        import io
        with open(str(path), "rb") as f:
            wb = openpyxl.load_workbook(io.BytesIO(f.read()))
    except Exception as e:
        return json.dumps({"error": f"无法打开文件: {e}"}, ensure_ascii=False)

    if sheet_name not in wb.sheetnames:
        return json.dumps(
            {"error": f"Sheet '{sheet_name}' 不存在，可用: {wb.sheetnames}"},
            ensure_ascii=False,
        )

    ws = wb[sheet_name]
    try:
        old_value = ws[cell_address].value
        ws[cell_address] = new_value
    except Exception as e:
        return json.dumps({"error": f"单元格操作失败: {e}"}, ensure_ascii=False)

    _ensure_writable(path)
    backup_path = _backup(path)
    wb.save(str(path))

    return json.dumps(
        {
            "status": "success",
            "filename": filename,
            "sheet": sheet_name,
            "cell": cell_address,
            "old_value": str(old_value) if old_value is not None else "",
            "new_value": new_value,
            "backup": backup_path.name,
        },
        ensure_ascii=False,
    )
