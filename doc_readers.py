"""文档读取模块：docx 分块读取 + xlsx/xls 双格式读取。"""
from __future__ import annotations
import json
from pathlib import Path
from typing import Any

from config import PROJECTS_DIR, CHUNK_SIZE_PARAGRAPHS


# ---------------------------------------------------------------------------
# 路径工具
# ---------------------------------------------------------------------------

def _resolve(project_name: str, filename: str) -> Path:
    path = PROJECTS_DIR / project_name / filename
    # 安全检查：防止路径穿越
    path = path.resolve()
    base = PROJECTS_DIR.resolve()
    if not str(path).startswith(str(base)):
        raise ValueError(f"路径不合法: {path}")
    return path


# ---------------------------------------------------------------------------
# list_projects / list_files
# ---------------------------------------------------------------------------

def list_projects() -> str:
    if not PROJECTS_DIR.exists():
        return json.dumps({"error": "projects/ 目录不存在"}, ensure_ascii=False)
    projects = [p.name for p in sorted(PROJECTS_DIR.iterdir()) if p.is_dir()]
    return json.dumps({"projects": projects}, ensure_ascii=False)


def list_files(project_name: str) -> str:
    project_dir = (PROJECTS_DIR / project_name).resolve()
    if not project_dir.exists():
        return json.dumps({"error": f"项目 '{project_name}' 不存在"}, ensure_ascii=False)
    exts = {".docx", ".xlsx", ".xls"}
    files = []
    for f in sorted(project_dir.iterdir()):
        if f.suffix.lower() in exts:
            size_mb = round(f.stat().st_size / 1024 / 1024, 2)
            files.append({"name": f.name, "size_mb": size_mb})
    return json.dumps({"project": project_name, "files": files}, ensure_ascii=False)


# ---------------------------------------------------------------------------
# read_docx
# ---------------------------------------------------------------------------

def read_docx(project_name: str, filename: str, chunk_index: int = 0) -> str:
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

    paragraphs = [p.text for p in doc.paragraphs]
    total = len(paragraphs)
    total_chunks = max(1, (total + CHUNK_SIZE_PARAGRAPHS - 1) // CHUNK_SIZE_PARAGRAPHS)

    if chunk_index < 0 or chunk_index >= total_chunks:
        return json.dumps(
            {"error": f"chunk_index 超出范围 [0, {total_chunks - 1}]"},
            ensure_ascii=False,
        )

    start = chunk_index * CHUNK_SIZE_PARAGRAPHS
    end = min(start + CHUNK_SIZE_PARAGRAPHS, total)
    chunk = paragraphs[start:end]

    lines = [f"[{start + i}] {text}" for i, text in enumerate(chunk)]

    return json.dumps(
        {
            "project": project_name,
            "filename": filename,
            "chunk_index": chunk_index,
            "total_chunks": total_chunks,
            "paragraph_range": [start, end - 1],
            "content": "\n".join(lines),
        },
        ensure_ascii=False,
    )


# ---------------------------------------------------------------------------
# read_excel
# ---------------------------------------------------------------------------

def read_excel(project_name: str, filename: str, sheet_name: str = "") -> str:
    path = _resolve(project_name, filename)
    if not path.exists():
        return json.dumps({"error": f"文件不存在: {filename}"}, ensure_ascii=False)

    suffix = path.suffix.lower()

    if suffix == ".xlsx":
        return _read_xlsx(path, sheet_name)
    elif suffix == ".xls":
        return _read_xls(path, sheet_name)
    else:
        return json.dumps({"error": f"不支持的格式: {suffix}"}, ensure_ascii=False)


def _read_xlsx(path: Path, sheet_name: str) -> str:
    try:
        import openpyxl
    except ImportError:
        return json.dumps({"error": "openpyxl 未安装"}, ensure_ascii=False)

    try:
        wb = openpyxl.load_workbook(str(path), data_only=True)
    except Exception as e:
        return json.dumps({"error": f"无法打开文件: {e}"}, ensure_ascii=False)

    sheets_to_read = [sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.sheetnames
    result: dict[str, Any] = {}

    for sname in sheets_to_read:
        ws = wb[sname]
        rows = []
        for row in ws.iter_rows(values_only=True):
            rows.append([str(c) if c is not None else "" for c in row])
        result[sname] = rows

    return json.dumps(
        {"filename": path.name, "sheets": result},
        ensure_ascii=False,
    )


def _read_xls(path: Path, sheet_name: str) -> str:
    try:
        import xlrd
    except ImportError:
        return json.dumps({"error": "xlrd 未安装"}, ensure_ascii=False)

    try:
        wb = xlrd.open_workbook(str(path))
    except Exception as e:
        return json.dumps({"error": f"无法打开文件: {e}"}, ensure_ascii=False)

    sheet_names = wb.sheet_names()
    sheets_to_read = [sheet_name] if sheet_name and sheet_name in sheet_names else sheet_names
    result: dict[str, Any] = {}

    for sname in sheets_to_read:
        ws = wb.sheet_by_name(sname)
        rows = []
        for r in range(ws.nrows):
            rows.append([str(ws.cell_value(r, c)) for c in range(ws.ncols)])
        result[sname] = rows

    return json.dumps(
        {"filename": path.name, "sheets": result, "note": ".xls 为只读格式"},
        ensure_ascii=False,
    )


# ---------------------------------------------------------------------------
# search_in_file
# ---------------------------------------------------------------------------

def search_in_file(project_name: str, filename: str, keyword: str) -> str:
    path = _resolve(project_name, filename)
    if not path.exists():
        return json.dumps({"error": f"文件不存在: {filename}"}, ensure_ascii=False)

    suffix = path.suffix.lower()

    if suffix == ".docx":
        return _search_docx(path, keyword)
    elif suffix in (".xlsx", ".xls"):
        return _search_excel(path, keyword, suffix)
    else:
        return json.dumps({"error": f"不支持的格式: {suffix}"}, ensure_ascii=False)


def _search_docx(path: Path, keyword: str) -> str:
    try:
        from docx import Document
    except ImportError:
        return json.dumps({"error": "python-docx 未安装"}, ensure_ascii=False)

    doc = Document(str(path))
    matches = []
    for i, para in enumerate(doc.paragraphs):
        if keyword in para.text:
            matches.append({"paragraph_index": i, "text": para.text})

    return json.dumps(
        {"filename": path.name, "keyword": keyword, "matches": matches},
        ensure_ascii=False,
    )


def _search_excel(path: Path, keyword: str, suffix: str) -> str:
    matches = []

    if suffix == ".xlsx":
        try:
            import openpyxl
            wb = openpyxl.load_workbook(str(path), data_only=True)
            for sname in wb.sheetnames:
                ws = wb[sname]
                for row in ws.iter_rows():
                    for cell in row:
                        if cell.value is not None and keyword in str(cell.value):
                            matches.append({
                                "sheet": sname,
                                "cell": cell.coordinate,
                                "value": str(cell.value),
                            })
        except Exception as e:
            return json.dumps({"error": str(e)}, ensure_ascii=False)
    else:
        try:
            import xlrd
            wb = xlrd.open_workbook(str(path))
            for sname in wb.sheet_names():
                ws = wb.sheet_by_name(sname)
                for r in range(ws.nrows):
                    for c in range(ws.ncols):
                        val = str(ws.cell_value(r, c))
                        if keyword in val:
                            col_letter = chr(ord("A") + c) if c < 26 else f"C{c}"
                            matches.append({
                                "sheet": sname,
                                "cell": f"{col_letter}{r + 1}",
                                "value": val,
                            })
        except Exception as e:
            return json.dumps({"error": str(e)}, ensure_ascii=False)

    return json.dumps(
        {"filename": path.name, "keyword": keyword, "matches": matches},
        ensure_ascii=False,
    )
