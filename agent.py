"""中文文书智能助手 - 主入口。"""
from __future__ import annotations
import sys
import json

# 确保 UTF-8 输出
sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

from openai import OpenAI
from rich.console import Console
from rich.panel import Panel
from rich.markdown import Markdown

import config
from context_manager import ContextManager
from doc_readers import list_projects, list_files, read_docx, read_excel, search_in_file, read_image, read_pdf, read_pptx, read_project_folder
from doc_writers import modify_docx_paragraph, modify_excel_cell, normalize_docx_style, set_docx_font_style, restructure_docx_paragraphs

console = Console()

# ---------------------------------------------------------------------------
# 工具注册表（Responses API 格式：name/description/parameters 不嵌套在 function 下）
# ---------------------------------------------------------------------------

TOOLS = [
    {
        "type": "function",
        "name": "list_projects",
        "description": "列出 projects/ 目录下所有项目文件夹",
        "parameters": {"type": "object", "properties": {}, "required": []},
    },
    {
        "type": "function",
        "name": "list_files",
        "description": "列出指定项目内的 docx/xlsx/xls 文件及大小",
        "parameters": {
            "type": "object",
            "properties": {
                "project_name": {"type": "string", "description": "项目文件夹名称"},
            },
            "required": ["project_name"],
        },
    },
    {
        "type": "function",
        "name": "read_docx",
        "description": (
            "分块读取 docx 文档，返回带段落编号的文本。"
            "当 total_chunks > 1 时，需依次调用 chunk_index 0 到 total_chunks-1 读取全部内容。"
        ),
        "parameters": {
            "type": "object",
            "properties": {
                "project_name": {"type": "string", "description": "项目文件夹名称"},
                "filename": {"type": "string", "description": "文件名，如 报告.docx"},
                "chunk_index": {
                    "type": "integer",
                    "description": "分块索引，从 0 开始",
                    "default": 0,
                },
            },
            "required": ["project_name", "filename"],
        },
    },
    {
        "type": "function",
        "name": "read_excel",
        "description": "读取 xlsx 或 xls 文件的 Sheet 内容",
        "parameters": {
            "type": "object",
            "properties": {
                "project_name": {"type": "string", "description": "项目文件夹名称"},
                "filename": {"type": "string", "description": "文件名"},
                "sheet_name": {
                    "type": "string",
                    "description": "Sheet 名称，留空则读取所有 Sheet",
                    "default": "",
                },
            },
            "required": ["project_name", "filename"],
        },
    },
    {
        "type": "function",
        "name": "search_in_file",
        "description": "在 docx/xlsx/xls 文件中搜索关键词，返回匹配的段落或单元格",
        "parameters": {
            "type": "object",
            "properties": {
                "project_name": {"type": "string", "description": "项目文件夹名称"},
                "filename": {"type": "string", "description": "文件名"},
                "keyword": {"type": "string", "description": "搜索关键词"},
            },
            "required": ["project_name", "filename", "keyword"],
        },
    },
    {
        "type": "function",
        "name": "read_image",
        "description": "读取图片文件（PNG/JPG/JPEG/BMP/WEBP/GIF），调用视觉AI描述或回答图片相关问题",
        "parameters": {
            "type": "object",
            "properties": {
                "project_name": {"type": "string", "description": "项目文件夹名称"},
                "filename": {"type": "string", "description": "图片文件名，如 设计图.png"},
                "question": {
                    "type": "string",
                    "description": "对图片的提问，留空则描述全部内容",
                    "default": "",
                },
            },
            "required": ["project_name", "filename"],
        },
    },
    {
        "type": "function",
        "name": "read_pdf",
        "description": "读取 PDF 文件的文字内容，支持指定页码范围",
        "parameters": {
            "type": "object",
            "properties": {
                "project_name": {"type": "string", "description": "项目文件夹名称"},
                "filename": {"type": "string", "description": "PDF 文件名，如 方案.pdf"},
                "page_range": {
                    "type": "string",
                    "description": "页码范围，如 '1-5' 或 '1,3,5'，留空读全部",
                    "default": "",
                },
            },
            "required": ["project_name", "filename"],
        },
    },
    {
        "type": "function",
        "name": "read_pptx",
        "description": "读取 PPT/PPTX 演示文稿的文字内容，支持指定幻灯片范围",
        "parameters": {
            "type": "object",
            "properties": {
                "project_name": {"type": "string", "description": "项目文件夹名称"},
                "filename": {"type": "string", "description": "PPT/PPTX 文件名"},
                "slide_range": {
                    "type": "string",
                    "description": "幻灯片范围，如 '1-10'，留空读全部",
                    "default": "",
                },
            },
            "required": ["project_name", "filename"],
        },
    },
    {
        "type": "function",
        "name": "read_project_folder",
        "description": (
            "批量读取项目文件夹下所有文件（docx/pdf/pptx/图片等），整体理解一个项目。"
            "图片会自动调用视觉AI生成描述。适合用户说'帮我了解整个xxx项目'时使用。"
        ),
        "parameters": {
            "type": "object",
            "properties": {
                "project_name": {"type": "string", "description": "项目文件夹名称"},
                "subfolder": {
                    "type": "string",
                    "description": "子文件夹路径，留空则读项目根目录",
                    "default": "",
                },
                "include_images": {
                    "type": "boolean",
                    "description": "是否包含图片文件（会调用视觉AI，速度较慢）",
                    "default": True,
                },
                "max_files": {
                    "type": "integer",
                    "description": "最多读取文件数，避免超出token限制，默认20",
                    "default": 20,
                },
            },
            "required": ["project_name"],
        },
    },
    {
        "type": "function",
        "name": "modify_docx_paragraph",
        "description": "修改 docx 文档中指定段落的文本内容，保留原有格式，修改前自动备份",
        "parameters": {
            "type": "object",
            "properties": {
                "project_name": {"type": "string", "description": "项目文件夹名称"},
                "filename": {"type": "string", "description": "文件名"},
                "paragraph_index": {
                    "type": "integer",
                    "description": "段落索引（read_docx 返回的编号）",
                },
                "new_text": {"type": "string", "description": "新的段落文本"},
            },
            "required": ["project_name", "filename", "paragraph_index", "new_text"],
        },
    },
    {
        "type": "function",
        "name": "restructure_docx_paragraphs",
        "description": (
            "批量重构 docx 文档的段落结构：为每个段落指定 heading1/heading2/heading3/body，"
            "设置对应的 Word 标准样式（Heading 1/2/3 / Normal），支持 Word 自动生成目录。"
            "标题顶格无缩进，正文首行缩进2字符。"
            "使用前必须先用 read_docx 读完全文，再根据内容判断每段的层级后一次性调用本工具。"
        ),
        "parameters": {
            "type": "object",
            "properties": {
                "project_name": {"type": "string", "description": "项目文件夹名称"},
                "filename": {"type": "string", "description": "docx 文件名"},
                "paragraph_styles": {
                    "type": "object",
                    "description": (
                        "段落索引 → 样式的映射，键为段落索引字符串，值为 heading1/heading2/heading3/body。"
                        "例：{\"0\": \"heading1\", \"1\": \"body\", \"2\": \"heading2\"}"
                    ),
                    "additionalProperties": {"type": "string", "enum": ["heading1", "heading2", "heading3", "body"]},
                },
            },
            "required": ["project_name", "filename", "paragraph_styles"],
        },
    },
    {
        "type": "function",
        "name": "set_docx_font_style",
        "description": (
            "对 docx 文档设置字体名称、字号、粗体。"
            "可通过 style_filter 指定作用范围：all=全文，heading=仅标题段落，body=仅正文段落；"
            "也可通过 paragraph_indices 指定具体段落（优先级更高）。修改前自动备份。"
        ),
        "parameters": {
            "type": "object",
            "properties": {
                "project_name": {"type": "string", "description": "项目文件夹名称"},
                "filename": {"type": "string", "description": "docx 文件名"},
                "font_name": {"type": "string", "description": "字体名称，如 宋体、微软雅黑、黑体"},
                "font_size": {"type": "string", "description": "字号，支持中文名（小四、四号）或数字（12、14）"},
                "bold": {"type": "boolean", "description": "true=加粗，false=取消加粗"},
                "style_filter": {
                    "type": "string",
                    "enum": ["all", "heading", "body"],
                    "description": "all=全文所有段落，heading=仅标题段落，body=仅正文段落。默认 all",
                },
                "paragraph_indices": {
                    "type": "array",
                    "items": {"type": "integer"},
                    "description": "指定段落索引列表（来自 read_docx 编号），优先级高于 style_filter",
                },
            },
            "required": ["project_name", "filename"],
        },
    },
    {
        "type": "function",
        "name": "normalize_docx_style",
        "description": (
            "统一 docx 文档的文字格式：清除所有 run 级别的格式覆盖（粗体、红色、斜体、下划线、高亮等），"
            "让文字统一继承段落样式。可指定段落范围，不指定则处理全文。修改前自动备份。"
        ),
        "parameters": {
            "type": "object",
            "properties": {
                "project_name": {"type": "string", "description": "项目文件夹名称"},
                "filename": {"type": "string", "description": "docx 文件名"},
                "paragraph_indices": {
                    "type": "array",
                    "items": {"type": "integer"},
                    "description": "要处理的段落索引列表（来自 read_docx 的编号）。留空则处理全文所有段落。",
                },
            },
            "required": ["project_name", "filename"],
        },
    },
    {
        "type": "function",
        "name": "modify_excel_cell",
        "description": "修改 xlsx 文件中指定单元格的值，修改前自动备份。注意：.xls 格式只读。",
        "parameters": {
            "type": "object",
            "properties": {
                "project_name": {"type": "string", "description": "项目文件夹名称"},
                "filename": {"type": "string", "description": "xlsx 文件名"},
                "sheet_name": {"type": "string", "description": "Sheet 名称"},
                "cell_address": {
                    "type": "string",
                    "description": "单元格地址，如 B3",
                },
                "new_value": {"type": "string", "description": "新的单元格值"},
            },
            "required": [
                "project_name",
                "filename",
                "sheet_name",
                "cell_address",
                "new_value",
            ],
        },
    },
]


# ---------------------------------------------------------------------------
# 工具调度
# ---------------------------------------------------------------------------

TOOL_MAP = {
    "list_projects": lambda args: list_projects(),
    "list_files": lambda args: list_files(args["project_name"]),
    "read_docx": lambda args: read_docx(
        args["project_name"], args["filename"], int(args.get("chunk_index", 0))
    ),
    "read_excel": lambda args: read_excel(
        args["project_name"], args["filename"], args.get("sheet_name", "")
    ),
    "search_in_file": lambda args: search_in_file(
        args["project_name"], args["filename"], args["keyword"]
    ),
    "read_image": lambda args: read_image(
        args["project_name"], args["filename"], args.get("question", "")
    ),
    "read_pdf": lambda args: read_pdf(
        args["project_name"], args["filename"], args.get("page_range", "")
    ),
    "read_pptx": lambda args: read_pptx(
        args["project_name"], args["filename"], args.get("slide_range", "")
    ),
    "read_project_folder": lambda args: read_project_folder(
        args["project_name"],
        args.get("subfolder", ""),
        args.get("include_images", True),
        int(args.get("max_files", 20)),
    ),
    "modify_docx_paragraph": lambda args: modify_docx_paragraph(
        args["project_name"],
        args["filename"],
        int(args["paragraph_index"]),
        args["new_text"],
    ),
    "restructure_docx_paragraphs": lambda args: restructure_docx_paragraphs(
        args["project_name"],
        args["filename"],
        args["paragraph_styles"],
    ),
    "set_docx_font_style": lambda args: set_docx_font_style(
        args["project_name"],
        args["filename"],
        args.get("font_name"),
        args.get("font_size"),
        args.get("bold"),
        args.get("style_filter", "all"),
        args.get("paragraph_indices") or None,
    ),
    "normalize_docx_style": lambda args: normalize_docx_style(
        args["project_name"],
        args["filename"],
        args.get("paragraph_indices") or None,
    ),
    "modify_excel_cell": lambda args: modify_excel_cell(
        args["project_name"],
        args["filename"],
        args["sheet_name"],
        args["cell_address"],
        args["new_value"],
    ),
}


def dispatch_tool(name: str, arguments: str) -> str:
    try:
        args = json.loads(arguments)
    except json.JSONDecodeError:
        return json.dumps({"error": "参数解析失败"}, ensure_ascii=False)

    handler = TOOL_MAP.get(name)
    if handler is None:
        return json.dumps({"error": f"未知工具: {name}"}, ensure_ascii=False)

    try:
        return handler(args)
    except Exception as e:
        return json.dumps({"error": str(e)}, ensure_ascii=False)


# ---------------------------------------------------------------------------
# 历史转换：内部存储 → Responses API input 格式
# ---------------------------------------------------------------------------

def _build_input(history: list) -> list:
    """将对话历史转换为 Responses API 的 input 列表。"""
    items = []
    for entry in history:
        entry_type = entry.get("type")
        role = entry.get("role")

        if entry_type == "function_call":
            # function_call 需要 status: completed
            items.append({**entry, "status": "completed"})
        elif entry_type == "function_call_output":
            items.append(entry)
        elif role == "user":
            content = entry.get("content", "")
            items.append({
                "role": "user",
                "content": [{"type": "input_text", "text": content}],
            })
        elif role == "assistant":
            content = entry.get("content", "")
            if content:
                # assistant 历史消息需要 type: message + status: completed
                items.append({
                    "type": "message",
                    "role": "assistant",
                    "content": [{"type": "output_text", "text": content}],
                    "status": "completed",
                })
    return items


# ---------------------------------------------------------------------------
# 模型选择
# ---------------------------------------------------------------------------

def _pick_model(history: list) -> str:
    total = sum(len(str(e.get("content", "") or e.get("output", "") or e.get("arguments", "")))
                for e in history)
    if total > config.LARGE_FILE_THRESHOLD_MB * 1024 * 1024:
        return config.MODEL_LARGE
    return config.MODEL_DEFAULT


# ---------------------------------------------------------------------------
# 对话循环
# ---------------------------------------------------------------------------

def chat_loop() -> None:
    client = OpenAI(api_key=config.API_KEY, base_url=config.BASE_URL)
    ctx = ContextManager()

    console.print(
        Panel(
            "[bold cyan]中文文书智能助手[/bold cyan]\n"
            "输入问题开始对话，输入 [bold]exit[/bold] 或 [bold]quit[/bold] 退出，"
            "输入 [bold]clear[/bold] 清空对话历史",
            title="欢迎",
        )
    )

    while True:
        try:
            user_input = console.input("[bold green]你>[/bold green] ").strip()
        except (EOFError, KeyboardInterrupt):
            console.print("\n[yellow]再见！[/yellow]")
            break

        if not user_input:
            continue
        if user_input.lower() in ("exit", "quit"):
            console.print("[yellow]再见！[/yellow]")
            break
        if user_input.lower() == "clear":
            ctx.clear()
            console.print("[dim]对话历史已清空。[/dim]")
            continue

        ctx.add({"role": "user", "content": user_input})

        # 多轮工具调用循环
        while True:
            history = ctx.get_history()
            model = _pick_model(history)
            input_items = _build_input(history)

            with console.status("[dim]思考中...[/dim]", spinner="dots"):
                try:
                    response = client.responses.create(
                        model=model,
                        instructions=config.SYSTEM_PROMPT,
                        input=input_items,
                        tools=TOOLS,
                        temperature=config.TEMPERATURE,
                    )
                except Exception as e:
                    console.print(f"[red]API 请求失败: {e}[/red]")
                    ctx._history.pop()  # 移除刚才加入的 user 消息
                    break

            # 解析输出
            has_tool_calls = False
            reply_text = ""

            for item in response.output:
                if item.type == "function_call":
                    has_tool_calls = True
                    console.print(f"[dim]  调用工具: {item.name}[/dim]")
                    result = dispatch_tool(item.name, item.arguments)

                    # 存入历史：function_call + function_call_output
                    ctx.add({
                        "type": "function_call",
                        "id": item.id,
                        "call_id": item.call_id,
                        "name": item.name,
                        "arguments": item.arguments,
                    })
                    ctx.add({
                        "type": "function_call_output",
                        "call_id": item.call_id,
                        "output": result,
                    })

                elif item.type == "message":
                    for part in item.content:
                        if hasattr(part, "text"):
                            reply_text += part.text

            if has_tool_calls:
                # 继续循环，让模型处理工具结果
                continue

            # 最终回复
            ctx.add({"role": "assistant", "content": reply_text})
            console.print()
            console.print(Panel(Markdown(reply_text), title="[bold blue]助手[/bold blue]"))
            console.print()
            break


if __name__ == "__main__":
    chat_loop()
