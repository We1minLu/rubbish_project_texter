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
from doc_readers import list_projects, list_files, read_docx, read_excel, search_in_file
from doc_writers import modify_docx_paragraph, modify_excel_cell

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
    "modify_docx_paragraph": lambda args: modify_docx_paragraph(
        args["project_name"],
        args["filename"],
        int(args["paragraph_index"]),
        args["new_text"],
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
