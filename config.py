import sys
import os
from pathlib import Path

# --- UTF-8 stdout ---
sys.stdout.reconfigure(encoding="utf-8")
sys.stderr.reconfigure(encoding="utf-8")

# --- 加载 .env（python-dotenv 未安装时手动解析）---
_env_file = Path(__file__).parent / ".env"
if _env_file.exists():
    with open(_env_file, encoding="utf-8") as _f:
        for _line in _f:
            _line = _line.strip()
            if _line and not _line.startswith("#") and "=" in _line:
                _k, _v = _line.split("=", 1)
                os.environ.setdefault(_k.strip(), _v.strip())

# --- Paths ---
BASE_DIR = Path(__file__).parent
PROJECTS_DIR = BASE_DIR / "projects"

# --- Doubao API ---
API_KEY = os.environ.get("ARK_API_KEY", "")
BASE_URL = "https://ark.cn-beijing.volces.com/api/v3"
# 使用火山引擎控制台「推理接入点」中的 endpoint ID，格式：ep-xxxxxxxx-xxxxx
MODEL_DEFAULT = os.environ.get("ARK_MODEL_DEFAULT", "")
MODEL_LARGE   = os.environ.get("ARK_MODEL_LARGE", MODEL_DEFAULT)
LARGE_FILE_THRESHOLD_MB = 5

TEMPERATURE = 0.3

# --- Chunking ---
CHUNK_SIZE_PARAGRAPHS = 200  # paragraphs per chunk for docx

# --- Context window ---
MAX_HISTORY_TURNS = 20       # keep last N user+assistant pairs
MAX_CONTEXT_TOKENS = 30_000  # rough token budget for history trimming

# --- System Prompt ---
SYSTEM_PROMPT = """你是一个专业的中文文书智能助手，负责帮助用户管理和分析存放在 projects/ 目录下的项目文档（.docx/.xlsx/.xls）。

## 能力
- 列出项目和文件
- 读取、分析、总结文档内容
- 在文档中搜索关键词
- 修改文档段落和单元格（修改前自动备份）

## 大文档处理流程
当文档包含多个分块（total_chunks > 1）时，你必须：
1. 依次调用 read_docx，chunk_index 从 0 递增到 total_chunks-1
2. 对每个分块生成局部摘要
3. 最后汇总所有局部摘要，输出完整的总结

## 注意事项
- .xls 格式为只读，无法修改，请提示用户另存为 .xlsx
- 修改操作执行前会自动备份，备份文件名含时间戳
- 始终用中文回复用户
"""

if not API_KEY:
    raise RuntimeError("未设置 ARK_API_KEY，请在 .env 文件中配置")
if not MODEL_DEFAULT:
    raise RuntimeError("未设置 ARK_MODEL_DEFAULT，请在 .env 文件中配置推理接入点 ID")
