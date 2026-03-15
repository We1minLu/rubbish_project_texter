# 中文文书智能助手

基于**豆包大模型**（火山引擎方舟平台）的本地中文文书管理助手，支持对话式读取、总结、搜索和修改 `.docx` / `.xlsx` / `.xls` 文档。

## 功能

| 功能 | 说明 |
|------|------|
| 列出项目/文件 | 列出 `projects/` 下所有项目及文档 |
| 读取文档 | 分块读取大 .docx，自动多轮处理 |
| 总结内容 | 对长文档进行 Map-Reduce 分块总结 |
| 搜索关键词 | 在文档中定位段落/单元格 |
| 修改文档 | 修改 .docx 段落或 .xlsx 单元格，**自动备份** |

## 环境要求

- Python 3.9+
- 火山引擎方舟平台账号及推理接入点

## 安装

```bash
git clone https://github.com/your-username/rubbish_project_texter.git
cd rubbish_project_texter
pip install -r requirements.txt
```

## 配置

1. 复制配置模板：
   ```bash
   cp .env.example .env
   ```

2. 编辑 `.env`，填入你的信息：
   ```
   ARK_API_KEY=你的API Key
   ARK_MODEL_DEFAULT=ep-xxxxxxxx-xxxxx
   ARK_MODEL_LARGE=ep-xxxxxxxx-xxxxx
   ```

   - **API Key**：登录 [火山引擎控制台](https://console.volcengine.com/ark) → API Key 管理
   - **推理接入点 ID**：控制台 → 方舟 → 推理接入点，格式为 `ep-xxxxxxxx-xxxxx`
   - `ARK_MODEL_LARGE` 可与 `ARK_MODEL_DEFAULT` 相同，或填写 128k 上下文版本的接入点

3. 在项目根目录创建 `projects/` 文件夹，按项目名建立子文件夹，将文档放入其中：
   ```
   projects/
   └── 我的项目/
       ├── 报告.docx
       └── 预算.xlsx
   ```

## 启动

**Windows：**
```bat
run.bat
```

**直接运行：**
```bash
python -X utf8 agent.py
```

## 使用示例

```
你> 有哪些项目？
你> 列出东洛岛的文件
你> 总结第一阶段LWM.docx
你> 搜索"预算"在报告.docx中的位置
你> 把预算表.xlsx的B3改为150000
你> clear        # 清空对话历史
你> exit         # 退出
```

## 文件结构

```
.
├── agent.py           # 主入口：对话循环 + 工具调用
├── config.py          # 配置加载（读取 .env）
├── doc_readers.py     # 文档读取（docx 分块 / xlsx / xls）
├── doc_writers.py     # 文档修改（自动备份）
├── context_manager.py # 对话历史管理
├── requirements.txt   # 依赖声明
├── run.bat            # Windows 一键启动
├── .env.example       # 配置模板
└── projects/          # 存放你的项目文档（不上传 Git）
```

## 注意事项

- `.xls` 格式**只读**，如需修改请在 Excel 中另存为 `.xlsx`
- 每次修改文档前会自动生成带时间戳的备份文件
- `.env` 文件含私密信息，已加入 `.gitignore`，**请勿手动上传**

## 依赖

```
openai>=2.28.0      # 豆包 API OpenAI 兼容层
python-docx>=1.1.2  # .docx 读写
openpyxl>=3.1.2     # .xlsx 读写
xlrd>=2.0.1         # .xls 只读
rich>=13.7.0        # 终端彩色 UI
```
