# Doc-Generator MCP Server

基于 MCP (Model Context Protocol) 的文档智能生成与编辑服务，支持 **Word / Excel / Visio / 图表** 的自动生成、定向编辑与质量验证。

---

## 功能概览

| 类别 | 工具 | 说明 |
|------|------|------|
| **高层编排** | `auto_generate` | 一句话需求自动生成文档 |
| | `auto_edit` | 一句话指令自动编辑已有文档 |
| | `analyze_requirement` | 需求分析，返回结构化 Schema |
| **Word** | `generate_word` | 根据 JSON Schema 生成 Word 文档 |
| | `read_document` | 读取文档结构与内容 |
| | `edit_document` | 定向编辑（替换文本/更新表格/插入段落等） |
| | `replace_image` | 替换文档中嵌入的图片 |
| | `list_images` | 列出文档中所有嵌入图片 |
| **Excel** | `generate_excel` | 生成 Excel（含数据、公式、图表、条件格式） |
| **Visio** | `generate_visio` | 生成 Visio 图表（流程图/组织架构图/数据流图等） |
| | `read_visio` | 读取 Visio 文件结构 |
| | `edit_visio` | 编辑已有 Visio 文件 |
| | `embed_visio_to_word` | 将 Visio 图嵌入 Word OLE 对象 |
| **图表** | `generate_diagram` | Mermaid / Graphviz 渲染图表 |
| **验证** | `validate_document` | 分级质量验证（评分 + 错误详情） |
| **查询** | `list_templates` | 列出可用模板 |
| | `get_template_info` | 获取模板详情 |
| | `get_schema_spec` | 获取 JSON Schema 规范 |
| **经验库** | `get_experience` | 查询历史经验 |
| | `log_experience` | 记录经验（含准入机制） |
| | `review_experience` | 审核待审经验 |

---

## 与同类工具的差异化优势

目前市面上的文档类 MCP 工具大多聚焦单一格式（如 `office-word-mcp-server` 仅支持 Word，`excel-mcp-server` 仅支持 Excel），**Doc-Generator 是首个覆盖 Word + Excel + Visio + 图表全格式的一站式文档生成 MCP 服务**，其核心优势包括：

### 1. 全格式覆盖，一站式服务

| 对比维度 | Doc-Generator | 同类工具 |
|---------|--------------|---------|
| **Word 生成** | ✅ 模板驱动 + Schema 驱动 | ✅ 部分支持 |
| **Excel 生成** | ✅ 含公式、图表、条件格式 | ✅ 部分支持 |
| **Visio 生成** | ✅ **唯一支持 Visio (.vsdx) 的 MCP 服务** | ❌ 均不支持 |
| **图表渲染** | ✅ Mermaid + Graphviz 双引擎 | ❌ 无 |
| **Visio 嵌入 Word** | ✅ **OLE 嵌入保留双击可编辑性** | ❌ 无 |
| **跨格式联动** | ✅ 如 Visio 图 → Word OLE 嵌入 | ❌ 无 |

### 2. 端到端自动编排（零代码）

不同于仅提供原子操作的工具，Doc-Generator 提供完整的 **"需求 → 分析 → 生成 → 验证 → 修复"** 闭环：

```
用户输入 → analyze_requirement → generate_* → validate_document → 自动修复 → 交付
```

- `auto_generate`：一句话需求直达成品文档
- `auto_edit`：自然语言描述修改意图，自动执行编辑
- 验证失败时 LLM 自动诊断并重试，无需人工干预

### 3. 内置质量验证与自动修复

- **三级验证体系**：critical / important / suggestion，评分 0-100
- **LLM 驱动自动修复**：验证不通过时自动诊断问题并重试生成
- 覆盖 Word 文档结构、Excel 数据完整性等多维度检查

### 4. Visio 深度集成（业界首创）

- **生成 Visio 图表**：支持流程图、组织架构图、数据流图等 15 种形状类型
- **编辑已有 Visio 文件**：修改文本、样式、添加/删除形状和连接线
- **OLE 嵌入 Word**：将 Visio 图嵌入 Word 文档，保留双击可编辑性
- 双引擎支持：COM 自动化（需 Visio 安装）和 vsdx 库（无需 Visio）

### 5. 经验知识库（持续学习）

- 每次操作自动记录经验教训
- 带准入机制：仅 verified / user_feedback 来源直接入库，inferred 来源需审核
- 自动去重：基于 SequenceMatcher 相似度检测，避免知识冗余
- 操作前自动查询相关经验，避免重复犯错

### 6. 智能备份与安全

- 单备份策略：每个文件只保留一份原始备份，避免文件膨胀
- 编辑前自动备份，支持恢复

### 7. 模板驱动 + 灵活编排

- **模板模式**：通过 Jinja2 模板控制格式，适合标准化文档
- **Schema 模式**：通过 Pydantic Schema 动态生成，适合灵活内容
- **混合模式**：两种使用模式自由组合

---

## 项目结构

```
doc-generator-mcp/
├── server.py              # MCP Server 入口（工具注册 + 路由）
├── config.yaml            # 服务配置（LLM、路径、备份策略等）
├── requirements.txt       # Python 依赖
├── start.bat              # Windows 启动脚本
├── schemas/               # Pydantic 数据模型
│   ├── word_schema.py     #   Word 文档 Schema
│   ├── excel_schema.py    #   Excel 文件 Schema
│   ├── diagram_schema.py  #   图表 Schema
│   ├── edit_schema.py     #   编辑指令 Schema
│   └── visio_schema.py    #   Visio 图表 Schema
├── engines/               # 渲染引擎
│   ├── llm_client.py      #   LLM 调用封装（OpenAI 兼容）
│   └── template_renderer.py #  Word 模板渲染引擎
├── tools/                 # MCP 工具实现
│   ├── auto_generate.py   #   自动生成编排
│   ├── auto_edit.py       #   自动编辑编排
│   ├── analyze.py         #   需求分析
│   ├── generate_word.py   #   Word 生成
│   ├── generate_excel.py  #   Excel 生成
│   ├── generate_diagram.py#   图表生成（Mermaid/Graphviz）
│   ├── generate_visio.py  #   Visio 生成（COM/vsdx）
│   ├── read_document.py   #   文档读取
│   ├── edit_document.py   #   文档编辑
│   ├── edit_visio.py      #   Visio 编辑
│   ├── embed_visio.py     #   Visio 嵌入 Word
│   ├── replace_image.py   #   图片替换
│   ├── validate.py        #   文档验证
│   └── query.py           #   模板/Schema 查询
├── utils/                 # 工具模块
│   ├── backup.py          #   智能备份（单备份策略）
│   ├── color.py           #   颜色工具函数
│   └── experience.py      #   经验知识库（准入 + 去重 + 审核）
├── validation/            # 验证模块
│   └── validator.py       #   分级验证器（评分 0-100）
├── templates/             # 文档模板目录
│   ├── word/
│   ├── excel/
│   └── diagram/
├── examples/              # Few-shot 示例
│   └── few_shot/
├── output/                # 默认输出目录
└── experiences/           # 经验数据存储
```

## 快速开始

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

> 注意：Visio 相关功能（`generate_visio` COM 模式、`edit_visio`、`embed_visio_to_word`）需要本机安装 Microsoft Visio 并安装 `pywin32`。

### 2. 配置 LLM

编辑 `config.yaml` 或设置环境变量：

```bash
set LLM_API_KEY=your-api-key
set LLM_BASE_URL=https://api.openai.com/v1   # 可选，用于本地模型
```

### 3. 启动服务

```bash
# Windows
start.bat

# 或直接运行
python server.py
```

### 4. 接入 MCP 客户端

在 TRAE / Claude Desktop 等 MCP 客户端的配置文件中添加：

```json
{
  "mcpServers": {
    "doc-generator": {
      "command": "<你的路径>\\doc-generator-mcp\\start.bat",
      "args": [],
      "env": {
        "START_MCP_TIMEOUT_MS": "30000",
        "RUN_MCP_TIMEOUT_MS": "120000"
      }
    }
  }
}
```

## 两种使用模式

### 模式一：自动编排（零代码）

直接描述需求，服务自动完成分析 -> 生成 -> 验证的全流程：

```
auto_generate: "生成一份包含5个钻孔数据的地质勘探报告"
auto_edit: "把第三章的表格数据更新为最新数据"
```

### 模式二：原子工具（Agent 编排）

外部 Agent 或模型按需调用各原子工具，灵活组合：

```
1. analyze_requirement -> 获取结构化 Schema
2. generate_word / generate_excel -> 按 Schema 生成
3. validate_document -> 验证质量
4. edit_document -> 定向修改
```

## 核心特性

- **全格式覆盖**：Word + Excel + Visio + 图表，一站式文档生成
- **模板驱动**：通过 Jinja2 模板控制格式，避免生成格式错误
- **分级验证**：自动生成后执行三级验证（critical / important / suggestion），评分 0-100
- **自动修复**：验证失败时 LLM 诊断问题并自动重试
- **智能备份**：编辑前自动备份原文件，单备份策略避免文件膨胀
- **经验知识库**：带准入机制的经验记录系统，防止错误知识污染
- **Visio OLE 嵌入**：支持将 Visio 图嵌入 Word 并保留双击可编辑性（业界首创）
- **跨环境访问**：通过 MCP 协议，任何支持 MCP 的客户端均可调用

## 依赖说明

| 包 | 用途 |
|----|------|
| `mcp` | MCP 协议框架 |
| `openai` / `pydantic` | LLM 调用 & 数据校验 |
| `python-docx` / `docxtpl` / `Jinja2` | Word 文档生成 |
| `xlsxwriter` / `openpyxl` / `pandas` | Excel 文件生成 |
| `graphviz` / `Pillow` | 图表渲染 |
| `vsdx` / `pywin32` | Visio 文件操作 |
| `pyyaml` | 配置文件解析 |

## 许可证

MIT License
