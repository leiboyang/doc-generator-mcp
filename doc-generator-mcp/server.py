"""
MCP Server: doc-generator
文档智能生成服务 —— 支持 Word / Excel / 图表的自动生成与编辑
"""

import json
import os
import sys
from pathlib import Path
from datetime import datetime

import yaml
from mcp.server import Server
from mcp.types import Tool, TextContent

# 添加项目根目录到路径
sys.path.insert(0, str(Path(__file__).parent))

from schemas import (
    WordSchema, ExcelSchema, DiagramSchema,
    EditInstruction, EditResult,
    VisioSchema,
)


# ========== 配置加载 ==========

def load_config() -> dict:
    config_path = Path(__file__).parent / "config.yaml"
    if config_path.exists():
        with open(config_path, "r", encoding="utf-8") as f:
            config = yaml.safe_load(f)
    else:
        config = {}

    # 环境变量覆盖
    config.setdefault("llm", {})
    config["llm"]["api_key"] = os.environ.get("LLM_API_KEY", config["llm"].get("api_key", ""))
    config["llm"]["base_url"] = os.environ.get("LLM_BASE_URL", config["llm"].get("base_url", ""))

    return config


CONFIG = load_config()


# ========== MCP Server 初始化 ==========

server = Server("doc-generator")


# ========== 工具注册 ==========

@server.list_tools()
async def list_tools() -> list[Tool]:
    return [
        # ---- 高层编排工具 ----
        Tool(
            name="auto_generate",
            description="一句话需求自动生成文档。输入自然语言描述，自动分析需求、选择模板、生成文档并验证。",
            inputSchema={
                "type": "object",
                "properties": {
                    "request": {
                        "type": "string",
                        "description": "自然语言需求描述，如'生成一份包含5个钻孔数据的地质勘探报告'"
                    },
                    "doc_type": {
                        "type": "string",
                        "enum": ["word", "excel", "diagram"],
                        "description": "文档类型提示（可选，不填自动判断）"
                    },
                    "template": {
                        "type": "string",
                        "description": "指定模板名称（可选，不填自动选择）"
                    },
                    "output_dir": {
                        "type": "string",
                        "description": "输出目录路径（可选）"
                    },
                    "max_retries": {
                        "type": "integer",
                        "default": 3,
                        "description": "最大重试次数"
                    }
                },
                "required": ["request"]
            }
        ),
        Tool(
            name="auto_edit",
            description="一句话描述修改意图，自动编辑已有文档。修改部分内容并保持其余格式不变。",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": {
                        "type": "string",
                        "description": "已有文档的文件路径"
                    },
                    "instruction": {
                        "type": "string",
                        "description": "自然语言修改指令，如'把第三章表格数据更新为最新数据'"
                    },
                    "backup": {
                        "type": "boolean",
                        "default": True,
                        "description": "是否自动备份原文件"
                    },
                    "output_path": {
                        "type": "string",
                        "description": "输出路径（可选，不填则覆盖原文件）"
                    }
                },
                "required": ["file_path", "instruction"]
            }
        ),
        Tool(
            name="analyze_requirement",
            description="分析需求但不生成文档。返回文档类型、推荐模板、结构化 Schema 等分析结果。",
            inputSchema={
                "type": "object",
                "properties": {
                    "request": {
                        "type": "string",
                        "description": "自然语言需求描述"
                    }
                },
                "required": ["request"]
            }
        ),

        # ---- 原子生成工具 ----
        Tool(
            name="generate_word",
            description="根据 JSON Schema 和模板生成 Word 文档",
            inputSchema={
                "type": "object",
                "properties": {
                    "schema": {"type": "object", "description": "符合 WordSchema 的 JSON 结构"},
                    "template": {"type": "string", "description": "模板名称"},
                    "output_path": {"type": "string", "description": "输出文件路径"}
                },
                "required": ["schema", "template", "output_path"]
            }
        ),
        Tool(
            name="generate_excel",
            description="根据 JSON Schema 生成 Excel 文件（含数据、公式、图表）",
            inputSchema={
                "type": "object",
                "properties": {
                    "schema": {"type": "object", "description": "符合 ExcelSchema 的 JSON 结构"},
                    "template": {"type": "string", "description": "模板名称（可选）"},
                    "output_path": {"type": "string", "description": "输出文件路径"}
                },
                "required": ["schema", "output_path"]
            }
        ),
        Tool(
            name="generate_diagram",
            description="根据节点/连接关系生成图表（流程图、组织架构图等）",
            inputSchema={
                "type": "object",
                "properties": {
                    "schema": {"type": "object", "description": "符合 DiagramSchema 的 JSON 结构"},
                    "engine": {
                        "type": "string",
                        "enum": ["mermaid", "graphviz", "com"],
                        "default": "mermaid",
                        "description": "渲染引擎"
                    },
                    "output_format": {
                        "type": "string",
                        "enum": ["svg", "png", "vsdx"],
                        "default": "svg",
                        "description": "输出格式"
                    },
                    "output_path": {"type": "string", "description": "输出文件路径"}
                },
                "required": ["schema", "output_path"]
            }
        ),

        # ---- 文档编辑工具 ----
        Tool(
            name="read_document",
            description="读取已有文档的结构和内容，返回结构化描述 JSON",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": {"type": "string", "description": "文档文件路径"},
                    "include_content": {
                        "type": "boolean", "default": True,
                        "description": "是否包含段落文本内容"
                    },
                    "include_style": {
                        "type": "boolean", "default": False,
                        "description": "是否包含样式详情"
                    },
                    "max_depth": {
                        "type": "integer", "default": 2,
                        "description": "结构解析深度（1=仅标题，2=标题+段落，3=完整）"
                    }
                },
                "required": ["file_path"]
            }
        ),
        Tool(
            name="edit_document",
            description="对已有文档执行定向编辑，精确修改指定内容，保留其余部分不变",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": {"type": "string", "description": "文档文件路径"},
                    "edits": {
                        "type": "array",
                        "items": {"type": "object"},
                        "description": "编辑指令列表"
                    },
                    "backup": {
                        "type": "boolean", "default": True,
                        "description": "是否自动备份原文件"
                    },
                    "output_path": {
                        "type": "string",
                        "description": "输出路径（可选，不填则覆盖原文件）"
                    }
                },
                "required": ["file_path", "edits"]
            }
        ),

        # ---- Visio 工具 ----
        Tool(
            name="generate_visio",
            description="根据结构化 Schema 生成 Visio 图表文件（.vsdx）。支持流程图、组织架构图、数据流图等。",
            inputSchema={
                "type": "object",
                "properties": {
                    "schema": {"type": "object", "description": "符合 VisioSchema 的 JSON 结构（含 pages/shapes/connectors）"},
                    "engine": {
                        "type": "string", "enum": ["com", "vsdx"], "default": "com",
                        "description": "渲染引擎: com(Visio COM自动化) 或 vsdx(Python库)"
                    },
                    "output_path": {"type": "string", "description": "输出文件路径（.vsdx）"}
                },
                "required": ["schema", "output_path"]
            }
        ),
        Tool(
            name="read_visio",
            description="读取 Visio 文件（.vsdx）的结构和内容，返回形状、连接线等信息。",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": {"type": "string", "description": "Visio 文件路径"}
                },
                "required": ["file_path"]
            }
        ),
        Tool(
            name="edit_visio",
            description="编辑已有 Visio 文件。支持修改形状文本/样式、添加/删除形状、通过 AutoConnect 添加连接线。修改前自动备份。",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": {"type": "string", "description": "Visio 文件路径"},
                    "edits": {
                        "type": "array",
                        "items": {"type": "object"},
                        "description": "编辑指令列表，支持: update_shape_text, update_shape_style, add_shape, delete_shape"
                    },
                    "backup": {"type": "boolean", "default": True, "description": "是否自动备份"},
                    "output_path": {"type": "string", "description": "输出路径（可选，不填则覆盖原文件）"}
                },
                "required": ["file_path", "edits"]
            }
        ),

        # ---- 图片替换工具 ----
        Tool(
            name="replace_image",
            description="替换 Word 文档中嵌入的图片。支持按图题文本、rId或图片索引定位。自动保护 Visio OLE 对象（防止误替换为图片丢失可编辑性）。修改前自动备份。",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": {"type": "string", "description": "Word 文档路径"},
                    "new_image_path": {"type": "string", "description": "新图片文件路径"},
                    "target": {
                        "type": "string",
                        "description": "定位方式: 图题文本(如'图 1')、rId(如'rId14')、图片索引(如'0')"
                    },
                    "backup": {"type": "boolean", "default": True, "description": "是否自动备份"},
                    "output_path": {"type": "string", "description": "输出路径（可选）"}
                },
                "required": ["file_path", "new_image_path", "target"]
            }
        ),
        Tool(
            name="list_images",
            description="列出 Word 文档中所有嵌入的图片，包括 rId、类型、大小、关联图题等信息。",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": {"type": "string", "description": "Word 文档路径"}
                },
                "required": ["file_path"]
            }
        ),

        # ---- Visio 嵌入工具 ----
        Tool(
            name="embed_visio_to_word",
            description="将 Visio 图嵌入 Word 文档的 OLE 对象中，保留双击可编辑性。支持已有 .vsdx 文件或自动创建。自动适配 Word 页面宽度。",
            inputSchema={
                "type": "object",
                "properties": {
                    "word_path": {"type": "string", "description": "Word 文档路径"},
                    "visio_path": {"type": "string", "description": "已有 .vsdx 文件路径（可选，不提供则根据 shapes 参数自动创建）"},
                    "target": {"type": "string", "description": "定位方式: 图题文本(如'图 1')或 rId(如'rId15')"},
                    "shapes": {
                        "type": "array",
                        "items": {"type": "object"},
                        "description": "形状列表 (不传 visio_path 时必填), 每项: {id, label, x, y, width, height, fill_color, line_color, style}"
                    },
                    "connectors": {
                        "type": "array",
                        "items": {"type": "object"},
                        "description": "连接线列表, 每项: {from_id, to_id, label}"
                    },
                    "layout": {
                        "type": "string",
                        "enum": ["auto", "custom"],
                        "default": "auto",
                        "description": "布局模式: auto(自动排列) / custom(使用提供的坐标)"
                    },
                    "backup": {"type": "boolean", "default": True, "description": "是否自动备份"},
                    "output_path": {"type": "string", "description": "输出路径（可选）"}
                },
                "required": ["word_path", "target"]
            }
        ),

        # ---- 经验知识库工具 ----
        Tool(
            name="get_experience",
            description="查询历史经验教训。支持按分类、关键词、重要度检索。建议在开始复杂任务前调用，避免重复犯错。",
            inputSchema={
                "type": "object",
                "properties": {
                    "category": {"type": "string", "description": "按分类过滤，如 'visio_com', 'word_ole', 'backup'（可选）"},
                    "keyword": {"type": "string", "description": "按关键词搜索，匹配经验内容（可选）"},
                    "days": {"type": "integer", "default": 30, "description": "回溯天数"},
                    "importance": {
                        "type": "string",
                        "enum": ["high", "normal", "low"],
                        "description": "按重要度过滤（可选）"
                    }
                }
            }
        ),
        Tool(
            name="log_experience",
            description="记录经验教训。必须有来源(source)和证据(evidence)。verified/user_feedback来源直接入库，inferred来源需审核后才可被查询到。自动去重。",
            inputSchema={
                "type": "object",
                "properties": {
                    "category": {
                        "type": "string",
                        "description": "经验分类，如 'visio_com', 'word_ole', 'word_format', 'backup', 'excel', 'general'"
                    },
                    "lessons": {
                        "type": "array",
                        "items": {"type": "string"},
                        "description": "经验教训列表，每条是一个简洁的知识点"
                    },
                    "source": {
                        "type": "string",
                        "enum": ["verified", "user_feedback", "inferred"],
                        "description": "来源: verified=经过执行验证的事实, user_feedback=用户明确反馈, inferred=大模型推理(需审核)"
                    },
                    "evidence": {
                        "type": "string",
                        "description": "证据: 错误信息、执行结果、用户原话等。verified/user_feedback来源必填"
                    },
                    "summary": {"type": "string", "description": "操作摘要（可选）"},
                    "context": {"type": "string", "description": "背景描述（可选）"},
                    "importance": {
                        "type": "string",
                        "enum": ["high", "normal", "low"],
                        "default": "normal",
                        "description": "重要程度"
                    }
                },
                "required": ["category", "lessons", "source"]
            }
        ),
        Tool(
            name="review_experience",
            description="审核待审经验。列出/批准/驳回 inferred 来源的经验条目。",
            inputSchema={
                "type": "object",
                "properties": {
                    "action": {
                        "type": "string",
                        "enum": ["list", "approve", "reject"],
                        "description": "操作: list=列出待审, approve=批准, reject=驳回"
                    },
                    "entry_id": {"type": "string", "description": "指定条目的时间戳（可选，不填则操作所有待审条目）"},
                    "category": {"type": "string", "description": "按分类过滤（可选）"}
                },
                "required": ["action"]
            }
        ),

        # ---- 查询辅助工具 ----
        Tool(
            name="validate_document",
            description="对已生成的文档执行分级验证，返回质量评分和错误详情",
            inputSchema={
                "type": "object",
                "properties": {
                    "file_path": {"type": "string", "description": "文档文件路径"},
                    "expected_schema": {
                        "type": "object",
                        "description": "预期的 Schema（可选）"
                    }
                },
                "required": ["file_path"]
            }
        ),
        Tool(
            name="list_templates",
            description="列出可用的文档模板",
            inputSchema={
                "type": "object",
                "properties": {
                    "doc_type": {
                        "type": "string",
                        "enum": ["word", "excel", "diagram", "all"],
                        "default": "all",
                        "description": "按文档类型过滤"
                    }
                }
            }
        ),
        Tool(
            name="get_template_info",
            description="获取指定模板的详细信息（Schema 结构、占位符、示例数据）",
            inputSchema={
                "type": "object",
                "properties": {
                    "template_name": {"type": "string", "description": "模板名称"}
                },
                "required": ["template_name"]
            }
        ),
        Tool(
            name="get_schema_spec",
            description="获取指定文档类型的 JSON Schema 规范，供外部模型构造输入时参考",
            inputSchema={
                "type": "object",
                "properties": {
                    "doc_type": {
                        "type": "string",
                        "enum": ["word", "excel", "diagram"],
                        "description": "文档类型"
                    }
                },
                "required": ["doc_type"]
            }
        ),
    ]


# ========== 工具路由 ==========

@server.call_tool()
async def call_tool(name: str, arguments: dict) -> list[TextContent]:
    """工具调用路由"""
    try:
        if name == "auto_generate":
            from tools.auto_generate import auto_generate
            result = await auto_generate(
                request=arguments["request"],
                doc_type=arguments.get("doc_type"),
                template=arguments.get("template"),
                output_dir=arguments.get("output_dir"),
                max_retries=arguments.get("max_retries", 3),
                config=CONFIG,
            )

        elif name == "auto_edit":
            from tools.auto_edit import auto_edit
            result = await auto_edit(
                file_path=arguments["file_path"],
                instruction=arguments["instruction"],
                backup=arguments.get("backup", True),
                output_path=arguments.get("output_path"),
                config=CONFIG,
            )

        elif name == "analyze_requirement":
            from tools.analyze import analyze_requirement
            result = await analyze_requirement(
                request=arguments["request"],
                config=CONFIG,
            )

        elif name == "generate_word":
            from tools.generate_word import generate_word
            schema = WordSchema(**arguments["schema"])
            result = generate_word(
                schema=schema,
                template=arguments["template"],
                output_path=arguments["output_path"],
                config=CONFIG,
            )

        elif name == "generate_excel":
            from tools.generate_excel import generate_excel
            schema = ExcelSchema(**arguments["schema"])
            result = generate_excel(
                schema=schema,
                template=arguments.get("template"),
                output_path=arguments["output_path"],
                config=CONFIG,
            )

        elif name == "generate_diagram":
            from tools.generate_diagram import generate_diagram
            schema = DiagramSchema(**arguments["schema"])
            result = generate_diagram(
                schema=schema,
                engine=arguments.get("engine", "mermaid"),
                output_format=arguments.get("output_format", "svg"),
                output_path=arguments["output_path"],
                config=CONFIG,
            )

        elif name == "read_document":
            from tools.read_document import read_document
            result = read_document(
                file_path=arguments["file_path"],
                include_content=arguments.get("include_content", True),
                include_style=arguments.get("include_style", False),
                max_depth=arguments.get("max_depth", 2),
            )

        elif name == "edit_document":
            from tools.edit_document import edit_document
            edits = [EditInstruction(**e) for e in arguments["edits"]]
            result = edit_document(
                file_path=arguments["file_path"],
                edits=edits,
                backup=arguments.get("backup", True),
                output_path=arguments.get("output_path"),
            )

        elif name == "validate_document":
            from tools.validate import validate_document
            result = validate_document(
                file_path=arguments["file_path"],
                expected_schema=arguments.get("expected_schema"),
            )

        elif name == "generate_visio":
            from tools.generate_visio import generate_visio
            schema = VisioSchema(**arguments["schema"])
            result = generate_visio(
                schema=schema,
                output_path=arguments["output_path"],
                engine=arguments.get("engine", "com"),
                config=CONFIG,
            )

        elif name == "read_visio":
            from tools.edit_visio import read_visio
            result = read_visio(
                file_path=arguments["file_path"],
                config=CONFIG,
            )

        elif name == "edit_visio":
            from tools.edit_visio import edit_visio
            result = edit_visio(
                file_path=arguments["file_path"],
                edits=arguments["edits"],
                backup=arguments.get("backup", True),
                output_path=arguments.get("output_path"),
                config=CONFIG,
            )

        elif name == "replace_image":
            from tools.replace_image import replace_image
            target = arguments.get("target", "0")
            # 尝试将 target 转为 int（如果是数字索引）
            try:
                target = int(target)
            except (ValueError, TypeError):
                pass
            result = replace_image(
                file_path=arguments["file_path"],
                new_image_path=arguments["new_image_path"],
                target=target,
                backup=arguments.get("backup", True),
                output_path=arguments.get("output_path"),
            )

        elif name == "list_images":
            from tools.replace_image import list_images
            result = list_images(file_path=arguments["file_path"])

        elif name == "embed_visio_to_word":
            from tools.embed_visio import embed_visio_to_word
            result = embed_visio_to_word(
                word_path=arguments["word_path"],
                visio_path=arguments.get("visio_path"),
                target=arguments.get("target"),
                shapes=arguments.get("shapes"),
                connectors=arguments.get("connectors"),
                layout=arguments.get("layout", "auto"),
                backup=arguments.get("backup", True),
                output_path=arguments.get("output_path"),
                config=CONFIG,
            )

        elif name == "get_experience":
            from utils.experience import get_experience_logger
            exp_logger = get_experience_logger()
            category = arguments.get("category")
            keyword = arguments.get("keyword")
            days = arguments.get("days", 30)
            importance = arguments.get("importance")
            entries = exp_logger.query(
                category=category, keyword=keyword,
                days=days, importance=importance,
            )
            categories = exp_logger.get_all_categories(days=days)
            result = {
                "success": True,
                "matched_entries": entries[-30:],
                "available_categories": categories,
                "total_matched": len(entries),
            }

        elif name == "log_experience":
            from utils.experience import get_experience_logger
            exp_logger = get_experience_logger()
            result = exp_logger.log(
                category=arguments["category"],
                lessons=arguments["lessons"],
                source=arguments.get("source", "inferred"),
                evidence=arguments.get("evidence", ""),
                summary=arguments.get("summary", ""),
                context=arguments.get("context", ""),
                importance=arguments.get("importance", "normal"),
            )

        elif name == "review_experience":
            from utils.experience import get_experience_logger
            exp_logger = get_experience_logger()
            result = exp_logger.review(
                action=arguments["action"],
                entry_id=arguments.get("entry_id"),
                category=arguments.get("category"),
            )

        elif name == "list_templates":
            from tools.query import list_templates
            result = list_templates(
                doc_type=arguments.get("doc_type", "all"),
                config=CONFIG,
            )

        elif name == "get_template_info":
            from tools.query import get_template_info
            result = get_template_info(
                template_name=arguments["template_name"],
                config=CONFIG,
            )

        elif name == "get_schema_spec":
            from tools.query import get_schema_spec
            result = get_schema_spec(doc_type=arguments["doc_type"])

        else:
            result = {"success": False, "error": f"未知工具: {name}"}

    except Exception as e:
        result = {"success": False, "error": str(e), "tool": name}

    return [TextContent(
        type="text",
        text=json.dumps(result, ensure_ascii=False, indent=2, default=str)
    )]


# ========== 入口 ==========

def main():
    """启动 MCP Server"""
    import asyncio
    from mcp.server.stdio import stdio_server

    async def run():
        async with stdio_server() as (read_stream, write_stream):
            await server.run(
                read_stream,
                write_stream,
                server.create_initialization_options()
            )

    asyncio.run(run())


if __name__ == "__main__":
    main()
