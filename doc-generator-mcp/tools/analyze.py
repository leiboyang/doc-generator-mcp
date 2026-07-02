"""需求分析工具 —— 将自然语言需求转为结构化分析结果"""

import json
import logging
from datetime import datetime

from engines.llm_client import LLMClient
from schemas.word_schema import WordSchema
from schemas.excel_schema import ExcelSchema
from schemas.diagram_schema import DiagramSchema

logger = logging.getLogger(__name__)

# 各类型的 JSON Schema（供 LLM 参考）
SCHEMA_SPECS = {
    "word": WordSchema.model_json_schema(),
    "excel": ExcelSchema.model_json_schema(),
    "diagram": DiagramSchema.model_json_schema(),
}

ANALYSIS_SYSTEM_PROMPT = """你是一个文档需求分析专家。根据用户的自然语言描述，分析并返回结构化的需求信息。

你必须返回严格的 JSON，包含以下字段：

{
  "doc_type": "word | excel | diagram",  // 判断文档类型
  "confidence": 0.0 ~ 1.0,              // 判断置信度
  "title": "文档标题",
  "description": "需求简述",
  "suggested_template": "模板名称或 null",
  "ambiguities": ["不明确的信息列表"],
  "estimated_complexity": "low | medium | high",
  "extracted_schema": { ... },           // 根据 doc_type 填充对应的 Schema
  "generation_hints": {                  // 生成建议
    "prefer_template": true/false,
    "needs_table": true/false,
    "needs_chart": true/false,
    "needs_formula": true/false,
    "estimated_sections": 0,
    "estimated_rows": 0
  }
}

判断规则：
- 包含"报告""文档""方案""总结"等 → word
- 包含"表格""数据""统计""对比"等 → excel
- 包含"流程图""架构图""组织图""关系图"等 → diagram
- 不确定时，优先选 word，降低 confidence

extracted_schema 必须严格符合对应类型的 JSON Schema 结构。
- word: 需要 title, sections (每个 section 有 heading 和 content)
- excel: 需要 data (headers + rows), sheet_name
- diagram: 需要 title, nodes (id + label + type), connections (from + to), diagram_type

如果需求中缺少具体数据，请根据上下文合理推断并填充，同时在 ambiguities 中说明。"""


async def analyze_requirement(request: str, config: dict = None) -> dict:
    """分析自然语言需求

    Args:
        request: 自然语言需求描述
        config: 配置字典

    Returns:
        结构化分析结果
    """
    config = config or {}
    llm = LLMClient(config)

    if not llm.is_available:
        # 无 LLM：返回基础分析（仅规则匹配）
        analysis = _rule_based_analysis(request)
        return {
            "success": True,
            "analysis": analysis,
            "raw_request": request,
        }

    # 有 LLM：完整分析
    try:
        messages = [
            {"role": "system", "content": ANALYSIS_SYSTEM_PROMPT},
            {
                "role": "user",
                "content": (
                    f"用户需求：{request}\n\n"
                    f"当前日期：{datetime.now().strftime('%Y-%m-%d')}\n\n"
                    f"请分析此需求并返回 JSON。"
                ),
            },
        ]

        result = llm.chat_json(messages)

        # 验证 extracted_schema 是否符合对应类型的 Pydantic 模型
        doc_type = result.get("doc_type", "word")
        schema_data = result.get("extracted_schema", {})
        validation_errors = _validate_extracted_schema(doc_type, schema_data)

        result["validation_errors"] = validation_errors
        result["analyzed_at"] = datetime.now().isoformat()

        return {
            "success": True,
            "analysis": result,
            "raw_request": request,
        }

    except Exception as e:
        logger.error(f"LLM 分析失败: {e}")
        # 降级到规则分析
        fallback = _rule_based_analysis(request)
        fallback["warning"] = f"LLM 分析失败，已降级为规则分析: {e}"
        return {
            "success": True,
            "analysis": fallback,
            "raw_request": request,
        }


def _rule_based_analysis(request: str) -> dict:
    """基于规则的基础分析（无 LLM 时使用）"""
    request_lower = request.lower()

    # 文档类型判断
    word_keywords = ["报告", "文档", "方案", "总结", "说明", "合同", "协议", "说明书", "report", "document"]
    excel_keywords = ["表格", "数据", "统计", "对比", "汇总", "清单", "报表", "spreadsheet", "excel"]
    diagram_keywords = ["流程图", "架构图", "组织图", "关系图", "拓扑", "流程", "diagram", "flowchart"]

    word_score = sum(1 for k in word_keywords if k in request_lower)
    excel_score = sum(1 for k in excel_keywords if k in request_lower)
    diagram_score = sum(1 for k in diagram_keywords if k in request_lower)

    scores = {"word": word_score, "excel": excel_score, "diagram": diagram_score}
    doc_type = max(scores, key=scores.get)
    max_score = scores[doc_type]

    if max_score == 0:
        doc_type = "word"
        confidence = 0.3
    else:
        confidence = min(0.5 + max_score * 0.15, 0.95)

    # 提取标题（简单规则：取"生成/创建/写"后面的内容）
    title = request
    for prefix in ["生成", "创建", "写一份", "写一个", "制作", "帮我生成", "请生成", "请创建"]:
        if request.startswith(prefix):
            title = request[len(prefix):].strip()
            break

    return {
        "doc_type": doc_type,
        "confidence": confidence,
        "title": title[:100],
        "description": request[:200],
        "suggested_template": None,
        "ambiguities": ["未使用 LLM 分析，建议配置 API Key 以获得更精准的结果"],
        "estimated_complexity": "medium",
        "extracted_schema": _build_default_schema(doc_type, title),
        "generation_hints": {
            "prefer_template": False,
            "needs_table": any(k in request_lower for k in ["表格", "数据", "表"]),
            "needs_chart": any(k in request_lower for k in ["图表", "图", "柱状", "饼"]),
            "needs_formula": any(k in request_lower for k in ["公式", "计算", "合计"]),
            "estimated_sections": 3,
            "estimated_rows": 5,
        },
        "analyzed_at": datetime.now().isoformat(),
    }


def _build_default_schema(doc_type: str, title: str) -> dict:
    """构建默认 Schema（占位）"""
    if doc_type == "word":
        return {
            "title": title,
            "date": datetime.now().strftime("%Y-%m-%d"),
            "sections": [
                {"heading": "概述", "content": ["（待填充）"]},
            ],
        }
    elif doc_type == "excel":
        return {
            "sheet_name": "Sheet1",
            "data": {
                "headers": ["项目", "数值"],
                "rows": [],
            },
        }
    elif doc_type == "diagram":
        return {
            "title": title,
            "diagram_type": "flowchart",
            "nodes": [
                {"id": "start", "label": "开始", "type": "terminator"},
                {"id": "end", "label": "结束", "type": "terminator"},
            ],
            "connections": [
                {"from": "start", "to": "end"},
            ],
        }
    return {}


def _validate_extracted_schema(doc_type: str, schema_data: dict) -> list:
    """验证 extracted_schema 是否符合 Pydantic 模型"""
    errors = []
    try:
        if doc_type == "word":
            WordSchema(**schema_data)
        elif doc_type == "excel":
            ExcelSchema(**schema_data)
        elif doc_type == "diagram":
            DiagramSchema(**schema_data)
    except Exception as e:
        errors.append(str(e))
    return errors
