"""自动编辑工具 —— 一句话指令 → 定向编辑已有文档"""

import json
import logging
from pathlib import Path

from engines.llm_client import LLMClient
from tools.read_document import read_document
from tools.edit_document import edit_document
from tools.validate import validate_document
from schemas.edit_schema import EditInstruction

logger = logging.getLogger(__name__)

EDIT_SYSTEM_PROMPT = """你是一个文档编辑专家。用户会提供一份文档的结构描述和一条修改指令。
你需要分析修改意图，生成精确的编辑操作列表。

可用的编辑操作类型：
- replace_text: 替换文本。参数: search(搜索文本), replace(替换文本), scope("all"=全文搜索)
- update_table: 更新表格数据。参数: table_index(表格索引), data(二维数组), start_row(起始行,默认1)
- insert_paragraph: 在指定段落后插入。参数: after_index(段落索引), text(新段落文本), inherit_style_from("adjacent")
- delete_paragraph: 删除段落。参数: index(段落索引)
- update_cell: 更新单个单元格。参数: table_index, row, column, value
- insert_row: 插入行(仅Excel)。参数: at_row

每条编辑操作的 JSON 格式：
{"action": "replace_text", "search": "原文本", "replace": "新文本", "scope": "all"}
{"action": "update_table", "table_index": 0, "data": [["A","B"],["1","2"]], "start_row": 1}
{"action": "insert_paragraph", "after_index": 5, "text": "新段落内容", "inherit_style_from": "adjacent"}
{"action": "delete_paragraph", "index": 3}
{"action": "update_cell", "table_index": 0, "row": 0, "column": 1, "value": "新值"}

重要原则：
1. 只修改需要改的部分，不要动其他内容（保持格式不变）
2. 替换文本时，search 必须是文档中确实存在的文本
3. 插入段落时，after_index 必须是有效的段落索引
4. 尽量用最少的操作完成修改

你必须返回严格的 JSON：
{
  "analysis": "对修改意图的分析",
  "edits": [ ... ],  // 编辑操作列表
  "confidence": 0.0 ~ 1.0,
  "warnings": ["注意事项"]
}"""


async def auto_edit(
    file_path: str,
    instruction: str,
    backup: bool = True,
    output_path: str = None,
    config: dict = None,
) -> dict:
    """一句话指令自动编辑文档

    流程：
    1. 读取文档结构
    2. LLM 分析修改意图，生成编辑指令
    3. 执行编辑
    4. 验证结果
    5. 失败则重试（带错误反馈）

    Args:
        file_path: 文档路径
        instruction: 自然语言修改指令
        backup: 是否备份
        output_path: 输出路径
        config: 配置字典

    Returns:
        编辑结果
    """
    config = config or {}
    llm = LLMClient(config)
    max_retries = config.get("generation", {}).get("max_retries", 3)

    # ===== Step 1: 读取文档结构 =====
    logger.info(f"[auto_edit] Step 1: 读取文档 - {file_path}")
    doc_info = read_document(file_path, include_content=True, include_style=False, max_depth=3)

    if not doc_info.get("success"):
        return {"success": False, "error": f"读取文档失败: {doc_info.get('error', '未知错误')}"}

    structure = doc_info.get("structure", {})

    # ===== Step 2-4: 分析 → 编辑 → 验证（带重试） =====
    attempts = []

    for attempt in range(1, max_retries + 1):
        logger.info(f"[auto_edit] Step 2: 第 {attempt} 次分析修改意图...")

        # LLM 分析修改意图
        if not llm.is_available:
            return {
                "success": False,
                "error": "自动编辑需要 LLM 支持，请配置 LLM_API_KEY 环境变量",
                "hint": "或者使用 edit_document 工具手动指定编辑指令",
            }

        try:
            edit_plan = _analyze_edit_intent(llm, instruction, doc_info, structure)
        except Exception as e:
            return {"success": False, "error": f"LLM 分析失败: {e}"}

        # 解析编辑指令
        edits_raw = edit_plan.get("edits", [])
        if not edits_raw:
            return {
                "success": False,
                "error": "LLM 未生成有效的编辑指令",
                "analysis": edit_plan.get("analysis", ""),
            }

        try:
            edits = [EditInstruction(**e) for e in edits_raw]
        except Exception as e:
            attempts.append({"attempt": attempt, "error": f"指令解析失败: {e}", "edits": edits_raw})
            if attempt >= max_retries:
                return {"success": False, "error": f"编辑指令解析失败: {e}", "attempts": attempts}
            continue

        # 执行编辑
        logger.info(f"[auto_edit] Step 3: 执行 {len(edits)} 条编辑指令...")
        edit_result = edit_document(
            file_path=file_path,
            edits=edits,
            backup=backup and attempt == 1,  # 只在第一次备份
            output_path=output_path,
        )

        if not edit_result.get("success"):
            errors = edit_result.get("errors", [])
            attempts.append({
                "attempt": attempt,
                "success": False,
                "errors": errors,
                "analysis": edit_plan.get("analysis", ""),
            })

            # 用错误反馈重试
            if attempt < max_retries:
                instruction = f"{instruction}\n\n[上次失败原因: {'; '.join(errors)}]"
            continue

        # 验证编辑结果
        logger.info(f"[auto_edit] Step 4: 验证编辑结果...")
        target_file = edit_result.get("file_path", file_path)
        val_result = validate_document(target_file)

        attempt_record = {
            "attempt": attempt,
            "success": True,
            "analysis": edit_plan.get("analysis", ""),
            "edits_count": len(edits),
            "edit_result": edit_result,
            "validation_score": val_result.get("quality_score", 0),
        }
        attempts.append(attempt_record)

        # 编辑后验证：确保文档结构没有被破坏
        post_edit_info = read_document(target_file, include_content=True, max_depth=2)
        if post_edit_info.get("success"):
            attempt_record["post_edit_structure"] = {
                "paragraphs": post_edit_info.get("structure", {}).get("total_paragraphs", 0),
                "tables": post_edit_info.get("structure", {}).get("total_tables", 0),
            }

        # 编辑成功
        return {
            "success": True,
            "file_path": target_file,
            "backup_path": edit_result.get("backup_path", ""),
            "analysis": edit_plan.get("analysis", ""),
            "changes_applied": edit_result.get("changes_applied", 0),
            "summary": edit_result.get("summary", ""),
            "validation_score": val_result.get("quality_score", 0),
            "confidence": edit_plan.get("confidence", 0),
            "warnings": edit_plan.get("warnings", []),
            "attempts": attempt,
        }

    return {
        "success": False,
        "error": f"经过 {max_retries} 次尝试仍无法完成编辑",
        "attempts": attempts,
    }


def _analyze_edit_intent(llm: LLMClient, instruction: str, doc_info: dict, structure: dict) -> dict:
    """使用 LLM 分析编辑意图并生成编辑指令"""

    # 构建文档结构摘要
    doc_summary = _summarize_structure(doc_info, structure)

    messages = [
        {"role": "system", "content": EDIT_SYSTEM_PROMPT},
        {
            "role": "user",
            "content": (
                f"修改指令：{instruction}\n\n"
                f"文档结构：\n{doc_summary}\n\n"
                f"请分析修改意图并生成编辑操作列表。"
            ),
        },
    ]

    return llm.chat_json(messages)


def _summarize_structure(doc_info: dict, structure: dict) -> str:
    """将文档结构转为 LLM 可读的摘要"""
    file_type = doc_info.get("file_type", "unknown")
    lines = [f"文件类型: {file_type}"]

    if file_type == "word":
        sections = structure.get("sections", [])
        lines.append(f"总段落数: {structure.get('total_paragraphs', 0)}")
        lines.append(f"表格数: {structure.get('total_tables', 0)}")
        lines.append("")
        lines.append("章节结构:")
        for sec in sections:
            lines.append(f"  [{sec.get('index', '?')}] {sec.get('heading', '(无标题)')} ({sec.get('style', '')})")
            for para in sec.get("paragraphs", [])[:5]:  # 最多展示 5 个段落
                text = para.get("text", "")[:80]
                lines.append(f"    [{para.get('index', '?')}] {text}")
            if len(sec.get("paragraphs", [])) > 5:
                lines.append(f"    ... 还有 {len(sec['paragraphs']) - 5} 段")

        # 表格摘要
        tables = structure.get("tables", [])
        if tables:
            lines.append("")
            lines.append("表格:")
            for t in tables:
                lines.append(f"  表格 {t.get('index', '?')}: {t.get('rows', 0)} 行 x {t.get('cols', 0)} 列")
                if "data" in t:
                    for i, row in enumerate(t["data"][:3]):
                        lines.append(f"    行{i}: {row}")

    elif file_type == "excel":
        sheets = structure.get("sheets", [])
        lines.append(f"工作表数: {len(sheets)}")
        for sheet in sheets:
            lines.append(f"  {sheet.get('name', '?')}: {sheet.get('max_row', 0)} 行 x {sheet.get('max_column', 0)} 列")
            if "preview_data" in sheet:
                for i, row in enumerate(sheet["preview_data"][:5]):
                    lines.append(f"    行{i}: {row}")

    return "\n".join(lines)
