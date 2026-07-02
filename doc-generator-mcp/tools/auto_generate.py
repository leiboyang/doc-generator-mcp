"""自动生成工具 —— 一句话需求 → 完整文档的全流程编排"""

import json
import logging
from pathlib import Path
from datetime import datetime

from engines.llm_client import LLMClient
from tools.analyze import analyze_requirement, _rule_based_analysis
from tools.generate_word import generate_word
from tools.generate_excel import generate_excel
from tools.generate_diagram import generate_diagram
from tools.validate import validate_document
from schemas.word_schema import WordSchema
from schemas.excel_schema import ExcelSchema
from schemas.diagram_schema import DiagramSchema

logger = logging.getLogger(__name__)

# Schema 修复提示模板
REPAIR_SYSTEM_PROMPT = """你是一个文档 Schema 修复专家。上一次生成的 Schema 在验证时发现了错误，请修复这些问题。

原始 Schema：
{original_schema}

验证错误：
{errors}

请返回修复后的完整 JSON Schema，确保：
1. 所有必填字段都存在
2. 数据类型正确
3. 表格行列数匹配
4. 节点 ID 引用正确

只返回修复后的 JSON，不要其他内容。"""


async def auto_generate(
    request: str,
    doc_type: str = None,
    template: str = None,
    output_dir: str = None,
    max_retries: int = 3,
    config: dict = None,
) -> dict:
    """一句话需求自动生成文档

    完整流程：
    1. 需求分析 → 确定文档类型和 Schema
    2. 模板选择 → 匹配最佳模板
    3. 文档生成 → 调用对应生成器
    4. 验证检查 → 质量评分
    5. 失败重试 → LLM 诊断修复（最多 max_retries 次）

    Args:
        request: 自然语言需求
        doc_type: 文档类型提示（可选）
        template: 模板名称（可选）
        output_dir: 输出目录（可选）
        max_retries: 最大重试次数
        config: 配置字典

    Returns:
        生成结果
    """
    config = config or {}
    llm = LLMClient(config)

    # 默认输出目录
    if not output_dir:
        output_dir = str(Path(config.get("paths", {}).get("output_dir", "./output")).resolve())
    Path(output_dir).mkdir(parents=True, exist_ok=True)

    # ===== Step 1: 需求分析 =====
    logger.info(f"[auto_generate] Step 1: 分析需求 - {request[:50]}...")
    analysis_result = await analyze_requirement(request, config)

    if not analysis_result.get("success"):
        return {"success": False, "error": f"需求分析失败: {analysis_result.get('error', '未知错误')}"}

    analysis = analysis_result.get("analysis", analysis_result)

    # 确定文档类型
    final_doc_type = doc_type or analysis.get("doc_type", "word")
    confidence = analysis.get("confidence", 0.5)

    # 确定输出文件名
    title = analysis.get("title", "未命名文档")
    safe_title = _sanitize_filename(title)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    ext_map = {"word": ".docx", "excel": ".xlsx", "diagram": ".svg"}
    ext = ext_map.get(final_doc_type, ".docx")
    output_path = str(Path(output_dir) / f"{safe_title}_{timestamp}{ext}")

    # ===== Step 2: 获取 Schema =====
    logger.info(f"[auto_generate] Step 2: 准备 Schema (doc_type={final_doc_type})")
    schema_data = analysis.get("extracted_schema", {})

    # 如果 LLM 不可用，使用规则分析的结果可能需要补充
    if not llm.is_available and not schema_data:
        schema_data = analysis.get("extracted_schema", {})

    # ===== Step 3-5: 生成 → 验证 → 重试循环 =====
    attempts = []
    best_result = None

    for attempt in range(1, max_retries + 1):
        logger.info(f"[auto_generate] Step 3: 第 {attempt} 次尝试生成...")

        # 生成文档
        gen_result = _generate_document(
            doc_type=final_doc_type,
            schema_data=schema_data,
            template=template or analysis.get("suggested_template"),
            output_path=output_path,
            config=config,
        )

        if not gen_result.get("success"):
            error_msg = gen_result.get("error", "生成失败")
            attempts.append({"attempt": attempt, "success": False, "error": error_msg})

            # 尝试用 LLM 修复 Schema
            if llm.is_available and attempt < max_retries:
                schema_data = _repair_schema_with_llm(llm, schema_data, error_msg, final_doc_type)
            continue

        # 验证文档
        logger.info(f"[auto_generate] Step 4: 验证文档...")
        val_result = validate_document(output_path)
        quality_score = val_result.get("quality_score", 0)
        critical_errors = val_result.get("critical_errors", [])

        attempt_record = {
            "attempt": attempt,
            "success": True,
            "generation": gen_result,
            "validation": val_result,
            "quality_score": quality_score,
        }
        attempts.append(attempt_record)

        # 记录最佳结果
        if best_result is None or quality_score > best_result.get("quality_score", 0):
            best_result = {
                "success": True,
                "file_path": gen_result.get("file_path", output_path),
                "doc_type": final_doc_type,
                "template_used": gen_result.get("template_used"),
                "validation_score": quality_score,
                "attempts": attempt,
                "attempt_details": attempts,
                "analysis": {
                    "title": title,
                    "confidence": confidence,
                    "ambiguities": analysis.get("ambiguities", []),
                },
            }

        # 质量达标（>= 70 分且无严重错误）
        if quality_score >= 70 and not critical_errors:
            logger.info(f"[auto_generate] 第 {attempt} 次尝试通过验证 (score={quality_score})")
            break

        # 需要重试：用 LLM 诊断修复
        if llm.is_available and attempt < max_retries:
            logger.info(f"[auto_generate] 质量不足 (score={quality_score})，尝试 LLM 修复...")
            schema_data = _repair_with_validation_feedback(
                llm, schema_data, val_result, final_doc_type
            )
        elif attempt < max_retries:
            # 无 LLM，无法自动修复，停止重试
            break

    if best_result is None:
        return {
            "success": False,
            "error": f"经过 {max_retries} 次尝试仍无法生成合格文档",
            "attempts": attempts,
        }

    # 添加警告信息
    warnings = []
    if confidence < 0.7:
        warnings.append(f"需求分析置信度较低 ({confidence:.0%})，建议检查生成结果")
    if analysis.get("ambiguities"):
        warnings.extend(analysis["ambiguities"])
    if best_result.get("validation_score", 0) < 80:
        warnings.append(f"文档质量评分 {best_result['validation_score']}，建议人工复核")

    if warnings:
        best_result["warnings"] = warnings

    return best_result


def _generate_document(
    doc_type: str,
    schema_data: dict,
    template: str,
    output_path: str,
    config: dict,
) -> dict:
    """根据文档类型调用对应的生成器"""
    try:
        if doc_type == "word":
            schema = WordSchema(**schema_data)
            return generate_word(
                schema=schema,
                template=template or "default",
                output_path=output_path,
                config=config,
            )
        elif doc_type == "excel":
            schema = ExcelSchema(**schema_data)
            return generate_excel(
                schema=schema,
                template=template,
                output_path=output_path,
                config=config,
            )
        elif doc_type == "diagram":
            schema = DiagramSchema(**schema_data)
            return generate_diagram(
                schema=schema,
                engine=config.get("generation", {}).get("default_diagram_engine", "mermaid"),
                output_path=output_path,
                config=config,
            )
        else:
            return {"success": False, "error": f"不支持的文档类型: {doc_type}"}

    except Exception as e:
        return {"success": False, "error": f"Schema 验证或生成失败: {e}"}


def _repair_schema_with_llm(
    llm: LLMClient, schema_data: dict, error_msg: str, doc_type: str
) -> dict:
    """使用 LLM 修复 Schema 错误"""
    try:
        messages = [
            {"role": "system", "content": REPAIR_SYSTEM_PROMPT.format(
                original_schema=json.dumps(schema_data, ensure_ascii=False, indent=2),
                errors=error_msg,
            )},
            {"role": "user", "content": "请修复上述 Schema 中的问题。"},
        ]
        repaired = llm.chat_json(messages)
        logger.info(f"[auto_generate] LLM 修复 Schema 成功")
        return repaired
    except Exception as e:
        logger.warning(f"[auto_generate] LLM 修复失败: {e}")
        return schema_data


def _repair_with_validation_feedback(
    llm: LLMClient, schema_data: dict, validation_result: dict, doc_type: str
) -> dict:
    """将验证反馈传给 LLM，让其改进 Schema"""
    try:
        critical = validation_result.get("critical_errors", [])
        important = validation_result.get("important_errors", [])
        suggestions = validation_result.get("suggestions", [])

        feedback = "验证结果：\n"
        if critical:
            feedback += f"\n严重错误 ({len(critical)}):\n" + "\n".join(f"- {e}" for e in critical)
        if important:
            feedback += f"\n重要问题 ({len(important)}):\n" + "\n".join(f"- {e}" for e in important)
        if suggestions:
            feedback += f"\n建议 ({len(suggestions)}):\n" + "\n".join(f"- {s}" for s in suggestions)

        messages = [
            {"role": "system", "content": REPAIR_SYSTEM_PROMPT.format(
                original_schema=json.dumps(schema_data, ensure_ascii=False, indent=2),
                errors=feedback,
            )},
            {"role": "user", "content": "请根据验证反馈修复 Schema。"},
        ]
        repaired = llm.chat_json(messages)
        logger.info(f"[auto_generate] 验证反馈修复成功")
        return repaired
    except Exception as e:
        logger.warning(f"[auto_generate] 验证反馈修复失败: {e}")
        return schema_data


def _sanitize_filename(name: str) -> str:
    """清理文件名中的非法字符"""
    import re
    # 移除 Windows 文件名非法字符
    safe = re.sub(r'[<>:"/\\|?*]', '', name)
    # 截断过长文件名
    if len(safe) > 50:
        safe = safe[:50]
    return safe.strip()
