"""查询辅助工具：list_templates, get_template_info, get_schema_spec"""

import json
from pathlib import Path

from schemas.word_schema import WordSchema
from schemas.excel_schema import ExcelSchema
from schemas.diagram_schema import DiagramSchema


def list_templates(doc_type: str = "all", config: dict = None) -> dict:
    """列出可用模板"""
    config = config or {}
    template_dir = Path(config.get("paths", {}).get("template_dir", "./templates"))

    templates = {"word": [], "excel": [], "diagram": []}

    # Word 模板
    if doc_type in ("all", "word"):
        word_dir = template_dir / "word"
        if word_dir.exists():
            for f in word_dir.glob("*.docx"):
                templates["word"].append({
                    "name": f.stem,
                    "file": str(f),
                    "type": "word",
                })

    # Excel 模板
    if doc_type in ("all", "excel"):
        excel_dir = template_dir / "excel"
        if excel_dir.exists():
            for f in excel_dir.glob("*.xltx"):
                templates["excel"].append({
                    "name": f.stem,
                    "file": str(f),
                    "type": "excel",
                })
            for f in excel_dir.glob("*.xlsx"):
                templates["excel"].append({
                    "name": f.stem,
                    "file": str(f),
                    "type": "excel",
                })

    # Diagram 配置
    if doc_type in ("all", "diagram"):
        diagram_dir = template_dir / "diagram"
        if diagram_dir.exists():
            for f in diagram_dir.glob("*.json"):
                templates["diagram"].append({
                    "name": f.stem,
                    "file": str(f),
                    "type": "diagram",
                })

    return {
        "success": True,
        "templates": templates,
        "total": sum(len(v) for v in templates.values()),
    }


def get_template_info(template_name: str, config: dict = None) -> dict:
    """获取模板详情"""
    config = config or {}
    template_dir = Path(config.get("paths", {}).get("template_dir", "./templates"))

    # 搜索模板文件
    for subdir in ["word", "excel", "diagram"]:
        for ext in ["*.docx", "*.xltx", "*.xlsx", "*.json"]:
            matches = list((template_dir / subdir).glob(f"{template_name}{ext}"))
            if matches:
                f = matches[0]
                return {
                    "success": True,
                    "name": template_name,
                    "type": subdir,
                    "file": str(f),
                    "size_bytes": f.stat().st_size,
                }

    return {
        "success": False,
        "error": f"模板 '{template_name}' 不存在",
    }


def get_schema_spec(doc_type: str) -> dict:
    """获取指定文档类型的 JSON Schema 规范"""
    schema_map = {
        "word": WordSchema,
        "excel": ExcelSchema,
        "diagram": DiagramSchema,
    }

    if doc_type not in schema_map:
        return {"success": False, "error": f"不支持的文档类型: {doc_type}"}

    model = schema_map[doc_type]
    json_schema = model.model_json_schema()

    return {
        "success": True,
        "doc_type": doc_type,
        "json_schema": json_schema,
    }
