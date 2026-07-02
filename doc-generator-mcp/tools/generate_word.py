"""Word 文档生成工具"""

import shutil
from pathlib import Path

from schemas.word_schema import WordSchema
from engines.template_renderer import (
    render_word_from_template,
    render_word_from_scratch,
    get_template_context,
)


def generate_word(
    schema: WordSchema,
    template: str,
    output_path: str,
    config: dict = None,
) -> dict:
    """生成 Word 文档

    Args:
        schema: Word 文档 Schema
        template: 模板名称
        output_path: 输出文件路径
        config: 配置字典

    Returns:
        生成结果字典
    """
    config = config or {}
    template_dir = Path(config.get("paths", {}).get("template_dir", "./templates"))
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    context = get_template_context(schema.model_dump())

    # 查找模板文件
    template_path = template_dir / "word" / f"{template}.docx"

    if template_path.exists():
        # 有模板：使用 docxtpl 渲染
        render_word_from_template(str(template_path), str(output), context)
        return {
            "success": True,
            "file_path": str(output),
            "mode": "template",
            "template_used": template,
            "message": f"使用模板 '{template}' 生成 Word 文档成功",
        }
    else:
        # 无模板：从零创建
        render_word_from_scratch(str(output), context)
        return {
            "success": True,
            "file_path": str(output),
            "mode": "scratch",
            "template_used": None,
            "message": f"无匹配模板，从零创建 Word 文档成功（建议使用模板以获得更好的格式控制）",
        }
