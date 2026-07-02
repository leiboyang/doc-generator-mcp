"""文档验证工具"""

from validation.validator import validate_document as _validate


def validate_document(file_path: str, expected_schema: dict = None) -> dict:
    """MCP 工具：验证文档"""
    return _validate(file_path, expected_schema)
