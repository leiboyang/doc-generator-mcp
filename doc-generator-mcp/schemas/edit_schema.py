"""编辑指令的 Pydantic Schema"""

from pydantic import BaseModel, Field, model_validator
from typing import Literal, Any
from enum import Enum


class EditAction(str, Enum):
    REPLACE_TEXT = "replace_text"
    UPDATE_TABLE = "update_table"
    INSERT_PARAGRAPH = "insert_paragraph"
    DELETE_PARAGRAPH = "delete_paragraph"
    UPDATE_CELL = "update_cell"
    INSERT_ROW = "insert_row"
    UPDATE_PARAGRAPH_FORMAT = "update_paragraph_format"


# 每种 action 必需的字段映射
_ACTION_REQUIRED_FIELDS: dict[EditAction, list[str]] = {
    EditAction.REPLACE_TEXT: ["search", "replace"],
    EditAction.UPDATE_TABLE: ["table_index", "data"],
    EditAction.INSERT_PARAGRAPH: ["after_index", "text"],
    EditAction.DELETE_PARAGRAPH: ["index"],
    EditAction.UPDATE_CELL: ["row", "column", "value"],
    EditAction.INSERT_ROW: ["table_index", "at_row"],
    EditAction.UPDATE_PARAGRAPH_FORMAT: ["index"],
}


class EditInstruction(BaseModel):
    """单条编辑指令"""
    action: EditAction = Field(..., description="操作类型")

    # replace_text 参数
    search: str | None = Field(None, description="要搜索的文本")
    replace: str | None = Field(None, description="替换为的文本")
    scope: Literal["paragraph", "table_cell", "all"] = Field(
        default="paragraph", description="搜索范围"
    )

    # update_table 参数
    table_index: int | None = Field(None, ge=0, description="表格索引")
    data: list[list] | None = Field(None, description="新数据")
    start_row: int = Field(default=1, ge=0, description="起始行")

    # insert_paragraph 参数
    after_index: int | None = Field(None, ge=0, description="在哪个段落之后插入")
    text: str | None = Field(None, description="新段落文本")
    inherit_style_from: Literal["adjacent", "normal"] = Field(
        default="adjacent", description="样式继承策略"
    )

    # delete_paragraph 参数
    index: int | None = Field(None, ge=0, description="段落索引")

    # update_cell 参数
    row: int | None = Field(None, ge=0, description="行号")
    column: int | None = Field(None, ge=0, description="列号")
    value: Any = Field(None, description="新值")

    # insert_row 参数
    at_row: int | None = Field(None, ge=0, description="在哪一行插入")

    # update_paragraph_format 参数
    alignment: Literal["left", "center", "right", "justify"] | None = Field(
        None, description="对齐方式"
    )
    space_before: float | None = Field(None, ge=0, description="段前间距（磅）")
    space_after: float | None = Field(None, ge=0, description="段后间距（磅）")
    line_spacing: float | None = Field(None, ge=0, description="行间距（倍数）")
    first_line_indent: float | None = Field(None, ge=0, description="首行缩进（厘米）")
    left_indent: float | None = Field(None, ge=0, description="左缩进（厘米）")
    right_indent: float | None = Field(None, ge=0, description="右缩进（厘米）")

    @model_validator(mode="after")
    def validate_action_fields(self) -> "EditInstruction":
        """校验当前 action 所需的字段是否已提供"""
        required = _ACTION_REQUIRED_FIELDS.get(self.action, [])
        missing = []
        for field_name in required:
            if getattr(self, field_name, None) is None:
                missing.append(field_name)
        if missing:
            raise ValueError(
                f"action='{self.action.value}' 缺少必需字段: {', '.join(missing)}"
            )
        return self


class EditResult(BaseModel):
    """编辑结果"""
    success: bool
    file_path: str = ""
    backup_path: str = ""
    changes_applied: int = 0
    summary: str = ""
    errors: list[str] = Field(default_factory=list, description="错误列表")
    integrity_check: dict = Field(default_factory=dict, description="完整性检查结果")
