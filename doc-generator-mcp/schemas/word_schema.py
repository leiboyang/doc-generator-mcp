"""Word 文档的 Pydantic Schema —— 严格验证 LLM 输出"""

from pydantic import BaseModel, Field, field_validator, ConfigDict
from typing import Literal


class WordSection(BaseModel):
    """文档章节"""
    heading: str = Field(..., min_length=1, max_length=200, description="章节标题")
    content: list[str] = Field(default_factory=list, description="章节段落列表")


class WordTable(BaseModel):
    """文档表格"""
    caption: str = Field(default="", max_length=200, description="表格标题")
    headers: list[str] = Field(..., min_length=1, description="表头列名")
    rows: list[list] = Field(default_factory=list, description="数据行")

    @field_validator("rows")
    @classmethod
    def validate_row_lengths(cls, v, info):
        if "headers" not in info.data:
            return v
        expected_cols = len(info.data["headers"])
        for i, row in enumerate(v):
            if len(row) != expected_cols:
                raise ValueError(
                    f"第 {i+1} 行列数 ({len(row)}) 与表头 ({expected_cols}) 不匹配"
                )
        return v


class WordSchema(BaseModel):
    """Word 文档的完整 Schema"""
    doc_type: Literal["word"] = "word"
    title: str = Field(..., min_length=1, max_length=300, description="文档标题")
    date: str = Field(default="", max_length=20, description="文档日期")
    sections: list[WordSection] = Field(
        ..., min_length=1, max_length=50, description="章节列表"
    )
    tables: list[WordTable] = Field(
        default_factory=list, max_length=20, description="表格列表"
    )

    @field_validator("sections")
    @classmethod
    def sections_not_empty(cls, v):
        for i, s in enumerate(v):
            if not s.heading.strip():
                raise ValueError(f"第 {i+1} 个章节标题不能为空白")
        return v

    model_config = ConfigDict(populate_by_name=True)
