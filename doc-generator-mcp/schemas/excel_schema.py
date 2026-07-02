"""Excel 文件的 Pydantic Schema"""

from pydantic import BaseModel, Field, field_validator, ConfigDict
from typing import Literal


class ExcelData(BaseModel):
    """表格数据"""
    headers: list[str] = Field(..., min_length=1, description="表头")
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


class ExcelFormula(BaseModel):
    """Excel 公式"""
    cell: str = Field(..., description="目标单元格，如 B5")
    formula: str = Field(..., description="公式字符串，如 =SUM(A1:A10)")
    description: str = Field(default="", max_length=200)


class ExcelChart(BaseModel):
    """Excel 图表"""
    type: Literal["column", "bar", "line", "pie", "area", "scatter"] = "column"
    title: str = Field(..., min_length=1, max_length=200)
    data_range: str = Field(..., description="数据区域，如 A1:B10")
    position: str = Field(default="F2", description="图表放置位置")


class ConditionalRule(BaseModel):
    """条件格式规则"""
    range: str = Field(..., description="应用范围，如 B2:B10")
    type: Literal["data_bar", "color_scale", "icon_set", "cell_value"] = "data_bar"
    color: str = Field(default="#4f46e5")


class ExcelFormatting(BaseModel):
    """格式配置"""
    header_style: str = Field(default="bold_with_border")
    number_format: str = Field(default="0.00")
    conditional_rules: list[ConditionalRule] = Field(default_factory=list)


class ExcelSchema(BaseModel):
    """Excel 文件的完整 Schema"""
    doc_type: Literal["excel"] = "excel"
    sheet_name: str = Field(default="Sheet1", max_length=31, description="工作表名称")
    data: ExcelData = Field(..., description="表格数据")
    formulas: list[ExcelFormula] = Field(default_factory=list, description="公式列表")
    formatting: ExcelFormatting = Field(default_factory=ExcelFormatting)
    charts: list[ExcelChart] = Field(default_factory=list, description="图表列表")

    model_config = ConfigDict(populate_by_name=True)
