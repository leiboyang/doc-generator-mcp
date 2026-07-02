"""Visio 图表 Pydantic Schema"""

from enum import Enum
from typing import Optional
from pydantic import BaseModel, Field, field_validator, model_validator


class VisioShapeType(str, Enum):
    """Visio 形状类型"""
    RECTANGLE = "rectangle"
    ROUNDED_RECT = "rounded_rect"
    ELLIPSE = "ellipse"
    DIAMOND = "diamond"
    PARALLELOGRAM = "parallelogram"
    CYLINDER = "cylinder"
    DOCUMENT = "document"
    PROCESS = "process"
    DECISION = "decision"
    TERMINATOR = "terminator"
    DATA = "data"
    PREDEFINED = "predefined"
    DELAY = "delay"
    ARROW = "arrow"
    CUSTOM = "custom"


class VisioConnectorType(str, Enum):
    """连接线类型"""
    STRAIGHT = "straight"
    ELBOW = "elbow"
    CURVED = "curved"


class VisioColor(str, Enum):
    """预定义颜色"""
    BLUE = "#4472C4"
    GREEN = "#548235"
    ORANGE = "#ED7D31"
    RED = "#C00000"
    GRAY = "#A5A5A5"
    LIGHT_BLUE = "#D6E4F0"
    LIGHT_GREEN = "#E2EFDA"
    LIGHT_ORANGE = "#FCE4D6"
    WHITE = "#FFFFFF"
    BLACK = "#000000"


class VisioShape(BaseModel):
    """Visio 形状"""
    id: str = Field(description="形状唯一标识", pattern=r"^[a-zA-Z0-9_-]+$")
    label: str = Field(description="形状文本标签")
    shape_type: VisioShapeType = Field(default=VisioShapeType.ROUNDED_RECT, description="形状类型")
    x: float = Field(default=0.0, description="X 坐标（英寸，从左到右）")
    y: float = Field(default=0.0, description="Y 坐标（英寸，从上到下）")
    width: float = Field(default=2.0, gt=0, description="宽度（英寸，须为正数）")
    height: float = Field(default=1.0, gt=0, description="高度（英寸，须为正数）")
    fill_color: str = Field(default=VisioColor.LIGHT_BLUE.value, description="填充颜色（十六进制）")
    line_color: str = Field(default=VisioColor.BLUE.value, description="边框颜色（十六进制）")
    font_size: int = Field(default=10, ge=1, le=72, description="字号")
    bold: bool = Field(default=False, description="是否加粗")

    @field_validator("id")
    @classmethod
    def validate_id(cls, v):
        if not v or not v.strip():
            raise ValueError("形状 ID 不能为空")
        return v.strip()


class VisioConnector(BaseModel):
    """Visio 连接线"""
    from_shape: str = Field(description="起始形状 ID")
    to_shape: str = Field(description="目标形状 ID")
    label: str = Field(default="", description="连接线标签")
    connector_type: VisioConnectorType = Field(default=VisioConnectorType.ELBOW, description="连接线类型")
    color: str = Field(default=VisioColor.BLACK.value, description="颜色（十六进制）")
    arrow_head: bool = Field(default=True, description="是否有箭头")

    @field_validator("from_shape", "to_shape")
    @classmethod
    def validate_shape_ref(cls, v):
        if not v or not v.strip():
            raise ValueError("形状引用不能为空")
        return v.strip()


class VisioPage(BaseModel):
    """Visio 页面"""
    name: str = Field(default="Page-1", description="页面名称")
    width: float = Field(default=11.0, gt=0, description="页面宽度（英寸）")
    height: float = Field(default=8.5, gt=0, description="页面高度（英寸）")
    shapes: list[VisioShape] = Field(default_factory=list, description="形状列表")
    connectors: list[VisioConnector] = Field(default_factory=list, description="连接线列表")

    @model_validator(mode="after")
    def validate_connector_refs(self) -> "VisioPage":
        """校验连接线引用的形状 ID 是否存在"""
        shape_ids = {s.id for s in self.shapes}
        for conn in self.connectors:
            if conn.from_shape not in shape_ids:
                raise ValueError(f"连接线 from_shape='{conn.from_shape}' 在 shapes 中不存在")
            if conn.to_shape not in shape_ids:
                raise ValueError(f"连接线 to_shape='{conn.to_shape}' 在 shapes 中不存在")
        return self


class VisioSchema(BaseModel):
    """Visio 文档 Schema"""
    doc_type: str = "visio"
    title: str = Field(description="图表标题")
    pages: list[VisioPage] = Field(default_factory=lambda: [VisioPage()], description="页面列表")
    description: str = Field(default="", description="图表描述")

    @field_validator("title")
    @classmethod
    def validate_title(cls, v):
        if not v or not v.strip():
            raise ValueError("标题不能为空")
        return v.strip()
