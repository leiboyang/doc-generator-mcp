"""图表的 Pydantic Schema"""

from pydantic import BaseModel, Field, field_validator, ConfigDict
from typing import Literal
from enum import Enum


class NodeType(str, Enum):
    TERMINATOR = "terminator"
    PROCESS = "process"
    DECISION = "decision"
    IO = "io"


class DiagramNode(BaseModel):
    """图表节点"""
    id: str = Field(
        ..., min_length=1, max_length=50,
        pattern=r"^[a-zA-Z0-9_-]+$",
        description="节点唯一标识"
    )
    label: str = Field(..., min_length=1, max_length=200, description="节点文本")
    type: NodeType = Field(default=NodeType.PROCESS, description="节点形状类型")

    @field_validator("label")
    @classmethod
    def label_not_empty(cls, v: str) -> str:
        if not v.strip():
            raise ValueError("节点标签不能为空白")
        return v.strip()


class DiagramConnection(BaseModel):
    """节点连接"""
    from_node: str = Field(..., alias="from", description="起始节点 ID")
    to_node: str = Field(..., alias="to", description="目标节点 ID")
    label: str | None = Field(None, max_length=50, description="连接线标签")


class LayoutConfig(BaseModel):
    """布局配置"""
    direction: Literal["top_to_bottom", "left_to_right"] = "top_to_bottom"
    spacing_x: float = Field(default=2.5, gt=0.5, lt=10.0)
    spacing_y: float = Field(default=1.5, gt=0.5, lt=10.0)


class DiagramSchema(BaseModel):
    """图表的完整 Schema"""
    doc_type: Literal["diagram"] = "diagram"
    diagram_type: Literal["flowchart", "org_chart", "network", "mind_map", "sequence"] = "flowchart"
    title: str = Field(..., min_length=1, max_length=200)
    nodes: list[DiagramNode] = Field(..., min_length=1, max_length=100)
    connections: list[DiagramConnection] = Field(default_factory=list)
    layout: LayoutConfig = Field(default_factory=LayoutConfig)

    @field_validator("connections")
    @classmethod
    def validate_connections(cls, v, info):
        """验证所有连接引用的节点 ID 都存在"""
        if "nodes" not in info.data:
            return v
        node_ids = {n.id for n in info.data["nodes"]}
        for conn in v:
            if conn.from_node not in node_ids:
                raise ValueError(f"连接引用了不存在的节点: {conn.from_node}")
            if conn.to_node not in node_ids:
                raise ValueError(f"连接引用了不存在的节点: {conn.to_node}")
        return v

    model_config = ConfigDict(populate_by_name=True)
