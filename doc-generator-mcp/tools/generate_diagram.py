"""图表生成工具 —— Mermaid + Graphviz"""

import subprocess
import tempfile
from pathlib import Path

from schemas.diagram_schema import DiagramSchema


# Shape 类型到 Mermaid 语法的映射
MERMAID_SHAPES = {
    "terminator": ("([", "])"),
    "process": ("[", "]"),
    "decision": ("{", "}"),
    "io": ("[/", "/]"),
}

# 图表类型到 Mermaid 声明的映射
MERMAID_DIAGRAM_TYPES = {
    "flowchart": "flowchart TD",
    "org_chart": "flowchart TD",
    "network": "flowchart LR",
    "mind_map": "mindmap",
    "sequence": "sequenceDiagram",
}


def generate_diagram(
    schema: DiagramSchema,
    engine: str = "mermaid",
    output_format: str = "svg",
    output_path: str = None,
    config: dict = None,
) -> dict:
    """生成图表

    Args:
        schema: 图表 Schema
        engine: 渲染引擎 (mermaid / graphviz)
        output_format: 输出格式 (svg / png)
        output_path: 输出文件路径
        config: 配置字典

    Returns:
        生成结果字典
    """
    config = config or {}
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    if engine == "graphviz":
        return _generate_graphviz(schema, output_format, output)
    else:
        return _generate_mermaid(schema, output_format, output)


def _generate_mermaid(schema: DiagramSchema, output_format: str, output: Path) -> dict:
    """使用 Mermaid 生成图表"""

    # 生成 Mermaid 代码
    mermaid_code = _schema_to_mermaid(schema)

    # 写入临时 .mmd 文件
    mmd_path = output.with_suffix(".mmd")
    mmd_path.write_text(mermaid_code, encoding="utf-8")

    # 尝试使用 mmdc 渲染
    try:
        cmd = [
            "mmdc",
            "-i", str(mmd_path),
            "-o", str(output),
            "-f", output_format,
            "--theme", "default",
        ]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)

        if result.returncode == 0:
            # 清理临时文件
            mmd_path.unlink(missing_ok=True)
            return {
                "success": True,
                "file_path": str(output),
                "engine": "mermaid",
                "format": output_format,
                "node_count": len(schema.nodes),
                "connection_count": len(schema.connections),
                "message": f"Mermaid 图表生成成功: {len(schema.nodes)} 个节点, {len(schema.connections)} 条连接",
            }
        else:
            # mmdc 失败，保存 .mmd 文件作为备选
            return {
                "success": False,
                "file_path": str(mmd_path),
                "engine": "mermaid",
                "format": "mmd",
                "node_count": len(schema.nodes),
                "connection_count": len(schema.connections),
                "error": f"mmdc 渲染失败 (exit code={result.returncode}): {result.stderr[:200]}",
                "mermaid_code": mermaid_code,
                "message": "mmdc 渲染失败，已保存 Mermaid 源文件（可复制到 mermaid.live 在线渲染）",
            }

    except FileNotFoundError:
        # mmdc 未安装，保存 .mmd 文件
        return {
            "success": False,
            "file_path": str(mmd_path),
            "engine": "mermaid",
            "format": "mmd",
            "node_count": len(schema.nodes),
            "connection_count": len(schema.connections),
            "mermaid_code": mermaid_code,
            "error": "mmdc 未安装，请先安装 mermaid-cli: npm install -g @mermaid-js/mermaid-cli",
            "message": "mmdc 未安装，已保存 Mermaid 源文件。可复制到 mermaid.live 在线渲染。",
        }
    except Exception as e:
        # 其他错误，保存 .mmd 文件
        return {
            "success": False,
            "file_path": str(mmd_path),
            "engine": "mermaid",
            "format": "mmd",
            "node_count": len(schema.nodes),
            "connection_count": len(schema.connections),
            "mermaid_code": mermaid_code,
            "error": f"渲染异常: {e}",
            "message": f"渲染异常，已保存 Mermaid 源文件: {e}",
        }


def _generate_graphviz(schema: DiagramSchema, output_format: str, output: Path) -> dict:
    """使用 Graphviz 生成图表"""
    try:
        import graphviz
    except ImportError:
        return {
            "success": False,
            "error": "graphviz 库未安装，请运行: pip install graphviz",
        }

    dot = graphviz.Digraph(
        name=schema.title,
        format=output_format,
        graph_attr={"rankdir": "TB" if schema.layout.direction == "top_to_bottom" else "LR"},
    )

    # 添加节点
    shape_map = {
        "terminator": "oval",
        "process": "box",
        "decision": "diamond",
        "io": "parallelogram",
    }

    for node in schema.nodes:
        dot.node(
            node.id,
            label=node.label,
            shape=shape_map.get(node.type.value, "box"),
        )

    # 添加边
    for conn in schema.connections:
        dot.edge(conn.from_node, conn.to_node, label=conn.label or "")

    # 渲染
    try:
        rendered_path = dot.render(
            filename=str(output.with_suffix("")),
            cleanup=True,
        )
        return {
            "success": True,
            "file_path": rendered_path,
            "engine": "graphviz",
            "format": output_format,
            "node_count": len(schema.nodes),
            "connection_count": len(schema.connections),
            "message": f"Graphviz 图表生成成功: {len(schema.nodes)} 个节点",
        }
    except Exception as e:
        return {
            "success": False,
            "error": f"Graphviz 渲染失败: {e}",
        }


def _schema_to_mermaid(schema: DiagramSchema) -> str:
    """将 Schema 转换为 Mermaid 语法"""
    diagram_decl = MERMAID_DIAGRAM_TYPES.get(schema.diagram_type, "flowchart TD")
    lines = [diagram_decl]

    # 节点定义
    for node in schema.nodes:
        left, right = MERMAID_SHAPES.get(node.type.value, ("[", "]"))
        label = node.label.replace("\n", "<br/>")
        lines.append(f'    {node.id}{left}"{label}"{right}')

    # 连接线
    for conn in schema.connections:
        label = f'|{conn.label}|' if conn.label else ''
        lines.append(f'    {conn.from_node} -->{label} {conn.to_node}')

    return "\n".join(lines)
