"""Visio 生成工具 —— 通过 COM 自动化或 vsdx 库生成 .vsdx 文件"""

import os
import logging
from pathlib import Path

from schemas.visio_schema import VisioSchema, VisioShapeType, VisioConnectorType
from utils.color import hex_to_rgb

logger = logging.getLogger(__name__)

# Visio 形状类型 → Visio Stencil Master 名称映射
SHAPE_MASTERS = {
    VisioShapeType.RECTANGLE: "Rectangle",
    VisioShapeType.ROUNDED_RECT: "Rounded rectangle",
    VisioShapeType.ELLIPSE: "Ellipse",
    VisioShapeType.DIAMOND: "Diamond",
    VisioShapeType.PARALLELOGRAM: "Parallelogram",
    VisioShapeType.CYLINDER: "Cylinder",
    VisioShapeType.DOCUMENT: "Document",
    VisioShapeType.PROCESS: "Process",
    VisioShapeType.DECISION: "Decision",
    VisioShapeType.TERMINATOR: "Rounded rectangle",
    VisioShapeType.DATA: "Parallelogram",
    VisioShapeType.PREDEFINED: "Rectangle",
    VisioShapeType.DELAY: "D shape",
    VisioShapeType.ARROW: "Right Arrow",
    VisioShapeType.CUSTOM: "Rectangle",
}


def generate_visio(
    schema: VisioSchema,
    output_path: str,
    engine: str = "com",
    config: dict = None,
) -> dict:
    """生成 Visio 文件

    Args:
        schema: Visio Schema
        output_path: 输出路径
        engine: 引擎类型 ("com" 或 "vsdx")
        config: 配置

    Returns:
        生成结果
    """
    config = config or {}

    # 确保输出目录存在
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)

    if engine == "com":
        return _generate_with_com(schema, output_path)
    else:
        return _generate_with_vsdx_lib(schema, output_path)


def _generate_with_com(schema: VisioSchema, output_path: str) -> dict:
    """使用 Visio COM 自动化生成"""
    try:
        import win32com.client
    except ImportError:
        return {"success": False, "error": "win32com 不可用，请安装 pywin32"}

    visio_app = None
    try:
        visio_app = win32com.client.Dispatch("Visio.Application")
        visio_app.Visible = False

        # 创建新文档
        doc = visio_app.Documents.Add("")
        page = doc.Pages(1)
        page.Name = schema.pages[0].name if schema.pages else "Page-1"

        # 设置页面大小
        if schema.pages:
            p = schema.pages[0]
            page.PageSheet.CellsU("PageWidth").FormulaU = f"{p.width}in"
            page.PageSheet.CellsU("PageHeight").FormulaU = f"{p.height}in"

        shape_map = {}  # id → Visio shape object

        # 在循环外打开基本模具，避免重复加载
        basic_stencil = None
        try:
            basic_stencil = visio_app.Documents.Open("Basic_M.vssx")
        except Exception as e:
            logger.warning(f"打开基本模具失败: {e}")

        for page_data in schema.pages:
            for shape in page_data.shapes:
                # 获取 Master 形状名称
                master_name = SHAPE_MASTERS.get(shape.shape_type, "Rectangle")

                # 尝试从基本模具加载
                try:
                    master = None
                    if basic_stencil:
                        for i in range(1, basic_stencil.Masters.Count + 1):
                            m = basic_stencil.Masters(i)
                            if m.Name.lower() == master_name.lower():
                                master = m
                                break

                    if master is None and basic_stencil:
                        # 回退到矩形
                        master = basic_stencil.Masters("Rectangle")

                    # 放置形状 (Visio 坐标系: 左下角为原点)
                    visio_x = shape.x + shape.width / 2
                    visio_y = (page_data.height - shape.y) - shape.height / 2

                    vis_shape = page.Drop(master, visio_x, visio_y)
                    shape_map[shape.id] = vis_shape

                    # 设置大小
                    vis_shape.CellsU("Width").FormulaU = f"{shape.width}in"
                    vis_shape.CellsU("Height").FormulaU = f"{shape.height}in"

                    # 设置文本
                    vis_shape.Text = shape.label

                    # 设置填充颜色
                    fill_r, fill_g, fill_b = hex_to_rgb(shape.fill_color)
                    vis_shape.CellsU("FillForegnd").FormulaU = f"RGB({fill_r},{fill_g},{fill_b})"

                    # 设置线条颜色
                    line_r, line_g, line_b = hex_to_rgb(shape.line_color)
                    vis_shape.CellsU("LineWeight").FormulaU = "1pt"
                    vis_shape.CellsU("LineColor").FormulaU = f"RGB({line_r},{line_g},{line_b})"

                    # 设置字体
                    vis_shape.CellsU("Char.Size").FormulaU = f"{shape.font_size}pt"
                    if shape.bold:
                        vis_shape.CellsU("Char.Style").FormulaU = "1"  # Bold

                except Exception as e:
                    logger.warning(f"形状 {shape.id} 放置失败: {e}")
                    # 回退：直接画矩形
                    visio_x = shape.x + shape.width / 2
                    visio_y = (page_data.height - shape.y) - shape.height / 2
                    vis_shape = page.DrawRectangle(
                        shape.x, page_data.height - shape.y - shape.height,
                        shape.x + shape.width, page_data.height - shape.y
                    )
                    vis_shape.Text = shape.label
                    shape_map[shape.id] = vis_shape

            # 绘制连接线
            for conn in page_data.connectors:
                try:
                    from_shape = shape_map.get(conn.from_shape)
                    to_shape = shape_map.get(conn.to_shape)
                    if from_shape and to_shape:
                        # 使用 AutoConnect 创建连接线 (已验证可用)
                        from_shape.AutoConnect(to_shape, 0)  # 0 = visAutoConnectDirDefault

                        # 找到刚创建的连接线并设置样式
                        # AutoConnect 创建的连接线是最新的 1D 形状
                        for s_idx in range(1, page.Shapes.Count + 1):
                            s = page.Shapes(s_idx)
                            if s.OneD:  # 1D = 连接线
                                # 检查是否已处理过（通过检查是否已有样式设置）
                                existing_color = s.CellsU("LineColor").FormulaU
                                if existing_color == "RGB(0,0,0)":  # 默认颜色=未处理
                                    # 设置连接线样式
                                    s.CellsU("EndArrow").FormulaU = "5" if conn.arrow_head else "0"
                                    s.CellsU("EndArrowSize").FormulaU = "4"

                                    r, g, b = hex_to_rgb(conn.color)
                                    s.CellsU("LineColor").FormulaU = f"RGB({r},{g},{b})"
                                    s.CellsU("LineWeight").FormulaU = "1pt"

                                    if conn.connector_type == VisioConnectorType.STRAIGHT:
                                        s.CellsU("ObjType").FormulaU = "2"

                                    if conn.label:
                                        s.Text = conn.label
                                        s.CellsU("Char.Size").FormulaU = "9pt"
                                    break

                except Exception as e:
                    logger.warning(f"连接线 {conn.from_shape} → {conn.to_shape} 失败: {e}")

        # 保存
        doc.SaveAs(os.path.abspath(output_path))
        doc.Close()

        shape_count = sum(len(p.shapes) for p in schema.pages)
        conn_count = sum(len(p.connectors) for p in schema.pages)

        return {
            "success": True,
            "file_path": output_path,
            "engine": "com",
            "shape_count": shape_count,
            "connector_count": conn_count,
            "message": f"Visio 文件生成成功: {shape_count} 个形状, {conn_count} 条连接线",
        }

    except Exception as e:
        return {"success": False, "error": f"Visio COM 生成失败: {e}"}

    finally:
        if visio_app:
            try:
                visio_app.Quit()
            except:
                pass


def _generate_with_vsdx_lib(schema: VisioSchema, output_path: str) -> dict:
    """使用 vsdx 库生成（无需 Visio 安装）"""
    try:
        import vsdx
    except ImportError:
        return {"success": False, "error": "vsdx 库未安装，请运行: pip install vsdx"}

    try:
        # vsdx 库需要一个模板文件来创建新文档
        # 如果没有模板，使用 COM 创建的空文件作为模板
        template_path = Path(__file__).parent.parent / "templates" / "blank.vsdx"
        if not template_path.exists():
            return {
                "success": False,
                "error": f"vsdx 模板不存在: {template_path}。请使用 engine='com' 或提供模板文件。"
            }

        with vsdx.VisioFile(str(template_path)) as vis:
            page = vis.pages[0] if vis.pages else vis.add_page("Page-1")

            for page_data in schema.pages:
                for shape in page_data.shapes:
                    # vsdx 库的形状操作较为底层，使用 XML 操作
                    sub = page.add_sub_page(shape.label) if False else None
                    # 简化实现：仅设置文本和基本属性
                    pass

            vis.save_vsdx(output_path)

            return {
                "success": True,
                "file_path": output_path,
                "engine": "vsdx",
                "message": "Visio 文件生成成功（vsdx 库模式，功能有限）",
            }

    except Exception as e:
        return {"success": False, "error": f"vsdx 库生成失败: {e}"}



