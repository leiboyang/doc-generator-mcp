"""Visio 编辑工具 —— 读取和编辑已有 .vsdx 文件"""

import os
import shutil
import logging
from pathlib import Path
from utils.backup import smart_backup
from utils.color import hex_to_rgb

logger = logging.getLogger(__name__)


def read_visio(file_path: str, config: dict = None) -> dict:
    """读取 Visio 文件结构

    Args:
        file_path: .vsdx 文件路径
        config: 配置

    Returns:
        文档结构信息
    """
    if not Path(file_path).exists():
        return {"success": False, "error": f"文件不存在: {file_path}"}

    try:
        import win32com.client
    except ImportError:
        return {"success": False, "error": "win32com 不可用"}

    visio_app = None
    try:
        visio_app = win32com.client.Dispatch("Visio.Application")
        visio_app.Visible = False

        doc = visio_app.Documents.Open(os.path.abspath(file_path))
        result = {
            "success": True,
            "file_path": file_path,
            "file_type": "visio",
            "page_count": doc.Pages.Count,
            "pages": [],
        }

        for page_idx in range(1, doc.Pages.Count + 1):
            page = doc.Pages(page_idx)
            page_info = {
                "name": page.Name,
                "shape_count": page.Shapes.Count,
                "shapes": [],
            }

            for shape_idx in range(1, page.Shapes.Count + 1):
                shape = page.Shapes(shape_idx)
                shape_info = {
                    "index": shape_idx,
                    "id": shape.ID,
                    "name": shape.Name,
                    "text": shape.Text.strip() if shape.Text else "",
                    "type": shape.Type,
                    "x": round(shape.CellsU("PinX").ResultIU, 2),
                    "y": round(shape.CellsU("PinY").ResultIU, 2),
                    "width": round(shape.CellsU("Width").ResultIU, 2),
                    "height": round(shape.CellsU("Height").ResultIU, 2),
                }

                # 检查是否有连接线
                try:
                    connects = shape.Connects
                    if connects and connects.Count > 0:
                        conn_info = []
                        for c in range(1, connects.Count + 1):
                            conn = connects.Item(c)
                            conn_info.append({
                                "from_sheet": conn.FromSheet.ID if conn.FromSheet else None,
                                "to_sheet": conn.ToSheet.ID if conn.ToSheet else None,
                            })
                        if conn_info:
                            shape_info["connects"] = conn_info
                except:
                    pass

                page_info["shapes"].append(shape_info)

            result["pages"].append(page_info)

        doc.Close()
        return result

    except Exception as e:
        return {"success": False, "error": f"读取 Visio 失败: {e}"}

    finally:
        if visio_app:
            try:
                visio_app.Quit()
            except:
                pass


def edit_visio(
    file_path: str,
    edits: list,
    backup: bool = True,
    output_path: str = None,
    config: dict = None,
) -> dict:
    """编辑 Visio 文件

    支持的编辑操作：
    - update_shape_text: 修改形状文本 {action, shape_id/shape_index, text}
    - update_shape_style: 修改形状样式 {action, shape_id, fill_color, line_color, font_size}
    - add_shape: 添加形状 {action, label, shape_type, x, y, width, height}
    - add_connector: 添加连接线 {action, from_id, to_id, label}
    - delete_shape: 删除形状 {action, shape_id}

    Args:
        file_path: .vsdx 文件路径
        edits: 编辑指令列表
        backup: 是否备份
        output_path: 输出路径
        config: 配置

    Returns:
        编辑结果
    """
    if not Path(file_path).exists():
        return {"success": False, "error": f"文件不存在: {file_path}"}

    try:
        import win32com.client
    except ImportError:
        return {"success": False, "error": "win32com 不可用"}

    # 备份
    backup_path = smart_backup(file_path, backup, output_path)

    target_path = output_path or file_path
    if output_path:
        shutil.copy2(file_path, output_path)

    visio_app = None
    changes = []

    try:
        visio_app = win32com.client.Dispatch("Visio.Application")
        visio_app.Visible = False

        doc = visio_app.Documents.Open(os.path.abspath(target_path))
        page = doc.Pages(1)  # 默认编辑第一页

        for edit in edits:
            action = edit.get("action", "")

            if action == "update_shape_text":
                shape = _find_shape(page, edit)
                if shape:
                    old_text = shape.Text
                    shape.Text = edit.get("text", "")
                    changes.append(f"形状 {edit.get('shape_id', edit.get('shape_index'))} 文本: '{old_text}' → '{edit.get('text', '')}'")
                else:
                    changes.append(f"[失败] 未找到形状: {edit.get('shape_id', edit.get('shape_index'))}")

            elif action == "update_shape_style":
                shape = _find_shape(page, edit)
                if shape:
                    if "fill_color" in edit:
                        r, g, b = hex_to_rgb(edit["fill_color"])
                        shape.CellsU("FillForegnd").FormulaU = f"RGB({r},{g},{b})"
                    if "line_color" in edit:
                        r, g, b = hex_to_rgb(edit["line_color"])
                        shape.CellsU("LineColor").FormulaU = f"RGB({r},{g},{b})"
                    if "font_size" in edit:
                        shape.CellsU("Char.Size").FormulaU = f"{edit['font_size']}pt"
                    changes.append(f"形状 {edit.get('shape_id')} 样式已更新")

            elif action == "add_shape":
                vis_shape = page.DrawRectangle(
                    edit.get("x", 0), edit.get("y", 0),
                    edit.get("x", 0) + edit.get("width", 2),
                    edit.get("y", 0) + edit.get("height", 1)
                )
                vis_shape.Text = edit.get("label", "")
                changes.append(f"添加形状: '{edit.get('label', '')}'")

            elif action == "add_connector":
                from_id = edit.get("from_id")
                to_id = edit.get("to_id")
                from_shp = None
                to_shp = None
                for i in range(1, page.Shapes.Count + 1):
                    s = page.Shapes(i)
                    if s.ID == from_id:
                        from_shp = s
                    if s.ID == to_id:
                        to_shp = s
                if from_shp and to_shp:
                    # 记录当前形状数量，以便后续识别新创建的连接线
                    shape_count_before = page.Shapes.Count
                    from_shp.AutoConnect(to_shp, 0)
                    label = edit.get("label", "")
                    if label:
                        # 仅处理新创建的连接线（从 shape_count_before+1 开始）
                        for i in range(shape_count_before + 1, page.Shapes.Count + 1):
                            s = page.Shapes(i)
                            if s.OneD:
                                s.Text = label
                                s.CellsU("Char.Size").FormulaU = "9pt"
                                break
                    changes.append(f"添加连接线: {from_id} → {to_id}")
                else:
                    changes.append(f"[失败] 未找到形状: from={from_id}, to={to_id}")

            elif action == "delete_shape":
                shape = _find_shape(page, edit)
                if shape:
                    shape.Delete()
                    changes.append(f"删除形状: {edit.get('shape_id')}")

        doc.Save()
        doc.Close()

        return {
            "success": True,
            "file_path": target_path,
            "backup_path": backup_path,
            "changes_applied": len(changes),
            "changes": changes,
        }

    except Exception as e:
        return {"success": False, "error": f"编辑 Visio 失败: {e}", "changes": changes}

    finally:
        if visio_app:
            try:
                visio_app.Quit()
            except:
                pass


def _find_shape(page, edit: dict):
    """根据 edit 中的 shape_id 或 shape_index 查找形状"""
    if "shape_id" in edit:
        target_id = edit["shape_id"]
        for i in range(1, page.Shapes.Count + 1):
            shape = page.Shapes(i)
            if shape.ID == target_id:
                return shape
    elif "shape_index" in edit:
        idx = edit["shape_index"]
        if 1 <= idx <= page.Shapes.Count:
            return page.Shapes(idx)
    elif "shape_name" in edit:
        name = edit["shape_name"]
        for i in range(1, page.Shapes.Count + 1):
            shape = page.Shapes(i)
            if shape.Name == name:
                return shape
    return None
