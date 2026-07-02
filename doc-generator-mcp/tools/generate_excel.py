"""Excel 文件生成工具"""

import logging
from pathlib import Path
import xlsxwriter

from schemas.excel_schema import ExcelSchema

logger = logging.getLogger(__name__)


def generate_excel(
    schema: ExcelSchema,
    template: str | None = None,
    output_path: str | None = None,
    config: dict | None = None,
) -> dict:
    """生成 Excel 文件

    Args:
        schema: Excel Schema
        template: 模板名称（可选，当前未使用）
        output_path: 输出文件路径
        config: 配置字典

    Returns:
        生成结果字典
    """
    config = config or {}
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    data = schema.data
    sheet_name = schema.sheet_name

    workbook = xlsxwriter.Workbook(str(output))
    sheet = workbook.add_worksheet(sheet_name)

    # ===== 样式定义 =====
    header_fmt = workbook.add_format({
        'bold': True,
        'border': 1,
        'bg_color': '#1e40af',
        'font_color': 'white',
        'align': 'center',
        'valign': 'vcenter',
    })
    number_fmt = workbook.add_format({
        'num_format': schema.formatting.number_format,
        'border': 1,
    })
    text_fmt = workbook.add_format({
        'border': 1,
        'text_wrap': True,
    })

    # ===== 写入表头 =====
    headers = data.headers
    # 设置列宽
    for col, h in enumerate(headers):
        sheet.write(0, col, h, header_fmt)
        # 自动列宽：取表头和数据的最大长度
        max_len = len(str(h))
        for row_data in data.rows:
            if col < len(row_data):
                max_len = max(max_len, len(str(row_data[col])))
        sheet.set_column(col, col, max(max_len + 4, 10))

    # ===== 写入数据 =====
    for row_idx, row in enumerate(data.rows, start=1):
        for col_idx, val in enumerate(row):
            fmt = number_fmt if isinstance(val, (int, float)) else text_fmt
            sheet.write(row_idx, col_idx, val, fmt)

    # ===== 写入公式 =====
    formula_errors = []
    for f in schema.formulas:
        try:
            sheet.write_formula(f.cell, f.formula)
        except Exception as e:
            msg = f"公式写入失败: cell={f.cell}, formula={f.formula}, error={e}"
            logger.warning(msg)
            formula_errors.append(msg)

    # ===== 写入条件格式 =====
    for rule in schema.formatting.conditional_rules:
        if rule.type == "data_bar":
            sheet.conditional_format(
                rule.range,
                {'type': 'data_bar', 'bar_color': rule.color}
            )
        elif rule.type == "color_scale":
            sheet.conditional_format(
                rule.range,
                {'type': '3_color_scale'}
            )

    # ===== 写入图表 =====
    for chart_def in schema.charts:
        chart = workbook.add_chart({"type": chart_def.type})

        # 解析数据范围
        data_range = chart_def.data_range
        if ":" in data_range:
            # 格式如 A1:B10
            start, end = data_range.split(":")
            start_col = _col_letter_to_num(start)
            end_col = _col_letter_to_num(end)
            last_row = int("".join(c for c in end if c.isdigit()))

            chart.add_series({
                "name": chart_def.title,
                "categories": [sheet_name, 1, start_col, last_row, start_col],
                "values": [sheet_name, 1, start_col + 1, last_row, start_col + 1],
            })
        else:
            # 简单情况：使用全部数据
            last_row = len(data.rows) + 1
            chart.add_series({
                "name": chart_def.title,
                "categories": [sheet_name, 1, 0, last_row, 0],
                "values": [sheet_name, 1, 1, last_row, 1],
            })

        chart.set_title({"name": chart_def.title})
        sheet.insert_chart(chart_def.position, chart)

    workbook.close()

    result = {
        "success": True,
        "file_path": str(output),
        "sheet_name": sheet_name,
        "row_count": len(data.rows),
        "column_count": len(headers),
        "formula_count": len(schema.formulas),
        "chart_count": len(schema.charts),
        "message": f"Excel 文件生成成功: {len(data.rows)} 行数据, {len(schema.formulas)} 个公式, {len(schema.charts)} 个图表",
    }
    if formula_errors:
        result["warning"] = f"{len(formula_errors)} 个公式写入失败"
        result["formula_errors"] = formula_errors
    return result


def _col_letter_to_num(ref: str) -> int:
    """将列字母转为数字索引 (A=0, B=1, ...)"""
    letters = "".join(c for c in ref if c.isalpha()).upper()
    result = 0
    for c in letters:
        result = result * 26 + (ord(c) - ord("A"))
    return result
