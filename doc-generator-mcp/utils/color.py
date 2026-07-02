"""颜色工具函数"""


def hex_to_rgb(hex_color: str) -> tuple[int, int, int]:
    """HEX 颜色转 RGB 元组

    Args:
        hex_color: 十六进制颜色字符串，如 "#4472C4" 或 "4472C4"

    Returns:
        (R, G, B) 元组，取值范围 0-255
    """
    hex_color = hex_color.lstrip("#")
    if len(hex_color) != 6:
        return (0, 0, 0)
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return (r, g, b)
