"""
智能备份工具 —— 单备份策略，避免重复生成备份文件

规则:
1. 每个原始文件只保留一份备份，命名为 {原文件名}_原始副本{扩展名}
2. 如果备份已存在，不再重复创建
3. 返回备份路径（无论是否新建）
"""

import shutil
from pathlib import Path


def smart_backup(
    file_path: str,
    backup: bool = True,
    output_path: str = None,
) -> str:
    """智能备份文件

    Args:
        file_path: 原始文件路径
        backup: 是否需要备份
        output_path: 输出路径（如果指定了 output_path 且不覆盖原文件，可选备份）

    Returns:
        备份文件路径（空字符串表示未备份）
    """
    if not backup:
        return ""

    path = Path(file_path)
    if not path.exists():
        return ""

    # 如果输出到新文件且不覆盖原文件，不需要备份原文件
    # （原文件不会被修改）
    if output_path and Path(output_path) != path:
        return ""

    # 备份命名: {原文件名}_原始副本{扩展名}
    backup_path = path.with_name(f"{path.stem}_原始副本{path.suffix}")

    # 如果备份已存在，不再重复创建
    if backup_path.exists():
        return str(backup_path)

    # 创建备份
    shutil.copy2(str(path), str(backup_path))
    return str(backup_path)
