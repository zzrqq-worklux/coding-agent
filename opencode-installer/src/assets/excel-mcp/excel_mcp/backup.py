"""写操作前的文件备份。"""
from __future__ import annotations

import shutil
import time
from pathlib import Path


def backup_file(filepath: str | Path, suffix: str = "") -> str | None:
    """Copy filepath to a timestamped .bak next to it. Returns the backup path."""
    src = Path(filepath)
    if not src.exists():
        return None
    stamp = time.strftime("%Y%m%d-%H%M%S")
    tag = f"-{suffix}" if suffix else ""
    dst = src.with_name(f"{src.stem}.bak-{stamp}{tag}{src.suffix}")
    shutil.copy2(src, dst)
    return str(dst)
