"""Registry-based orchestration and atomic output replacement."""

import os
from pathlib import Path
import tempfile

from .legacy import LEGACY, convert
from .models import Cleaner, CleanError, CleanReport
from .ooxml import OOXMLCleaner

DEFAULT_CLEANERS: tuple[Cleaner, ...] = (OOXMLCleaner(),)


def clean_file(path: str | Path, *, cleaners=DEFAULT_CLEANERS) -> CleanReport:
    source = Path(path).expanduser().resolve(strict=True)
    if not source.is_file() or source.name.startswith("~$"):
        raise CleanError("输入必须是有效的 Office 文件，不能是临时锁文件。")
    report = CleanReport(source=str(source), output="")
    with tempfile.TemporaryDirectory(prefix=".office-clean-", dir=source.parent) as temp:
        working = convert(source, Path(temp), report) if source.suffix.lower() in LEGACY else source
        cleaner = next((c for c in cleaners if working.suffix.lower() in c.extensions), None)
        if cleaner is None:
            raise CleanError(f"暂不支持 {source.suffix or '无扩展名'} 格式；支持新版 Word、Excel、PowerPoint，旧版需本机 Office。")
        # Every input is immutable, including files already ending in _clean.
        stem = source.stem + "_clean"
        target = source.with_name(stem + working.suffix.lower())
        report.output = str(target)
        data = cleaner.clean(working, report)
        temporary = None
        try:
            with tempfile.NamedTemporaryFile(dir=target.parent, prefix=".office-clean-", suffix=".tmp", delete=False) as out:
                temporary = Path(out.name)
                out.write(data)
                out.flush()
                os.fsync(out.fileno())
            os.replace(temporary, target)
        finally:
            if temporary is not None:
                temporary.unlink(missing_ok=True)
    return report
