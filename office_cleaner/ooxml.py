from pathlib import Path

from .models import CleanError, CleanReport
from .package import Package, parse_xml
from .rules import DEFAULT_RULES


class OOXMLCleaner:
    extensions = frozenset({".docx", ".docm", ".dotx", ".dotm", ".xlsx", ".xlsm", ".xltx", ".xltm", ".pptx", ".pptm", ".potx", ".potm", ".ppsx", ".ppsm"})

    def __init__(self, rules=DEFAULT_RULES):
        self.rules = tuple(rules)

    def clean(self, source: Path, report: CleanReport) -> bytes:
        if source.stat().st_size > Package.MAX_INPUT:
            raise CleanError("文件超过 256 MiB 的处理上限。")
        package = Package(source.read_bytes())
        expected = "word/document.xml" if source.suffix.lower().startswith((".do",)) else "xl/workbook.xml" if source.suffix.lower().startswith(".xl") else "ppt/presentation.xml"
        if expected not in package.parts:
            raise CleanError("文件扩展名与实际文档类型不一致。")
        package.validate()
        for rule in self.rules:
            rule.apply(package, report)
        data = package.dump()
        # Reopen the result; CRC, XML and relationship checks run before replacement.
        verified = Package(data)
        verified.validate()
        return data
