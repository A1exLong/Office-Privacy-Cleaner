"""Independent package rules; register new rules in DEFAULT_RULES."""

from typing import Protocol
from lxml import etree as ET

from .models import CleanReport
from .package import Package


class Rule(Protocol):
    def apply(self, package: Package, report: CleanReport) -> None: ...


def local(node) -> str:
    return ET.QName(node).localname if isinstance(node.tag, str) else ""


class MetadataRule:
    def apply(self, package: Package, report: CleanReport) -> None:
        removed = set()
        metadata_ns = {
            "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
            "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
            "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
        }
        for name in list(package.parts):
            root = package.xml.get(name)
            if name.startswith(("docProps/", "_xmlsignatures/")) or (
                root is not None and ET.QName(root).namespace in metadata_ns
            ):
                removed.add(name)
        for name in removed:
            package.remove(name)
        # Remove relationship sidecars belonging to removed parts too.
        for name in list(package.parts):
            if name.endswith(".rels"):
                from .package import owner_for_rels
                if owner_for_rels(name) in removed:
                    removed.add(name)
                    package.remove(name)
        if removed:
            package.repair_references(removed)
            report.count("元数据、缩略图或签名部件", len(removed))


class CollaborationRule:
    def apply(self, package: Package, report: CleanReport) -> None:
        removed = set()
        # Relationship types identify comment/people parts even at nonstandard paths.
        identity_types = {
            "comments", "commentAuthors", "commentsExtended", "commentsExtensible",
            "commentsIds", "people", "person", "persons", "threadedComment",
            "threadedComments", "authors", "commentAuthor", "revisionLog", "revisionHeaders",
        }
        for name, root in list(package.xml.items()):
            if name.endswith(".rels"):
                from .package import owner_for_rels, resolve_target
                owner = owner_for_rels(name)
                for rel in root:
                    if rel.get("TargetMode") != "External" and rel.get("Type", "").rsplit("/", 1)[-1] in identity_types:
                        removed.add(resolve_target(owner, rel.get("Target", "")))
        for name in package.parts:
            lower = name.lower()
            if lower.startswith(("word/comments", "word/people", "ppt/comments/", "ppt/commentauthors", "ppt/authors/", "ppt/persons/", "xl/comments", "xl/threadedcomments/", "xl/persons/", "xl/revisions/")):
                removed.add(name)
        for name in list(package.parts):
            if name.endswith(".rels"):
                from .package import owner_for_rels
                if owner_for_rels(name) in removed:
                    removed.add(name)
        for name in removed:
            package.remove(name)
        if removed:
            package.repair_references(removed)
            report.count("批注、作者或协作历史部件", len(removed))
        for name, root in list(package.xml.items()):
            for element in list(root.iter()):
                tag = local(element)
                namespace = ET.QName(element).namespace if isinstance(element.tag, str) else ""
                is_word = namespace in {"http://schemas.openxmlformats.org/wordprocessingml/2006/main", "http://purl.oclc.org/ooxml/wordprocessingml/main"}
                if is_word and tag in {"commentRangeStart", "commentRangeEnd", "commentReference", "rsids"}:
                    parent = element.getparent()
                    if parent is not None:
                        parent.remove(element)
                        report.count("Word 批注标记与编辑标识")
                    continue
                if is_word:
                    for attr in list(element.attrib):
                        key = ET.QName(attr).localname
                        if key.startswith("rsid") or key in {"date", "dateUtc"}:
                            del element.attrib[attr]
                            report.count("Word 编辑者时间与会话标识")
                        elif key in {"author", "initials"} and element.get(attr) != "Anonymous":
                            element.set(attr, "Anonymous")
                            report.count("Word 修订者匿名化")
                    if tag in {"ins", "del", "moveFrom", "moveTo"} or tag.endswith("PrChange"):
                        report.warn("Word 修订内容已保留，作者身份已匿名化；历史删除文字可能仍含正文个人信息。")
                if name.startswith("xl/"):
                    if tag == "fileSharing" and "userName" in element.attrib:
                        del element.attrib["userName"]
                        report.count("Excel 共享用户名")
                    if tag == "customWorkbookViews":
                        parent = element.getparent()
                        if parent is not None:
                            parent.remove(element)
                            report.count("Excel 个人视图")
                if name.endswith(".vml") and tag == "shape":
                    note = any(local(c) == "ClientData" and c.get("ObjectType") == "Note" for c in element.iter())
                    if note and element.getparent() is not None:
                        element.getparent().remove(element)
                        report.count("Excel 批注形状")


class RiskInventoryRule:
    def apply(self, package: Package, report: CleanReport) -> None:
        for name in package.parts:
            lower = name.lower()
            if lower.endswith("vbaproject.bin"):
                report.warn("宏已保留，宏代码及二进制内容未匿名化。")
            if "/embeddings/" in lower or lower.endswith(".bin") and "vba" not in lower:
                report.warn("存在嵌入对象或二进制部件，内部内容未递归清理。")
            if lower.startswith("customxml/"):
                report.warn("自定义 XML 数据已保留，可能包含业务或个人信息。")
            if "/media/" in lower:
                report.warn("媒体内容已保留；图片 EXIF、截图文字、音视频内容未匿名化。")
            if lower.startswith("ppt/notesslides/notesSlide".lower()) and lower.endswith(".xml"):
                report.warn("演讲者备注已保留，其中的姓名等正文信息未匿名化。")
        for name, root in package.xml.items():
            if name.endswith(".rels") and any(r.get("TargetMode") == "External" for r in root):
                report.warn("外部链接已保留以维持数据关系；地址可能含用户名或本地路径。")
        report.warn("正文、单元格、公式、幻灯片、页眉页脚、文件名及文件系统信息不做个人信息识别。")


DEFAULT_RULES: tuple[Rule, ...] = (MetadataRule(), CollaborationRule(), RiskInventoryRule())
