"""OPC package I/O, secure XML parsing and relationship consistency."""

from io import BytesIO
from pathlib import PurePosixPath
import posixpath
from zipfile import BadZipFile, ZipFile, ZipInfo, ZIP_DEFLATED

from lxml import etree as ET

from .models import CleanError

REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
R_NAMESPACES = {
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "http://purl.oclc.org/ooxml/officeDocument/relationships",
}


def parse_xml(data: bytes, name: str):
    try:
        root = ET.fromstring(data, parser=ET.XMLParser(
            resolve_entities=False, load_dtd=False, no_network=True,
            huge_tree=False, remove_blank_text=False, remove_comments=True, remove_pis=True,
        ))
        if root.getroottree().docinfo.doctype:
            raise CleanError(f"不支持含 DTD 的 XML：{name}")
        return root
    except ET.XMLSyntaxError as exc:
        raise CleanError(f"XML 损坏：{name}") from exc


def serialize(root) -> bytes:
    return ET.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)


def owner_for_rels(name: str) -> str:
    if name == "_rels/.rels":
        return ""
    path = PurePosixPath(name)
    return str(path.parent.parent / path.name[:-5])


def resolve_target(owner: str, target: str) -> str:
    # OPC targets are URI paths, not Windows paths.
    from urllib.parse import unquote, urlsplit
    path = unquote(urlsplit(target).path)
    return posixpath.normpath(posixpath.join(posixpath.dirname(owner), path)).lstrip("/")


class Package:
    MAX_INPUT = 256 * 1024 * 1024
    MAX_TOTAL = 1024 * 1024 * 1024
    MAX_PART = 128 * 1024 * 1024
    MAX_PARTS = 20000

    def __init__(self, data: bytes):
        self.parts: dict[str, bytes] = {}
        self.xml: dict[str, object] = {}
        if len(data) > self.MAX_INPUT:
            raise CleanError("文件超过 256 MiB 的处理上限。")
        try:
            with ZipFile(BytesIO(data)) as archive:
                infos = archive.infolist()
                if len(infos) > self.MAX_PARTS:
                    raise CleanError("文档包含过多部件。")
                total = 0
                seen = set()
                for info in infos:
                    name = info.filename
                    path = PurePosixPath(name)
                    if name.lower() in seen or ".." in path.parts or name.startswith("/") or "\\" in name:
                        raise CleanError("文档包含重复或非法 ZIP 路径。")
                    seen.add(name.lower())
                    if info.flag_bits & 1:
                        raise CleanError("不支持加密的 ZIP 文档。")
                    if info.is_dir():
                        continue
                    total += info.file_size
                    if info.file_size > self.MAX_PART or total > self.MAX_TOTAL:
                        raise CleanError("文档解压大小超过处理上限。")
                    if info.file_size > max(info.compress_size, 1) * 1000:
                        raise CleanError("文档压缩比例异常。")
                    self.parts[name] = archive.read(info)
        except (BadZipFile, RuntimeError, OSError, NotImplementedError) as exc:
            raise CleanError("无法读取文档包，文件可能损坏或已加密。") from exc
        if "[Content_Types].xml" not in self.parts or "_rels/.rels" not in self.parts:
            raise CleanError("不是完整的 Office Open XML 文档。")
        for name, data in self.parts.items():
            if name.endswith((".xml", ".rels", ".vml")):
                self.xml[name] = parse_xml(data, name)

    def remove(self, name: str) -> None:
        self.parts.pop(name, None)
        self.xml.pop(name, None)

    def repair_references(self, removed: set[str]) -> None:
        for name, root in list(self.xml.items()):
            if not name.endswith(".rels"):
                continue
            owner = owner_for_rels(name)
            removed_ids = set()
            for relationship in list(root):
                if relationship.get("TargetMode") == "External":
                    continue
                target = resolve_target(owner, relationship.get("Target", ""))
                if target in removed:
                    removed_ids.add(relationship.get("Id"))
                    root.remove(relationship)
            if removed_ids and owner in self.xml:
                for element in list(self.xml[owner].iter()):
                    for attr, value in list(element.attrib.items()):
                        q = ET.QName(attr)
                        if q.namespace in R_NAMESPACES and value in removed_ids:
                            parent = element.getparent()
                            if parent is not None:
                                parent.remove(element)
                            break
        types = self.xml["[Content_Types].xml"]
        for element in list(types):
            if element.get("PartName", "").lstrip("/") in removed:
                types.remove(element)

    def validate(self) -> None:
        for name, root in self.xml.items():
            if not name.endswith(".rels"):
                continue
            owner = owner_for_rels(name)
            if owner and owner not in self.parts:
                raise CleanError(f"关系所属部件不存在：{name}")
            for relation in root:
                if relation.get("TargetMode") != "External":
                    target = resolve_target(owner, relation.get("Target", ""))
                    if target not in self.parts:
                        raise CleanError(f"存在无效的内部文档关系：{name}")

    def dump(self) -> bytes:
        self.validate()
        output = BytesIO()
        with ZipFile(output, "w", compression=ZIP_DEFLATED) as archive:
            # Recreate every ZIP header to discard original timestamps, comments and extra fields.
            for name in self.parts:
                info = ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
                info.compress_type = ZIP_DEFLATED
                info.external_attr = 0o600 << 16
                archive.writestr(info, serialize(self.xml[name]) if name in self.xml else self.parts[name])
        return output.getvalue()
