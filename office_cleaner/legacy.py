"""Isolated, optional Microsoft Office conversion for old binary formats."""

from pathlib import Path
import sys

from .models import CleanError, CleanReport

LEGACY = {".doc", ".dot", ".xls", ".xlt", ".ppt", ".pps", ".pot"}


def convert(source: Path, directory: Path, report: CleanReport) -> Path:
    if sys.platform != "win32":
        raise CleanError("旧版 Office 文件需要 Windows 和本机 Microsoft Office；请先另存为新版格式。")
    try:
        import pythoncom
        from win32com.client import DispatchEx
    except ImportError as exc:
        raise CleanError("旧版格式需要 pywin32 和本机 Microsoft Office。") from exc
    kind = "word" if source.suffix.lower() in {".doc", ".dot"} else "excel" if source.suffix.lower() in {".xls", ".xlt"} else "ppt"
    app = document = None
    pythoncom.CoInitialize()
    try:
        app = DispatchEx({"word": "Word.Application", "excel": "Excel.Application", "ppt": "PowerPoint.Application"}[kind])
        app.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
        if kind == "word":
            app.Visible = False
            app.DisplayAlerts = 0
            document = app.Documents.Open(str(source), ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False, PasswordDocument="", WritePasswordDocument="")
            macro = bool(document.HasVBProject)
            target = directory / ("converted.docm" if macro else "converted.docx")
            document.SaveAs2(str(target), FileFormat=13 if macro else 12, AddToRecentFiles=False)
        elif kind == "excel":
            app.Visible = False
            app.DisplayAlerts = False
            app.AskToUpdateLinks = False
            app.EnableEvents = False
            document = app.Workbooks.Open(str(source), UpdateLinks=0, ReadOnly=True, Password="", WriteResPassword="", IgnoreReadOnlyRecommended=True, AddToMru=False)
            macro = bool(document.HasVBProject)
            target = directory / ("converted.xlsm" if macro else "converted.xlsx")
            document.SaveAs(str(target), FileFormat=52 if macro else 51, AddToMru=False)
        else:
            document = app.Presentations.Open(str(source), ReadOnly=True, Untitled=False, WithWindow=False)
            macro = bool(document.HasVBProject)
            target = directory / ("converted.pptm" if macro else "converted.pptx")
            document.SaveAs(str(target), FileFormat=25 if macro else 24)
        report.warn("旧版文件经 Microsoft Office 转为新版格式；请核对转换后的版式、公式与功能。")
        return target
    except Exception as exc:
        raise CleanError("旧版文件转换失败：请确认安装了对应 Microsoft Office，文件未加密且可正常打开。") from exc
    finally:
        if document is not None:
            try:
                document.Close(False) if kind == "word" else document.Close(SaveChanges=False) if kind == "excel" else document.Close()
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
