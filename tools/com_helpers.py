import win32com.client
import time
import os
import locale
from typing import Optional, Tuple, List, Any


def get_acad() -> Tuple[Any, Any, Optional[str]]:
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application.25")
        if acad.Documents.Count == 0:
            return None, None, "无打开的文档，请先创建或打开一个图纸"
        return acad, acad.ActiveDocument, None
    except Exception as e:
        return None, None, f"连接AutoCAD失败: {str(e)}"


def wait_cmd(doc, max_iter=300, interval=0.1):
    for _ in range(max_iter):
        try:
            if doc.GetVariable("CMDACTIVE") == 0:
                break
        except:
            pass
        time.sleep(interval)


def make_variant(points: List[float]):
    return win32com.client.VARIANT(
        win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8,
        points
    )


def enable_log_capture(doc):
    log_path = None
    logfile_was = None
    log_before = ""
    try:
        logfile_was = doc.GetVariable("LOGFILEMODE")
        doc.SetVariable("LOGFILEMODE", 1)
        time.sleep(0.5)
        log_path = str(doc.GetVariable("LOGFILENAME"))
        if os.path.exists(log_path):
            for enc in [locale.getpreferredencoding(), "gbk", "gb2312", "utf-8"]:
                try:
                    with open(log_path, "r", encoding=enc) as f:
                        log_before = f.read()
                    break
                except:
                    continue
    except:
        pass
    return log_path, log_before, logfile_was


def disable_log_capture(doc, logfile_was, log_path, log_before) -> List[str]:
    output_lines = []
    if log_path and os.path.exists(log_path):
        try:
            log_after = None
            for enc in [locale.getpreferredencoding(), "gbk", "gb2312", "utf-8"]:
                try:
                    with open(log_path, "r", encoding=enc) as f:
                        log_after = f.read()
                    break
                except:
                    continue
            if log_after is not None:
                new_part = log_after[len(log_before):].strip()
                if new_part:
                    for line in new_part.split("\n"):
                        s = line.strip()
                        if s:
                            output_lines.append(s)
        except:
            pass
    if logfile_was is not None:
        try:
            doc.SetVariable("LOGFILEMODE", logfile_was)
        except:
            pass
    return output_lines


def set_active_layer(doc, layer_name):
    if not layer_name:
        return
    try:
        doc.Layers.Item(layer_name)
    except:
        doc.Layers.Add(layer_name)
    doc.ActiveLayer = doc.Layers.Item(layer_name)
