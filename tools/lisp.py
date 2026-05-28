from mcp.server.fastmcp import Context
from typing import Optional, List, Any
import json
import time
import os
import locale
import math

from tools.com_helpers import get_acad, wait_cmd, enable_log_capture, disable_log_capture


def register_tools(mcp):
    @mcp.tool()
    def execute_command(ctx: Context, command: str) -> str:
        """在AutoCAD中执行命令（如 LINE, CIRCLE, ERASE, TRIM 等）并获取命令行输出"""
        acad, doc, err = get_acad()
        if err:
            return err

        wait_cmd(doc)
        log_path, log_before, logfile_was = enable_log_capture(doc)

        cmd = command.strip()
        if not cmd.endswith("\n"):
            cmd += "\n"
        doc.SendCommand(cmd)

        wait_cmd(doc, 300)
        time.sleep(0.5)
        output_lines = disable_log_capture(doc, logfile_was, log_path, log_before)

        result_parts = [f"命令: {command}"]
        if output_lines:
            result_parts.append("输出:")
            for line in output_lines[-30:]:
                result_parts.append(f"  {line}")
        else:
            result_parts.append("输出: (命令未产生文本输出)")
        return "\n".join(result_parts)

    @mcp.tool()
    def create_block(ctx: Context, name: str, handles: List[str], base_point: Optional[List[float]] = None, insert: bool = True) -> str:
        """创建图块定义，从指定实体生成块，可选择是否插入块参照"""
        acad, doc, err = get_acad()
        if err:
            return json.dumps({"success": False, "error": err, "data": None}, ensure_ascii=False)

        bp = base_point or [0.0, 0.0, 0.0]
        if len(bp) < 3:
            bp = [bp[0], bp[1], 0.0]

        wait_cmd(doc)

        handle_list = " ".join(f'"{h}"' for h in handles)
        lisp_code = f"""
(progn
  (setq _name "{name}")
  (setq _bp (list {bp[0]} {bp[1]} {bp[2]}))
  (setq _handles (list {handle_list}))
  (if (tblsearch "BLOCK" _name)
    (setvar "USERS1" (strcat "块 \\"" _name "\\" 已存在"))
    (progn
      (entmake (list (cons 0 "BLOCK") (cons 2 _name) (cons 70 2) (cons 10 _bp)))
      (foreach _h _handles
        (setq _ent (handent _h))
        (if _ent
          (progn
            (setq _ed (entget _ent))
            (foreach _x '(-1 5 330 360 -2)
              (setq _ed (vl-remove (assoc _x _ed) _ed))
            )
            (entmake _ed)
          )
        )
      )
      (entmake (list (cons 0 "ENDBLK")))
      {f'(entmake (list (cons 0 "INSERT") (cons 2 "{name}") (cons 10 (list {bp[0]} {bp[1]} {bp[2]}))))' if insert else 'nil'}
      (setvar "USERS1" (strcat "块 \\"" _name "\\" 已创建，包含 " (itoa (length _handles)) " 个实体"))
    )
  )
)
""".strip()

        log_path, log_before, logfile_was = enable_log_capture(doc)
        wrapped = f'(progn (setvar "USERS1" (vl-princ-to-string (progn {lisp_code}))))\n'
        doc.SendCommand(wrapped)
        wait_cmd(doc)
        time.sleep(0.3)
        output_lines = disable_log_capture(doc, logfile_was, log_path, log_before)

        result = ""
        try:
            result = str(doc.GetVariable("USERS1"))
        except:
            pass

        return json.dumps({
            "success": True,
            "error": None,
            "data": {
                "name": name,
                "handles": handles,
                "base_point": bp,
                "message": result or "块已创建",
            }
        }, ensure_ascii=False)

    @mcp.tool()
    def autocad_select_objects(ctx: Context, filter_list: Optional[str] = None, mode: str = "X", points: Optional[List[float]] = None) -> str:
        """在AutoCAD中通过ssget选择对象，支持多种模式和过滤条件，返回handle列表（自动避免交互弹窗）"""
        acad, doc, err = get_acad()
        if err:
            return err

        wait_cmd(doc)

        mode_upper = (mode or "X").upper()
        no_pick_modes = {"X", "A", "L", "P", "I"}

        if mode_upper in no_pick_modes:
            ss_args = f'"{mode_upper}"'
        elif mode_upper in ("C", "W"):
            if not points or len(points) < 4:
                return f"模式{mode}需要至少4个坐标(两个点)，传入: {points}"
            ss_args = f'"{mode_upper}" \'({points[0]} {points[1]} 0) \'({points[2]} {points[3]} 0)'
        elif mode_upper in ("F", "CP", "WP"):
            if not points or len(points) < 2 or len(points) % 2 != 0:
                return f"模式{mode}需要至少2个坐标(偶数个)，传入: {points}"
            pt_list = " ".join(f'({points[i]} {points[i+1]} 0)' for i in range(0, len(points), 2))
            ss_args = f'"{mode_upper}" \'({pt_list})'
        else:
            mode_upper = "X"
            ss_args = '"X"'

        if filter_list:
            filter_str = filter_list.strip()
            if filter_str.startswith("'"):
                filter_str = filter_str[1:]
            ss_lisp = f'(ssget {ss_args} \'{filter_str})'
        else:
            ss_lisp = f'(ssget {ss_args})'

        lisp_code = f"""
(progn
  (setq _ss {ss_lisp})
  (setq _handles '())
  (if _ss
    (repeat (setq _i (sslength _ss))
      (setq _handles (cons (cdr (assoc 5 (entget (ssname _ss (setq _i (1- _i)))))) _handles))
    )
  )
  (setvar "USERS1" (vl-princ-to-string _handles))
)
""".strip()

        log_path, log_before, logfile_was = enable_log_capture(doc)
        wrapped = f'(progn (setvar "USERS1" (vl-princ-to-string (progn {lisp_code}))))\n'
        doc.SendCommand(wrapped)
        wait_cmd(doc)
        time.sleep(0.3)
        output_lines = disable_log_capture(doc, logfile_was, log_path, log_before)

        result = ""
        try:
            result = str(doc.GetVariable("USERS1"))
        except:
            pass

        parts = [f"模式: {mode_upper}", f"过滤: {filter_list or '(全部)'}"]
        if result:
            parts.append(f"结果: {result}")
        if output_lines:
            parts.append("交互输出:")
            for line in output_lines[-10:]:
                parts.append(f"  {line}")
        if not result and not output_lines:
            parts.append("(无结果)")
        return "\n".join(parts)

    @mcp.tool()
    def autocad_eval_lisp(ctx: Context, lisp_code: str) -> str:
        """在AutoCAD中执行LISP代码，原子化执行（解决entsel/getpoint在while循环中代码泄漏问题），返回完整结果"""
        acad, doc, err = get_acad()
        if err:
            return err

        wait_cmd(doc)
        log_path, log_before, logfile_was = enable_log_capture(doc)

        wrapped = f'(progn (setvar "USERS1" (vl-princ-to-string (progn {lisp_code}))))\n'
        doc.SendCommand(wrapped)

        for _ in range(3000):
            try:
                if doc.GetVariable("CMDACTIVE") == 0:
                    break
            except:
                pass
            time.sleep(0.1)

        output_lines = disable_log_capture(doc, logfile_was, log_path, log_before)

        result = ""
        try:
            result = str(doc.GetVariable("USERS1"))
        except:
            pass

        parts = [f"LISP: {lisp_code}"]
        if result:
            parts.append(f"结果: {result}")
        if output_lines:
            parts.append("交互输出:")
            for line in output_lines[-30:]:
                parts.append(f"  {line}")
        if not result and not output_lines:
            parts.append("(无返回值)")
        return "\n".join(parts)

    @mcp.tool()
    def autocad_get_sysvar(ctx: Context, name: str) -> str:
        """读取AutoCAD系统变量"""
        acad, doc, err = get_acad()
        if err:
            return json.dumps({"success": False, "error": err, "data": None}, ensure_ascii=False)
        try:
            value = doc.GetVariable(name)
            return json.dumps({"success": True, "error": None, "data": str(value)}, ensure_ascii=False)
        except Exception as e:
            return json.dumps({"success": False, "error": str(e), "data": None}, ensure_ascii=False)

    @mcp.tool()
    def autocad_set_sysvar(ctx: Context, name: str, value: Any) -> str:
        """设置AutoCAD系统变量"""
        acad, doc, err = get_acad()
        if err:
            return json.dumps({"success": False, "error": err, "data": None}, ensure_ascii=False)
        try:
            doc.SetVariable(name, value)
            return json.dumps({"success": True, "error": None, "data": f"{name} = {value}"}, ensure_ascii=False)
        except Exception as e:
            return json.dumps({"success": False, "error": str(e), "data": None}, ensure_ascii=False)
