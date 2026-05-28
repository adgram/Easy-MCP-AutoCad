from mcp.server.fastmcp import Context
from typing import Optional, List
import win32com.client
import json
import math

from tools.com_helpers import get_acad, make_variant, set_active_layer
from tools.database import db_connect


def register_tools(mcp):
    @mcp.tool()
    def create_new_drawing(ctx: Context, template: Optional[str] = None) -> str:
        """创建新的 AutoCAD 图纸"""
        try:
            acad = win32com.client.Dispatch("AutoCAD.Application.25")
            acad.Visible = True
            if template:
                acad.Documents.Add(template)
            else:
                acad.Documents.Add()
            return "成功创建新图纸"
        except Exception as e:
            return f"创建图纸失败: {str(e)}"

    @mcp.tool()
    def draw_line(ctx: Context, start_x: float, start_y: float, end_x: float, end_y: float, layer: Optional[str] = None) -> str:
        """在AutoCAD中绘制直线"""
        acad, doc, err = get_acad()
        if err:
            return err
        try:
            model_space = doc.ModelSpace
            set_active_layer(doc, layer)
            line = model_space.AddLine(
                make_variant([start_x, start_y, 0]),
                make_variant([end_x, end_y, 0])
            )
            with db_connect() as conn:
                cursor = conn.cursor()
                props = {"start_point": [start_x, start_y, 0], "end_point": [end_x, end_y, 0]}
                cursor.execute(
                    "INSERT INTO cad_elements (handle, name, type, layer, properties) VALUES (?, ?, ?, ?, ?)",
                    (line.Handle, "Line", "AcDbLine", doc.ActiveLayer.Name, json.dumps(props))
                )
                conn.commit()
            return f"已创建直线，Handle: {line.Handle}, 图层: {doc.ActiveLayer.Name}"
        except Exception as e:
            return f"创建直线失败: {str(e)}"

    @mcp.tool()
    def draw_circle(ctx: Context, center_x: float, center_y: float, radius: float, layer: Optional[str] = None) -> str:
        """在AutoCAD中绘制圆"""
        acad, doc, err = get_acad()
        if err:
            return err
        try:
            model_space = doc.ModelSpace
            set_active_layer(doc, layer)
            circle = model_space.AddCircle(make_variant([center_x, center_y, 0]), radius)
            with db_connect() as conn:
                cursor = conn.cursor()
                props = {"center_point": [center_x, center_y, 0], "radius": radius}
                cursor.execute(
                    "INSERT INTO cad_elements (handle, name, type, layer, properties) VALUES (?, ?, ?, ?, ?)",
                    (circle.Handle, "Circle", "AcDbCircle", doc.ActiveLayer.Name, json.dumps(props))
                )
                conn.commit()
            return f"已创建圆，Handle: {circle.Handle}, 半径: {radius}, 图层: {doc.ActiveLayer.Name}"
        except Exception as e:
            return f"创建圆失败: {str(e)}"

    @mcp.tool()
    def draw_polyline(ctx: Context, points: List[float], closed: bool = False, layer: Optional[str] = None) -> str:
        """绘制多段线 (Polyline)"""
        acad, doc, err = get_acad()
        if err:
            return err
        try:
            model_space = doc.ModelSpace
            set_active_layer(doc, layer)
            pline = model_space.AddLightWeightPolyline(make_variant(points))
            pline.Closed = closed
            return f"已创建多段线，Handle: {pline.Handle}"
        except Exception as e:
            return f"创建多段线失败: {str(e)}"

    @mcp.tool()
    def draw_rectangle(ctx: Context, x1: float, y1: float, x2: float, y2: float, layer: Optional[str] = None) -> str:
        """绘制矩形"""
        points = [x1, y1, x2, y1, x2, y2, x1, y2]
        return draw_polyline(ctx, points, closed=True, layer=layer)

    @mcp.tool()
    def draw_text(ctx: Context, text_string: str, insert_x: float, insert_y: float, height: float = 2.5, rotation: float = 0, layer: Optional[str] = None) -> str:
        """绘制单行文字"""
        acad, doc, err = get_acad()
        if err:
            return err
        try:
            model_space = doc.ModelSpace
            set_active_layer(doc, layer)
            text_obj = model_space.AddText(text_string, make_variant([insert_x, insert_y, 0]), height)
            if rotation != 0:
                text_obj.Rotation = rotation * math.pi / 180
            return f"已创建文字 '{text_string}'，Handle: {text_obj.Handle}"
        except Exception as e:
            return f"创建文字失败: {str(e)}"

    @mcp.tool()
    def draw_device_connection(ctx: Context, start_device: str, end_device: str, start_x: Optional[float] = None, start_y: Optional[float] = None, end_x: Optional[float] = None, end_y: Optional[float] = None, layer: Optional[str] = None) -> str:
        """绘制设备之间的连接线"""
        acad, doc, err = get_acad()
        if err:
            return err
        try:
            model_space = doc.ModelSpace
            set_active_layer(doc, layer)

            if start_x is None or start_y is None or end_x is None or end_y is None:
                with db_connect() as conn:
                    cursor = conn.cursor()
                    cursor.execute(
                        "SELECT properties FROM cad_elements WHERE type = 'CustomDevice' AND json_extract(properties, '$.label') = ?",
                        (start_device,)
                    )
                    start_result = cursor.fetchone()
                    cursor.execute(
                        "SELECT properties FROM cad_elements WHERE type = 'CustomDevice' AND json_extract(properties, '$.label') = ?",
                        (end_device,)
                    )
                    end_result = cursor.fetchone()

                if not start_result:
                    return f"未找到标签为 {start_device} 的设备"
                if not end_result:
                    return f"未找到标签为 {end_device} 的设备"

                start_props = json.loads(start_result[0])
                end_props = json.loads(end_result[0])
                start_x = start_props["position"][0] - 5
                start_y = start_props["position"][1]
                end_x = end_props["position"][0] - 5
                end_y = end_props["position"][1]

            line1 = model_space.AddLine(make_variant([start_x, start_y, 0]), make_variant([start_x - 10, start_y, 0]))
            line2 = model_space.AddLine(make_variant([start_x - 10, start_y, 0]), make_variant([start_x - 10, end_y, 0]))
            line3 = model_space.AddLine(make_variant([start_x - 10, end_y, 0]), make_variant([end_x, end_y, 0]))
            return f"已创建从 {start_device} 到 {end_device} 的连接线"
        except Exception as e:
            return f"创建连接线失败: {str(e)}"

    @mcp.tool()
    def create_layer(ctx: Context, name: str, color_index: int = 7) -> str:
        """创建或修改图层"""
        acad, doc, err = get_acad()
        if err:
            return err
        try:
            try:
                layer = doc.Layers.Item(name)
            except:
                layer = doc.Layers.Add(name)
            layer.Color = color_index
            return f"图层 '{name}' 已设置，颜色: {color_index}"
        except Exception as e:
            return f"设置图层失败: {str(e)}"

    @mcp.tool()
    def autocad_get_layers(ctx: Context) -> str:
        """获取图纸中所有图层信息（名称、颜色、状态等）"""
        acad, doc, err = get_acad()
        if err:
            return json.dumps({"success": False, "error": err, "data": None}, ensure_ascii=False)
        try:
            layers = []
            for i in range(doc.Layers.Count):
                lay = doc.Layers.Item(i)
                layers.append({
                    "name": lay.Name,
                    "color": lay.Color,
                    "linetype": lay.Linetype,
                    "on": lay.LayerOn,
                    "frozen": lay.Freeze,
                    "locked": lay.Lock,
                })
            return json.dumps({"success": True, "error": None, "data": layers}, ensure_ascii=False)
        except Exception as e:
            return json.dumps({"success": False, "error": str(e), "data": None}, ensure_ascii=False)

    @mcp.tool()
    def autocad_get_blocks(ctx: Context) -> str:
        """获取图纸中所有块定义信息（名称、数量等）"""
        acad, doc, err = get_acad()
        if err:
            return json.dumps({"success": False, "error": err, "data": None}, ensure_ascii=False)
        try:
            blocks = []
            for i in range(doc.Blocks.Count):
                blk = doc.Blocks.Item(i)
                blocks.append({
                    "name": blk.Name,
                    "count": blk.Count,
                    "origin": [blk.Origin[0], blk.Origin[1], blk.Origin[2]] if hasattr(blk, "Origin") else None,
                })
            return json.dumps({"success": True, "error": None, "data": blocks}, ensure_ascii=False)
        except Exception as e:
            return json.dumps({"success": False, "error": str(e), "data": None}, ensure_ascii=False)

    @mcp.tool()
    def autocad_get_linetypes(ctx: Context) -> str:
        """获取图纸中所有线型信息"""
        acad, doc, err = get_acad()
        if err:
            return json.dumps({"success": False, "error": err, "data": None}, ensure_ascii=False)
        try:
            linetypes = []
            for i in range(doc.Linetypes.Count):
                lt = doc.Linetypes.Item(i)
                linetypes.append({
                    "name": lt.Name,
                    "description": lt.Description if hasattr(lt, "Description") else "",
                })
            return json.dumps({"success": True, "error": None, "data": linetypes}, ensure_ascii=False)
        except Exception as e:
            return json.dumps({"success": False, "error": str(e), "data": None}, ensure_ascii=False)
