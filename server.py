from mcp.server.fastmcp import FastMCP, Context
from typing import Optional, List, Dict, Any
import win32com.client
import sqlite3
import json
import random
import re

# 创建服务器
mcp = FastMCP("AutoCAD-DB-Server")

# 初始化 SQLite 数据库
def init_db():
    try:
        conn = sqlite3.connect("autocad_data.db")
        cursor = conn.cursor()
        # 创建实体表
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS cad_elements (
            id INTEGER PRIMARY KEY,
            handle TEXT UNIQUE,
            name TEXT NOT NULL,
            type TEXT NOT NULL,
            layer TEXT,
            properties TEXT
        )
        ''')
        # 创建文字内容统计表
        cursor.execute('''
        CREATE TABLE IF NOT EXISTS text_patterns (
            id INTEGER PRIMARY KEY,
            pattern TEXT UNIQUE,
            count INTEGER DEFAULT 0,
            drawing TEXT
        )
        ''')
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"数据库初始化失败: {str(e)}")
        return False

# 确保数据库初始化
init_db()

# ======= AutoCAD 基础工具 =======

@mcp.tool()
def create_new_drawing(ctx: Context, template: Optional[str] = None) -> str:
    """创建新的 AutoCAD 图纸"""
    try:
        # 尝试连接到 AutoCAD
        acad = win32com.client.Dispatch("AutoCAD.Application.25")
        acad.Visible = True
        
        # 创建新文档
        if template:
            doc = acad.Documents.Add(template)
        else:
            doc = acad.Documents.Add()
            
        return f"成功创建新图纸"
    except Exception as e:
        return f"创建图纸失败: {str(e)}"

@mcp.tool()
def draw_line(ctx: Context, start_x: float, start_y: float, end_x: float, end_y: float, layer: Optional[str] = None) -> str:
    """在AutoCAD中绘制直线
    
    Args:
        start_x: 起点X坐标
        start_y: 起点Y坐标
        end_x: 终点X坐标
        end_y: 终点Y坐标
        layer: 可选的图层名称
    """
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application.25")
        if acad.Documents.Count == 0:
            return "无打开的文档，请先创建或打开一个图纸"
        
        doc = acad.ActiveDocument
        model_space = doc.ModelSpace
        
        # 如果指定了图层，先切换或创建图层
        if layer:
            try:
                # 尝试获取图层
                doc.Layers.Item(layer)
            except:
                # 图层不存在，创建新图层
                doc.Layers.Add(layer)
            
            # 设置当前图层
            doc.ActiveLayer = doc.Layers.Item(layer)
        
        # 创建直线
        line = model_space.AddLine(
            win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, [start_x, start_y, 0]),
            win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, [end_x, end_y, 0])
        )
        
        # 将线条信息存入数据库
        conn = sqlite3.connect("autocad_data.db")
        cursor = conn.cursor()
        props = {
            "start_point": [start_x, start_y, 0],
            "end_point": [end_x, end_y, 0]
        }
        cursor.execute(
            "INSERT INTO cad_elements (handle, name, type, layer, properties) VALUES (?, ?, ?, ?, ?)",
            (line.Handle, "Line", "AcDbLine", doc.ActiveLayer.Name, json.dumps(props))
        )
        conn.commit()
        conn.close()
        
        return f"已创建直线，Handle: {line.Handle}, 图层: {doc.ActiveLayer.Name}"
    except Exception as e:
        return f"创建直线失败: {str(e)}"

@mcp.tool()
def draw_circle(ctx: Context, center_x: float, center_y: float, radius: float, layer: Optional[str] = None) -> str:
    """在AutoCAD中绘制圆
    
    Args:
        center_x: 圆心X坐标
        center_y: 圆心Y坐标  
        radius: 半径
        layer: 可选的图层名称
    """
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application.25")
        if acad.Documents.Count == 0:
            return "无打开的文档，请先创建或打开一个图纸"
        
        doc = acad.ActiveDocument
        model_space = doc.ModelSpace
        
        # 如果指定了图层，先切换或创建图层
        if layer:
            try:
                # 尝试获取图层
                doc.Layers.Item(layer)
            except:
                # 图层不存在，创建新图层
                doc.Layers.Add(layer)
            
            # 设置当前图层
            doc.ActiveLayer = doc.Layers.Item(layer)
        
        # 创建圆
        center_point = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, [center_x, center_y, 0])
        circle = model_space.AddCircle(center_point, radius)
        
        # 将圆信息存入数据库
        conn = sqlite3.connect("autocad_data.db")
        cursor = conn.cursor()
        props = {
            "center_point": [center_x, center_y, 0],
            "radius": radius
        }
        cursor.execute(
            "INSERT INTO cad_elements (handle, name, type, layer, properties) VALUES (?, ?, ?, ?, ?)",
            (circle.Handle, "Circle", "AcDbCircle", doc.ActiveLayer.Name, json.dumps(props))
        )
        conn.commit()
        conn.close()
        
        return f"已创建圆，Handle: {circle.Handle}, 半径: {radius}, 图层: {doc.ActiveLayer.Name}"
    except Exception as e:
        return f"创建圆失败: {str(e)}"

# ======= 实体扫描和数据库交互 =======

@mcp.tool()
def scan_all_entities(ctx: Context) -> str:
    """扫描当前图纸中的所有实体并保存到数据库"""
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application.25")
        if acad.Documents.Count == 0:
            return "无打开的文档，请先创建或打开一个图纸"
        
        doc = acad.ActiveDocument
        model_space = doc.ModelSpace
        
        # 连接数据库
        conn = sqlite3.connect("autocad_data.db")
        cursor = conn.cursor()
        
        # 清空现有记录（可选）
        cursor.execute("DELETE FROM cad_elements")
        
        # 统计信息
        count = 0
        entity_types = {}
        
        # 遍历所有实体
        for i in range(model_space.Count):
            try:
                entity = model_space.Item(i)
                entity_type = entity.ObjectName
                
                # 统计类型数量
                if entity_type in entity_types:
                    entity_types[entity_type] += 1
                else:
                    entity_types[entity_type] = 1
                
                # 获取基本属性
                properties = {}
                if entity_type == "AcDbLine":
                    properties = {
                        "start_point": [entity.StartPoint[0], entity.StartPoint[1], entity.StartPoint[2]],
                        "end_point": [entity.EndPoint[0], entity.EndPoint[1], entity.EndPoint[2]]
                    }
                elif entity_type == "AcDbCircle":
                    properties = {
                        "center": [entity.Center[0], entity.Center[1], entity.Center[2]],
                        "radius": entity.Radius
                    }
                elif entity_type == "AcDbText" or entity_type == "AcDbMText":
                    properties = {
                        "text": entity.TextString,
                        "position": [entity.InsertionPoint[0], entity.InsertionPoint[1], entity.InsertionPoint[2]] if hasattr(entity, "InsertionPoint") else None,
                        "height": entity.Height if hasattr(entity, "Height") else None
                    }
                
                # 将实体信息存入数据库
                cursor.execute(
                    "INSERT OR REPLACE INTO cad_elements (handle, name, type, layer, properties) VALUES (?, ?, ?, ?, ?)",
                    (entity.Handle, entity.ObjectName.replace("AcDb", ""), entity.ObjectName, entity.Layer, json.dumps(properties))
                )
                
                count += 1
            except Exception as e:
                print(f"处理实体 {i} 时出错: {str(e)}")
        
        conn.commit()
        conn.close()
        
        # 格式化类型统计
        type_summary = "\n".join([f"{t}: {c}" for t, c in entity_types.items()])
        
        return f"已扫描并保存 {count} 个实体到数据库。\n\n实体类型统计:\n{type_summary}"
    except Exception as e:
        return f"扫描实体失败: {str(e)}"

@mcp.tool()
def highlight_entity(ctx: Context, handle: str, color: int = 1) -> str:
    """通过Handle在AutoCAD中高亮显示指定实体
    
    Args:
        handle: 实体的Handle值
        color: 高亮颜色码（1=红色, 2=黄色, 3=绿色, 4=青色, 5=蓝色, 6=洋红色）
    """
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application.25")
        if acad.Documents.Count == 0:
            return "无打开的文档，请先创建或打开一个图纸"
        
        doc = acad.ActiveDocument
        
        # 根据Handle直接获取实体
        try:
            entity = doc.HandleToObject(handle)
        except:
             return f"未找到Handle为 {handle} 的实体"

        original_color = entity.Color
        entity.Color = color
        
        return f"已高亮实体 {handle}，颜色从 {original_color} 改为 {color}"
    except Exception as e:
        return f"高亮实体失败: {str(e)}"

# ======= 文本分析工具 =======

@mcp.tool()
def count_text_patterns(ctx: Context, pattern: str = "PMC-3M") -> str:
    """统计图纸中文本实体中特定模式的出现次数
    
    Args:
        pattern: 要搜索的文本模式，默认为"PMC-3M"
    """
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application.25")
        if acad.Documents.Count == 0:
            return "无打开的文档，请先创建或打开一个图纸"
        
        doc = acad.ActiveDocument
        model_space = doc.ModelSpace
        drawing_name = doc.Name
        
        # 连接数据库
        conn = sqlite3.connect("autocad_data.db")
        cursor = conn.cursor()
        
        # 计数器
        count = 0
        matching_entities = []
        
        # 遍历所有实体
        for i in range(model_space.Count):
            try:
                entity = model_space.Item(i)
                
                # 检查是否为文本实体
                if hasattr(entity, "TextString"):
                    text = entity.TextString
                    
                    # 搜索模式
                    if pattern in text:
                        count += 1
                        matching_entities.append({
                            "handle": entity.Handle,
                            "text": text,
                            "layer": entity.Layer,
                            "position": [entity.InsertionPoint[0], entity.InsertionPoint[1]] if hasattr(entity, "InsertionPoint") else None
                        })
            except Exception as e:
                print(f"处理文本实体 {i} 时出错: {str(e)}")
        
        # 保存统计结果到数据库
        cursor.execute(
            "INSERT OR REPLACE INTO text_patterns (pattern, count, drawing) VALUES (?, ?, ?)",
            (pattern, count, drawing_name)
        )
        conn.commit()
        conn.close()
        
        result = f"在图纸 '{drawing_name}' 中找到 {count} 处匹配模式 '{pattern}' 的文本。"
        
        # 如果有匹配项，显示详细信息
        if count > 0:
            details = "\n\n匹配详情："
            for i, match in enumerate(matching_entities[:10]):  # 限制显示前10个
                details += f"\n{i+1}. 文本: '{match['text']}', 图层: {match['layer']}, Handle: {match['handle']}"
            
            if len(matching_entities) > 10:
                details += f"\n... 以及其他 {len(matching_entities) - 10} 个匹配项"
                
            result += details
        
        return result
    except Exception as e:
        return f"统计文本模式失败: {str(e)}"

@mcp.tool()
def highlight_text_matches(ctx: Context, pattern: str = "PMC-3M", color: int = 1) -> str:
    """高亮显示包含指定文本模式的所有文本实体
    
    Args:
        pattern: 要搜索的文本模式，默认为"PMC-3M"
        color: 高亮颜色码（1=红色, 2=黄色, 3=绿色, 4=青色, 5=蓝色, 6=洋红色）
    """
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application.25")
        if acad.Documents.Count == 0:
            return "无打开的文档，请先创建或打开一个图纸"
        
        doc = acad.ActiveDocument
        model_space = doc.ModelSpace
        
        # 创建选择集
        try:
            # 尝试删除可能存在的选择集
            doc.SelectionSets.Item("TextMatches").Delete()
        except:
            pass
        
        selection = doc.SelectionSets.Add("TextMatches")
        
        # 计数器
        count = 0
        
        # 遍历所有实体
        for i in range(model_space.Count):
            try:
                entity = model_space.Item(i)
                
                # 检查是否为文本实体
                if hasattr(entity, "TextString"):
                    text = entity.TextString
                    
                    # 搜索模式
                    if pattern in text:
                        # 保存原始颜色
                        original_color = entity.Color
                        
                        # 修改颜色
                        entity.Color = color
                        
                        # 添加到选择集
                        selection.AddItems([entity])
                        
                        count += 1
            except Exception as e:
                print(f"处理文本实体 {i} 时出错: {str(e)}")
        
        if count > 0:
            # 缩放到选择集
            doc.ActiveView.ZoomAll()
            return f"已高亮显示 {count} 个包含 '{pattern}' 的文本实体"
        else:
            selection.Delete()
            return f"未找到包含 '{pattern}' 的文本实体"
    except Exception as e:
        return f"高亮文本匹配失败: {str(e)}"

# ======= 数据库查询工具 =======

@mcp.tool()
def get_all_tables(ctx: Context) -> str:
    """获取数据库中的所有表"""
    try:
        conn = sqlite3.connect("autocad_data.db")
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
        tables = cursor.fetchall()
        conn.close()
        
        table_list = [table[0] for table in tables]
        return json.dumps(table_list, indent=2)
    except Exception as e:
        return f"获取表列表失败: {str(e)}"

@mcp.tool()
def get_table_schema(ctx: Context, table_name: str) -> str:
    """获取指定表的结构信息"""
    try:
        conn = sqlite3.connect("autocad_data.db")
        cursor = conn.cursor()
        cursor.execute(f"PRAGMA table_info({table_name});")
        columns = cursor.fetchall()
        conn.close()
        
        schema = []
        for col in columns:
            schema.append({
                "cid": col[0],
                "name": col[1],
                "type": col[2],
                "notnull": col[3],
                "default_value": col[4],
                "pk": col[5]
            })
        
        return json.dumps(schema, indent=2)
    except Exception as e:
        return f"获取表结构失败: {str(e)}"

@mcp.tool()
def execute_query(ctx: Context, query: str) -> str:
    """执行自定义数据库查询"""
    try:
        conn = sqlite3.connect("autocad_data.db")
        cursor = conn.cursor()
        cursor.execute(query)
        
        # 如果是SELECT查询，获取结果
        if query.strip().upper().startswith("SELECT"):
            columns = [desc[0] for desc in cursor.description]
            rows = cursor.fetchall()
            
            # 将结果转换为字典列表
            result = []
            for row in rows:
                result.append(dict(zip(columns, row)))
                
            conn.commit()
            conn.close()
            return json.dumps(result, indent=2)
        else:
            # 非SELECT查询，返回影响的行数
            conn.commit()
            affected = cursor.rowcount
            conn.close()
            return f"执行成功，影响了 {affected} 行"
    except Exception as e:
        return f"执行查询失败: {str(e)}"

@mcp.tool()
def query_and_highlight(ctx: Context, sql_query: str, highlight_color: int = 1) -> str:
    """根据SQL查询结果高亮显示AutoCAD实体
    
    Args:
        sql_query: 必须是返回handle列的SQL查询
        highlight_color: 高亮颜色码（1-255）
    """
    try:
        # 执行查询
        conn = sqlite3.connect("autocad_data.db")
        cursor = conn.cursor()
        cursor.execute(sql_query)
        rows = cursor.fetchall()
        conn.close()
        
        if not rows:
            return "查询未返回任何结果"
        
        # 获取列名
        column_names = [description[0] for description in cursor.description]
        
        # 查找handle列
        handle_index = -1
        for i, name in enumerate(column_names):
            if name.lower() == 'handle':
                handle_index = i
                break
        
        if handle_index == -1:
            return "查询结果中未找到handle列"
        
        # 提取所有handle
        handles = [row[handle_index] for row in rows]
        
        # 高亮实体
        acad = win32com.client.Dispatch("AutoCAD.Application.25")
        if acad.Documents.Count == 0:
            return "无打开的文档，请先创建或打开一个图纸"
        
        doc = acad.ActiveDocument
        
        # 创建选择集
        try:
            doc.SelectionSets.Item("QueryResults").Delete()
        except:
            pass
        
        selection = doc.SelectionSets.Add("QueryResults")
        
        # 高亮找到的实体
        highlighted_count = 0
        entities_to_add = []
        
        for handle in handles:
            try:
                # 直接获取实体
                entity = doc.HandleToObject(handle)
                entity.Color = highlight_color
                entities_to_add.append(entity)
                highlighted_count += 1
            except Exception as e:
                print(f"处理实体 {handle} 时出错: {str(e)}")
        
        # 将实体添加到选择集（方便在CAD中查看属性等）
        if entities_to_add:
            try:
                # AddItems需要变体数组，comtypes会自动处理list，pywin32有时需要特别处理
                # 这里简单尝试，如果失败不影响高亮结果
                selection.AddItems(entities_to_add)
            except:
                pass
        
        if highlighted_count > 0:
            # 缩放到选择集
            doc.ActiveView.ZoomAll()
            return f"已高亮显示 {highlighted_count} 个实体（共 {len(handles)} 个结果）"
        else:
            return f"未能高亮任何实体"
    except Exception as e:
        return f"查询并高亮失败: {str(e)}"

@mcp.tool()
def draw_polyline(ctx: Context, points: List[float], closed: bool = False, layer: Optional[str] = None) -> str:
    """绘制多段线 (Polyline)
    
    Args:
        points: 坐标列表 [x1, y1, x2, y2, ...]
        closed: 是否闭合
        layer: 图层名称
    """
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application.25")
        if acad.Documents.Count == 0:
            return "无打开的文档"
            
        doc = acad.ActiveDocument
        model_space = doc.ModelSpace
        
        # 处理图层
        if layer:
            try:
                doc.Layers.Item(layer)
            except:
                doc.Layers.Add(layer)
            doc.ActiveLayer = doc.Layers.Item(layer)

        # 转换坐标点
        # LightweightPolyline 需要 2D 坐标数组 (double)
        pt_array = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, points)
        
        # 创建多段线
        pline = model_space.AddLightWeightPolyline(pt_array)
        pline.Closed = closed
        
        return f"已创建多段线，Handle: {pline.Handle}"
    except Exception as e:
        return f"创建多段线失败: {str(e)}"

@mcp.tool()
def draw_rectangle(ctx: Context, x1: float, y1: float, x2: float, y2: float, layer: Optional[str] = None) -> str:
    """绘制矩形
    
    Args:
        x1, y1: 角点1坐标
        x2, y2: 对角点2坐标
        layer: 图层名称
    """
    # 构造矩形的4个顶点坐标
    points = [x1, y1, x2, y1, x2, y2, x1, y2]
    return draw_polyline(ctx, points, closed=True, layer=layer)

@mcp.tool()
def draw_text(ctx: Context, text_string: str, insert_x: float, insert_y: float, height: float = 2.5, rotation: float = 0, layer: Optional[str] = None) -> str:
    """绘制单行文字
    
    Args:
        text_string: 文字内容
        insert_x, insert_y: 插入点坐标
        height: 文字高度
        rotation: 旋转角度 (度)
        layer: 图层名称
    """
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application.25")
        if acad.Documents.Count == 0:
            return "无打开的文档"
            
        doc = acad.ActiveDocument
        model_space = doc.ModelSpace
        
        # 处理图层
        if layer:
            try:
                doc.Layers.Item(layer)
            except:
                doc.Layers.Add(layer)
            doc.ActiveLayer = doc.Layers.Item(layer)
            
        insert_pnt = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, [insert_x, insert_y, 0])
        text_obj = model_space.AddText(text_string, insert_pnt, height)
        
        if rotation != 0:
            import math
            text_obj.Rotation = rotation * math.pi / 180
            
        return f"已创建文字 '{text_string}'，Handle: {text_obj.Handle}"
    except Exception as e:
        return f"创建文字失败: {str(e)}"

@mcp.tool()
def create_layer(ctx: Context, name: str, color_index: int = 7) -> str:
    """创建或修改图层
    
    Args:
        name: 图层名称
        color_index: 颜色索引 (1-255)
    """
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application.25")
        doc = acad.ActiveDocument
        
        try:
            layer = doc.Layers.Item(name)
        except:
            layer = doc.Layers.Add(name)
            
        layer.Color = color_index
        return f"图层 '{name}' 已设置，颜色: {color_index}"
    except Exception as e:
        return f"设置图层失败: {str(e)}"

@mcp.tool()
def draw_device_connection(ctx: Context, start_device: str, end_device: str, start_x: Optional[float] = None, start_y: Optional[float] = None, end_x: Optional[float] = None, end_y: Optional[float] = None, layer: Optional[str] = None) -> str:
    """绘制设备之间的连接线
    
    Args:
        start_device: 起始设备标签，如"P14"
        end_device: 结束设备标签，如"P02"
        start_x: 可选的起始点X坐标（如果不提供则自动查找设备）
        start_y: 可选的起始点Y坐标
        end_x: 可选的结束点X坐标
        end_y: 可选的结束点Y坐标
        layer: 可选的图层名称
    """
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application.25")
        if acad.Documents.Count == 0:
            return "无打开的文档，请先创建或打开一个图纸"
        
        doc = acad.ActiveDocument
        model_space = doc.ModelSpace
        
        # 如果指定了图层，先切换或创建图层
        if layer:
            try:
                # 尝试获取图层
                doc.Layers.Item(layer)
            except:
                # 图层不存在，创建新图层
                doc.Layers.Add(layer)
            
            # 设置当前图层
            doc.ActiveLayer = doc.Layers.Item(layer)

        # 如果没有提供坐标，尝试从数据库中查找设备
        if start_x is None or start_y is None or end_x is None or end_y is None:
            conn = sqlite3.connect("autocad_data.db")
            cursor = conn.cursor()
            
            # 查找起始设备
            cursor.execute(
                "SELECT properties FROM cad_elements WHERE type = 'CustomDevice' AND json_extract(properties, '$.label') = ?",
                (start_device,)
            )
            start_result = cursor.fetchone()
            
            # 查找结束设备
            cursor.execute(
                "SELECT properties FROM cad_elements WHERE type = 'CustomDevice' AND json_extract(properties, '$.label') = ?",
                (end_device,)
            )
            end_result = cursor.fetchone()
            
            conn.close()
            
            if not start_result:
                return f"未找到标签为 {start_device} 的设备"
                
            if not end_result:
                return f"未找到标签为 {end_device} 的设备"
            
            # 解析设备位置和尺寸
            start_props = json.loads(start_result[0])
            end_props = json.loads(end_result[0])
            
            start_pos = start_props["position"]
            end_pos = end_props["position"]
            
            # 设置连接线起点和终点（设备的左侧连接点）
            start_x = start_pos[0] - 5  # 设备左侧
            start_y = start_pos[1]
            end_x = end_pos[0] - 5
            end_y = end_pos[1]
            
        # 创建连接线（水平线段 + 垂直线段 + 水平线段）
        # 首先创建起始水平线段
        line1_start = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, [start_x, start_y, 0])
        line1_end = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, [start_x - 10, start_y, 0])
        line1 = model_space.AddLine(line1_start, line1_end)
        
        # 创建垂直连接线
        line2_start = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, [start_x - 10, start_y, 0])
        line2_end = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, [start_x - 10, end_y, 0])
        line2 = model_space.AddLine(line2_start, line2_end)
        
        # 创建结束水平线段
        line3_start = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, [start_x - 10, end_y, 0])
        line3_end = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, [end_x, end_y, 0])
        line3 = model_space.AddLine(line3_start, line3_end)
        
        return f"已创建从 {start_device} 到 {end_device} 的连接线"
    except Exception as e:
        return f"创建连接线失败: {str(e)}"
    
@mcp.tool()
def move_entity(ctx: Context, handle: str, start_point: List[float], end_point: List[float]) -> str:
    """移动实体
    
    Args:
        handle: 实体句柄
        start_point: 基干点 [x, y, z]（通常都是[0,0,0]或者实体的某个点）
        end_point: 目标点 [x, y, z]
    """
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application.25")
        doc = acad.ActiveDocument
        
        try:
            entity = doc.HandleToObject(handle)
        except:
             return f"未找到Handle为 {handle} 的实体"
             
        p1 = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, start_point)
        p2 = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, end_point)
        
        entity.Move(p1, p2)
        return f"已移动实体 {handle}"
    except Exception as e:
        return f"移动实体失败: {str(e)}"

@mcp.tool()
def rotate_entity(ctx: Context, handle: str, base_point: List[float], angle: float) -> str:
    """旋转实体
    
    Args:
        handle: 实体句柄
        base_point: 旋转中心点 [x, y, z]
        angle: 旋转角度（度）
    """
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application.25")
        doc = acad.ActiveDocument
        
        try:
            entity = doc.HandleToObject(handle)
        except:
             return f"未找到Handle为 {handle} 的实体"
             
        p_base = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, base_point)
        
        import math
        angle_rad = angle * math.pi / 180
        
        entity.Rotate(p_base, angle_rad)
        return f"已旋转实体 {handle} {angle} 度"
    except Exception as e:
        return f"旋转实体失败: {str(e)}"

@mcp.tool()
def copy_entity(ctx: Context, handle: str, start_point: List[float], end_point: List[float]) -> str:
    """复制实体
    
    Args:
        handle: 源实体句柄
        start_point: 基干点 [x, y, z]
        end_point: 目标点 [x, y, z]
    """
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application.25")
        doc = acad.ActiveDocument
        
        try:
            entity = doc.HandleToObject(handle)
        except:
             return f"未找到Handle为 {handle} 的实体"
             
        # Copy()方法返回新对象
        new_entity = entity.Copy()
        
        p1 = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, start_point)
        p2 = win32com.client.VARIANT(win32com.client.pythoncom.VT_ARRAY | win32com.client.pythoncom.VT_R8, end_point)
        
        new_entity.Move(p1, p2)
        return f"已复制实体，新Handle: {new_entity.Handle}"
    except Exception as e:
        return f"复制实体失败: {str(e)}"

@mcp.tool()
def execute_command(ctx: Context, command: str) -> str:
    """在AutoCAD中执行命令（如 LINE, CIRCLE, ERASE, TRIM 等）并获取命令行输出
    
    Args:
        command: AutoCAD命令字符串（如 "LINE 0,0 100,100 ", "CIRCLE 50,50 30", "ERASE ALL "），注意命令结束后需要加空格或回车
    """
    import time, os, locale
    acad = win32com.client.Dispatch("AutoCAD.Application.25")
    if acad.Documents.Count == 0:
        return "无打开的文档"
    
    doc = acad.ActiveDocument

    # 等待之前的命令完成
    for _ in range(300):
        try:
            if doc.GetVariable("CMDACTIVE") == 0:
                break
        except:
            pass
        time.sleep(0.1)

    # 启用日志文件捕获完整输出
    log_path = None
    logfile_was = None
    log_content_before = ""
    try:
        logfile_was = doc.GetVariable("LOGFILEMODE")
        doc.SetVariable("LOGFILEMODE", 1)
        time.sleep(0.5)
        log_path = str(doc.GetVariable("LOGFILENAME"))
        if os.path.exists(log_path):
            for enc in [locale.getpreferredencoding(), "gbk", "gb2312", "utf-8"]:
                try:
                    with open(log_path, "r", encoding=enc) as f:
                        log_content_before = f.read()
                    break
                except:
                    continue
    except:
        pass

    # 发送命令
    cmd = command.strip()
    if not cmd.endswith("\n"):
        cmd += "\n"
    doc.SendCommand(cmd)

    # 等待命令完成
    for _ in range(300):
        try:
            if doc.GetVariable("CMDACTIVE") == 0:
                break
        except:
            pass
        time.sleep(0.1)
    time.sleep(0.5)

    # 从日志文件读取新增内容（全文对比）
    output_lines = []
    if log_path and os.path.exists(log_path):
        try:
            log_content_after = None
            for enc in [locale.getpreferredencoding(), "gbk", "gb2312", "utf-8"]:
                try:
                    with open(log_path, "r", encoding=enc) as f:
                        log_content_after = f.read()
                    break
                except:
                    continue
            if log_content_after is not None:
                new_part = log_content_after[len(log_content_before):].strip()
                if new_part:
                    for line in new_part.split("\n"):
                        s = line.strip()
                        if s:
                            output_lines.append(s)
        except:
            pass

    # 恢复日志设置
    if logfile_was is not None:
        try:
            doc.SetVariable("LOGFILEMODE", logfile_was)
        except:
            pass

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
    """创建图块定义，从指定实体生成块，可选择是否插入块参照
    
    Args:
        name: 块名称
        handles: 要纳入块的实体 handle 列表，如 ["377", "378", "379"]
        base_point: 块基点 [x, y, z]，默认 [0,0,0]
        insert: 是否在创建后插入块参照，默认 True
    """
    import time, os, locale, json
    acad = win32com.client.Dispatch("AutoCAD.Application.25")
    if acad.Documents.Count == 0:
        return json.dumps({"success": False, "error": "无打开的文档", "data": None}, ensure_ascii=False)

    doc = acad.ActiveDocument
    bp = base_point or [0.0, 0.0, 0.0]
    if len(bp) < 3:
        bp = [bp[0], bp[1], 0.0]

    for _ in range(300):
        try:
            if doc.GetVariable("CMDACTIVE") == 0:
                break
        except:
            pass
        time.sleep(0.1)

    # 构建 LISP：遍历 handle，用 entmake 复制到块定义
    handle_list = " ".join(f'"{h}"' for h in handles)
    lisp_code = f"""
(progn
  (setq _name "{name}")
  (setq _bp (list {bp[0]} {bp[1]} {bp[2]}))
  (setq _handles (list {handle_list}))
  
  ;; 检查块名是否已存在
  (if (tblsearch "BLOCK" _name)
    (setvar "USERS1" (strcat "块 \\"" _name "\\" 已存在"))
    (progn
      ;; 开始块定义
      (entmake (list (cons 0 "BLOCK") (cons 2 _name) (cons 70 2) (cons 10 _bp)))
      
      ;; 遍历 handle，将实体 DXF 数据复制到块中
      (foreach _h _handles
        (setq _ent (handent _h))
        (if _ent
          (progn
            (setq _ed (entget _ent))
            ;; 移除 handle、所有者等内部字段（保留图层、颜色等显示属性）
            (foreach _x '(-1 5 330 360 -2)
              (setq _ed (vl-remove (assoc _x _ed) _ed))
            )
            (entmake _ed)
          )
        )
      )
      
      ;; 结束块定义
      (entmake (list (cons 0 "ENDBLK")))
      
      ;; 插入块参照
      {f'(entmake (list (cons 0 "INSERT") (cons 2 "{name}") (cons 10 (list {bp[0]} {bp[1]} {bp[2]}))))' if insert else 'nil'}
      
      (setvar "USERS1" (strcat "块 \\"" _name "\\" 已创建，包含 " (itoa (length _handles)) " 个实体"))
    )
  )
)
""".strip()

    # 执行 LISP
    log_path = None
    logfile_was = None
    log_content_before = ""
    try:
        logfile_was = doc.GetVariable("LOGFILEMODE")
        doc.SetVariable("LOGFILEMODE", 1)
        time.sleep(0.5)
        log_path = str(doc.GetVariable("LOGFILENAME"))
        if os.path.exists(log_path):
            for enc in [locale.getpreferredencoding(), "gbk", "gb2312", "utf-8"]:
                try:
                    with open(log_path, "r", encoding=enc) as f:
                        log_content_before = f.read()
                    break
                except:
                    continue
    except:
        pass

    wrapped = f'(progn (setvar "USERS1" (vl-princ-to-string (progn {lisp_code}))))\n'
    doc.SendCommand(wrapped)

    for _ in range(300):
        try:
            if doc.GetVariable("CMDACTIVE") == 0:
                break
        except:
            pass
        time.sleep(0.1)
    time.sleep(0.3)

    result = ""
    try:
        result = str(doc.GetVariable("USERS1"))
    except:
        pass

    output_lines = []
    if log_path and os.path.exists(log_path):
        try:
            log_content_after = None
            for enc in [locale.getpreferredencoding(), "gbk", "gb2312", "utf-8"]:
                try:
                    with open(log_path, "r", encoding=enc) as f:
                        log_content_after = f.read()
                    break
                except:
                    continue
            if log_content_after is not None:
                new_part = log_content_after[len(log_content_before):].strip()
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
    """在AutoCAD中通过ssget选择对象，支持多种模式和过滤条件，返回handle列表（自动避免交互弹窗）
    
    Args:
        filter_list: 可选的过滤条件，LISP关联列表格式，如 '((0 . "LINE")(8 . "0"))'，不传则选择所有
        mode: 选择模式 — "X"=全部(默认), "A"=当前空间全部, "L"=最后创建的, "P"=前一个选集, "I"=当前pickfirst, "C"=窗交(points需4个坐标), "W"=窗选, "F"=栏选, "CP"/ "WP"=多边形
        points: 坐标列表，用于"C"/"W"/"F"/"CP"/"WP"模式，eg [x1,y1,x2,y2,...]
    """
    import time, os, locale
    acad = win32com.client.Dispatch("AutoCAD.Application.25")
    if acad.Documents.Count == 0:
        return "无打开的文档"

    doc = acad.ActiveDocument

    for _ in range(300):
        try:
            if doc.GetVariable("CMDACTIVE") == 0:
                break
        except:
            pass
        time.sleep(0.1)

    # 构建 ssget 调用
    mode_upper = (mode or "X").upper()
    
    # 无交互模式：不需要用户操作
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
        no_pick_modes.add("X")
        mode_upper = "X"
        ss_args = '"X"'

    if filter_list:
        filter_str = filter_list.strip()
        if filter_str.startswith("'"):
            filter_str = filter_str[1:]
        ss_lisp = f'(ssget {ss_args} \'{filter_str})'
    else:
        ss_lisp = f'(ssget {ss_args})'

    # 构建完整的 LISP：提取 handle 列表
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

    # 执行 LISP（复用 eval_lisp 逻辑）
    log_path = None
    logfile_was = None
    log_content_before = ""
    try:
        logfile_was = doc.GetVariable("LOGFILEMODE")
        doc.SetVariable("LOGFILEMODE", 1)
        time.sleep(0.5)
        log_path = str(doc.GetVariable("LOGFILENAME"))
        if os.path.exists(log_path):
            for enc in [locale.getpreferredencoding(), "gbk", "gb2312", "utf-8"]:
                try:
                    with open(log_path, "r", encoding=enc) as f:
                        log_content_before = f.read()
                    break
                except:
                    continue
    except:
        pass

    wrapped = f'(progn (setvar "USERS1" (vl-princ-to-string (progn {lisp_code}))))\n'
    doc.SendCommand(wrapped)

    for _ in range(300):
        try:
            if doc.GetVariable("CMDACTIVE") == 0:
                break
        except:
            pass
        time.sleep(0.1)
    time.sleep(0.3)

    result = ""
    try:
        result = str(doc.GetVariable("USERS1"))
    except:
        pass

    output_lines = []
    if log_path and os.path.exists(log_path):
        try:
            log_content_after = None
            for enc in [locale.getpreferredencoding(), "gbk", "gb2312", "utf-8"]:
                try:
                    with open(log_path, "r", encoding=enc) as f:
                        log_content_after = f.read()
                    break
                except:
                    continue
            if log_content_after is not None:
                new_part = log_content_after[len(log_content_before):].strip()
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
def autocad_get_layers(ctx: Context) -> str:
    """获取图纸中所有图层信息（名称、颜色、状态等）"""
    import json
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application.25")
        if acad.Documents.Count == 0:
            return json.dumps({"success": False, "error": "无打开的文档", "data": None}, ensure_ascii=False)
        doc = acad.ActiveDocument
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
    import json
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application.25")
        if acad.Documents.Count == 0:
            return json.dumps({"success": False, "error": "无打开的文档", "data": None}, ensure_ascii=False)
        doc = acad.ActiveDocument
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
    import json
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application.25")
        if acad.Documents.Count == 0:
            return json.dumps({"success": False, "error": "无打开的文档", "data": None}, ensure_ascii=False)
        doc = acad.ActiveDocument
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

@mcp.tool()
def autocad_get_sysvar(ctx: Context, name: str) -> str:
    """读取AutoCAD系统变量
    
    Args:
        name: 系统变量名，如 "CMDACTIVE"、"DWGNAME"、"LASTPROMPT"
    """
    import json
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application.25")
        if acad.Documents.Count == 0:
            return json.dumps({"success": False, "error": "无打开的文档", "data": None}, ensure_ascii=False)
        doc = acad.ActiveDocument
        value = doc.GetVariable(name)
        return json.dumps({"success": True, "error": None, "data": str(value)}, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"success": False, "error": str(e), "data": None}, ensure_ascii=False)

@mcp.tool()
def autocad_set_sysvar(ctx: Context, name: str, value: Any) -> str:
    """设置AutoCAD系统变量
    
    Args:
        name: 系统变量名，如 "CMDECHO"、"LOGFILEMODE"
        value: 值（整数、浮点数或字符串）
    """
    import json
    try:
        acad = win32com.client.Dispatch("AutoCAD.Application.25")
        if acad.Documents.Count == 0:
            return json.dumps({"success": False, "error": "无打开的文档", "data": None}, ensure_ascii=False)
        doc = acad.ActiveDocument
        doc.SetVariable(name, value)
        return json.dumps({"success": True, "error": None, "data": f"{name} = {value}"}, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"success": False, "error": str(e), "data": None}, ensure_ascii=False)

@mcp.tool()
def autocad_eval_lisp(ctx: Context, lisp_code: str) -> str:
    """在AutoCAD中执行LISP代码，原子化执行（解决entsel/getpoint在while循环中代码泄漏问题），返回完整结果
    
    Args:
        lisp_code: LISP表达式，如 '(setq a (+ 1 2))' 或带 entsel 的循环
    """
    import time, os, locale
    acad = win32com.client.Dispatch("AutoCAD.Application.25")
    if acad.Documents.Count == 0:
        return "无打开的文档"

    doc = acad.ActiveDocument

    # 等待之前的命令完成
    for _ in range(300):
        try:
            if doc.GetVariable("CMDACTIVE") == 0:
                break
        except:
            pass
        time.sleep(0.1)

    # 启用日志文件捕获完整交互输出
    log_path = None
    logfile_was = None
    log_content_before = ""
    try:
        logfile_was = doc.GetVariable("LOGFILEMODE")
        doc.SetVariable("LOGFILEMODE", 1)
        time.sleep(0.5)
        log_path = str(doc.GetVariable("LOGFILENAME"))
        if os.path.exists(log_path):
            for enc in [locale.getpreferredencoding(), "gbk", "gb2312", "utf-8"]:
                try:
                    with open(log_path, "r", encoding=enc) as f:
                        log_content_before = f.read()
                    break
                except:
                    continue
    except:
        pass

    # 用 (progn ...) 原子化包裹代码，结果存入 USERS1
    # 使用 vl-princ-to-string 转换任意 LISP 类型为字符串
    wrapped = f'(progn (setvar "USERS1" (vl-princ-to-string (progn {lisp_code}))))\n'

    doc.SendCommand(wrapped)

    # 等待命令完成（交互式命令如 entsel 需要用户操作，使用更长超时）
    for _ in range(3000):
        try:
            if doc.GetVariable("CMDACTIVE") == 0:
                break
        except:
            pass
        time.sleep(0.1)

    # 从日志文件读取新增内容
    output_lines = []
    if log_path and os.path.exists(log_path):
        try:
            log_content_after = None
            for enc in [locale.getpreferredencoding(), "gbk", "gb2312", "utf-8"]:
                try:
                    with open(log_path, "r", encoding=enc) as f:
                        log_content_after = f.read()
                    break
                except:
                    continue
            if log_content_after is not None:
                new_part = log_content_after[len(log_content_before):].strip()
                if new_part:
                    for line in new_part.split("\n"):
                        s = line.strip()
                        if s:
                            output_lines.append(s)
        except:
            pass

    # 恢复日志设置
    if logfile_was is not None:
        try:
            doc.SetVariable("LOGFILEMODE", logfile_was)
        except:
            pass

    # 读取 USERS1 中的返回结果
    result = ""
    try:
        result = str(doc.GetVariable("USERS1"))
    except:
        pass

    # 构建返回
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

# 启动服务器
if __name__ == "__main__":
    mcp.run()