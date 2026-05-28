from mcp.server.fastmcp import Context
from typing import Optional, List
import json

from tools.com_helpers import get_acad, make_variant
from tools.database import db_connect


def register_tools(mcp):
    @mcp.tool()
    def scan_all_entities(ctx: Context) -> str:
        """扫描当前图纸中的所有实体并保存到数据库"""
        acad, doc, err = get_acad()
        if err:
            return err
        try:
            model_space = doc.ModelSpace
            with db_connect() as conn:
                cursor = conn.cursor()
                cursor.execute("DELETE FROM cad_elements")
                count = 0
                entity_types = {}
                for i in range(model_space.Count):
                    try:
                        entity = model_space.Item(i)
                        entity_type = entity.ObjectName
                        entity_types[entity_type] = entity_types.get(entity_type, 0) + 1

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
                        elif entity_type in ("AcDbText", "AcDbMText"):
                            properties = {
                                "text": entity.TextString,
                                "position": [entity.InsertionPoint[0], entity.InsertionPoint[1], entity.InsertionPoint[2]] if hasattr(entity, "InsertionPoint") else None,
                                "height": entity.Height if hasattr(entity, "Height") else None
                            }

                        cursor.execute(
                            "INSERT OR REPLACE INTO cad_elements (handle, name, type, layer, properties) VALUES (?, ?, ?, ?, ?)",
                            (entity.Handle, entity.ObjectName.replace("AcDb", ""), entity.ObjectName, entity.Layer, json.dumps(properties))
                        )
                        count += 1
                    except Exception as e:
                        print(f"处理实体 {i} 时出错: {str(e)}")
                conn.commit()
            type_summary = "\n".join([f"{t}: {c}" for t, c in entity_types.items()])
            return f"已扫描并保存 {count} 个实体到数据库。\n\n实体类型统计:\n{type_summary}"
        except Exception as e:
            return f"扫描实体失败: {str(e)}"

    @mcp.tool()
    def highlight_entity(ctx: Context, handle: str, color: int = 1) -> str:
        """通过Handle在AutoCAD中高亮显示指定实体"""
        acad, doc, err = get_acad()
        if err:
            return err
        try:
            try:
                entity = doc.HandleToObject(handle)
            except:
                return f"未找到Handle为 {handle} 的实体"
            original_color = entity.Color
            entity.Color = color
            return f"已高亮实体 {handle}，颜色从 {original_color} 改为 {color}"
        except Exception as e:
            return f"高亮实体失败: {str(e)}"

    @mcp.tool()
    def count_text_patterns(ctx: Context, pattern: str = "PMC-3M") -> str:
        """统计图纸中文本实体中特定模式的出现次数"""
        acad, doc, err = get_acad()
        if err:
            return err
        try:
            model_space = doc.ModelSpace
            drawing_name = doc.Name
            count = 0
            matching_entities = []
            for i in range(model_space.Count):
                try:
                    entity = model_space.Item(i)
                    if hasattr(entity, "TextString") and pattern in entity.TextString:
                        count += 1
                        matching_entities.append({
                            "handle": entity.Handle,
                            "text": entity.TextString,
                            "layer": entity.Layer,
                            "position": [entity.InsertionPoint[0], entity.InsertionPoint[1]] if hasattr(entity, "InsertionPoint") else None
                        })
                except Exception as e:
                    print(f"处理文本实体 {i} 时出错: {str(e)}")

            with db_connect() as conn:
                cursor = conn.cursor()
                cursor.execute(
                    "INSERT OR REPLACE INTO text_patterns (pattern, count, drawing) VALUES (?, ?, ?)",
                    (pattern, count, drawing_name)
                )
                conn.commit()

            result = f"在图纸 '{drawing_name}' 中找到 {count} 处匹配模式 '{pattern}' 的文本。"
            if count > 0:
                details = "\n\n匹配详情："
                for i, match in enumerate(matching_entities[:10]):
                    details += f"\n{i+1}. 文本: '{match['text']}', 图层: {match['layer']}, Handle: {match['handle']}"
                if len(matching_entities) > 10:
                    details += f"\n... 以及其他 {len(matching_entities) - 10} 个匹配项"
                result += details
            return result
        except Exception as e:
            return f"统计文本模式失败: {str(e)}"

    @mcp.tool()
    def highlight_text_matches(ctx: Context, pattern: str = "PMC-3M", color: int = 1) -> str:
        """高亮显示包含指定文本模式的所有文本实体"""
        acad, doc, err = get_acad()
        if err:
            return err
        try:
            model_space = doc.ModelSpace
            try:
                doc.SelectionSets.Item("TextMatches").Delete()
            except:
                pass
            selection = doc.SelectionSets.Add("TextMatches")
            count = 0
            for i in range(model_space.Count):
                try:
                    entity = model_space.Item(i)
                    if hasattr(entity, "TextString") and pattern in entity.TextString:
                        entity.Color = color
                        selection.AddItems([entity])
                        count += 1
                except Exception as e:
                    print(f"处理文本实体 {i} 时出错: {str(e)}")
            if count > 0:
                doc.ActiveView.ZoomAll()
                return f"已高亮显示 {count} 个包含 '{pattern}' 的文本实体"
            else:
                selection.Delete()
                return f"未找到包含 '{pattern}' 的文本实体"
        except Exception as e:
            return f"高亮文本匹配失败: {str(e)}"

    @mcp.tool()
    def move_entity(ctx: Context, handle: str, start_point: List[float], end_point: List[float]) -> str:
        """移动实体"""
        acad, doc, err = get_acad()
        if err:
            return err
        try:
            try:
                entity = doc.HandleToObject(handle)
            except:
                return f"未找到Handle为 {handle} 的实体"
            entity.Move(make_variant(start_point), make_variant(end_point))
            return f"已移动实体 {handle}"
        except Exception as e:
            return f"移动实体失败: {str(e)}"

    @mcp.tool()
    def rotate_entity(ctx: Context, handle: str, base_point: List[float], angle: float) -> str:
        """旋转实体"""
        acad, doc, err = get_acad()
        if err:
            return err
        try:
            try:
                entity = doc.HandleToObject(handle)
            except:
                return f"未找到Handle为 {handle} 的实体"
            import math
            entity.Rotate(make_variant(base_point), angle * math.pi / 180)
            return f"已旋转实体 {handle} {angle} 度"
        except Exception as e:
            return f"旋转实体失败: {str(e)}"

    @mcp.tool()
    def copy_entity(ctx: Context, handle: str, start_point: List[float], end_point: List[float]) -> str:
        """复制实体"""
        acad, doc, err = get_acad()
        if err:
            return err
        try:
            try:
                entity = doc.HandleToObject(handle)
            except:
                return f"未找到Handle为 {handle} 的实体"
            new_entity = entity.Copy()
            new_entity.Move(make_variant(start_point), make_variant(end_point))
            return f"已复制实体，新Handle: {new_entity.Handle}"
        except Exception as e:
            return f"复制实体失败: {str(e)}"

    @mcp.tool()
    def get_all_tables(ctx: Context) -> str:
        """获取数据库中的所有表"""
        try:
            with db_connect() as conn:
                cursor = conn.cursor()
                cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
                tables = cursor.fetchall()
            table_list = [table[0] for table in tables]
            return json.dumps(table_list, indent=2)
        except Exception as e:
            return f"获取表列表失败: {str(e)}"

    @mcp.tool()
    def get_table_schema(ctx: Context, table_name: str) -> str:
        """获取指定表的结构信息"""
        try:
            with db_connect() as conn:
                cursor = conn.cursor()
                cursor.execute(f"PRAGMA table_info({table_name});")
                columns = cursor.fetchall()
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
            with db_connect() as conn:
                cursor = conn.cursor()
                cursor.execute(query)
                if query.strip().upper().startswith("SELECT"):
                    columns = [desc[0] for desc in cursor.description]
                    rows = cursor.fetchall()
                    result = []
                    for row in rows:
                        result.append(dict(zip(columns, row)))
                    conn.commit()
                    return json.dumps(result, indent=2)
                else:
                    conn.commit()
                    return f"执行成功，影响了 {cursor.rowcount} 行"
        except Exception as e:
            return f"执行查询失败: {str(e)}"

    @mcp.tool()
    def query_and_highlight(ctx: Context, sql_query: str, highlight_color: int = 1) -> str:
        """根据SQL查询结果高亮显示AutoCAD实体"""
        try:
            with db_connect() as conn:
                cursor = conn.cursor()
                cursor.execute(sql_query)
                rows = cursor.fetchall()
                column_names = [description[0] for description in cursor.description]

            if not rows:
                return "查询未返回任何结果"

            handle_index = -1
            for i, name in enumerate(column_names):
                if name.lower() == 'handle':
                    handle_index = i
                    break
            if handle_index == -1:
                return "查询结果中未找到handle列"

            handles = [row[handle_index] for row in rows]

            acad, doc, err = get_acad()
            if err:
                return err

            try:
                doc.SelectionSets.Item("QueryResults").Delete()
            except:
                pass
            selection = doc.SelectionSets.Add("QueryResults")

            highlighted_count = 0
            entities_to_add = []
            for handle in handles:
                try:
                    entity = doc.HandleToObject(handle)
                    entity.Color = highlight_color
                    entities_to_add.append(entity)
                    highlighted_count += 1
                except Exception as e:
                    print(f"处理实体 {handle} 时出错: {str(e)}")

            if entities_to_add:
                try:
                    selection.AddItems(entities_to_add)
                except:
                    pass

            if highlighted_count > 0:
                doc.ActiveView.ZoomAll()
                return f"已高亮显示 {highlighted_count} 个实体（共 {len(handles)} 个结果）"
            else:
                return "未能高亮任何实体"
        except Exception as e:
            return f"查询并高亮失败: {str(e)}"
