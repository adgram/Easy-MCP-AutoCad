# AutoCAD MCP Server

基于 **Model Context Protocol (MCP)** 的 AutoCAD 集成服务器，允许通过 AI 与 AutoCAD 进行自然语言交互。

## 功能特点

- 自然语言控制 AutoCAD 创建和修改图纸
- 基础绘图（直线、圆、多段线、矩形、文字）
- 图层管理（创建、修改）
- 图纸元素扫描与分析
- 文本模式查询与高亮（如 `PMC-3M`）
- 实体操作（移动、旋转、复制）
- 执行 AutoCAD 命令并捕获命令行输出
- 内置 SQLite 数据库，支持 CAD 元素存储和 SQL 查询

## 系统要求

- Python 3.10+
- AutoCAD 2018+（需支持 COM 接口）
- Windows 操作系统

## 安装

```bash
pip install -r requirements.txt
```

## 使用

```bash
python server.py
```

### 集成 opencode

编辑 `~/.config/opencode/opencode.jsonc`：

```json
{
  "mcp": {
    "autocad": {
      "type": "local",
      "command": ["path/to/venv/Scripts/python.exe", "path/to/server.py"],
      "enabled": true
    }
  }
}
```

## 可用工具

| 功能 | 说明 |
|------|------|
| `create_new_drawing` | 创建新的 AutoCAD 图纸 |
| `draw_line` | 画直线 |
| `draw_circle` | 画圆 |
| `draw_polyline` | 画多段线 |
| `draw_rectangle` | 画矩形 |
| `draw_text` | 添加文字 |
| `draw_device_connection` | 绘制设备连接线 |
| `create_block` | 从指定实体 handle 列表创建图块定义，可选插入块参照 |
| `create_layer` | 创建或修改图层 |
| `execute_command` | 执行 AutoCAD 命令并捕获命令行输出 |
| `autocad_eval_lisp` | 原子化执行 LISP 代码（解决 `entsel`/`getpoint` 在循环中代码泄漏），返回完整结果 |
| `autocad_get_sysvar` | 读取系统变量，返回 `{"success": bool, "data": value, "error": str}` |
| `autocad_set_sysvar` | 设置系统变量，返回 `{"success": bool, "data": str, "error": str}` |
| `autocad_get_layers` | 获取图纸所有图层信息（名称、颜色、开关、冻结、锁定） |
| `autocad_get_blocks` | 获取图纸所有块定义信息（名称、实体数、基点） |
| `autocad_get_linetypes` | 获取图纸所有线型信息（名称、描述） |
| `autocad_select_objects` | 通过 `ssget` 选择对象，支持 `X/A/L/P/I/C/W/F/CP/WP` 多种模式 + 过滤条件，返回 handle 列表（自动避免交互弹窗） |
| `move_entity` | 移动实体 |
| `rotate_entity` | 旋转实体 |
| `copy_entity` | 复制实体 |
| `highlight_entity` | 高亮显示实体 |
| `highlight_text_matches` | 高亮匹配文本 |
| `scan_all_entities` | 扫描图纸全部实体 |
| `count_text_patterns` | 统计文本模式 |
| `execute_query` | 执行 SQL 查询 |
| `query_and_highlight` | 根据 SQL 查询结果高亮实体 |
| `get_all_tables` | 获取数据库表列表 |
| `get_table_schema` | 获取表结构 |

## 架构原理

### 工作流程

```
AI -- 自然语言 --> MCP 客户端 -- JSON-RPC --> MCP 服务器(server.py)
                                                   |
                                          COM 自动化调用
                                                   |
                                              AutoCAD
```

1. **指令解析**：AI 分析意图并匹配到 MCP 工具
2. **工具调用**：AI 向 `server.py` 发送 JSON-RPC 请求（工具名 + 参数）
3. **驱动 CAD**：`server.py` 通过 Windows COM 接口操作 AutoCAD
4. **即时反馈**：AutoCAD 返回执行结果，服务器反馈给 AI

### 技术栈

| 组件 | 作用 |
|:-----|:-----|
| **Python** | 核心编程语言 |
| **mcp (FastMCP)** | MCP 协议 SDK，将 Python 函数转换为 AI 可调用的工具 |
| **pywin32 (win32com)** | Windows COM 桥梁，操作 AutoCAD |
| **SQLite** | 本地数据库，存储 CAD 实体数据 |

### AutoCAD 对象模型

AutoCAD 的 COM 接口采用层级结构：

1. **Application** — 整个 AutoCAD 程序
2. **ActiveDocument** — 当前编辑的图纸
3. **ModelSpace** — 模型空间（绘图区域）
4. **Entity** — 具体实体（直线、圆、文字等）

```python
acad = win32com.client.Dispatch("AutoCAD.Application.25")
doc = acad.ActiveDocument
model = doc.ModelSpace
line = model.AddLine(start_point, end_point)
line.Move(base, target)
```

### 命令执行原理

`execute_command` 通过 AutoCAD 日志文件捕获完整命令行输出：

1. 开启 `LOGFILEMODE`，读取日志基线
2. 通过 `SendCommand` 发送命令
3. 等待 `CMDACTIVE` 归零（命令执行完毕）
4. 再次读取日志，diff 增量即为输出
5. 恢复 `LOGFILEMODE` 原值

### 图块创建原理

`create_block` 通过 LISP `entmake` 创建块定义：

1. `(entmake '((0 . "BLOCK") (2 . name) (70 . 2) (10 . base_point)))` — 开始块定义
2. 遍历 handle 列表，用 `handent` 获取实体名，`entget` 获取 DXF 数据，移除 handle/owner 后 `entmake` 到块中
3. `(entmake '((0 . "ENDBLK")))` — 结束块定义
4. 可选：`(entmake '((0 . "INSERT") ...))` — 插入块参照

保留原始实体的图层、颜色、线型等属性，仅剥离内部 handle/owner 字段。返回结构化 JSON。

### 系统变量与字典查询

`autocad_get_sysvar` / `autocad_set_sysvar` 直接通过 COM 的 `GetVariable`/`SetVariable` 操作 AutoCAD 系统变量，不经过命令行。返回结构化 JSON：

```json
{"success": true, "data": "1", "error": null}
```

`autocad_get_layers` / `autocad_get_blocks` / `autocad_get_linetypes` 通过 COM 对象模型遍历图纸字典，返回完整的结构化信息，无需解析命令行输出。`get_layers` 返回图层名称、颜色、开关/冻结/锁定状态；`get_blocks` 返回块名称、包含实体数、原点坐标；`get_linetypes` 返回线型名称和描述。

### 对象安全选择原理

`autocad_select_objects` 解决 `(ssget)` 无参数弹窗挂起问题：

- **问题**：`(ssget)` 无参数时弹出交互选择框，MCP 服务器无法进行 UI 操作导致无限挂起。
- **解决**：提供多种无交互选择模式：
  - `X` — 全库选择（默认）
  - `A` — 当前空间全部对象
  - `L` — 最后创建的实体
  - `P` — 前一个选集
  - `I` — 当前 Pickfirst 选中的实体
  - `C`/`W` — 窗交/窗选（需 4 个坐标 `[x1,y1,x2,y2]`）
  - `F` — 栏选（需 2n 个坐标）
  - `CP`/`WP` — 多边形窗交/窗选（需 2n 个坐标）
- **过滤**：支持传入 DXF 过滤列表（如 `'((0 . "LINE")(8 . "0"))`），不传则选择全部对象。
- **handle 提取**：通过 `entget` + `assoc 5` 提取每个对象的 handle，返回 handle 字符串列表。

### LISP 原子化执行原理

`autocad_eval_lisp` 解决 `SendCommand` 逐行 feed 导致的代码泄漏问题：

- **问题**：`SendCommand` 将文本逐行发送到命令行。遇到 `entsel`/`getpoint` 时 AutoCAD 等待用户输入，后续代码会作为这些函数的输入参数被消费，而非作为代码执行。
- **解决**：将代码包裹在 `(progn ...)` 中作为一个表达式整体发送，AutoCAD 先完整解析整个表达式再开始执行，交互函数暂停时后续代码不会泄漏。
- **结果捕获**：使用 `(vl-princ-to-string ...)` 将任意 LISP 返回值转为字符串存入 `USERS1` 系统变量，同时通过日志文件捕获完整交互过程。超时设为 5 分钟以支持交互式操作。

### 扩展方案对比

| 方案 | 优点 | 缺点 |
|:----|:-----|:-----|
| **Python + COM + MCP**（本方案） | 开发快，交互好，支持读写 | 依赖 Windows + AutoCAD |
| **生成 AutoLISP** | 无需 MCP 服务器 | 单向交互，无法读取图纸 |
| **生成 DXF 文件** | 脱离 AutoCAD | 无法修改已有 DWG |
| **.NET API + MCP** | 性能最强，功能最全 | 开发难度大，需编译 |
| **Python + ezdxf** | 无需 AutoCAD | 有限 DWG 支持，无交互 |
