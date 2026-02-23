# IFCPropGetter

IFCPropGetter 是一个轻量级桌面工具，用于从 IFC 文件中批量提取指定的属性集/属性值，并导出为格式化的 Excel 或 CSV 文件。提供直观的图形界面，支持属性列表的动态增删改查、文件大小显示、实时日志等功能。

## 📌 核心用途

- 快速提取 IFC 构件中的自定义属性（如 `Assembly/Cast unit Mark` 等）
- 自动跳过不需要的实体类型（如 `IfcProject`, `IfcSpace`, `IfcAnnotation`）
- 支持导出带样式的 Excel（字体、对齐、边框、列宽自动适配）或纯 CSV
- 可选择性包含 `GlobalId` 和 `Name` 字段
- 实时显示处理进度和日志，支持取消任务

## 🔧 环境依赖与安装

### 系统要求
- Python 3.7 或更高版本

### 依赖库
项目依赖以下 Python 包：
```
ifcopenshell
pandas
openpyxl
customtkinter
```
推荐使用 `pip` 一次性安装：
```bash
pip install ifcopenshell pandas openpyxl customtkinter
```
> 注意：`ifcopenshell` 在某些平台可能需要从 [官方渠道](https://blenderbim.org/download-ifcopenshell.html) 下载对应的 wheel 文件安装。

## 🚀 快速启动

1. 克隆或下载本项目到本地
2. 安装上述依赖
3. 在项目根目录下运行：
   ```bash
   python run.py
   ```
4. 程序主界面随即打开，即可开始使用

## 📁 项目目录结构

```
IFCPropGetter/
├── ifc_prop_getter/          # 核心模块包
│   ├── __init__.py
│   ├── constants.py          # 全局常量（默认属性、跳过实体类型等）
│   ├── extractor.py          # IFC 属性提取逻辑（线程任务）
│   ├── gui.py                 # 图形界面（customtkinter）
│   ├── main.py                # 程序入口
│   └── utils.py               # 工具函数（时间戳、文件名清理、Excel 样式等）
├── resources/                 # 资源文件
│   └── IFCPropGetter.ico      # 程序图标
└── run.py                     # 启动脚本
```

## 🧩 核心功能使用示例

### 1. 选择 IFC 文件
- 点击主界面 **“浏览”** 按钮，选择一个 `.ifc` 文件
- 文件路径下方会显示文件大小

### 2. 管理提取属性
- 在 **“属性名称”** 输入框中键入属性名（例如 `Assembly/Cast unit Mark`）
- 点击 **“添加”** 将属性加入列表
- 可使用 **上移/下移** 调整顺序，**删除选中** 移除条目，**清空列表** 一键清除
- 勾选 **包含 GlobalId** 和 **包含 Name** 决定是否输出这两个系统字段

### 3. 设置输出选项
- **输出文件名**：默认取 IFC 文件名后缀 `_data`，可手动修改
- **输出文件夹**：默认为桌面，可点击 **“浏览”** 更改
- **输出格式**：选择 Excel (`.xlsx`) 或 CSV (`.csv`)

### 4. 执行提取
- 点击 **“开始提取”**，弹出进度窗口
- 程序会在后台扫描 IFC 文件，提取所有包含指定属性的构件
- 日志区实时显示处理进度和警告信息
- 完成后自动弹出成功提示，并显示输出文件路径

### 5. 取消任务
- 处理过程中可点击进度窗口的 **“取消”** 按钮终止任务

## 📝 注意事项
- 属性名支持点号分隔的格式 `属性集.属性名`（例如 `Pset_WallCommon.Reference`），提高提取精准度
- Excel 输出会自动应用样式：标题行加粗、灰色背景，内容居中对齐，列宽自动调整（`GlobalId` 列 32，其余列 24）
- 支持的大型 IFC 文件可能耗时较长，请耐心等待

## 📄 许可证
[MIT](LICENSE)

---

如有问题或建议，欢迎提交 [Issue](https://github.com/your-repo/issues) 或 Pull Request。
