# 万能 PDF 转换器

支持 Word/Excel/PPT/图片/CAD 批量转换为 PDF 的桌面工具。

## 功能特性

- 📄 Word/Excel/PPT 转 PDF
- 🖼️ 图片 (JPG/PNG/BMP) 转 PDF
- 📐 CAD (DWG/DXF) 转 PDF
- 🖱️ 原生拖拽支持
- ⚙️ 可调节边框边距（毫米单位）

## 前置依赖

- Windows 10/11
- AutoCAD 2020+（用于转换 DWG/DXF 文件）
- Microsoft Office 2016+（用于转换 Word/Excel/PPT）

## 工作原理

本软件通过以下方式自动调用：
- **AutoCAD**: 自动搜索系统注册表或常见安装目录查找 `accoreconsole.exe`
- **Office**: 通过 Windows COM 接口自动调用（无需指定路径）

## 使用方法

1. 下载 release 中的 `.exe` 文件
2. 双击运行
3. 选择要转换的文件或文件夹
4. 设置输出目录
5. 点击"开始执行转换"

## 开发相关

### 安装依赖
```bash
pip install customtkinter Pillow comtypes pymupdf
```

### 打包 EXE
```bash
pyinstaller 万能PDF转换器.spec
```

## 许可证

本项目基于 AGPLv3 许可证开源。
