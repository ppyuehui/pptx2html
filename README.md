# PPTX 转 HTML 工具

将 PowerPoint (.pptx) 转换为单个 HTML 文件，布局完全还原。

## 功能

- **布局还原**：与 PPT 完全一致的显示效果
- **WebP 压缩**：图片自动压缩，体积减小 30-50%
- **单文件输出**：所有资源（图片、视频）嵌入 HTML，无需额外文件
- **视频支持**：自动提取 PPT 中的视频并在 HTML 中播放
- **激光笔**：支持鼠标轨迹显示，可调整轨迹时间和颜色
- **全屏演示**：支持全屏模式，适合投影演示
- **ffmpeg 集成**：自动下载，视频压缩减小文件体积

## 系统要求

- Windows 10/11
- Microsoft PowerPoint（需要 COM 接口）
- Python 3.8+（源码运行时）

## 安装

### 直接使用 exe（推荐）

1. 从 [Releases](../../releases) 下载最新版 `PPTX转HTML工具.exe`
2. 双击运行，首次使用会自动下载 ffmpeg

### 源码运行

```bash
# 克隆项目
git clone https://github.com/ppyuehui/pptx2html.git
cd pptx2html

# 安装依赖
pip install Pillow jinja2 pywin32

# 运行
python pptx2html.py
```

## 使用方法

1. 打开工具，点击「选择文件」选择 PPTX 文件
2. 设置 WebP 图片质量（默认 80，越大质量越好但体积越大）
3. 点击「开始转换」
4. 转换完成后自动打开生成的 HTML 文件

## 视频处理

- 工具会自动提取 PPT 中嵌入的视频（mp4/avi/wmv/mov）
- 视频会被压缩后嵌入 HTML，支持在演示中直接播放
- 首次使用需要下载 ffmpeg（约 90MB）

## 配置文件

配置文件保存在：

```
%APPDATA%\pptx2html\config.json
```

包含 ffmpeg 保存路径等设置。

## 目录结构

```
pptx2html/
├── pptx2html.py      # 主程序
├── .gitignore
├── README.md
├── build/            # PyInstaller 构建目录（git 忽略）
├── dist/             # 打包输出目录（git 忽略）
└── .venv/            # 虚拟环境（git 忽略）
```

## 打包 exe

```bash
# 安装 PyInstaller
pip install pyinstaller

# 打包
.venv\Scripts\pyinstaller.exe --onefile --windowed --name "PPTX转HTML工具" pptx2html.py
```

打包后的 exe 在 `dist/` 目录下。

## 许可证

MIT License

## 问题反馈

如有问题，请提交 [Issue](../../issues)
