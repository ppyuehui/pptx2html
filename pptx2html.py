import sys
import traceback
import os
import time
import gc
from pathlib import Path
import webbrowser
import comtypes.client
from jinja2 import Template
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import shutil
import base64
import subprocess
from PIL import Image

# 添加异常日志记录
def log_error(message):
    log_path = os.path.join(os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(__file__), "error_log.txt")
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} - {message}\n")

try:
    """
    PPTX to HTML 转换工具 - WebP 优化版 + 激光笔 + 视频提取
    包含：WebP压缩导出、单文件HTML、智能隐藏UI、全屏控制、激光笔轨迹、视频嵌入
    """

    class PPTXToHTMLConverter:
        def __init__(self):
            self.WINDOW_WIDTH = 500
            self.WINDOW_HEIGHT = 700
            self.root = tk.Tk()
            self.root.title("PPTX 转 HTML 工具")
            self.root.geometry(f"{self.WINDOW_WIDTH}x{self.WINDOW_HEIGHT}")
            self.root.resizable(False, False)

            # 新增：先隐藏窗口
            self.root.withdraw()

            # 新增：窗口居中显示
            self.root.update_idletasks()
            width = self.root.winfo_width()
            height = self.root.winfo_height()
            screen_width = self.root.winfo_screenwidth()
            screen_height = self.root.winfo_screenheight()
            x = (screen_width // 2) - (width // 2)
            y = (screen_height // 2) - (height // 2)
            self.root.geometry(f"{self.WINDOW_WIDTH}x{self.WINDOW_HEIGHT}+{x}+{y}")
            
            self.pptx_file = None
            self.output_dir = None
            self.is_converting = False
            self.output_html_path = None
            self._ffmpeg_process = None
            
            self.setup_ui()
            # 新增：设置好后再显示
            self.root.deiconify()
            
        def setup_ui(self):
            main_frame = ttk.Frame(self.root, padding="20")
            main_frame.pack(fill=tk.BOTH, expand=True)
            
            title_label = ttk.Label(main_frame, text="PPTX 转 HTML 工具", font=("Arial", 16, "bold"))
            title_label.pack(pady=(0, 20))
            
            # 文件选择区域
            file_frame = ttk.LabelFrame(main_frame, text="选择文件", padding="10")
            file_frame.pack(fill=tk.X, pady=(0, 10))
            
            self.file_path_var = tk.StringVar()
            file_entry = ttk.Entry(file_frame, textvariable=self.file_path_var, state="readonly")
            file_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
            
            browse_btn = ttk.Button(file_frame, text="浏览...", command=self.browse_file)
            browse_btn.pack(side=tk.RIGHT)
            
            # 输出目录区域
            output_frame = ttk.LabelFrame(main_frame, text="输出目录", padding="10")
            output_frame.pack(fill=tk.X, pady=(0, 10))
            
            self.output_dir_var = tk.StringVar()
            output_entry = ttk.Entry(output_frame, textvariable=self.output_dir_var, state="readonly")
            output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
            
            browse_dir_btn = ttk.Button(output_frame, text="浏览...", command=self.browse_directory)
            browse_dir_btn.pack(side=tk.RIGHT)
            
            # 选项区域
            options_frame = ttk.LabelFrame(main_frame, text="选项", padding="10")
            options_frame.pack(fill=tk.X, pady=(0, 10))
            
            self.delete_images_var = tk.BooleanVar(value=True)
            delete_check = ttk.Checkbutton(options_frame, text="转换完成后删除图片文件夹", variable=self.delete_images_var)
            delete_check.pack(anchor=tk.W)
            
            # WebP 质量设置
            quality_frame = ttk.Frame(options_frame)
            quality_frame.pack(fill=tk.X, pady=(5, 0))
            quality_label = ttk.Label(quality_frame, text="WebP 质量：")
            quality_label.pack(side=tk.LEFT)
            self.webp_quality_var = tk.IntVar(value=80)
            quality_scale = ttk.Scale(quality_frame, from_=50, to=100, variable=self.webp_quality_var, orient=tk.HORIZONTAL)
            quality_scale.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
            self.quality_value_label = ttk.Label(quality_frame, text="80")
            self.quality_value_label.pack(side=tk.LEFT)
            self.webp_quality_var.trace_add("write", lambda *args: self.quality_value_label.config(text=str(self.webp_quality_var.get())))
            
            # 视频处理选项
            video_frame = ttk.Frame(options_frame)
            video_frame.pack(fill=tk.X, pady=(5, 0))
            ttk.Label(video_frame, text="视频处理：").pack(side=tk.LEFT)
            self.video_mode_var = tk.StringVar(value="compress")
            ttk.Radiobutton(video_frame, text="不嵌入", variable=self.video_mode_var, value="none").pack(side=tk.LEFT, padx=(5, 0))
            ttk.Radiobutton(video_frame, text="压缩嵌入", variable=self.video_mode_var, value="compress").pack(side=tk.LEFT, padx=(5, 0))
            ttk.Radiobutton(video_frame, text="原始嵌入", variable=self.video_mode_var, value="original").pack(side=tk.LEFT, padx=(5, 0))
            
            # ffmpeg 状态
            ffmpeg_frame = ttk.LabelFrame(options_frame, text="ffmpeg 视频压缩工具", padding="5")
            ffmpeg_frame.pack(fill=tk.X, pady=(5, 0))
            
            # 状态行
            ffmpeg_status_row = ttk.Frame(ffmpeg_frame)
            ffmpeg_status_row.pack(fill=tk.X)
            ttk.Label(ffmpeg_status_row, text="状态：").pack(side=tk.LEFT)
            self.ffmpeg_status_var = tk.StringVar(value="检测中...")
            self.ffmpeg_status_label = ttk.Label(ffmpeg_status_row, textvariable=self.ffmpeg_status_var, foreground="gray")
            self.ffmpeg_status_label.pack(side=tk.LEFT, padx=(5, 0))
            self.ffmpeg_download_btn = ttk.Button(ffmpeg_status_row, text="下载安装", command=self.start_download_ffmpeg)
            self.ffmpeg_download_btn.pack(side=tk.RIGHT)
            self.ffmpeg_download_btn.pack_forget()
            
            # 保存位置（可自选）
            ffmpeg_dir_row = ttk.Frame(ffmpeg_frame)
            ffmpeg_dir_row.pack(fill=tk.X, pady=(2, 0))
            ttk.Label(ffmpeg_dir_row, text="保存位置：").pack(side=tk.LEFT)
            self.ffmpeg_dir_var = tk.StringVar(value=str(Path(__file__).parent / ".ffmpeg"))
            ffmpeg_dir_entry = ttk.Entry(ffmpeg_dir_row, textvariable=self.ffmpeg_dir_var, state="readonly")
            ffmpeg_dir_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
            ttk.Button(ffmpeg_dir_row, text="选择", command=self.browse_ffmpeg_dir).pack(side=tk.RIGHT)
            
            ttk.Label(ffmpeg_frame, text="下载地址：gyan.dev/ffmpeg/builds (精简版 ~90MB)", foreground="gray").pack(anchor=tk.W)
            
            # 进度条（初始隐藏）
            self.ffmpeg_progress_frame = ttk.Frame(ffmpeg_frame)
            self.ffmpeg_progress_frame.pack(fill=tk.X, pady=(3, 0))
            self.ffmpeg_progress_var = tk.DoubleVar(value=0)
            self.ffmpeg_progress_bar = ttk.Progressbar(self.ffmpeg_progress_frame, variable=self.ffmpeg_progress_var, maximum=100, mode='determinate')
            self.ffmpeg_progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True)
            self.ffmpeg_progress_label = ttk.Label(self.ffmpeg_progress_frame, text="0%", width=10)
            self.ffmpeg_progress_label.pack(side=tk.RIGHT, padx=(5, 0))
            self.ffmpeg_progress_frame.pack_forget()
            
            # 异步检测 ffmpeg
            threading.Thread(target=self.check_ffmpeg_status, daemon=True).start()
            
            # 激光笔设置区域
            laser_frame = ttk.LabelFrame(main_frame, text="激光笔默认设置（可在HTML中调整）", padding="10")
            laser_frame.pack(fill=tk.X, pady=(0, 10))
            
            # 轨迹时间设置
            trace_time_label = ttk.Label(laser_frame, text="轨迹显示时间（秒）：")
            trace_time_label.grid(row=0, column=0, sticky=tk.W, pady=2)
            
            self.trace_time_var = tk.StringVar(value="2")
            trace_time_entry = ttk.Entry(laser_frame, textvariable=self.trace_time_var, width=10)
            trace_time_entry.grid(row=0, column=1, sticky=tk.W, padx=(5, 0), pady=2)
            
            trace_time_hint = ttk.Label(laser_frame, text="(默认值，HTML中可实时调整)")
            trace_time_hint.grid(row=0, column=2, sticky=tk.W, padx=(10, 0), pady=2)
            
            # 轨迹宽度设置
            trace_width_label = ttk.Label(laser_frame, text="轨迹宽度（像素）：")
            trace_width_label.grid(row=1, column=0, sticky=tk.W, pady=2)
            
            self.trace_width_var = tk.StringVar(value="5")
            trace_width_entry = ttk.Entry(laser_frame, textvariable=self.trace_width_var, width=10)
            trace_width_entry.grid(row=1, column=1, sticky=tk.W, padx=(5, 0), pady=2)
            
            trace_width_hint = ttk.Label(laser_frame, text="(默认值，HTML中可实时调整)")
            trace_width_hint.grid(row=1, column=2, sticky=tk.W, padx=(10, 0), pady=2)
            
            # 进度区域
            progress_frame = ttk.Frame(main_frame)
            progress_frame.pack(fill=tk.X, pady=(10, 0))
            
            self.progress_var = tk.StringVar(value="准备就绪")
            progress_label = ttk.Label(progress_frame, textvariable=self.progress_var)
            progress_label.pack(side=tk.LEFT)
            
            self.progress_bar = ttk.Progressbar(progress_frame, mode='determinate', length=300)
            self.progress_bar.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(10, 0))
            
            # 按钮区域
            button_frame = ttk.Frame(main_frame)
            button_frame.pack(fill=tk.X, pady=(20, 0))
            
            self.stop_btn = ttk.Button(button_frame, text="停止", command=self.stop_conversion)
            self.stop_btn.pack(side=tk.RIGHT, padx=(10, 0))
            self.stop_btn.pack_forget()
            
            self.convert_btn = ttk.Button(button_frame, text="开始转换", command=self.start_conversion)
            self.convert_btn.pack(side=tk.RIGHT, padx=(10, 0))
            
            # 状态栏
            self.status_var = tk.StringVar(value="就绪")
            status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
            status_bar.pack(side=tk.BOTTOM, fill=tk.X)
            
        def browse_file(self):
            filename = filedialog.askopenfilename(
                title="选择 PowerPoint 文件",
                filetypes=[
                    ("PowerPoint 文件", "*.pptx"),
                    ("所有文件", "*.*")
                ]
            )
            if filename:
                self.pptx_file = filename
                self.file_path_var.set(filename)
                output_dir = str(Path(filename).parent)
                self.output_dir_var.set(output_dir)
                # 重置状态
                self.output_html_path = None
                self.convert_btn.config(text="开始转换", state=tk.NORMAL)
                self.progress_bar['value'] = 0
                self.progress_var.set("准备就绪")
                self.status_var.set("就绪")
        
        def browse_directory(self):
            directory = filedialog.askdirectory(title="选择输出目录")
            if directory:
                self.output_dir_var.set(directory)
        
        def start_conversion(self):
            if self.convert_btn.cget("text") == "打开文件":
                self.open_output_file()
                return
            
            if not self.pptx_file:
                messagebox.showerror("错误", "请先选择PPTX文件！")
                return
            
            if not self.output_dir_var.get():
                messagebox.showerror("错误", "请先选择输出目录！")
                return
            
            try:
                trace_time = float(self.trace_time_var.get())
                if trace_time <= 0:
                    messagebox.showerror("错误", "轨迹显示时间必须大于0！")
                    return
            except ValueError:
                messagebox.showerror("错误", "请输入有效的数字作为轨迹显示时间！")
                return
            
            try:
                trace_width = float(self.trace_width_var.get())
                if trace_width <= 0:
                    messagebox.showerror("错误", "轨迹宽度必须大于0！")
                    return
            except ValueError:
                messagebox.showerror("错误", "请输入有效的数字作为轨迹宽度！")
                return
            
            self.is_converting = True
            self.convert_btn.pack_forget()
            self.stop_btn.pack(side=tk.RIGHT, padx=(10, 0))
            self.progress_bar['value'] = 0
            self.progress_var.set("正在转换...")
            self.status_var.set("正在转换...")
            
            conversion_thread = threading.Thread(target=self.convert_pptx)
            conversion_thread.daemon = True
            conversion_thread.start()
        
        def stop_conversion(self):
            """停止转换，清理临时文件"""
            self.is_converting = False
            # 杀掉 ffmpeg 进程
            if self._ffmpeg_process and self._ffmpeg_process.poll() is None:
                self._ffmpeg_process.kill()
                self._ffmpeg_process.wait()
                self._ffmpeg_process = None
            # 清理临时文件（稍等一下确保文件句柄释放）
            time.sleep(0.3)
            output_dir = self.output_dir_var.get()
            if output_dir:
                for subdir in ["images", "temp_videos"]:
                    d = Path(output_dir) / subdir
                    if d.exists():
                        try:
                            shutil.rmtree(d)
                        except Exception:
                            pass
                # 删除可能已生成的 HTML
                if self.pptx_file:
                    html_path = Path(output_dir) / f"{Path(self.pptx_file).stem}.html"
                    if html_path.exists():
                        try:
                            html_path.unlink()
                        except Exception:
                            pass
            # 恢复 UI
            self.stop_btn.pack_forget()
            self.convert_btn.pack(side=tk.RIGHT, padx=(10, 0))
            self.convert_btn.config(text="开始转换", state=tk.NORMAL)
            self.progress_bar['value'] = 0
            self.progress_var.set("已停止")
            self.status_var.set("已停止")
        
        def convert_pptx(self):
            try:
                self.root.after(0, lambda: self.progress_var.set("正在导出图片..."))
                images_dir, total_slides = self.export_ppt_to_images(self.pptx_file, self.output_dir_var.get())
                
                # 检查是否被停止
                if not self.is_converting:
                    return
                
                if not images_dir:
                    self.show_error("图片导出失败")
                    return
                
                if total_slides == 0:
                    self.show_error("PPT文件没有幻灯片！")
                    return
                
                self.root.after(0, lambda: self.progress_var.set("正在生成HTML..."))
                self.root.after(0, lambda: self.progress_bar.config(value=90))
                
                title = Path(self.pptx_file).stem
                output_html = self.create_single_html_file(images_dir, self.output_dir_var.get(), title, total_slides)
                
                # 再次检查是否被停止
                if not self.is_converting:
                    return
                
                if output_html:
                    self.output_html_path = str(output_html)
                    
                    if self.delete_images_var.get():
                        self.root.after(0, lambda: self.progress_var.set("正在清理..."))
                        self.root.after(0, lambda: self.progress_bar.config(value=95))
                        if images_dir.exists():
                            shutil.rmtree(images_dir)
                        # 清理视频临时目录
                        temp_videos_dir = Path(self.output_dir_var.get()) / "temp_videos"
                        if temp_videos_dir.exists():
                            shutil.rmtree(temp_videos_dir)
                    
                    self.root.after(0, self.on_conversion_success)
                else:
                    self.show_error("生成HTML失败")
                    
            except Exception as e:
                error_msg = f"转换失败: {str(e)}"
                log_error(f"{error_msg}\n{traceback.format_exc()}")
                self.show_error(error_msg)
            finally:
                self.is_converting = False
        
        def on_conversion_success(self):
            self.progress_var.set("转换完成！")
            self.progress_bar['value'] = 100
            self.status_var.set(f"成功生成: {self.output_html_path}")
            self.stop_btn.pack_forget()
            self.convert_btn.pack(side=tk.RIGHT, padx=(10, 0))
            self.convert_btn.config(text="打开文件", state=tk.NORMAL)
            
            if messagebox.askyesno("完成", "转换完成！\n\n是否立即打开查看？"):
                webbrowser.open(self.output_html_path)
        
        def open_output_file(self):
            if self.output_html_path and Path(self.output_html_path).exists():
                webbrowser.open(self.output_html_path)
            else:
                messagebox.showerror("错误", "HTML文件不存在！")
        
        def export_ppt_to_images(self, pptx_path, output_dir):
            pptx_path = Path(pptx_path)
            output_dir = Path(output_dir)
            output_dir.mkdir(parents=True, exist_ok=True)
            
            images_dir = output_dir / "images"
            images_dir.mkdir(exist_ok=True)
            
            self.powerpoint = None
            self.presentation = None
            
            try:
                self.powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
                self.powerpoint.Visible = True
                
                self.presentation = self.powerpoint.Presentations.Open(str(pptx_path))
                self.powerpoint.WindowState = 2  # ppWindowMinimized - 打开后再最小化
                total_slides = self.presentation.Slides.Count
                
                if total_slides == 0:
                    return images_dir, 0
                
                for i in range(1, total_slides + 1):
                    if not self.is_converting:
                        break
                    
                    progress = int((i / total_slides) * 80)
                    slide_num = i
                    self.root.after(0, lambda p=progress: self.progress_bar.config(value=p))
                    self.root.after(0, lambda s=slide_num, t=total_slides: self.progress_var.set(f"正在导出第 {s}/{t} 页..."))
                    
                    img_path = images_dir / f"slide_{i}.png"
                    
                    self.presentation.Slides[i].Export(
                        str(img_path),
                        "PNG",
                        1920,
                        1080
                    )
                    
                    time.sleep(0.1)
                
                # PNG → WebP 转换
                quality = self.webp_quality_var.get()
                self.root.after(0, lambda: self.progress_var.set("正在转换为WebP..."))
                for i in range(1, total_slides + 1):
                    if not self.is_converting:
                        break
                    png_path = images_dir / f"slide_{i}.png"
                    webp_path = images_dir / f"slide_{i}.webp"
                    if png_path.exists():
                        img = Image.open(png_path)
                        img.save(str(webp_path), "WEBP", quality=quality)
                        img.close()
                        png_path.unlink()  # 删除 PNG 临时文件
                
                return images_dir, total_slides
                
            except Exception as e:
                print(f"❌ 导出失败: {e}")
                traceback.print_exc()
                return None, 0
                
            finally:
                self.close_powerpoint()
                gc.collect()
        
        def close_powerpoint(self):
            """关闭 PowerPoint COM 对象"""
            if self.presentation:
                try:
                    self.presentation.Close()
                except Exception:
                    pass
                self.presentation = None
            if self.powerpoint:
                try:
                    self.powerpoint.Quit()
                except Exception:
                    pass
                self.powerpoint = None
            gc.collect()
        
        def extract_videos_from_pptx(self, pptx_path, output_dir):
            """从 PPTX 中提取嵌入的视频文件，保存到输出目录的 temp_videos 子目录"""
            import zipfile
            videos = []
            temp_dir = Path(output_dir) / "temp_videos"
            temp_dir.mkdir(parents=True, exist_ok=True)
            try:
                with zipfile.ZipFile(str(pptx_path), 'r') as zf:
                    video_exts = ('.mp4', '.avi', '.wmv', '.mov', '.flv', '.mkv', '.webm')
                    for name in zf.namelist():
                        if name.startswith('ppt/media/') and name.lower().endswith(video_exts):
                            video_data = zf.read(name)
                            ext = Path(name).suffix.lower()
                            mime_map = {
                                '.mp4': 'video/mp4', '.avi': 'video/avi',
                                '.wmv': 'video/x-ms-wmv', '.mov': 'video/quicktime',
                                '.flv': 'video/x-flv', '.mkv': 'video/x-matroska',
                                '.webm': 'video/webm'
                            }
                            mime = mime_map.get(ext, 'video/mp4')
                            # 保存到临时文件
                            temp_path = temp_dir / Path(name).name
                            temp_path.write_bytes(video_data)
                            video_b64 = base64.b64encode(video_data).decode('utf-8')
                            videos.append({
                                'name': Path(name).name,
                                'mime': mime,
                                'path': str(temp_path),
                                'data': f"data:{mime};base64,{video_b64}",
                                'size_mb': len(video_data) / (1024 * 1024)
                            })
            except Exception as e:
                log_error(f"视频提取失败: {e}")
            return videos
        
        def find_ffmpeg(self):
            """查找 ffmpeg：PATH → 自选目录"""
            # 1. PATH 中查找
            try:
                result = subprocess.run(["ffmpeg", "-version"],
                                      capture_output=True, timeout=5)
                if result.returncode == 0:
                    return "ffmpeg"
            except Exception:
                pass
            # 2. 自选目录
            custom_dir = Path(self.ffmpeg_dir_var.get())
            ffmpeg_path = custom_dir / "ffmpeg.exe"
            if ffmpeg_path.exists():
                return str(ffmpeg_path)
            return None
        
        def download_ffmpeg(self, target_dir, progress_cb=None):
            """自动下载 ffmpeg 精简版（Windows），支持进度回调"""
            import zipfile
            target_dir = Path(target_dir)
            target_dir.mkdir(parents=True, exist_ok=True)
            # gyan.dev 精简版（约 90MB，比 BtbN 的 197MB 小一半）
            url = "https://www.gyan.dev/ffmpeg/builds/ffmpeg-release-essentials.zip"
            zip_path = target_dir / "ffmpeg.zip"
            
            def _reporthook(block_num, block_size, total_size):
                if progress_cb and total_size > 0:
                    downloaded = block_num * block_size
                    pct = min(100, downloaded / total_size * 100)
                    mb_done = downloaded / (1024 * 1024)
                    mb_total = total_size / (1024 * 1024)
                    progress_cb(pct, f"{mb_done:.1f}/{mb_total:.1f} MB")
            
            try:
                import urllib.request
                self.root.after(0, lambda: self._set_ffmpeg_status("⬇️ 正在下载精简版 (~90MB)...", "blue"))
                urllib.request.urlretrieve(url, str(zip_path), _reporthook)
                
                # 解压
                self.root.after(0, lambda: self._set_ffmpeg_status("📦 正在解压...", "blue"))
                with zipfile.ZipFile(str(zip_path), 'r') as zf:
                    for name in zf.namelist():
                        if name.endswith("bin/ffmpeg.exe"):
                            data = zf.read(name)
                            ffmpeg_path = target_dir / "ffmpeg.exe"
                            ffmpeg_path.write_bytes(data)
                            break
                zip_path.unlink(missing_ok=True)
                return str(target_dir / "ffmpeg.exe")
            except Exception as e:
                zip_path.unlink(missing_ok=True)
                raise RuntimeError(f"ffmpeg 下载失败: {e}")
        
        def compress_video(self, input_path, quality="medium"):
            """用 ffmpeg 压缩视频（可中断）"""
            ffmpeg = self.find_ffmpeg()
            if not ffmpeg:
                local_dir = Path(self.ffmpeg_dir_var.get())
                try:
                    self.root.after(0, lambda: self.progress_var.set("首次使用，正在下载 ffmpeg..."))
                    ffmpeg = self.download_ffmpeg(local_dir)
                except Exception as e:
                    log_error(str(e))
                    return None
            
            output_path = Path(input_path).with_suffix(".compressed.mp4")
            crf_map = {"low": "35", "medium": "28", "high": "20"}
            crf = crf_map.get(quality, "28")
            
            try:
                self._ffmpeg_process = subprocess.Popen([
                    ffmpeg, "-y", "-i", str(input_path),
                    "-c:v", "libx264", "-crf", crf,
                    "-preset", "fast",
                    "-c:a", "aac", "-b:a", "128k",
                    str(output_path)
                ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                
                # 等待完成，同时检查停止标志
                while self._ffmpeg_process.poll() is None:
                    if not self.is_converting:
                        self._ffmpeg_process.kill()
                        self._ffmpeg_process.wait()
                        output_path.unlink(missing_ok=True)
                        return None
                    time.sleep(0.2)
                
                self._ffmpeg_process = None
                
                if output_path.exists():
                    Path(input_path).unlink()
                    output_path.rename(input_path)
                    return str(input_path)
                return None
            except Exception as e:
                log_error(f"视频压缩失败: {e}")
                output_path.unlink(missing_ok=True)
                self._ffmpeg_process = None
                return None
        
        def browse_ffmpeg_dir(self):
            directory = filedialog.askdirectory(title="选择 ffmpeg 保存目录")
            if directory:
                self.ffmpeg_dir_var.set(directory)
                # 重新检测
                threading.Thread(target=self.check_ffmpeg_status, daemon=True).start()
        
        def check_ffmpeg_status(self):
            """异步检测 ffmpeg 是否可用"""
            result = self.find_ffmpeg()
            if result:
                source = "系统 PATH" if result == "ffmpeg" else "本地缓存"
                self.root.after(0, lambda: self._set_ffmpeg_status(f"✅ 已安装 ({source})", "green"))
            else:
                self.root.after(0, lambda: self._set_ffmpeg_status("❌ 未安装", "red"))
                self.root.after(0, lambda: self.ffmpeg_download_btn.pack(side=tk.RIGHT))
        
        def _set_ffmpeg_status(self, text, color):
            self.ffmpeg_status_var.set(text)
            self.ffmpeg_status_label.config(foreground=color)
        
        def start_download_ffmpeg(self):
            """点击下载按钮"""
            self.ffmpeg_download_btn.config(state=tk.DISABLED, text="下载中...")
            threading.Thread(target=self._do_download_ffmpeg, daemon=True).start()
        
        def _do_download_ffmpeg(self):
            """后台下载 ffmpeg（带进度条）"""
            def _update_progress(pct, text):
                self.root.after(0, lambda: self.ffmpeg_progress_var.set(pct))
                self.root.after(0, lambda: self.ffmpeg_progress_label.config(text=text))
            
            try:
                # 显示进度条
                self.root.after(0, lambda: self.ffmpeg_progress_frame.pack(fill=tk.X, pady=(3, 0)))
                self.root.after(0, lambda: self.ffmpeg_progress_var.set(0))
                
                local_dir = Path(self.ffmpeg_dir_var.get())
                self.download_ffmpeg(local_dir, progress_cb=_update_progress)
                
                self.root.after(0, lambda: self.ffmpeg_progress_frame.pack_forget())
                self.root.after(0, lambda: self._set_ffmpeg_status("✅ 已安装 (本地缓存)", "green"))
                self.root.after(0, lambda: self.ffmpeg_download_btn.pack_forget())
            except Exception as e:
                self.root.after(0, lambda: self.ffmpeg_progress_frame.pack_forget())
                self.root.after(0, lambda: self._set_ffmpeg_status(f"❌ 下载失败: {e}", "red"))
                self.root.after(0, lambda: self.ffmpeg_download_btn.config(state=tk.NORMAL, text="重试下载"))
        
        def create_single_html_file(self, images_dir, output_dir, title, total_slides):
            images_base64 = []
            for i in range(1, total_slides + 1):
                img_path = images_dir / f"slide_{i}.webp"
                if img_path.exists():
                    with open(img_path, "rb") as f:
                        img_data = f.read()
                        img_base64 = base64.b64encode(img_data).decode('utf-8')
                        images_base64.append(f"data:image/webp;base64,{img_base64}")
            
            # 提取嵌入视频
            videos_data = []
            video_mode = self.video_mode_var.get()
            if video_mode != "none":
                videos_data = self.extract_videos_from_pptx(self.pptx_file, output_dir)
                if video_mode == "compress" and videos_data:
                    for v in videos_data:
                        self.root.after(0, lambda n=v['name']: self.progress_var.set(f"正在压缩视频: {n}"))
                        self.compress_video(v['path'], "medium")
                    # 重新读取压缩后的文件
                    for v in videos_data:
                        if Path(v['path']).exists():
                            with open(v['path'], 'rb') as f:
                                data = f.read()
                            v['data'] = f"data:{v['mime']};base64,{base64.b64encode(data).decode('utf-8')}"
                            v['size_mb'] = len(data) / (1024 * 1024)
            
            trace_time = float(self.trace_time_var.get())
            trace_width = float(self.trace_width_var.get())
            
            HTML_TEMPLATE = """
            <!DOCTYPE html>
            <html lang="zh-CN">
            <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <title>{{ title }} - PPT 在线浏览</title>
                <style>
                    * { margin: 0; padding: 0; box-sizing: border-box; }
                    body { 
                        background: #1a1a1a; 
                        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
                        overflow: hidden; 
                        height: 100vh; 
                        display: flex; 
                        flex-direction: column; 
                    }
                    .slides-container { 
                        flex: 1; 
                        position: relative; 
                        background: #000; 
                        overflow: hidden; 
                    }
                    .slide { 
                        position: absolute; 
                        top: 0; 
                        left: 0; 
                        width: 100%; 
                        height: 100%; 
                        display: flex; 
                        justify-content: center; 
                        align-items: center; 
                        opacity: 0; 
                        transition: opacity 0.3s ease; 
                        background: #1a1a1a; 
                    }
                    .slide.active { opacity: 1; z-index: 1; }
                    .slide img { 
                        max-width: 100%; 
                        max-height: 100%; 
                        object-fit: contain; 
                        box-shadow: 0 4px 20px rgba(0,0,0,0.3); 
                    }
                    
                    /* 缩略图侧边栏 */
                    .sidebar {
                        position: fixed;
                        left: 0;
                        top: 0;
                        width: 200px;
                        height: 100vh;
                        background: rgba(30, 30, 30, 0.95);
                        backdrop-filter: blur(10px);
                        z-index: 200;
                        transform: translateX(-100%);
                        transition: transform 0.3s ease;
                        overflow-y: auto;
                        padding: 10px;
                        border-right: 1px solid rgba(255,255,255,0.1);
                    }
                    .sidebar.show { transform: translateX(0); }
                    
                    .sidebar-title {
                        color: white;
                        font-size: 14px;
                        margin-bottom: 15px;
                        padding: 5px 10px;
                        background: rgba(255,255,255,0.1);
                        border-radius: 5px;
                    }
                    
                    .thumbnail {
                        position: relative;
                        margin-bottom: 10px;
                        cursor: pointer;
                        border-radius: 5px;
                        overflow: hidden;
                        transition: transform 0.2s ease, box-shadow 0.2s ease;
                    }
                    .thumbnail:hover {
                        transform: scale(1.05);
                        box-shadow: 0 0 10px rgba(0,120,212,0.5);
                    }
                    .thumbnail.active {
                        box-shadow: 0 0 15px rgba(0,120,212,0.8);
                        border: 2px solid #0078D4;
                    }
                    .thumbnail img {
                        width: 100%;
                        height: auto;
                        display: block;
                    }
                    .thumbnail-number {
                        position: absolute;
                        bottom: 5px;
                        right: 5px;
                        background: rgba(0,0,0,0.7);
                        color: white;
                        padding: 2px 6px;
                        border-radius: 3px;
                        font-size: 12px;
                    }
                    
                    /* 导航栏 */
                    .nav-bar {
                        position: fixed;
                        bottom: 20px;
                        left: 50%;
                        transform: translateX(-50%);
                        background: rgba(0, 0, 0, 0.8);
                        backdrop-filter: blur(10px);
                        border-radius: 50px;
                        padding: 8px 20px;
                        display: flex;
                        gap: 15px;
                        z-index: 100;
                        border: 1px solid rgba(255,255,255,0.2);
                        opacity: 0;
                        transition: opacity 0.3s ease;
                    }
                    .nav-bar.show { opacity: 1; }
                    .nav-bar button {
                        background: rgba(255,255,255,0.1);
                        border: none;
                        color: white;
                        font-size: 18px;
                        padding: 8px 20px;
                        border-radius: 40px;
                        cursor: pointer;
                        transition: all 0.2s;
                    }
                    .nav-bar button:hover {
                        background: #0078D4;
                        transform: scale(1.05);
                    }
                    .page-info {
                        color: white;
                        font-size: 16px;
                        padding: 8px 15px;
                        background: rgba(255,255,255,0.1);
                        border-radius: 40px;
                    }
                    
                    /* 右上角控制区 */
                    .top-controls {
                        position: fixed;
                        top: 20px;
                        right: 20px;
                        display: flex;
                        gap: 10px;
                        z-index: 300;
                        opacity: 0;
                        transition: opacity 0.3s ease;
                    }
                    .top-controls.show { opacity: 1; }
                    
                    .control-btn {
                        background: rgba(0, 0, 0, 0.7);
                        color: white;
                        border: 1px solid rgba(255,255,255,0.3);
                        padding: 10px 20px;
                        border-radius: 25px;
                        cursor: pointer;
                        font-size: 14px;
                        transition: all 0.2s;
                    }
                    .control-btn:hover {
                        background: #0078D4;
                        transform: scale(1.05);
                    }
                    
                    /* 设置面板 */
                    .settings-panel {
                        position: fixed;
                        top: 70px;
                        right: 20px;
                        width: 280px;
                        background: rgba(30, 30, 30, 0.95);
                        backdrop-filter: blur(10px);
                        border-radius: 15px;
                        padding: 20px;
                        z-index: 250;
                        border: 1px solid rgba(255,255,255,0.2);
                        opacity: 0;
                        visibility: hidden;
                        transform: translateY(-10px);
                        transition: all 0.3s ease;
                    }
                    .settings-panel.show {
                        opacity: 1;
                        visibility: visible;
                        transform: translateY(0);
                    }
                    
                    .settings-title {
                        color: white;
                        font-size: 16px;
                        font-weight: bold;
                        margin-bottom: 20px;
                        padding-bottom: 10px;
                        border-bottom: 1px solid rgba(255,255,255,0.2);
                    }
                    
                    .setting-item {
                        margin-bottom: 15px;
                    }
                    
                    .setting-label {
                        color: rgba(255,255,255,0.8);
                        font-size: 13px;
                        margin-bottom: 8px;
                        display: flex;
                        justify-content: space-between;
                        align-items: center;
                    }
                    
                    .setting-value {
                        color: #0078D4;
                        font-weight: bold;
                    }
                    
                    .setting-slider {
                        width: 100%;
                        height: 6px;
                        -webkit-appearance: none;
                        background: rgba(255,255,255,0.2);
                        border-radius: 3px;
                        outline: none;
                        cursor: pointer;
                    }
                    
                    .setting-slider::-webkit-slider-thumb {
                        -webkit-appearance: none;
                        width: 18px;
                        height: 18px;
                        background: #0078D4;
                        border-radius: 50%;
                        cursor: pointer;
                        transition: all 0.2s;
                    }
                    
                    .setting-slider::-webkit-slider-thumb:hover {
                        transform: scale(1.2);
                        box-shadow: 0 0 10px rgba(0,120,212,0.5);
                    }
                    
                    .setting-hint {
                        color: rgba(255,255,255,0.5);
                        font-size: 11px;
                        margin-top: 5px;
                    }
                    
                    .setting-reset {
                        width: 100%;
                        margin-top: 10px;
                        padding: 8px;
                        background: rgba(255,255,255,0.1);
                        border: 1px solid rgba(255,255,255,0.3);
                        color: white;
                        border-radius: 8px;
                        cursor: pointer;
                        font-size: 13px;
                        transition: all 0.2s;
                    }
                    .setting-reset:hover {
                        background: rgba(255,100,100,0.3);
                    }
                    
                    /* 激光笔画布 */
                    #laserCanvas {
                        position: fixed;
                        top: 0;
                        left: 0;
                        width: 100%;
                        height: 100%;
                        pointer-events: none;
                        z-index: 150;
                    }
                    
                    body.laser-mode {
                        cursor: url('data:image/svg+xml,<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24"><circle cx="12" cy="12" r="5" fill="red" stroke="white" stroke-width="2"/></svg>') 12 12, crosshair;
                    }
                    
                    .loading-hint {
                        position: fixed;
                        bottom: 80px;
                        left: 50%;
                        transform: translateX(-50%);
                        color: rgba(255,255,255,0.6);
                        font-size: 12px;
                        z-index: 50;
                    }
                    
                    /* 视频播放器 */
                    .slide video {
                        max-width: 100%;
                        max-height: 100%;
                        object-fit: contain;
                        box-shadow: 0 4px 20px rgba(0,0,0,0.3);
                    }
                    .video-indicator {
                        position: absolute;
                        top: 10px;
                        right: 10px;
                        background: rgba(255,0,0,0.8);
                        color: white;
                        padding: 4px 10px;
                        border-radius: 15px;
                        font-size: 12px;
                        z-index: 5;
                    }
                </style>
            </head>
            <body>
                <!-- 缩略图侧边栏 -->
                <div class="sidebar" id="sidebar">
                    <div class="sidebar-title">幻灯片缩略图</div>
                    {% for img_base64 in images_base64 %}
                    <div class="thumbnail" data-index="{{ loop.index0 }}">
                        <img src="{{ img_base64 }}" alt="第 {{ loop.index }} 页">
                        <div class="thumbnail-number">{{ loop.index }}</div>
                    </div>
                    {% endfor %}
                </div>
                
                <!-- 主幻灯片区域 -->
                <div class="slides-container" id="slidesContainer">
                    {% for img_base64 in images_base64 %}
                    <div class="slide" data-index="{{ loop.index0 }}">
                        <img src="{{ img_base64 }}" alt="第 {{ loop.index }} 页">
                    </div>
                    {% endfor %}
                </div>
                
                <!-- 视频列表 -->
                {% if videos %}
                <div id="videoSection" style="display:none; position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.95); z-index:400; display:none; justify-content:center; align-items:center; flex-direction:column;">
                    <div style="position:absolute; top:20px; right:20px; z-index:401;">
                        <button onclick="closeVideo()" style="background:rgba(255,255,255,0.2); border:none; color:white; padding:10px 20px; border-radius:25px; cursor:pointer; font-size:14px;">✕ 关闭视频</button>
                    </div>
                    <div style="position:absolute; top:20px; left:20px; color:white; font-size:16px; z-index:401;">
                        📹 PPT 内嵌视频 ({{ videos|length }} 个)
                    </div>
                    <div id="videoContainer" style="max-width:90%; max-height:80vh;"></div>
                    <div style="margin-top:15px; display:flex; gap:10px; flex-wrap:wrap; justify-content:center;">
                        {% for video in videos %}
                        <button onclick="playVideo({{ loop.index0 }})" style="background:rgba(255,255,255,0.15); border:1px solid rgba(255,255,255,0.3); color:white; padding:8px 16px; border-radius:20px; cursor:pointer; font-size:13px;">
                            ▶ {{ video.name }} ({{ "%.1f"|format(video.size_mb) }}MB)
                        </button>
                        {% endfor %}
                    </div>
                </div>
                {% endif %}
                <canvas id="laserCanvas"></canvas>
                
                <!-- 导航栏 -->
                <div class="nav-bar" id="navBar">
                    <button id="prevBtn">◀ 上一页</button>
                    <span class="page-info" id="pageInfo">1 / {{ total }}</span>
                    <button id="nextBtn">下一页 ▶</button>
                </div>
                
                <!-- 右上角控制区 -->
                <div class="top-controls" id="topControls">
                    <button class="control-btn" id="settingsBtn">⚙️ 设置</button>
                    {% if videos %}
                    <button class="control-btn" id="videoBtn" onclick="openVideo()">📹 视频</button>
                    {% endif %}
                    <button class="control-btn" id="enterFullscreenBtn">🖥️ 全屏</button>
                    <button class="control-btn" id="exitFullscreenBtn" style="display: none;">❌ 退出全屏</button>
                </div>
                
                <!-- 设置面板 -->
                <div class="settings-panel" id="settingsPanel">
                    <div class="settings-title">🎯 激光笔设置</div>
                    
                    <div class="setting-item">
                        <div class="setting-label">
                            <span>轨迹显示时间</span>
                            <span class="setting-value" id="timeValue">{{ trace_time }} 秒</span>
                        </div>
                        <input type="range" class="setting-slider" id="traceTimeSlider" 
                               min="0.5" max="10" step="0.1" value="{{ trace_time }}">
                        <div class="setting-hint">轨迹消失速度，时间越长消失越慢</div>
                    </div>
                    
                    <div class="setting-item">
                        <div class="setting-label">
                            <span>轨迹宽度</span>
                            <span class="setting-value" id="widthValue">{{ trace_width }} 像素</span>
                        </div>
                        <input type="range" class="setting-slider" id="traceWidthSlider" 
                               min="1" max="20" step="1" value="{{ trace_width }}">
                        <div class="setting-hint">建议 3-10 像素</div>
                    </div>
                    
                    <button class="setting-reset" id="resetSettings">恢复默认设置</button>
                </div>
                
                <!-- 操作提示 -->
                <!-- <div class="loading-hint">← 左侧显示缩略图 | 底部显示导航 | F11 全屏 | 全屏后鼠标可画 →</div>-->
                
                <script>
                    // DOM 元素
                    const slides = document.querySelectorAll('.slide');
                    const thumbnails = document.querySelectorAll('.thumbnail');
                    const prevBtn = document.getElementById('prevBtn');
                    const nextBtn = document.getElementById('nextBtn');
                    const pageInfo = document.getElementById('pageInfo');
                    const sidebar = document.getElementById('sidebar');
                    const navBar = document.getElementById('navBar');
                    const topControls = document.getElementById('topControls');
                    const enterFullscreenBtn = document.getElementById('enterFullscreenBtn');
                    const exitFullscreenBtn = document.getElementById('exitFullscreenBtn');
                    const settingsBtn = document.getElementById('settingsBtn');
                    const settingsPanel = document.getElementById('settingsPanel');
                    const laserCanvas = document.getElementById('laserCanvas');
                    const ctx = laserCanvas.getContext('2d');
                    
                    // 设置相关
                    const traceTimeSlider = document.getElementById('traceTimeSlider');
                    const traceWidthSlider = document.getElementById('traceWidthSlider');
                    const timeValueDisplay = document.getElementById('timeValue');
                    const widthValueDisplay = document.getElementById('widthValue');
                    const resetSettingsBtn = document.getElementById('resetSettings');
                    
                    // 默认值
                    const defaultTraceTime = {{ trace_time }};
                    const defaultTraceWidth = {{ trace_width }};
                    
                    // 当前设置
                    let traceTime = defaultTraceTime;
                    let traceWidth = defaultTraceWidth;
                    
                    // 状态变量
                    let currentIndex = 0;
                    const totalSlides = {{ total }};
                    let sidebarVisible = false;
                    let navBarVisible = false;
                    let topControlsVisible = false;
                    let settingsVisible = false;
                    let sidebarTimeout;
                    let navBarTimeout;
                    let topControlsTimeout;
                    
                    // 激光笔相关变量
                    let isDrawing = false;
                    let laserPoints = [];
                    let laserMode = false;
                    let animationRunning = false;
                    let cleanupInterval = null;
                    
                    // 设置画布大小
                    function resizeCanvas() {
                        laserCanvas.width = window.innerWidth;
                        laserCanvas.height = window.innerHeight;
                    }
                    resizeCanvas();
                    window.addEventListener('resize', resizeCanvas);
                    
                    // ========== 设置面板功能 ==========
                    
                    // 更新轨迹时间设置
                    traceTimeSlider.addEventListener('input', (e) => {
                        traceTime = parseFloat(e.target.value);
                        timeValueDisplay.textContent = traceTime.toFixed(1) + ' 秒';
                    });
                    
                    // 更新轨迹宽度设置
                    traceWidthSlider.addEventListener('input', (e) => {
                        traceWidth = parseInt(e.target.value);
                        widthValueDisplay.textContent = traceWidth + ' 像素';
                    });
                    
                    // 恢复默认设置
                    resetSettingsBtn.addEventListener('click', () => {
                        traceTime = defaultTraceTime;
                        traceWidth = defaultTraceWidth;
                        traceTimeSlider.value = traceTime;
                        traceWidthSlider.value = traceWidth;
                        timeValueDisplay.textContent = traceTime.toFixed(1) + ' 秒';
                        widthValueDisplay.textContent = traceWidth + ' 像素';
                    });
                    
                    // 切换设置面板
                    function toggleSettings() {
                        settingsVisible = !settingsVisible;
                        if (settingsVisible) {
                            settingsPanel.classList.add('show');
                        } else {
                            settingsPanel.classList.remove('show');
                        }
                    }
                    
                    settingsBtn.addEventListener('click', (e) => {
                        e.stopPropagation();
                        toggleSettings();
                    });
                    
                    // 点击外部关闭设置面板
                    document.addEventListener('click', (e) => {
                        if (settingsVisible && 
                            !settingsPanel.contains(e.target) && 
                            !settingsBtn.contains(e.target)) {
                            settingsVisible = false;
                            settingsPanel.classList.remove('show');
                        }
                    });
                    
                    // ========== 幻灯片导航功能 ==========
                    
                    function showSlide(index) {
                        if (index < 0) index = 0;
                        if (index >= totalSlides) index = totalSlides - 1;
                        
                        slides[currentIndex].classList.remove('active');
                        thumbnails[currentIndex].classList.remove('active');
                        
                        currentIndex = index;
                        
                        slides[currentIndex].classList.add('active');
                        thumbnails[currentIndex].classList.add('active');
                        
                        pageInfo.innerText = (currentIndex + 1) + ' / ' + totalSlides;
                        
                        thumbnails[currentIndex].scrollIntoView({ behavior: 'smooth', block: 'center' });
                    }
                    
                    // 侧边栏显示/隐藏
                    function showSidebar() {
                        clearTimeout(sidebarTimeout);
                        sidebar.classList.add('show');
                        sidebarVisible = true;
                    }
                    function hideSidebar() {
                        sidebarTimeout = setTimeout(() => {
                            sidebar.classList.remove('show');
                            sidebarVisible = false;
                        }, 300);
                    }
                    
                    // 导航栏显示/隐藏
                    function showNavBar() {
                        clearTimeout(navBarTimeout);
                        navBar.classList.add('show');
                        navBarVisible = true;
                    }
                    function hideNavBar() {
                        navBarTimeout = setTimeout(() => {
                            navBar.classList.remove('show');
                            navBarVisible = false;
                        }, 300);
                    }
                    
                    // 顶部控制区显示/隐藏
                    function showTopControls() {
                        clearTimeout(topControlsTimeout);
                        topControls.classList.add('show');
                        topControlsVisible = true;
                    }
                    function hideTopControls() {
                        topControlsTimeout = setTimeout(() => {
                            topControls.classList.remove('show');
                            topControlsVisible = false;
                        }, 300);
                    }
                    
                    // 全屏功能
                    function enterFullscreen() {
                        document.documentElement.requestFullscreen();
                    }
                    function exitFullscreen() {
                        if (document.exitFullscreen) {
                            document.exitFullscreen();
                        }
                    }
                    
                    // ========== 激光笔功能 ==========
                    
                    function startLaserMode() {
                        laserMode = true;
                        document.body.classList.add('laser-mode');
                        startCleanup();
                        startDrawLoop();
                    }
                    
                    function stopLaserMode() {
                        laserMode = false;
                        isDrawing = false;
                        document.body.classList.remove('laser-mode');
                        
                        if (cleanupInterval) {
                            clearInterval(cleanupInterval);
                            cleanupInterval = null;
                        }
                        
                        animationRunning = false;
                        ctx.clearRect(0, 0, laserCanvas.width, laserCanvas.height);
                        laserPoints = [];
                    }
                    
                    function startCleanup() {
                        if (cleanupInterval) clearInterval(cleanupInterval);
                        cleanupInterval = setInterval(() => {
                            const now = Date.now();
                            while (laserPoints.length > 0 && 
                                (now - laserPoints[0].timestamp) > traceTime * 1000) {
                                laserPoints.shift();
                            }
                        }, 50);
                    }
                    
                    function startDrawLoop() {
                        if (!animationRunning) {
                            animationRunning = true;
                            drawLaser();
                        }
                    }
                    
                    function drawLaser() {
                        if (!animationRunning) return;
                        
                        ctx.clearRect(0, 0, laserCanvas.width, laserCanvas.height);
                        
                        if (laserPoints.length > 0) {
                            ctx.lineCap = 'round';
                            ctx.lineJoin = 'round';
                            ctx.lineWidth = traceWidth;
                            
                            const now = Date.now();
                            
                            for (let i = 1; i < laserPoints.length; i++) {
                                const point = laserPoints[i];
                                const prevPoint = laserPoints[i - 1];
                                
                                if (point === null || prevPoint === null) continue;
                                
                                const age = (now - point.timestamp) / 1000;
                                const remainingTime = Math.max(0, traceTime - age);
                                const alpha = Math.min(0.8, remainingTime / traceTime);
                                
                                ctx.beginPath();
                                ctx.moveTo(prevPoint.x, prevPoint.y);
                                ctx.lineTo(point.x, point.y);
                                ctx.strokeStyle = `rgba(255, 0, 0, ${alpha})`;
                                ctx.stroke();
                            }
                            
                            const lastPoint = laserPoints[laserPoints.length - 1];
                            if (lastPoint !== null) {
                                const age = (now - lastPoint.timestamp) / 1000;
                                const remainingTime = Math.max(0, traceTime - age);
                                const alpha = Math.min(0.9, remainingTime / traceTime * 0.9);
                                
                                ctx.beginPath();
                                ctx.arc(lastPoint.x, lastPoint.y, traceWidth / 2, 0, Math.PI * 2);
                                ctx.fillStyle = `rgba(255, 0, 0, ${alpha})`;
                                ctx.fill();
                            }
                        }
                        
                        requestAnimationFrame(drawLaser);
                    }
                    
                    // ========== 事件监听 ==========
                    
                    prevBtn.addEventListener('click', () => showSlide(currentIndex - 1));
                    nextBtn.addEventListener('click', () => showSlide(currentIndex + 1));
                    enterFullscreenBtn.addEventListener('click', enterFullscreen);
                    exitFullscreenBtn.addEventListener('click', exitFullscreen);
                    
                    thumbnails.forEach((thumb, index) => {
                        thumb.addEventListener('click', () => showSlide(index));
                    });
                    
                    // 键盘导航
                    document.addEventListener('keydown', (e) => {
                        if (e.key === 'ArrowLeft') { 
                            showSlide(currentIndex - 1); 
                            e.preventDefault(); 
                        }
                        else if (e.key === 'ArrowRight') { 
                            showSlide(currentIndex + 1); 
                            e.preventDefault(); 
                        }
                        else if (e.key === 'Escape') { 
                            // 优先关闭视频
                            const videoSection = document.getElementById('videoSection');
                            if (videoSection && videoSection.style.display === 'flex') {
                                closeVideo();
                            } else {
                                exitFullscreen(); 
                                stopLaserMode();
                                settingsVisible = false;
                                settingsPanel.classList.remove('show');
                            }
                        }
                        else if (e.key === 'F11') { 
                            e.preventDefault(); 
                            enterFullscreen(); 
                        }
                    });
                    
                    // 鼠标滚轮翻页
                    document.getElementById('slidesContainer').addEventListener('wheel', (e) => {
                        e.preventDefault();
                        if (e.deltaY > 0) {
                            showSlide(currentIndex + 1);
                        } else {
                            showSlide(currentIndex - 1);
                        }
                    });
                    
                    // 鼠标移动事件
                    document.addEventListener('mousemove', (e) => {
                        const x = e.clientX;
                        const y = e.clientY;
                        
                        // 左侧缩略图
                        if (x < 30) {
                            showSidebar();
                        } else if (x > 200 && !sidebar.contains(document.elementFromPoint(x, y))) {
                            hideSidebar();
                        }
                        
                        // 底部导航栏
                        if (y > window.innerHeight - 50) {
                            showNavBar();
                        } else if (y < window.innerHeight - 100) {
                            hideNavBar();
                        }
                        
                        // 右上角控制区
                        if (x > window.innerWidth - 250 && y < 100) {
                            showTopControls();
                        } else if (x < window.innerWidth - 300 || y > 150) {
                            hideTopControls();
                        }
                        
                        // 激光笔绘制
                        if (laserMode && isDrawing) {
                            laserPoints.push({ 
                                x: e.clientX, 
                                y: e.clientY, 
                                timestamp: Date.now() 
                            });
                        }
                    });
                    
                    // 鼠标按下事件
                    document.addEventListener('mousedown', (e) => {
                        // 检查是否点击了设置面板或按钮
                        if (settingsPanel.contains(e.target) || 
                            settingsBtn.contains(e.target) ||
                            topControls.contains(e.target)) {
                            return;
                        }
                        
                        if (laserMode) {
                            isDrawing = true;
                            laserPoints.push(null);
                            laserPoints.push({ 
                                x: e.clientX, 
                                y: e.clientY, 
                                timestamp: Date.now() 
                            });
                            e.preventDefault();
                        }
                    });
                    
                    // 鼠标释放事件
                    document.addEventListener('mouseup', () => {
                        isDrawing = false;
                    });
                    
                    // 全屏状态变化
                    document.addEventListener('fullscreenchange', () => {
                        if (document.fullscreenElement) {
                            startLaserMode();
                            exitFullscreenBtn.style.display = 'block';
                            enterFullscreenBtn.style.display = 'none';
                        } else {
                            stopLaserMode();
                            exitFullscreenBtn.style.display = 'none';
                            enterFullscreenBtn.style.display = 'block';
                        }
                    });
                    
                    // 触摸滑动支持
                    let touchStartX = 0;
                    document.getElementById('slidesContainer').addEventListener('touchstart', (e) => {
                        touchStartX = e.changedTouches[0].screenX;
                    });
                    document.getElementById('slidesContainer').addEventListener('touchend', (e) => {
                        let endX = e.changedTouches[0].screenX;
                        if (endX < touchStartX - 50) showSlide(currentIndex + 1);
                        if (endX > touchStartX + 50) showSlide(currentIndex - 1);
                    });
                    
                    // ========== 视频播放功能 ==========
                    const videosData = [
                        {% for video in videos %}
                        { name: "{{ video.name }}", mime: "{{ video.mime }}", data: "{{ video.data }}" }{% if not loop.last %},{% endif %}
                        {% endfor %}
                    ];
                    
                    function openVideo() {
                        const section = document.getElementById('videoSection');
                        if (section) {
                            section.style.display = 'flex';
                            if (videosData.length === 1) playVideo(0);
                        }
                    }
                    
                    function closeVideo() {
                        const section = document.getElementById('videoSection');
                        const container = document.getElementById('videoContainer');
                        if (section) section.style.display = 'none';
                        if (container) container.innerHTML = '';
                    }
                    
                    function playVideo(index) {
                        const container = document.getElementById('videoContainer');
                        if (!container || !videosData[index]) return;
                        const v = videosData[index];
                        container.innerHTML = '<video controls autoplay style="max-width:100%;max-height:70vh;border-radius:8px;" src="' + v.data + '"></video>';
                    }
                    
                    // 初始化
                    showSlide(0);
                </script>
            </body>
            </html>
            """
            
            template = Template(HTML_TEMPLATE)
            final_html = template.render(
                title=title,
                images_base64=images_base64,
                total=total_slides,
                trace_time=trace_time,
                trace_width=trace_width,
                videos=videos_data
            )
            
            output_html = Path(output_dir) / f"{title}.html"
            with open(output_html, "w", encoding="utf-8") as f:
                f.write(final_html)
            
            return output_html
        
        def show_error(self, message):
            def _show():
                messagebox.showerror("错误", message)
                self.progress_var.set("转换失败")
                self.status_var.set("错误: " + message)
                self.stop_btn.pack_forget()
                self.convert_btn.pack(side=tk.RIGHT, padx=(10, 0))
                self.convert_btn.config(text="开始转换", state=tk.NORMAL)
            # 确保在主线程执行（兼容从子线程直接调用）
            try:
                self.root.after(0, _show)
            except Exception:
                pass
        
        def run(self):
            self.root.mainloop()


    if __name__ == "__main__":
        app = PPTXToHTMLConverter()
        app.run()

except Exception as e:
    log_error(f"启动错误: {str(e)}\n{traceback.format_exc()}")
    print(f"启动错误: {str(e)}")
    traceback.print_exc()
    input("按回车键退出...")
    sys.exit(1)