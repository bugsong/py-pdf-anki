import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
import fitz  # PyMuPDF库，用于处理PDF
import requests
import io
from PIL import Image, ImageTk
import os
import uuid
import json

class PDFAnkiTool:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF到Anki问答题工具")
        self.root.geometry("1000x700")
        
        # 设置中文字体支持
        self.root.option_add("*Font", "SimHei 10")
        
        # 初始化变量
        self.pdf_path = None
        self.pdf_document = None
        self.current_page = 0
        self.scale_factor = 1.0
        self.canvas_width = 800
        self.canvas_height = 600
        
        # 选择区域相关变量
        self.selection_start = None
        self.selection_end = None
        self.selection_rect = None
        self.is_selecting = False
        self.is_dragging = False
        self.drag_start = None
        self.screenshot = None
        self.screenshot_path = None
        self.canvas_image = None
        self.tk_image = None
        
        # 右键拖动相关变量
        self.is_panning = False
        self.pan_start = None
        self.canvas_offset_x = 0
        self.canvas_offset_y = 0
        
        # 图片存储路径设置
        self.custom_image_path = None  # 自定义图片存储路径
        
        # 图片另存设置
        self.save_image_var = tk.BooleanVar(value=False)
        self.save_image_locally = False  # 是否在本地保存图片，默认关闭
        
        # 创建界面控件
        self.create_widgets()
        
        # PDF显示区域
        self.canvas_frame = tk.Frame(self.root, relief=tk.SUNKEN, bd=2)
        self.canvas_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        self.canvas = tk.Canvas(self.canvas_frame, bg="white")
        self.canvas.pack(fill=tk.BOTH, expand=True)
        
        # 绑定鼠标事件
        self.canvas.bind("<ButtonPress-1>", self.on_mouse_down)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_mouse_up)
        self.canvas.bind("<MouseWheel>", self.on_mouse_wheel)
        self.canvas.bind("<ButtonPress-3>", self.on_right_mouse_down)  # 右键按下
        self.canvas.bind("<B3-Motion>", self.on_right_mouse_drag)    # 右键拖动
        self.canvas.bind("<ButtonRelease-3>", self.on_right_mouse_up)  # 右键释放
        
        # 底部状态栏
        status_frame = tk.Frame(self.root)
        status_frame.pack(side=tk.BOTTOM, fill=tk.X, padx=10, pady=5)
        
        self.status_bar = tk.Label(status_frame, text="请选择PDF文件并输入问题", bd=1, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 清除选择按钮
        self.clear_btn = tk.Button(status_frame, text="清除选择", command=self.clear_selection)
        self.clear_btn.pack(side=tk.RIGHT, padx=5)
    
    def create_widgets(self):
        """创建界面控件"""
        # 顶部控制区
        control_frame = tk.Frame(self.root)
        control_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 问题输入
        tk.Label(control_frame, text="问题:").pack(side=tk.LEFT, padx=5)
        self.question_entry = tk.Entry(control_frame, width=40)
        self.question_entry.pack(side=tk.LEFT, padx=5)
        self.question_entry.bind("<KeyRelease>", self.check_add_button_state)
        
        # 选择PDF按钮
        self.select_pdf_btn = tk.Button(control_frame, text="选择PDF", command=self.select_pdf)
        self.select_pdf_btn.pack(side=tk.LEFT, padx=5)
        
        # 页面导航
        page_nav_frame = tk.Frame(control_frame)
        page_nav_frame.pack(side=tk.LEFT, padx=20)
        
        self.prev_page_btn = tk.Button(page_nav_frame, text="上一页", command=self.prev_page)
        self.prev_page_btn.pack(side=tk.LEFT, padx=2)
        self.prev_page_btn.config(state=tk.DISABLED)
        
        self.page_label = tk.Label(page_nav_frame, text="页码: 0/0")
        self.page_label.pack(side=tk.LEFT, padx=10)
        
        self.next_page_btn = tk.Button(page_nav_frame, text="下一页", command=self.next_page)
        self.next_page_btn.pack(side=tk.LEFT, padx=2)
        self.next_page_btn.config(state=tk.DISABLED)
        
        # 缩放控制
        zoom_frame = tk.Frame(control_frame)
        zoom_frame.pack(side=tk.LEFT, padx=20)
        
        tk.Label(zoom_frame, text="缩放:").pack(side=tk.LEFT)
        self.zoom_in_btn = tk.Button(zoom_frame, text="+", width=3, command=self.zoom_in)
        self.zoom_in_btn.pack(side=tk.LEFT, padx=2)
        
        self.zoom_label = tk.Label(zoom_frame, text="100%")
        self.zoom_label.pack(side=tk.LEFT, padx=5)
        
        self.zoom_out_btn = tk.Button(zoom_frame, text="-", width=3, command=self.zoom_out)
        self.zoom_out_btn.pack(side=tk.LEFT, padx=2)
        
        self.fit_page_btn = tk.Button(zoom_frame, text="适应页面", command=self.fit_to_page)
        self.fit_page_btn.pack(side=tk.LEFT, padx=5)
        
        # 图片另存复选框
        self.save_image_checkbox = tk.Checkbutton(
            control_frame, 
            text="保存图片", 
            variable=self.save_image_var,
            command=self.toggle_save_image
        )
        self.save_image_checkbox.pack(side=tk.RIGHT, padx=5)
        
        # 设置图片存储路径按钮
        self.set_path_btn = tk.Button(control_frame, text="设置路径", command=self.set_image_path)
        self.set_path_btn.pack(side=tk.RIGHT, padx=5)
        
        # 添加到Anki按钮
        self.add_to_anki_btn = tk.Button(control_frame, text="添加到Anki", command=self.add_to_anki)
        self.add_to_anki_btn.pack(side=tk.RIGHT, padx=5)
        self.add_to_anki_btn.config(state=tk.DISABLED)
    
    def on_window_resize(self, event):
        """窗口大小改变时的处理"""
        if event.widget == self.root:
            # 延迟更新以避免频繁重绘
            self.root.after(100, self.update_page_display)
    
    def select_pdf(self):
        """选择PDF文件并加载"""
        file_path = filedialog.askopenfilename(
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )
        
        if file_path:
            try:
                self.pdf_path = file_path
                self.pdf_document = fitz.open(self.pdf_path)
                self.current_page = 0
                self.scale_factor = 1.0
                self.clear_selection()
                self.update_page_display()
                self.update_page_controls()
                self.status_bar.config(text=f"已加载PDF: {os.path.basename(self.pdf_path)}")
                self.check_add_button_state()
            except Exception as e:
                messagebox.showerror("错误", f"无法打开PDF文件: {str(e)}")
                self.reset_pdf()
    
    def reset_pdf(self):
        """重置PDF相关状态"""
        self.pdf_path = None
        self.pdf_document = None
        self.current_page = 0
        self.scale_factor = 1.0
        self.clear_selection()
        self.canvas.delete("all")
        self.update_page_controls()
        self.status_bar.config(text="请选择PDF文件并输入问题")
        self.check_add_button_state()
    
    def update_page_controls(self):
        """更新页面控制按钮状态"""
        if self.pdf_document:
            total_pages = len(self.pdf_document)
            self.page_label.config(text=f"页码: {self.current_page + 1}/{total_pages}")
            self.prev_page_btn.config(state=tk.NORMAL if self.current_page > 0 else tk.DISABLED)
            self.next_page_btn.config(state=tk.NORMAL if self.current_page < total_pages - 1 else tk.DISABLED)
            self.zoom_in_btn.config(state=tk.NORMAL)
            self.zoom_out_btn.config(state=tk.NORMAL)
            self.fit_page_btn.config(state=tk.NORMAL)
        else:
            self.page_label.config(text="页码: 0/0")
            self.prev_page_btn.config(state=tk.DISABLED)
            self.next_page_btn.config(state=tk.DISABLED)
            self.zoom_in_btn.config(state=tk.DISABLED)
            self.zoom_out_btn.config(state=tk.DISABLED)
            self.fit_page_btn.config(state=tk.DISABLED)
    
    def update_page_display(self):
        """更新当前页面显示"""
        if not self.pdf_document:
            return
        
        try:
            # 清除画布
            self.canvas.delete("all")
            
            # 获取当前页并渲染为图片
            page = self.pdf_document[self.current_page]
            pix = page.get_pixmap()
            
            # 获取画布实际大小
            self.canvas.update_idletasks()
            canvas_width = self.canvas.winfo_width()
            canvas_height = self.canvas.winfo_height()
            
            if canvas_width <= 1 or canvas_height <= 1:
                canvas_width = 800
                canvas_height = 600
            
            # 计算缩放后的尺寸
            display_width = int(pix.width * self.scale_factor)
            display_height = int(pix.height * self.scale_factor)
            
            # 转换为PIL图像
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            
            # 如果需要，调整图像大小
            if self.scale_factor != 1.0:
                img = img.resize((display_width, display_height), Image.LANCZOS)
            
            # 创建photoimage对象
            self.tk_image = ImageTk.PhotoImage(image=img)
            
            # 重置画布偏移量
            self.canvas_offset_x = 0
            self.canvas_offset_y = 0
            
            # 计算居中位置
            x_offset = max(0, (canvas_width - display_width) // 2)
            y_offset = max(0, (canvas_height - display_height) // 2)
            
            # 在画布上显示图像
            self.canvas_image = self.canvas.create_image(
                x_offset, y_offset, anchor=tk.NW, image=self.tk_image
            )
            
            # 更新缩放标签
            self.zoom_label.config(text=f"{int(self.scale_factor * 100)}%")
            
            # 重新绘制选择区域（如果存在）
            if self.selection_start and self.selection_end:
                self.draw_selection_rect()
            
        except Exception as e:
            self.status_bar.config(text=f"更新页面显示时出错: {str(e)}")
    
    def fit_to_page(self):
        """适应页面大小"""
        if not self.pdf_document:
            return
        
        try:
            page = self.pdf_document[self.current_page]
            pix = page.get_pixmap()
            
            self.canvas.update_idletasks()
            canvas_width = self.canvas.winfo_width()
            canvas_height = self.canvas.winfo_height()
            
            if canvas_width <= 1 or canvas_height <= 1:
                return
            
            # 计算适应缩放比例
            scale_x = canvas_width / pix.width
            scale_y = canvas_height / pix.height
            self.scale_factor = min(scale_x, scale_y) * 0.95  # 留5%边距
            
            self.update_page_display()
            
        except Exception as e:
            self.status_bar.config(text=f"适应页面时出错: {str(e)}")
    
    def zoom_in(self):
        """放大"""
        if self.scale_factor < 3.0:
            self.scale_factor *= 1.2
            self.update_page_display()
    
    def zoom_out(self):
        """缩小"""
        if self.scale_factor > 0.2:
            self.scale_factor /= 1.2
            self.update_page_display()
    
    def on_mouse_wheel(self, event):
        """鼠标滚轮缩放"""
        if not self.pdf_document:
            return
        
        # 根据滚轮方向调整缩放
        if event.delta > 0:
            self.zoom_in()
        else:
            self.zoom_out()
    
    def prev_page(self):
        """显示上一页"""
        if self.current_page > 0:
            self.current_page -= 1
            self.clear_selection()
            self.update_page_display()
            self.update_page_controls()
    
    def next_page(self):
        """显示下一页"""
        if self.pdf_document and self.current_page < len(self.pdf_document) - 1:
            self.current_page += 1
            self.clear_selection()
            self.update_page_display()
            self.update_page_controls()
    
    def on_mouse_down(self, event):
        """鼠标按下事件"""
        if not self.pdf_document:
            return
        
        # 检查是否点击在现有选择区域内
        if self.selection_start and self.selection_end:
            x1, y1 = self.selection_start
            x2, y2 = self.selection_end
            
            # 确保坐标有序
            min_x, max_x = min(x1, x2), max(x1, x2)
            min_y, max_y = min(y1, y2), max(y1, y2)
            
            # 检查点击是否在选择区域内
            if (min_x <= event.x <= max_x and min_y <= event.y <= max_y):
                # 开始拖动
                self.is_dragging = True
                self.drag_start = (event.x, event.y)
                self.canvas.config(cursor="fleur")
                return
        
        # 开始新的选择
        self.is_selecting = True
        self.selection_start = (event.x, event.y)
        self.selection_end = (event.x, event.y)
        self.draw_selection_rect()
    
    def on_mouse_drag(self, event):
        """鼠标拖动事件"""
        if not self.pdf_document:
            return
        
        if self.is_dragging and self.drag_start:
            # 拖动现有选择区域
            dx = event.x - self.drag_start[0]
            dy = event.y - self.drag_start[1]
            
            if self.selection_start and self.selection_end:
                # 移动选择区域
                self.selection_start = (self.selection_start[0] + dx, self.selection_start[1] + dy)
                self.selection_end = (self.selection_end[0] + dx, self.selection_end[1] + dy)
                self.drag_start = (event.x, event.y)
                self.draw_selection_rect()
                
        elif self.is_selecting:
            # 更新选择区域
            self.selection_end = (event.x, event.y)
            self.draw_selection_rect()
    
    def on_mouse_up(self, event):
        """鼠标释放事件"""
        if self.is_selecting:
            self.is_selecting = False
            # 确保选择区域有效
            if self.selection_start and self.selection_end:
                x1, y1 = self.selection_start
                x2, y2 = self.selection_end
                
                # 如果选择区域太小，清除选择
                if abs(x2 - x1) < 5 or abs(y2 - y1) < 5:
                    self.clear_selection()
                else:
                    self.capture_selected_area()
                    self.status_bar.config(text="已选择区域，右键或点击'清除选择'可重新选择")
        
        if self.is_dragging:
            self.is_dragging = False
            self.drag_start = None
            self.canvas.config(cursor="")
            if self.selection_start and self.selection_end:
                self.capture_selected_area()
        
        self.check_add_button_state()
    
    def on_right_mouse_down(self, event):
        """右键按下事件"""
        if not self.pdf_document:
            return
        
        # 开始拖动画布
        self.is_panning = True
        self.pan_start = (event.x, event.y)
        self.canvas.config(cursor="fleur")
        self.status_bar.config(text="拖动中...")
    
    def on_right_mouse_drag(self, event):
        """右键拖动事件"""
        if not self.is_panning or not self.pan_start:
            return
        
        # 计算移动距离
        dx = event.x - self.pan_start[0]
        dy = event.y - self.pan_start[1]
        
        # 移动画布上的图像
        if self.canvas_image:
            self.canvas.move(self.canvas_image, dx, dy)
            # 更新偏移量
            self.canvas_offset_x += dx
            self.canvas_offset_y += dy
            
            # 如果有选择区域，也一起移动
            if self.selection_start and self.selection_end:
                self.selection_start = (self.selection_start[0] + dx, self.selection_start[1] + dy)
                self.selection_end = (self.selection_end[0] + dx, self.selection_end[1] + dy)
                self.draw_selection_rect()
        
        # 更新拖动起点
        self.pan_start = (event.x, event.y)
    
    def on_right_mouse_up(self, event):
        """右键释放事件"""
        if self.is_panning:
            self.is_panning = False
            self.pan_start = None
            self.canvas.config(cursor="")
            self.status_bar.config(text="拖动完成")
    
    def draw_selection_rect(self):
        """绘制选择矩形"""
        if not self.selection_start or not self.selection_end:
            return
        
        # 删除旧的选择矩形
        self.canvas.delete("selection")
        
        x1, y1 = self.selection_start
        x2, y2 = self.selection_end
        
        # 绘制半透明填充
        self.canvas.create_rectangle(
            x1, y1, x2, y2,
            fill="lightblue", stipple="gray25",
            tags="selection"
        )
        
        # 绘制边框
        self.canvas.create_rectangle(
            x1, y1, x2, y2,
            outline="blue", width=2,
            tags="selection"
        )
        
        # 确保选择矩形在最上层
        self.canvas.tag_raise("selection")
    
    def clear_selection(self, event=None):
        """清除选择"""
        self.selection_start = None
        self.selection_end = None
        self.is_selecting = False
        self.is_dragging = False
        self.drag_start = None
        self.screenshot = None
        self.screenshot_path = None
        # 重置右键拖动状态
        self.is_panning = False
        self.pan_start = None
        self.canvas_offset_x = 0
        self.canvas_offset_y = 0
        self.canvas.delete("selection")
        self.canvas.config(cursor="")
        self.status_bar.config(text="已清除选择")
        self.check_add_button_state()
    
    def capture_selected_area(self):
        """截取选中的PDF区域并保存到本地"""
        if not self.selection_start or not self.selection_end or not self.pdf_document:
            return
        
        try:
            # 获取画布上的图像位置
            if not self.canvas_image:
                return
            
            canvas_coords = self.canvas.coords(self.canvas_image)
            if not canvas_coords:
                return
            
            img_x, img_y = canvas_coords
            
            # 转换为相对于图像的坐标
            x1, y1 = self.selection_start
            x2, y2 = self.selection_end
            
            # 确保坐标在图像范围内
            x1 = max(img_x, min(x1, img_x + self.tk_image.width()))
            y1 = max(img_y, min(y1, img_y + self.tk_image.height()))
            x2 = max(img_x, min(x2, img_x + self.tk_image.width()))
            y2 = max(img_y, min(y2, img_y + self.tk_image.height()))
            
            # 转换为相对于图像的坐标
            rel_x1 = (x1 - img_x) / self.scale_factor
            rel_y1 = (y1 - img_y) / self.scale_factor
            rel_x2 = (x2 - img_x) / self.scale_factor
            rel_y2 = (y2 - img_y) / self.scale_factor
            
            # 确保坐标有序
            min_x, max_x = min(rel_x1, rel_x2), max(rel_x1, rel_x2)
            min_y, max_y = min(rel_y1, rel_y2), max(rel_y1, rel_y2)
            
            # 检查选择区域是否有效
            if max_x - min_x < 1 or max_y - min_y < 1:
                self.status_bar.config(text="选择区域太小")
                return
            
            # 获取PDF页面
            page = self.pdf_document[self.current_page]
            
            # 创建裁剪区域
            clip_rect = fitz.Rect(min_x, min_y, max_x, max_y)
            
            # 创建新页面并裁剪
            new_page = fitz.open()
            new_page.new_page(width=clip_rect.width, height=clip_rect.height)
            new_page[0].show_pdf_page(
                fitz.Rect(0, 0, clip_rect.width, clip_rect.height),
                self.pdf_document, self.current_page, clip=clip_rect
            )
            
            # 渲染为图像，使用高DPI提升清晰度
            pix = new_page[0].get_pixmap(dpi=300)  # 使用300 DPI提高图像质量
            self.screenshot = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            
            new_page.close()
            
            # 确定图片存储路径
            if self.custom_image_path:
                # 使用自定义路径
                images_dir = self.custom_image_path
            else:
                # 使用默认路径：PDF所在目录下的images文件夹
                images_dir = os.path.join(os.path.dirname(self.pdf_path), "images")
            
            # 创建文件夹（如果不存在）
            os.makedirs(images_dir, exist_ok=True)
            
            # 生成图片文件名：PDF文件名 + 时间戳
            pdf_name = os.path.splitext(os.path.basename(self.pdf_path))[0]
            import time
            timestamp = time.strftime("%Y%m%d_%H%M%S")
            img_filename = f"{pdf_name}_{timestamp}.png"
            img_path = os.path.join(images_dir, img_filename)
            
            # 根据用户设置决定是否保存图片到本地
            if self.save_image_locally:
                # 保存图片到本地
                self.screenshot.save(img_path, "PNG")
                self.screenshot_path = img_path  # 保存图片路径
                
                # 显示保存路径信息
                path_info = f"自定义路径: {self.custom_image_path}" if self.custom_image_path else f"默认路径"
                self.status_bar.config(text=f"已截取区域并保存: {int(clip_rect.width)}x{int(clip_rect.height)} 像素 -> {img_filename} ({path_info})")
            else:
                # 不保存图片到本地，但保留截图数据用于Anki
                self.screenshot_path = None
                self.status_bar.config(text=f"已截取区域: {int(clip_rect.width)}x{int(clip_rect.height)} 像素 (未保存到本地)")
            
        except Exception as e:
            self.status_bar.config(text=f"截取区域时出错: {str(e)}")
            self.screenshot = None
            self.screenshot_path = None
    
    def check_add_button_state(self, event=None):
        """检查添加到Anki按钮的状态"""
        has_pdf = self.pdf_document is not None
        has_selection = self.selection_start is not None and self.selection_end is not None
        has_question = self.question_entry.get().strip() != ""
        has_screenshot = self.screenshot is not None  # 只要有截图数据即可，不要求必须保存到本地
        
        if has_pdf and has_selection and has_question and has_screenshot:
            self.add_to_anki_btn.config(state=tk.NORMAL)
        else:
            self.add_to_anki_btn.config(state=tk.DISABLED)
    
    def toggle_save_image(self):
        """切换图片保存状态"""
        self.save_image_locally = self.save_image_var.get()
        if self.save_image_locally:
            self.status_bar.config(text="图片保存功能已开启")
        else:
            self.status_bar.config(text="图片保存功能已关闭")
    
    def set_image_path(self):
        """设置自定义图片存储路径"""
        path = filedialog.askdirectory(
            title="选择图片存储路径",
            initialdir=os.path.dirname(self.pdf_path) if self.pdf_path else os.getcwd()
        )
        
        if path:
            self.custom_image_path = path
            self.status_bar.config(text=f"图片存储路径已设置为: {path}")
        else:
            # 如果用户取消选择，重置为默认路径
            self.custom_image_path = None
            self.status_bar.config(text="使用默认图片存储路径")
    
    def add_to_anki(self):
        """将问题和截取的PDF区域添加到Anki"""
        if not self.screenshot or not self.question_entry.get().strip():
            messagebox.showwarning("警告", "请输入问题并选择PDF区域")
            return
        
        try:
            # 根据是否保存到本地来处理图片数据
            if self.save_image_locally and self.screenshot_path:
                # 检查本地图片文件是否存在
                if not os.path.exists(self.screenshot_path):
                    raise Exception(f"本地图片文件不存在: {self.screenshot_path}")
                
                # 读取本地图片文件
                with open(self.screenshot_path, 'rb') as img_file:
                    img_data = img_file.read()
                
                # 获取图片文件名
                img_filename = os.path.basename(self.screenshot_path)
            else:
                # 直接使用内存中的截图数据
                import io
                img_buffer = io.BytesIO()
                self.screenshot.save(img_buffer, format="PNG")
                img_data = img_buffer.getvalue()
                
                # 生成唯一的文件名
                import uuid
                img_filename = f"screenshot_{uuid.uuid4().hex[:8]}.png"
            
            # 首先通过AnkiConnect添加图片
            import base64
            response = requests.post(
                "http://localhost:8765",
                json={
                    "action": "storeMediaFile",
                    "version": 6,
                    "params": {
                        "filename": img_filename,
                        "data": base64.b64encode(img_data).decode('utf-8')
                    }
                },
                timeout=10
            )
            
            response.raise_for_status()
            result = response.json()
            
            if result.get("error") is not None:
                raise Exception(f"添加图片失败: {result['error']}")
            
            # 然后创建Anki卡片
            question = self.question_entry.get().strip()
            answer = f'<img src="{img_filename}">'  # 使用HTML格式插入图片
            
            # 验证问题和答案不为空
            if not question:
                raise Exception("问题不能为空")
            if not answer:
                raise Exception("答案不能为空")

            # 首先获取可用的模型名称
            response = requests.post(
                "http://localhost:8765",
                json={
                    "action": "modelNames",
                    "version": 6
                },
                timeout=10
            )
            response.raise_for_status()
            models_result = response.json()
            
            if models_result.get("error") is not None:
                raise Exception(f"获取模型列表失败: {models_result['error']}")
            
            available_models = models_result.get("result", [])
            
            # 尝试找到合适的模型
            model_name = None
            preferred_models = ["问答题", "Basic", "基本", "Cloze", "填空"]
            
            for preferred in preferred_models:
                if preferred in available_models:
                    model_name = preferred
                    break
            
            if not model_name and available_models:
                model_name = available_models[0]  # 使用第一个可用模型
            
            if not model_name:
                raise Exception("Anki中没有找到可用的卡片模型")
            
            # 获取模型的字段信息
            response = requests.post(
                "http://localhost:8765",
                json={
                    "action": "modelFieldNames",
                    "version": 6,
                    "params": {
                        "modelName": model_name
                    }
                },
                timeout=10
            )
            response.raise_for_status()
            fields_result = response.json()
            
            if fields_result.get("error") is not None:
                raise Exception(f"获取字段信息失败: {fields_result['error']}")
            
            field_names = fields_result.get("result", [])
            
            # 根据可用字段确定字段映射
            fields = {}
            if len(field_names) >= 2:
                # 如果有至少两个字段，使用前两个字段
                fields[field_names[0]] = question
                fields[field_names[1]] = answer
            elif len(field_names) == 1:
                # 如果只有一个字段，合并问题和答案
                fields[field_names[0]] = f"{question}\n\n{answer}"
            else:
                raise Exception("模型没有可用的字段")
            
            # 使用找到的模型添加卡片
            response = requests.post(
                "http://localhost:8765",
                json={
                    "action": "addNote",
                    "version": 6,
                    "params": {
                        "note": {
                            "deckName": "默认",
                            "modelName": model_name,
                            "fields": fields,
                            "tags": ["PDF截取"]
                        }
                    }
                },
                timeout=10
            )
            
            response.raise_for_status()
            result = response.json()
            
            if result.get("error") is not None:
                error_msg = result['error']
                if 'model was not found' in error_msg:
                    raise Exception(
                        f"添加卡片失败: {error_msg}\n\n"
                        f"使用的模型: {model_name}\n"
                        f"可用模型: {', '.join(available_models)}\n\n"
                        "请确保Anki中存在正确的卡片模型。"
                    )
                elif 'cannot create note because it is empty' in error_msg:
                    raise Exception(
                        f"添加卡片失败: {error_msg}\n\n"
                        f"使用的模型: {model_name}\n"
                        f"字段映射: {fields}\n"
                        f"可用字段: {', '.join(field_names)}\n\n"
                        "请检查模型字段配置是否正确。"
                    )
                else:
                    raise Exception(f"添加卡片失败: {error_msg}\n\n使用的模型: {model_name}")
            
            # 成功添加卡片
            save_path_info = f"图片已保存到: {self.screenshot_path}" if self.screenshot_path else "图片未保存到本地"
            messagebox.showinfo("成功", f"卡片已成功添加到Anki!\n使用的模型: {model_name}\n{save_path_info}")
            
            # 清除问题输入和选择区域，让用户可以继续框选
            self.question_entry.delete(0, tk.END)
            self.clear_selection()
            self.status_bar.config(text="卡片添加成功，可以继续框选新区域")
            
        except requests.exceptions.ConnectionError:
            messagebox.showerror(
                "连接错误",
                "无法连接到AnkiConnect。\n\n"
                "请确保：\n"
                "1. Anki已打开\n"
                "2. AnkiConnect插件已安装并启用\n"
                "3. AnkiConnect运行在localhost:8765"
            )
        except requests.exceptions.Timeout:
            messagebox.showerror("超时错误", "连接AnkiConnect超时，请重试。")
        except Exception as e:
            messagebox.showerror("错误", f"添加到Anki时出错: {str(e)}")

if __name__ == "__main__":
    # 检查是否安装了必要的库
    required_libraries = {
        "fitz": "PyMuPDF",
        "requests": "requests",
        "PIL": "Pillow"
    }
    
    missing_libraries = []
    for lib, package in required_libraries.items():
        try:
            __import__(lib)
        except ImportError:
            missing_libraries.append(package)
    
    if missing_libraries:
        print(f"请先安装以下缺失的库: {' '.join([f'pip install {lib}' for lib in missing_libraries])}")
    else:
        root = tk.Tk()
        app = PDFAnkiTool(root)
        root.mainloop()
