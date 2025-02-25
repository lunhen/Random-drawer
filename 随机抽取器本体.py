import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import random
import pyttsx3
from PIL import Image, ImageTk
import json
import os
import shutil


class RandomPicker:
    def __init__(self, root):
        self.root = root
        self.root.title("随机抽取器")
        self.root.geometry("800x600")

        # 设置当前工作目录为程序所在目录
        os.chdir(os.path.dirname(os.path.abspath(__file__)))

        # 初始化变量
        self.background_image = None
        self.background_image_path = None  # 背景图片路径
        self.names = {}
        self.original_names = {}  # 用于保存分组初始名单
        self.current_group = None
        self.enable_tts = tk.BooleanVar(value=False)
        self.single_pick = tk.BooleanVar(value=False)  # 允许单次点名，默认为False（允许重复点名）
        self.is_picking = False  # 是否正在抽取
        self.pick_interval = 100  # 抽取间隔时间（毫秒）
        self.last_picked_name = None  # 记录最后一次抽中的姓名

        # 图片保存目录
        self.image_dir = os.path.join(os.path.dirname(__file__), "images")
        os.makedirs(self.image_dir, exist_ok=True)  # 创建图片目录
        self.default_image_path = os.path.join(self.image_dir, "background.jpg")  # 默认图片路径

        # 加载用户设置
        self.settings_file = "user_settings.json"
        self.load_settings()

        # 创建界面
        self.create_widgets()

        # 绑定窗口关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def create_widgets(self):
        # 背景图片设置
        self.bg_label = tk.Label(self.root)
        self.bg_label.place(relwidth=1, relheight=1)  # 背景图片铺满窗口

        # 如果背景图片路径存在，加载背景图片
        if self.background_image_path and os.path.exists(self.background_image_path):
            self.load_background(self.background_image_path)

        # 顶部功能区
        self.top_frame = tk.Frame(self.root)
        self.top_frame.pack(fill=tk.X, padx=10, pady=10)

        self.load_bg_button = tk.Button(self.top_frame, text="加载背景图片", command=self.load_background_dialog)
        self.load_bg_button.pack(side=tk.LEFT, padx=5)

        self.add_group_button = tk.Button(self.top_frame, text="添加分组", command=self.add_group)
        self.add_group_button.pack(side=tk.LEFT, padx=5)

        self.load_names_button = tk.Button(self.top_frame, text="加载姓名文件", command=self.load_names)
        self.load_names_button.pack(side=tk.LEFT, padx=5)

        self.group_label = tk.Label(self.top_frame, text="分组:")
        self.group_label.pack(side=tk.LEFT, padx=5)

        self.group_combobox = ttk.Combobox(self.top_frame, state="readonly")
        self.group_combobox.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.group_combobox.bind("<<ComboboxSelected>>", self.on_group_select)

        # 如果分组信息存在，加载分组
        if self.names:
            self.group_combobox['values'] = list(self.names.keys())
            if self.current_group:
                self.group_combobox.current(list(self.names.keys()).index(self.current_group))

        # 抽取结果显示
        self.result_label = tk.Label(self.root, text="", font=("Arial", 200))
        self.result_label.pack(pady=20)

        # 控制按钮
        self.control_frame = tk.Frame(self.root)
        self.control_frame.pack(pady=10)

        self.start_stop_button = tk.Button(self.control_frame, text="开始抽取", command=self.toggle_pick)
        self.start_stop_button.pack(side=tk.LEFT, padx=5)

        self.reset_button = tk.Button(self.control_frame, text="重置", command=self.reset_group)
        self.reset_button.pack(side=tk.LEFT, padx=5)

        # 设置
        self.settings_frame = tk.Frame(self.root)
        self.settings_frame.pack(pady=10)

        self.tts_checkbutton = tk.Checkbutton(self.settings_frame, text="启用语音朗读", variable=self.enable_tts)
        self.tts_checkbutton.pack(side=tk.LEFT, padx=5)

        self.single_pick_checkbutton = tk.Checkbutton(
            self.settings_frame, text="允许单次点名", variable=self.single_pick
        )
        self.single_pick_checkbutton.pack(side=tk.LEFT, padx=5)

    def load_background_dialog(self):
        file_path = filedialog.askopenfilename(filetypes=[("Image Files", "*.png;*.jpg;*.jpeg")])
        if file_path:
            self.load_background(file_path)

    def load_background(self, file_path):
        try:
            # 检查用户选择的文件路径是否与目标路径相同
            if os.path.abspath(file_path) == os.path.abspath(self.default_image_path):
                # 如果相同，直接加载图片
                self.background_image_path = self.default_image_path
            else:
                # 如果不同，将图片复制到程序目录下的images文件夹
                shutil.copy(file_path, self.default_image_path)
                self.background_image_path = self.default_image_path
                print("图片已复制到:", self.default_image_path)

            # 使用PIL加载图片并缩放至窗口大小
            image = Image.open(self.default_image_path)
            image = image.resize((self.root.winfo_width(), self.root.winfo_height()), Image.Resampling.LANCZOS)
            self.background_image = ImageTk.PhotoImage(image)
            self.bg_label.config(image=self.background_image)
        except Exception as e:
            messagebox.showerror("错误", f"无法加载图片: {e}")

    def add_group(self):
        group_name = simpledialog.askstring("添加分组", "请输入分组名称:")
        if group_name:
            self.names[group_name] = []
            self.original_names[group_name] = []  # 初始化原始名单
            self.group_combobox['values'] = list(self.names.keys())
            self.group_combobox.current(len(self.names) - 1)
            self.current_group = group_name

    def load_names(self):
        if not self.current_group:
            messagebox.showerror("错误", "请先选择一个分组")
            return

        file_path = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt")])
        if file_path:
            with open(file_path, 'r', encoding='utf-8') as file:
                self.names[self.current_group] = file.read().splitlines()
                self.original_names[self.current_group] = self.names[self.current_group].copy()  # 保存原始名单
            messagebox.showinfo("成功", f"已加载 {len(self.names[self.current_group])} 个姓名到分组 {self.current_group}")

    def on_group_select(self, event):
        self.current_group = self.group_combobox.get()

    def toggle_pick(self):
        if not self.current_group or not self.names[self.current_group]:
            messagebox.showerror("错误", "请先加载姓名并选择一个分组")
            return

        self.is_picking = not self.is_picking  # 切换抽取状态
        if self.is_picking:
            self.start_stop_button.config(text="停止抽取")
            self.pick_name()  # 开始抽取
        else:
            self.start_stop_button.config(text="开始抽取")
            if self.enable_tts.get():  # 仅在停止时朗读
                engine = pyttsx3.init()
                engine.say(f"{self.result_label.cget('text')}")
                engine.runAndWait()

            # 如果允许单次点名，则在抽取结束后移除被抽中的姓名
            if self.single_pick.get() and self.last_picked_name:
                self.names[self.current_group].remove(self.last_picked_name)
                self.last_picked_name = None  # 重置记录

    def pick_name(self):
        if self.is_picking:
            if not self.names[self.current_group]:
                messagebox.showinfo("提示", "当前分组已无更多姓名")
                self.is_picking = False
                self.start_stop_button.config(text="开始抽取")
                return

            name = random.choice(self.names[self.current_group])
            self.result_label.config(text=name)
            self.last_picked_name = name  # 记录最后一次抽中的姓名

            # 继续抽取
            self.root.after(self.pick_interval, self.pick_name)

    def reset_group(self):
        if not self.current_group:
            messagebox.showerror("错误", "请先选择一个分组")
            return

        # 将当前分组的名单恢复到初始状态
        if self.current_group in self.original_names:
            self.names[self.current_group] = self.original_names[self.current_group].copy()
            messagebox.showinfo("成功", f"已重置分组 {self.current_group} 的姓名列表")
        else:
            messagebox.showerror("错误", "当前分组没有初始名单")

    def load_settings(self):
        if os.path.exists(self.settings_file):
            try:
                with open(self.settings_file, 'r', encoding='utf-8') as file:
                    settings = json.load(file)
                    self.background_image_path = settings.get("background_image_path")
                    self.names = settings.get("names", {})
                    self.original_names = settings.get("original_names", {})
                    self.current_group = settings.get("current_group")
                    self.enable_tts.set(settings.get("enable_tts", False))
                    self.single_pick.set(settings.get("single_pick", False))
                print("加载设置成功，图片路径:", self.background_image_path)
            except Exception as e:
                messagebox.showerror("错误", f"加载设置失败: {e}")

    def save_settings(self):
        settings = {
            "background_image_path": self.background_image_path,
            "names": self.names,
            "original_names": self.original_names,
            "current_group": self.current_group,
            "enable_tts": self.enable_tts.get(),
            "single_pick": self.single_pick.get(),
        }
        try:
            with open(self.settings_file, 'w', encoding='utf-8') as file:
                json.dump(settings, file, ensure_ascii=False, indent=4)
            print("保存设置成功，图片路径:", self.background_image_path)
        except Exception as e:
            messagebox.showerror("错误", f"保存设置失败: {e}")

    def on_close(self):
        # 保存设置并退出
        self.save_settings()
        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = RandomPicker(root)
    root.mainloop()