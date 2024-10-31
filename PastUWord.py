import tkinter as tk
from tkinter import ttk
import keyboard
import win32clipboard
import win32gui
import win32api
import win32con
import threading
import time
import sys
from collections import deque


class PasteUWord:
    def __init__(self):
        self.clipboard_history = deque(maxlen=10)
        self.window = None
        self.is_window_showing = False
        self.last_active_window = None

        # 创建主窗口
        self.create_window()
        # 初始隐藏窗口
        self.window.withdraw()

        # 启动剪贴板监听线程
        self.clipboard_thread = threading.Thread(target=self.monitor_clipboard, daemon=True)
        self.clipboard_thread.start()

        # 注册全局快捷键
        keyboard.add_hotkey('ctrl+`', self.toggle_window)

    def get_clipboard_content(self):
        try:
            win32clipboard.OpenClipboard()
            data = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
            win32clipboard.CloseClipboard()
            return data
        except:
            return ""

    def set_clipboard_content(self, text):
        try:
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardText(text, win32clipboard.CF_UNICODETEXT)
            win32clipboard.CloseClipboard()
        except:
            pass

    def monitor_clipboard(self):
        last_content = ""
        while True:
            current_content = self.get_clipboard_content()
            if current_content and current_content != last_content:
                self.clipboard_history.append(current_content)
                last_content = current_content
                self.window.after(0, self.update_text_display)
            time.sleep(0.5)

    def create_window(self):
        self.window = tk.Tk()
        self.window.title("PasteUWord")
        self.window.geometry("600x800")

        # 创建框架来容纳文本区域和滚动条
        frame = ttk.Frame(self.window)
        frame.pack(expand=True, fill='both', padx=5, pady=5)

        # 创建文本显示区域
        self.text_area = tk.Text(frame, wrap=tk.WORD)
        self.text_area.pack(side=tk.LEFT, expand=True, fill='both')

        # 创建滚动条
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.text_area.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.text_area.configure(yscrollcommand=scrollbar.set)

        # 创建按钮框架
        button_frame = ttk.Frame(self.window)
        button_frame.pack(pady=5)

        # 创建粘贴按钮
        self.paste_button = ttk.Button(button_frame, text="PasteSelected", command=self.paste_selected)
        self.paste_button.pack(side=tk.LEFT, padx=5)

        # 创建退出按钮
        self.quit_button = ttk.Button(button_frame, text="QUIT", command=self.quit_application)
        self.quit_button.pack(side=tk.LEFT, padx=5)

        # 绑定F9快捷键
        self.window.bind('<F9>', lambda e: self.paste_selected())

        # 绑定窗口关闭事件
        self.window.protocol("WM_DELETE_WINDOW", self.hide_window)

    def update_text_display(self):
        self.text_area.delete(1.0, tk.END)
        for item in reversed(self.clipboard_history):
            self.text_area.insert(tk.END, item + "\n\n")

    def paste_selected(self):
        try:
            # 获取选中的文本
            selected_text = self.text_area.get(tk.SEL_FIRST, tk.SEL_LAST)

            # 隐藏当前窗口
            self.hide_window()

            # 等待一小段时间确保窗口隐藏
            time.sleep(0.1)

            # 激活之前记录的窗口
            if self.last_active_window:
                win32gui.SetForegroundWindow(self.last_active_window)

                # 保存当前剪贴板内容
                original_clipboard = self.get_clipboard_content()

                # 设置新的剪贴板内容
                self.set_clipboard_content(selected_text)

                # 模拟Ctrl+V粘贴
                keyboard.press_and_release('ctrl+v')

                # 等待一小段时间确保粘贴完成
                time.sleep(0.1)

                # 恢复原来的剪贴板内容
                self.set_clipboard_content(original_clipboard)

        except tk.TclError:  # 没有选中文本的情况
            pass

    def toggle_window(self):
        if not self.is_window_showing:
            self.show_window()
        else:
            self.hide_window()

    def show_window(self):
        # 记录当前活动窗口
        self.last_active_window = win32gui.GetForegroundWindow()

        # 更新并显示窗口
        self.update_text_display()
        self.window.deiconify()
        self.is_window_showing = True
        self.window.lift()  # 将窗口置顶
        self.window.focus_force()  # 强制获取焦点

    def hide_window(self):
        self.window.withdraw()
        self.is_window_showing = False

    def quit_application(self):
        # 注销快捷键
        keyboard.unhook_all()
        # 销毁窗口
        self.window.destroy()
        # 退出程序
        sys.exit()

    def run(self):
        self.window.mainloop()


if __name__ == "__main__":
    app = PasteUWord()
    app.run()