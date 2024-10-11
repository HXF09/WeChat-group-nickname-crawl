import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
from pywinauto.application import Application
import pywinauto
import time
import psutil
import pandas as pd
import numpy as np


def get_wechat_pid():
    pids = psutil.pids()
    for pid in pids:
        p = psutil.Process(pid)
        if p.name() == 'WeChat.exe':
            return pid
    return None


def get_name_list(pid, label_var):
    message = '>>> WeChat.exe pid: {}'.format(pid)
    label_var.set(message)
    print(message)  # 控制台输出

    message = '>>> 按照提示操作，打开【微信------->目标群聊------->聊天成员】，然后点击【查看更多】，否则查找不全！'
    label_var.set(message)
    print(message)  # 控制台输出

    for i in range(20):
        message = '请等待... ({:2d} 秒)'.format(20 - i)
        label_var.set(message)
        print(message)  # 控制台输出
        time.sleep(1)

    app = Application(backend='uia').connect(process=pid)
    win_main_Dialog = app.window(class_name='WeChatMainWndForPC')
    chat_list = win_main_Dialog.child_window(control_type='List', title='聊天成员')

    name_list = []
    all_members = []

    for i in chat_list.items():
        p = i.descendants()
        if p and len(p) > 5:
            if p[5].texts() and p[5].texts()[0].strip() != '' and (
                    p[5].texts()[0].strip() != '添加' and p[5].texts()[0].strip() != '移出'):
                name_list.append(p[5].texts()[0].strip())
                all_members.append([p[5].texts()[0].strip(), p[3].texts()[0].strip()])

    save_path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel files', '*.xlsx')],
                                             title="保存群成员列表")

    # 将群成员信息保存为Excel文件
    df = pd.DataFrame(np.array(all_members), columns=['群昵称', '微信昵称'])
    df.to_excel(save_path, index=False)

    message = '>>> 群成员共 {} 人，结果已保存至 {}'.format(len(name_list), save_path)
    label_var.set(message)
    print(message)  # 控制台输出


def match(label_var):
    pid = get_wechat_pid()
    if pid is None:
        error_message = '>>> 找不到 WeChat.exe，请先打开 WeChat.exe 再运行此程序！'
        messagebox.showerror('错误', error_message)
        print(error_message)  # 控制台输出
        return

    try:
        get_name_list(pid, label_var)
    except pywinauto.findwindows.ElementNotFoundError:
        error_message = '>>> 未找到【聊天成员】窗口，程序终止！\n请重启微信或检查群聊界面。'
        messagebox.showerror('错误', error_message)
        print(error_message)  # 控制台输出


def create_gui():
    # 创建主窗口
    root = tk.Tk()
    root.title('微信群成员获取工具')

    # 标签用于显示提示信息
    label_var = tk.StringVar()
    label = tk.Label(root, textvariable=label_var, justify='left')
    label.pack(pady=10)

    # 按钮用于启动匹配
    start_button = tk.Button(root, text='获取群成员', command=lambda: match(label_var))
    start_button.pack(pady=20)

    # 退出按钮
    exit_button = tk.Button(root, text='退出', command=root.quit)
    exit_button.pack(pady=10)

    # 运行主循环
    root.mainloop()


if __name__ == '__main__':
    create_gui()
