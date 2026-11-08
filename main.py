# 示例主程序：本地离线，不联网
# 把这个文件替换成你自己的程序源码（比如 Mini Assistant 的 main.py）
# 保持文件名 main.py 不变，打包脚本和 GitHub Actions 都不用改

import tkinter as tk
from tkinter import filedialog, messagebox


def main():
    root = tk.Tk()
    root.title("Mini Assistant - 本地版")
    root.geometry("520x320")

    label = tk.Label(root, text="本地离线工具，不联网也能用", font=("Microsoft YaHei", 12))
    label.pack(pady=20)

    def choose_file():
        path = filedialog.askopenfilename(title="选择文件")
        if path:
            messagebox.showinfo("已选择", f"已选择：\n{path}\n\n这里接你原来的文件处理逻辑。")

    btn = tk.Button(root, text="选择文件", command=choose_file, width=20, height=2)
    btn.pack(pady=10)

    tip = tk.Label(root, text="把 main.py 换成你自己的源码，\n再推送到 GitHub，就会自动打出 exe。", fg="gray")
    tip.pack(pady=10)

    root.mainloop()


if __name__ == "__main__":
    main()
