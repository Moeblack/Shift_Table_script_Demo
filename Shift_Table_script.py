import os
import re
import pandas as pd
from datetime import datetime, timedelta
from PIL import Image
from photoshop import Session
from psd_tools import PSDImage
import tkinter as tk
from tkinter import filedialog
import tkinter.ttk as ttk

def process_name(name):
    match = re.search(r"\s*([\u4e00-\u9fa5]{1,3}\s*[\u4e00-\u9fa5]{0,3}\s*[\u4e00-\u9fa5]{0,3})\（\d{11}\）", name)
    if match:
        return match.group(1).strip().replace(' ', '')
    return None

def main(excel_file, template_path, output_path, progress_var, root):
    df = pd.read_excel(excel_file, header=None)

    valid_rows = sum(1 for _, row in df.iterrows() if process_name(row[1]) and process_name(row[2]))

    processed_rows = 0
    total_rows = len(df)
    for _, row in df.iterrows():

        EXCEL_EPOCH = datetime(1899, 12, 30)
        date = EXCEL_EPOCH + timedelta(days=int(row[0]))
        day_shift_name = process_name(row[1])
        night_shift_name = process_name(row[2])

        if not day_shift_name or not night_shift_name:
            continue

        with Session(template_path) as psd:
            def find_layer_by_name(layers, layer_name):
                for layer in layers:
                    if layer.name == layer_name:
                        return layer
                return None

            day_shift_layer = find_layer_by_name(psd.active_document.layers, "白班")
            night_shift_layer = find_layer_by_name(psd.active_document.layers, "夜班")

            if not day_shift_layer or not night_shift_layer:
                print("未找到白班或夜班图层，请检查Photoshop模板")
                return

            if day_shift_layer.kind == 2:
                day_shift_layer.textItem.contents = day_shift_name
                day_shift_layer.textItem.size = 48

            if night_shift_layer.kind == 2:
                night_shift_layer.textItem.contents = night_shift_name
                night_shift_layer.textItem.size = 48

            if not os.path.exists(output_path):
                os.makedirs(output_path)

            filename = f"{date.strftime('%Y%m%d')}.png"
            output_file = os.path.join(output_path, filename)
            psd.active_document.saveAs(output_file, psd.JPEGSaveOptions())

        processed_rows += 1
        progress_var.set(int(processed_rows / valid_rows * 100))
        root.update_idletasks()
        root.update()


def browse_file(entry):
    file_path = filedialog.askopenfilename()
    entry.delete(0, tk.END)
    entry.insert(0, file_path)

def browse_folder(entry):
    folder_path = filedialog.askdirectory()
    entry.delete(0, tk.END)
    entry.insert(0, folder_path)

def create_gui():
    root = tk.Tk()
    root.title("排班表处理")

    excel_label = tk.Label(root, text="Excel文件路径：")
    excel_label.grid(row=0, column=0, sticky="e")
    excel_entry = tk.Entry(root, width=50)
    excel_entry.grid(row=0, column=1)
    excel_browse = tk.Button(root, text="浏览", command=lambda: browse_file(excel_entry))
    excel_browse.grid(row=0, column=2)

    template_label = tk.Label(root, text="模板文件路径：")
    template_label.grid(row=1, column=0, sticky="e")
    template_entry = tk.Entry(root, width=50)
    template_entry.grid(row=1, column=1)
    template_browse = tk.Button(root, text="浏览", command=lambda: browse_file(template_entry))
    template_browse.grid(row=1, column=2)

    output_label = tk.Label(root, text="输出文件夹路径：")
    output_label.grid(row=2, column=0, sticky="e")
    output_entry = tk.Entry(root, width=50)
    output_entry.grid(row=2, column=1)
    output_browse = tk.Button(root, text="浏览", command=lambda: browse_folder(output_entry))
    output_browse.grid(row=2, column=2)

    progress_var = tk.IntVar()
    progress_label = tk.Label(root, text="进度：")
    progress_label.grid(row=3, column=0, sticky="e")
    progress_bar = tk.ttk.Progressbar(root, variable=progress_var, maximum=100)
    progress_bar.grid(row=3, column=1, columnspan=2, sticky="we")

    start_button = tk.Button(root, text="开始", command=lambda: main(
        excel_entry.get(),
        template_entry.get(),
        output_entry.get(),
        progress_var,
        root
    ))
    start_button.grid(row=4, column=0, columnspan=3)
    tips_label = tk.Label(
        root,
        text="Tips：在使用本软件之前，请打开Photoshop，注意格式化排班表，修改模板的带班领导名称",
        fg="red",
    )
    tips_label.grid(row=5, column=0, columnspan=3)

    root.mainloop()

if __name__ == "__main__":
    create_gui()

