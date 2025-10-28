import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import os
import csv


class CsvBatchConverter:
    def __init__(self, root):
        self.root = root
        self.root.title('CSV 批次轉換工具')
        self.root.geometry('500x390')
        self.root.configure(bg='#f4f6fa')
        self.style = ttk.Style()
        self._set_theme()
        self.selected_files = []
        self.remove_columns_var = tk.StringVar()
        self.remove_columns_var.set('')
        self.file_list_var = tk.StringVar()
        self.file_list_var.set('尚未選取檔案')
        self.output_dir_var = tk.StringVar()
        self.output_dir_var.set('')
        self.output_format_var = tk.StringVar()
        self.output_format_var.set('txt')
        self.delimiter_var = tk.StringVar()
        self.delimiter_var.set('逗號(,)')
        self._build_gui()

    def _build_gui(self):
        font_title = ('Microsoft JhengHei', 18, 'bold')
        font_label = ('Microsoft JhengHei', 12)
        font_entry = ('Microsoft JhengHei', 12)

        # 標題
        title = tk.Label(self.root, text='CSV 批次轉換工具', font=font_title, bg='#f4f6fa', fg='#22223b')
        title.pack(pady=(18, 10))

        # 選檔案
        select_btn = ttk.Button(self.root, text='選取 CSV 檔案', command=self.select_files)
        select_btn.pack(pady=(5, 2))

        file_list_label = tk.Label(self.root, textvariable=self.file_list_var, font=font_label, bg='#f4f6fa', fg='#4b5563')
        file_list_label.pack(pady=(0, 10))

        # 輸出資料夾選擇
        output_dir_frame = tk.Frame(self.root, bg='#f4f6fa')
        output_dir_frame.pack(pady=5, fill='x', padx=40)
        output_dir_label = tk.Label(output_dir_frame, text='輸出資料夾:', font=font_label, bg='#f4f6fa')
        output_dir_label.pack(side=tk.LEFT)
        output_dir_entry = ttk.Entry(output_dir_frame, textvariable=self.output_dir_var, width=28, font=font_entry)
        output_dir_entry.pack(side=tk.LEFT, padx=4)
        output_dir_btn = ttk.Button(output_dir_frame, text='選擇', command=self.select_output_dir)
        output_dir_btn.pack(side=tk.LEFT, padx=4)

        # 欄位刪除
        remove_frame = tk.Frame(self.root, bg='#f4f6fa')
        remove_frame.pack(pady=5, fill='x', padx=40)
        remove_label = tk.Label(remove_frame, text='要刪除的欄位標題（逗號分隔）:', font=font_label, bg='#f4f6fa')
        remove_label.pack(side=tk.LEFT)
        remove_entry = ttk.Entry(remove_frame, textvariable=self.remove_columns_var, width=22, font=font_entry)
        remove_entry.pack(side=tk.LEFT, padx=4)

        # 輸出格式
        format_frame = tk.Frame(self.root, bg='#f4f6fa')
        format_frame.pack(pady=5, fill='x', padx=40)
        format_label = tk.Label(format_frame, text='輸出格式:', font=font_label, bg='#f4f6fa')
        format_label.pack(side=tk.LEFT)
        format_options = ['txt', 'tsv', 'csv']
        format_menu = ttk.Combobox(format_frame, textvariable=self.output_format_var, values=format_options, font=font_entry, width=8, state='readonly')
        format_menu.pack(side=tk.LEFT, padx=4)

        # 分隔符號
        delim_frame = tk.Frame(self.root, bg='#f4f6fa')
        delim_frame.pack(pady=5, fill='x', padx=40)
        delim_label = tk.Label(delim_frame, text='分隔符號:', font=font_label, bg='#f4f6fa')
        delim_label.pack(side=tk.LEFT)
        delim_options = ['逗號(,)', 'Tab(\t)', '分號(;)']
        delim_menu = ttk.Combobox(delim_frame, textvariable=self.delimiter_var, values=delim_options, font=font_entry, width=10, state='readonly')
        delim_menu.pack(side=tk.LEFT, padx=4)

        # 轉換按鈕
        convert_btn = ttk.Button(self.root, text='開始轉換', command=self.convert_files)
        convert_btn.pack(pady=(18, 10))

    def _set_theme(self):
        # 設定 ttk 樣式
        self.style.theme_use('clam')
        self.style.configure('TButton', font=('Microsoft JhengHei', 12, 'bold'), background='#3b82f6', foreground='white', borderwidth=0, focusthickness=3, focuscolor='none', padding=6)
        self.style.map('TButton', background=[('active', '#2563eb')])
        self.style.configure('TEntry', font=('Microsoft JhengHei', 12))
        self.style.configure('TCombobox', font=('Microsoft JhengHei', 12))

    def select_output_dir(self):
        dir_path = filedialog.askdirectory(title='選擇輸出資料夾')
        if dir_path:
            self.output_dir_var.set(dir_path)

    def select_files(self):
        files = filedialog.askopenfilenames(
            title='選取 CSV 檔案',
            filetypes=[('CSV Files', '*.csv')]
        )
        self.selected_files = list(files)
        self.file_list_var.set(f'已選取 {len(self.selected_files)} 個檔案')

    def convert_files(self):
        if not self.selected_files:
            messagebox.showwarning('警告', '請先選取檔案！')
            return
        output_dir = self.output_dir_var.get().strip()
        if not output_dir:
            messagebox.showwarning('警告', '請先選擇輸出資料夾！')
            return
        remove_columns = [col.strip() for col in self.remove_columns_var.get().split(',') if col.strip()]
        output_format = self.output_format_var.get()
        delim_map = {'逗號(,)': ',', 'Tab(\t)': '\t', '分號(;)': ';'}
        delimiter = delim_map.get(self.delimiter_var.get(), ',')
        ext_map = {'txt': '.txt', 'tsv': '.tsv', 'csv': '.csv'}
        ext = ext_map.get(output_format, '.txt')
        success_count = 0
        for file_path in self.selected_files:
            try:
                with open(file_path, 'r', encoding='utf-8', errors='ignore', newline='') as csvfile:
                    reader = csv.DictReader(csvfile)
                    if not reader.fieldnames:
                        csvfile.seek(0)
                        content = csvfile.read()
                        base_name = os.path.splitext(os.path.basename(file_path))[0]
                        out_path = os.path.join(output_dir, base_name + ext)
                        with open(out_path, 'w', encoding='utf-8') as f:
                            f.write(content)
                        success_count += 1
                        continue
                    keep_fields = [f for f in reader.fieldnames if f not in remove_columns]
                    rows = [ {k: row[k] for k in keep_fields} for row in reader ]
                base_name = os.path.splitext(os.path.basename(file_path))[0]
                out_path = os.path.join(output_dir, base_name + ext)
                with open(out_path, 'w', encoding='utf-8', newline='') as out_file:
                    writer = csv.DictWriter(out_file, fieldnames=keep_fields, delimiter=delimiter)
                    writer.writeheader()
                    writer.writerows(rows)
                success_count += 1
            except Exception as e:
                print(f'轉換失敗: {file_path}\n錯誤: {e}')
        messagebox.showinfo('完成', f'成功轉換 {success_count} 個檔案！')


def main():
    root = tk.Tk()
    app = CsvBatchConverter(root)
    root.mainloop()


if __name__ == '__main__':
    main()



# 舊的全域 GUI 變數與程式碼已移除，請使用 main() 入口
