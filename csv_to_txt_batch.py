
import ttkbootstrap as tb
from ttkbootstrap.dialogs import Messagebox
from tkinter import filedialog
import os
import csv
import pandas as pd


class CsvBatchConverter:
    def __init__(self, root):
        self.root = root
        self.root.title('CSV to TXT 批次轉換工具')
        self.root.geometry('500x540')
        self.selected_files = []
        self.remove_columns_var = tb.StringVar()
        self.remove_columns_var.set('')
        self.file_list_var = tb.StringVar()
        self.file_list_var.set('尚未選取檔案')
        self.output_dir_var = tb.StringVar()
        self.output_dir_var.set('Export_here')
        self.output_format_var = tb.StringVar()
        self.output_format_var.set('txt')
        self.delimiter_var = tb.StringVar()
        self.delimiter_var.set('逗號(,)')
        self.sheet_name_var = tb.StringVar()
        self.sheet_name_var.set('')
        self.merge_sheets_var = tb.BooleanVar()
        self.merge_sheets_var.set(False)
        self._setup_styles()
        self._build_gui()

    def _setup_styles(self):
        style = tb.Style.get_instance()
        style.configure('TCombobox', font=('Microsoft JhengHei', 12))
        style.configure('TCombobox.Listbox', background='#ffffff', font=('Microsoft JhengHei', 12))



    def _build_gui(self):
        # 標題
        title = tb.Label(self.root, text='CSV 批次轉換工具', font=('Microsoft JhengHei', 20, 'bold'))
        title.pack(pady=(18, 10))

        # 選檔案
        select_btn = tb.Button(self.root, text='選取 CSV/XLSX 檔案', command=self.select_files)
        select_btn.pack(pady=(5, 2))

        # 檔案清單區塊
        file_list_frame = tb.Labelframe(self.root, text='已選取檔案')
        file_list_frame.pack(pady=(0, 10), padx=40, fill='x')
        self.file_listbox = tb.ScrolledText(file_list_frame, height=4, font=('Microsoft JhengHei', 12), wrap='none', state='disabled')
        self.file_listbox.pack(fill='x', padx=4, pady=2)

        # 輸出資料夾選擇
        output_dir_frame = tb.Frame(self.root)
        output_dir_frame.pack(pady=5, fill='x', padx=40)
        output_dir_label = tb.Label(output_dir_frame, text='輸出資料夾:', font=('Microsoft JhengHei', 12))
        output_dir_label.pack(side='left')
        output_dir_entry = tb.Entry(output_dir_frame, textvariable=self.output_dir_var, width=28, font=('Microsoft JhengHei', 12))
        output_dir_entry.pack(side='left', padx=4)
        output_dir_btn = tb.Button(output_dir_frame, text='選擇', command=self.select_output_dir)
        output_dir_btn.pack(side='left', padx=4)

        # 欄位刪除
        remove_frame = tb.Frame(self.root)
        remove_frame.pack(pady=5, fill='x', padx=40)
        remove_label = tb.Label(remove_frame, text='要刪除的欄位標題（逗號分隔）:', font=('Microsoft JhengHei', 12))
        remove_label.pack(side='left')
        remove_entry = tb.Entry(remove_frame, textvariable=self.remove_columns_var, width=22, font=('Microsoft JhengHei', 12))
        remove_entry.pack(side='left', padx=4)

        # Sheet 選擇區塊
        sheet_frame = tb.Frame(self.root)
        sheet_frame.pack(pady=5, fill='x', padx=40)
        sheet_label = tb.Label(sheet_frame, text='Excel Sheet 名稱/索引 (空=預設):', font=('Microsoft JhengHei', 12))
        sheet_label.pack(side='left')
        sheet_entry = tb.Entry(sheet_frame, textvariable=self.sheet_name_var, width=18, font=('Microsoft JhengHei', 12))
        sheet_entry.pack(side='left', padx=4)

        merge_frame = tb.Frame(self.root)
        merge_frame.pack(pady=2, fill='x', padx=40)
        merge_check = tb.Checkbutton(merge_frame, text='合併所有 sheet', variable=self.merge_sheets_var)
        merge_check.pack(side='left')

        # 輸出格式與分隔符號同一行
        format_delim_frame = tb.Frame(self.root)
        format_delim_frame.pack(pady=5, fill='x', padx=40)

        format_label = tb.Label(format_delim_frame, text='輸出格式:', font=('Microsoft JhengHei', 12))
        format_label.pack(side='left')
        format_options = ['txt', 'tsv', 'csv']
        format_menu = tb.Combobox(format_delim_frame, textvariable=self.output_format_var, values=format_options, width=10, state='readonly')
        format_menu.pack(side='left', padx=4)

        delim_label = tb.Label(format_delim_frame, text='分隔符號:', font=('Microsoft JhengHei', 12))
        delim_label.pack(side='left', padx=(20,0))
        delim_options = ['逗號(,)', 'Tab(\t)', '分號(;)']
        delim_menu = tb.Combobox(format_delim_frame, textvariable=self.delimiter_var, values=delim_options, width=10, state='readonly')
        delim_menu.pack(side='left', padx=4)

        # 轉換按鈕
        convert_btn = tb.Button(self.root, text='開始轉換', command=self.convert_files)
        convert_btn.pack(pady=(18, 10))

    def select_output_dir(self):
        dir_path = filedialog.askdirectory(title='選擇輸出資料夾')
        if dir_path:
            self.output_dir_var.set(dir_path)


    def select_files(self):
        files = filedialog.askopenfilenames(
            title='選取 CSV/XLSX/XLS 檔案',
            filetypes=[('CSV/XLSX/XLS Files', '*.csv *.xlsx *.xls'), ('CSV Files', '*.csv'), ('Excel Files', '*.xlsx *.xls')]
        )
        self.selected_files = list(files)
        self._update_file_listbox()

    def _update_file_listbox(self):
        self.file_listbox.config(state='normal')
        self.file_listbox.delete('1.0', 'end')
        if self.selected_files:
            for f in self.selected_files:
                filename = os.path.basename(f)
                self.file_listbox.insert('end', filename + '\n')
        else:
            self.file_listbox.insert('end', '（尚未選取檔案）')
        self.file_listbox.config(state='disabled')


    def convert_files(self):
        if not self.selected_files:
            Messagebox.show_warning('請先選取檔案！', title='警告')
            return
        output_dir = self.output_dir_var.get().strip()
        if not output_dir:
            Messagebox.show_warning('請先選擇輸出資料夾！', title='警告')
            return
        # 若資料夾不存在則自動建立
        if not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir)
            except Exception as e:
                Messagebox.show_error(f'無法建立資料夾：{output_dir}\n錯誤: {e}', title='錯誤')
                return
        remove_columns = [col.strip() for col in self.remove_columns_var.get().split(',') if col.strip()]
        output_format = self.output_format_var.get()
        delim_map = {'逗號(,)': ',', 'Tab(\t)': '\t', '分號(;)': ';'}
        delimiter = delim_map.get(self.delimiter_var.get(), ',')
        ext_map = {'txt': '.txt', 'tsv': '.tsv', 'csv': '.csv'}
        ext = ext_map.get(output_format, '.txt')
        sheet_name = self.sheet_name_var.get().strip()
        merge_sheets = self.merge_sheets_var.get()
        success_count = 0
        for file_path in self.selected_files:
            try:
                file_ext = os.path.splitext(file_path)[1].lower()
                base_name = os.path.splitext(os.path.basename(file_path))[0]
                if file_ext == '.csv':
                    with open(file_path, 'r', encoding='utf-8', errors='ignore', newline='') as csvfile:
                        reader = csv.DictReader(csvfile)
                        if not reader.fieldnames:
                            csvfile.seek(0)
                            content = csvfile.read()
                            out_path = os.path.join(output_dir, base_name + ext)
                            with open(out_path, 'w', encoding='utf-8') as f:
                                f.write(content)
                            success_count += 1
                            continue
                        keep_fields = [f for f in reader.fieldnames if f not in remove_columns]
                        rows = [ {k: row[k] for k in keep_fields} for row in reader ]
                    out_path = os.path.join(output_dir, base_name + ext)
                    with open(out_path, 'w', encoding='utf-8', newline='') as out_file:
                        writer = csv.DictWriter(out_file, fieldnames=keep_fields, delimiter=delimiter)
                        writer.writeheader()
                        writer.writerows(rows)
                    success_count += 1
                elif file_ext in ('.xlsx', '.xls'):
                    # Excel 處理，根據副檔名選擇 engine
                    engine = 'openpyxl' if file_ext == '.xlsx' else 'xlrd'
                    if merge_sheets:
                        # 合併所有 sheet
                        xls = pd.ExcelFile(file_path, engine=engine)
                        all_dfs = []
                        for sheet in xls.sheet_names:
                            df = pd.read_excel(xls, sheet_name=sheet, dtype=str, engine=engine)
                            if not df.empty:
                                keep_fields = [col for col in df.columns if col not in remove_columns]
                                df2 = df[keep_fields]
                                df2['__sheetname__'] = sheet
                                all_dfs.append(df2)
                        if all_dfs:
                            merged = pd.concat(all_dfs, ignore_index=True)
                            out_path = os.path.join(output_dir, base_name + ext)
                            merged.to_csv(out_path, sep=delimiter, index=False, encoding='utf-8')
                            success_count += 1
                        else:
                            print(f'Excel 檔案無資料: {file_path}')
                    else:
                        # 指定 sheet 名稱或索引
                        sn = 0 if not sheet_name else sheet_name
                        try:
                            df = pd.read_excel(file_path, sheet_name=sn, dtype=str, engine=engine)
                        except Exception as e:
                            print(f'無法讀取 sheet: {sheet_name} in {file_path}，錯誤: {e}')
                            continue
                        if not df.empty:
                            keep_fields = [col for col in df.columns if col not in remove_columns]
                            df2 = df[keep_fields]
                            out_path = os.path.join(output_dir, base_name + ext)
                            df2.to_csv(out_path, sep=delimiter, index=False, encoding='utf-8')
                            success_count += 1
                        else:
                            print(f'Excel 檔案無資料: {file_path}')
                else:
                    print(f'不支援的檔案格式: {file_path}')
            except Exception as e:
                print(f'轉換失敗: {file_path}\n錯誤: {e}')
        Messagebox.show_info(f'成功轉換 {success_count} 個檔案！', title='完成')


def main():
    app = tb.Window(themename='flatly')
    CsvBatchConverter(app)
    app.mainloop()


if __name__ == '__main__':
    main()



# 舊的全域 GUI 變數與程式碼已移除，請使用 main() 入口
