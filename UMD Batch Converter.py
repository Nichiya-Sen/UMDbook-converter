import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import subprocess
import os
import threading
import shutil
import time
import datetime
import random
import string
import ctypes # 新增 ctypes 模組

try:
    import opencc
    OPENCC_AVAILABLE = True
except ImportError:
    OPENCC_AVAILABLE = False

LANGUAGES = {
    "zh_TW": {
        "window_title": "UMD 批次轉換器",
        "list_frame_title": "待轉換 UMD 檔案清單",
        "add_files": "加入檔案",
        "add_folder": "加入資料夾",
        "remove_selected": "移除所選",
        "clear_all": "全部清除",
        "output_settings_title": "輸出設定",
        "save_to": "儲存到:",
        "browse": "瀏覽...",
        "lang_conversion_title": "繁簡轉換:",
        "lang_s2t": "簡體中文 -> 繁體中文",
        "lang_t2s": "繁體中文 -> 簡體中文",
        "lang_none": "不轉換",
        "output_format_title": "輸出格式:",
        "start_conversion": "開始批次轉換",
        "status_initial": "請加入 UMD 檔案或資料夾以開始...",
        "status_file_count": "清單中共有 {count} 個檔案。",
        "status_processing": "處理中({i}/{total}): {filename}",
        "status_done": "全部轉換完成！",
        "welcome_title": "程式夥伴提示",
        "welcome_message_base": "歡迎使用！\n\n",
        "welcome_opencc_fail": "⚠️ [警告] 缺少 OpenCC 函式庫，繁簡轉換功能將不可用。\n請執行 'pip install opencc-python-reimplemented' 來安裝。\n\n",
        "welcome_calibre_fail": "⚠️ [警告] 找不到 Calibre 命令列工具。\nEPUB/MOBI/AZW3 轉換功能將不可用，僅能輸出 TXT。\n",
        "welcome_ok": "✅ 所有功能已就緒！\n本程式的電子書與繁簡轉換功能由 Calibre 與 OpenCC 提供支援。",
        "task_complete_title": "任務完成",
        "task_complete_message": "全部 {total} 個檔案已處理完畢。\n成功: {success} 個\n失敗: {fail} 個",
        "task_fail_message": "\n\n部分檔案轉換失敗，請查看 conversion_log.txt 檔案以了解詳細原因。"
    },
    "en_US": {
        "window_title": "UMD Batch Converter",
        "list_frame_title": "UMD Files to Convert",
        "add_files": "Add Files",
        "add_folder": "Add Folder",
        "remove_selected": "Remove",
        "clear_all": "Clear All",
        "output_settings_title": "Output Settings",
        "save_to": "Save to:",
        "browse": "Browse...",
        "lang_conversion_title": "Language:",
        "lang_s2t": "Simplified -> Traditional Chinese",
        "lang_t2s": "Traditional -> Simplified Chinese",
        "lang_none": "No Conversion",
        "output_format_title": "Output Format:",
        "start_conversion": "Start Batch Conversion",
        "status_initial": "Please add files or a folder to start...",
        "status_file_count": "{count} files in the list.",
        "status_processing": "Processing ({i}/{total}): {filename}",
        "status_done": "All conversions finished!",
        "welcome_title": "Assistant's Tip",
        "welcome_message_base": "Welcome!\n\n",
        "welcome_opencc_fail": "⚠️ [Warning] OpenCC library not found. Language conversion is disabled.\nPlease run 'pip install opencc-python-reimplemented' to install it.\n\n",
        "welcome_calibre_fail": "⚠️ [Warning] Calibre command-line tool not found.\nEPUB/MOBI/AZW3 conversion is disabled. Only TXT output is available.\n",
        "welcome_ok": "✅ All features are ready!\nEbook and language conversions are supported by Calibre and OpenCC.",
        "task_complete_title": "Task Complete",
        "task_complete_message": "All {total} files have been processed.\nSuccess: {success}\nFail: {fail}",
        "task_fail_message": "\n\nSome files failed to convert. Please check conversion_log.txt for details."
    },
    "ja_JP": {
        "window_title": "UMD一括コンバーター",
        "list_frame_title": "変換待ちのUMDファイルリスト",
        "add_files": "ファイルを追加",
        "add_folder": "フォルダを追加",
        "remove_selected": "削除",
        "clear_all": "すべてクリア",
        "output_settings_title": "出力設定",
        "save_to": "保存先:",
        "browse": "参照...",
        "lang_conversion_title": "言語変換:",
        "lang_s2t": "簡体字 -> 繁体字",
        "lang_t2s": "繁体字 -> 簡体字",
        "lang_none": "変換しない",
        "output_format_title": "出力形式:",
        "start_conversion": "一括変換を開始",
        "status_initial": "ファイルまたはフォルダを追加してください...",
        "status_file_count": "リストに {count} 個のファイルがあります。",
        "status_processing": "処理中 ({i}/{total}): {filename}",
        "status_done": "すべての変換が完了しました！",
        "welcome_title": "アシスタントのヒント",
        "welcome_message_base": "ようこそ！\n\n",
        "welcome_opencc_fail": "⚠️ [警告] OpenCCライブラリが見つかりません。言語変換機能は無効です。\n'pip install opencc-python-reimplemented' を実行してインストールしてください。\n\n",
        "welcome_calibre_fail": "⚠️ [警告] Calibreのコマンドラインツールが見つかりません。\nEPUB/MOBI/AZW3への変換は無効です。TXT出力のみ可能です。\n",
        "welcome_ok": "✅ すべての機能が準備完了です！\n電子書籍及び言語変換はCalibreとOpenCCによってサポートされています。",
        "task_complete_title": "タスク完了",
        "task_complete_message": "全 {total} ファイルの処理が完了しました。\n成功: {success}\n失敗: {fail}",
        "task_fail_message": "\n\n一部のファイルの変換に失敗しました。詳細は conversion_log.txt を確認してください。"
    }
}

def write_log(message, clear_log=False):
    log_file = "conversion_log.txt"
    mode = "w" if clear_log else "a"
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_file, mode, encoding="utf-8") as f:
        f.write(f"[{timestamp}] {message}\n")

class UMDConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.withdraw()
        self.current_lang = "zh_TW"
        self.umd_converter_exe = "umd2ebook_x64.exe"
        self.calibre_converter_exe = "ebook-convert"
        self.output_dir_path = tk.StringVar(value=os.getcwd())
        self.output_format = tk.StringVar(value="epub")
        self.lang_display_text = tk.StringVar()
        self.lang_conversion_options = {}
        self.setup_ui()
        self.switch_language(init=True)
        self.center_window(650, 600)
        self.root.deiconify()
        self.root.after(100, self.show_startup_info)
        
    def center_window(self, width, height):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        self.root.geometry(f'{width}x{height}+{int(x)}+{int(y)}')

    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.pack(fill="both", expand=True)
        self.list_frame = ttk.LabelFrame(main_frame)
        self.list_frame.pack(fill="both", expand=True, padx=5, pady=5)
        file_button_frame = ttk.Frame(main_frame)
        file_button_frame.pack(fill="x", pady=5)
        self.output_frame = ttk.LabelFrame(main_frame)
        self.output_frame.pack(fill="x", padx=5, pady=10)
        convert_frame = ttk.Frame(main_frame)
        convert_frame.pack(fill="x", pady=5, padx=5)
        lang_switch_frame = ttk.Frame(main_frame)
        lang_switch_frame.pack(side="bottom", fill="x", pady=(10, 0))
        ttk.Label(lang_switch_frame, text="Language:").pack(side="left")
        self.lang_switcher = ttk.Combobox(lang_switch_frame, values=["繁體中文 (zh_TW)", "English (en_US)", "日本語 (ja_JP)"], state="readonly", width=20)
        self.lang_switcher.pack(side="left", padx=5)
        self.lang_switcher.set("繁體中文 (zh_TW)")
        self.lang_switcher.bind("<<ComboboxSelected>>", self.switch_language)
        scrollbar = ttk.Scrollbar(self.list_frame)
        scrollbar.pack(side="right", fill="y")
        self.file_listbox = tk.Listbox(self.list_frame, selectmode="extended", yscrollcommand=scrollbar.set)
        self.file_listbox.pack(fill="both", expand=True, padx=5, pady=5)
        scrollbar.config(command=self.file_listbox.yview)
        self.add_files_button = ttk.Button(file_button_frame, command=self.add_files)
        self.add_files_button.pack(side="left", padx=5)
        self.add_folder_button = ttk.Button(file_button_frame, text="加入資料夾", command=self.add_folder)
        self.add_folder_button.pack(side="left", padx=5)
        self.remove_button = ttk.Button(file_button_frame, command=self.remove_selected)
        self.remove_button.pack(side="left", padx=5)
        self.clear_button = ttk.Button(file_button_frame, command=self.clear_all)
        self.clear_button.pack(side="left", padx=5)
        self.save_to_label = ttk.Label(self.output_frame)
        self.save_to_label.grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.entry_output_dir = ttk.Entry(self.output_frame, textvariable=self.output_dir_path, width=60, state="readonly")
        self.entry_output_dir.grid(row=0, column=1, padx=5)
        self.browse_output_button = ttk.Button(self.output_frame, command=self.browse_output_dir)
        self.browse_output_button.grid(row=0, column=2, padx=5)
        self.lang_conversion_label = ttk.Label(self.output_frame)
        self.lang_conversion_label.grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.lang_menu = ttk.Combobox(self.output_frame, textvariable=self.lang_display_text, state="readonly", width=30)
        self.lang_menu.grid(row=1, column=1, columnspan=2, padx=5, sticky="w")
        self.output_format_label = ttk.Label(convert_frame)
        self.output_format_label.pack(side="left", padx=(0,10))
        self.format_menu = ttk.Combobox(convert_frame, textvariable=self.output_format, values=["epub", "mobi", "azw3", "txt"], state="readonly", width=10)
        self.format_menu.pack(side="left")
        self.convert_button = ttk.Button(convert_frame, command=self.start_conversion_thread)
        self.convert_button.pack(side="right", ipady=5)
        self.status_label = ttk.Label(main_frame)
        self.status_label.pack(side="left", fill="x", expand=True, padx=(5,0), pady=(10,0))
        self.controls = [self.add_files_button, self.add_folder_button, self.remove_button, self.clear_button, self.browse_output_button, self.format_menu, self.convert_button, self.lang_switcher]

    def switch_language(self, event=None, init=False):
        if not init:
            selected_lang_str = self.lang_switcher.get()
            self.current_lang = selected_lang_str.split('(')[-1].strip(')')
        lang_dict = LANGUAGES[self.current_lang]
        self.lang_conversion_options = {
            lang_dict["lang_s2t"]: "s2t",
            lang_dict["lang_t2s"]: "t2s",
            lang_dict["lang_none"]: "none"
        }
        self.lang_menu.config(values=list(self.lang_conversion_options.keys()))
        self.lang_display_text.set(lang_dict["lang_s2t"])
        self.root.title(lang_dict["window_title"])
        self.list_frame.config(text=lang_dict["list_frame_title"])
        self.add_files_button.config(text=lang_dict["add_files"])
        self.add_folder_button.config(text=lang_dict["add_folder"])
        self.remove_button.config(text=lang_dict["remove_selected"])
        self.clear_button.config(text=lang_dict["clear_all"])
        self.output_frame.config(text=lang_dict["output_settings_title"])
        self.save_to_label.config(text=lang_dict["save_to"])
        self.browse_output_button.config(text=lang_dict["browse"])
        self.lang_conversion_label.config(text=lang_dict["lang_conversion_title"])
        self.output_format_label.config(text=lang_dict["output_format_title"])
        self.convert_button.config(text=lang_dict["start_conversion"])
        self.update_status()

    def show_startup_info(self):
        lang_dict = LANGUAGES[self.current_lang]
        info_message = lang_dict["welcome_message_base"]
        calibre_ok = shutil.which(self.calibre_converter_exe)
        if not OPENCC_AVAILABLE:
            info_message += lang_dict["welcome_opencc_fail"]
            self.lang_menu.config(state="disabled")
            self.lang_display_text.set(lang_dict["lang_none"])
        if not calibre_ok:
            info_message += lang_dict["welcome_calibre_fail"]
        if OPENCC_AVAILABLE and calibre_ok:
            info_message += lang_dict["welcome_ok"]
        messagebox.showinfo(lang_dict["welcome_title"], info_message)

    def update_status(self, text_key=None, **kwargs):
        lang_dict = LANGUAGES[self.current_lang]
        if text_key:
            self.status_label.config(text=lang_dict[text_key].format(**kwargs))
        else:
            count = self.file_listbox.size()
            if count > 0:
                self.status_label.config(text=lang_dict["status_file_count"].format(count=count))
            else:
                self.status_label.config(text=lang_dict["status_initial"])
    
    def browse_output_dir(self):
        dir_path = filedialog.askdirectory(title="選擇輸出資料夾")
        if dir_path:
            self.output_dir_path.set(dir_path)
        
    def add_files(self, file_paths=None):
        if not file_paths:
            file_paths = filedialog.askopenfilenames(title="選擇一個或多個 UMD 檔案", filetypes=(("UMD 檔案", "*.umd"), ("所有檔案", "*.*")))
        for file_path in file_paths:
            if file_path.lower().endswith(".umd") and file_path not in self.file_listbox.get(0, "end"):
                self.file_listbox.insert("end", file_path)
        self.update_status()
        
    def add_folder(self):
        folder_path = filedialog.askdirectory(title="選擇包含 UMD 檔案的資料夾")
        if not folder_path:
            return
        found_files = [os.path.join(root, file) for root, _, files in os.walk(folder_path) if file.lower().endswith(".umd")]
        if found_files:
            self.add_files(file_paths=found_files)
        else:
            messagebox.showinfo("提示", "在選定的資料夾中沒有找到 .umd 檔案。")
        
    def remove_selected(self):
        for index in reversed(self.file_listbox.curselection()):
            self.file_listbox.delete(index)
        self.update_status()
        
    def clear_all(self):
        self.file_listbox.delete(0, "end")
        self.update_status()
        
    def set_controls_state(self, state):
        for control in self.controls:
            control.config(state=state)
        
    def start_conversion_thread(self):
        lang_dict = LANGUAGES[self.current_lang]
        if self.file_listbox.size() == 0:
            messagebox.showwarning(lang_dict["task_complete_title"], "轉換清單是空的！")
            return
        
        selected_lang_text = self.lang_display_text.get()
        lang_mode = self.lang_conversion_options.get(selected_lang_text, "none")
        if lang_mode != "none" and not OPENCC_AVAILABLE:
            messagebox.showerror("功能鎖定", "缺少 OpenCC 函式庫，無法進行繁簡轉換。")
            return
        
        output_fmt = self.output_format.get()
        if output_fmt != 'txt' and not shutil.which(self.calibre_converter_exe):
            messagebox.showerror("功能鎖定", f"缺少 Calibre 工具，無法轉換為 {output_fmt.upper()}。")
            return
        
        write_log("="*20 + " 開始新的轉換任務 " + "="*20, clear_log=True)
        self.set_controls_state("disabled")
        thread = threading.Thread(target=self.run_batch_conversion)
        thread.start()

    def run_batch_conversion(self):
        files_to_convert = self.file_listbox.get(0, "end")
        total_files = len(files_to_convert)
        success_count, fail_count = 0, 0
        output_dir = self.output_dir_path.get()
        lang_dict = LANGUAGES[self.current_lang]
        selected_lang_text = self.lang_display_text.get()
        lang_mode = self.lang_conversion_options.get(selected_lang_text, "none")
        
        if lang_mode != "none" and OPENCC_AVAILABLE:
            cc = opencc.OpenCC(lang_mode)

        for i, input_file in enumerate(files_to_convert):
            self.update_status("status_processing", i=i+1, total=total_files, filename=os.path.basename(input_file))
            base_name = os.path.splitext(os.path.basename(input_file))[0]
            
            rand_str = ''.join(random.choices(string.ascii_lowercase + string.digits, k=6))
            temp_txt_path = os.path.join(os.getcwd(), f"_{base_name}_{rand_str}_temp.txt")

            try:
                write_log(f"步驟 1: '{input_file}' -> TXT")
                cmd_umd2txt = [self.umd_converter_exe, "-i", input_file, "-e", "txt"]
                subprocess.run(cmd_umd2txt, check=True, capture_output=True, timeout=60)
                
                generated_txt_path = os.path.join(os.getcwd(), f"{base_name}.txt")
                if os.path.exists(generated_txt_path):
                    os.rename(generated_txt_path, temp_txt_path)
                else:
                    raise FileNotFoundError("umd2ebook 未能產生預期的 TXT 檔案。")
                
                if lang_mode != "none":
                    write_log(f"步驟 2: 語言轉換 ({lang_mode})")
                    with open(temp_txt_path, 'r', encoding='utf-8') as f:
                        original_text = f.read()
                    converted_text = cc.convert(original_text)
                    with open(temp_txt_path, 'w', encoding='utf-8') as f:
                        f.write(converted_text)

                final_output_format = self.output_format.get()
                final_output_path = os.path.join(output_dir, f"{base_name}.{final_output_format}")

                if final_output_format == 'txt':
                    if os.path.normcase(temp_txt_path) != os.path.normcase(final_output_path):
                        shutil.move(temp_txt_path, final_output_path)
                    write_log(f"步驟 3: TXT 輸出完成 -> {final_output_path}")
                else:
                    write_log(f"步驟 3: Calibre: TXT -> {final_output_format.upper()}")
                    cmd_txt2ebook = [self.calibre_converter_exe, temp_txt_path, final_output_path, "--base-font-size", "16"]
                    subprocess.run(cmd_txt2ebook, check=True, capture_output=True, timeout=120)
                    os.remove(temp_txt_path)
                    write_log(f"成功清理暫存檔並輸出 -> {final_output_path}")
                
                success_count += 1
            except Exception as e:
                write_log(f"!!!!!!!!!! 處理 '{input_file}' 時發生錯誤 !!!!!!!!!!")
                if isinstance(e, subprocess.CalledProcessError):
                    write_log(f"錯誤來自於: {e.cmd}")
                    write_log(f"錯誤訊息:\n{e.stderr.decode('utf-8', 'ignore') if e.stderr else e.stdout.decode('utf-8', 'ignore')}")
                else:
                    write_log(f"未知錯誤: {e}")
                fail_count += 1
                if os.path.exists(temp_txt_path):
                    os.remove(temp_txt_path)
        
        self.update_status("status_done")
        summary_message = lang_dict["task_complete_message"].format(total=total_files, success=success_count, fail=fail_count)
        if fail_count > 0:
            summary_message += lang_dict["task_fail_message"]
        messagebox.showinfo(lang_dict["task_complete_title"], summary_message)
        self.set_controls_state("normal")

if __name__ == "__main__":
    # 使用 ctypes 模組設定 Windows 專用的 AppUserModelID，解決工作列圖示問題
    try:
        myappid = 'mycompany.myproduct.subproduct.version' # 可自訂
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    except (AttributeError, NameError):
        pass

    root = tk.Tk()
    app = UMDConverterApp(root)
    try:
        root.iconbitmap(default="UMD ebook converter.ico")
    except tk.TclError:
        pass
    root.mainloop()