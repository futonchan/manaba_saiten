import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import glob, os
import openpyxl

report_index = 0 # 読み込んだレポートリストのインデックス。グローバル変数

def browse_folder_entry(path_entry):
    path_entry.insert(0, "") # リセット
    selected_folder = filedialog.askdirectory() # フォルダの参照ダイアログを表示する
    path_entry.insert(0, selected_folder) # 選択されたフォルダのパスをテキストボックスに表示する
def do_nothing():
    pass

def write_excel(report_path_list, score_entry, wb):
    global report_index
    path = report_path_list[report_index]
    _, user =  os.path.split(path)
    user_id = user.split("@")[1]
    sheet = wb['Sheet1']
    score = score_entry.get()
    print(score)
    for row in sheet.rows:
        for cell in row:
            if cell.value == user_id:
                user_row = cell.row
                sheet[str("J") + str(user_row)] = score
                print(sheet[str("J") + str(user_row)])
    wb.save("reportlist2.xlsx")

def load_report_index(report_path_list, out_text, name_label):
    global report_index
    suffix = "report.txt"
    path = report_path_list[report_index]
    file_path = os.path.join(path, suffix)
    with open(file_path, "r", encoding="utf-8") as f:
        text = f.read()
    out_text.delete("1.0", tk.END) # リセット
    out_text.insert(tk.END, text)

    _, user =  os.path.split(path)
    size = str(os.path.getsize(file_path))
    name_label.config(text="user: " + user + ", size: " + size + ", files: " + str(report_index + 1) + "/" + str(len(report_path_list)))

# 前に戻るボタン、次に進むボタンを押すときにExcelファイルへの書き込みを行う
def load_report_previous_index(report_path_list, out_text, name_label, score_entry, wb):
    global report_index
    write_excel(report_path_list, score_entry, wb)
    report_index = report_index - 1
    if report_index < 0:
        report_index = 0
    load_report_index(report_path_list, out_text, name_label)

def load_report_next_index(report_path_list, out_text, name_label, score_entry, wb):
    global report_index
    write_excel(report_path_list, score_entry, wb)
    report_index = report_index + 1
    if report_index >= len(report_path_list):
        report_index = len(report_path_list) - 1
    load_report_index(report_path_list, out_text, name_label)

def load_report(src_entry):
    global report_index
    src_folder = src_entry.get()
    src_folder = src_folder.replace("/", "\\")
    report_index = 0
    print("load")
    folder_suffix = "*"
    report_path_list = glob.glob(os.path.join(src_folder, folder_suffix))
    report_path_list.remove(os.path.join(src_folder, "reportlist.xlsx"))
    # Excelファイルを読み込む
    wb = openpyxl.load_workbook(os.path.join(src_folder, "reportlist.xlsx"))
    # out_text.insert(tk.END, "入力元フォルダ: " + src_folder + "\n")
    window = tk.Toplevel(root)
    window.protocol("WM_DELETE_WINDOW", do_nothing)

    # ラベル
    name_label = ttk.Label(window, text="")
    name_label.pack()

    # ログ用テキストボックス
    log_frame = tk.Frame(window)
    log_text = tk.Text(log_frame, width=100, height=60)
    scrollbar = ttk.Scrollbar(log_frame)
    log_text.configure(yscrollcommand=scrollbar.set)
    scrollbar.configure(command=log_text.yview)
    log_frame.pack()
    log_text.pack(side="left")
    scrollbar.pack(side="right", fill=tk.Y)

    # 点数
    score_frame = tk.Frame(window)
    score_label = ttk.Label(score_frame, text="点数")
    score_values = ["60", "70", "80", "90", "100"]
    # score_entry = ttk.Entry(score_frame, width=100) # スコア手入力用
    score_combo = ttk.Combobox(score_frame, values=score_values)
    score_frame.pack()
    score_label.pack(side="left")
    # score_entry.pack(side="left") # スコア手入力用
    score_combo.pack(side="left")

    # 前に戻る、次に進むボタン
    button_frame = tk.Frame(window)
    previous_button = ttk.Button(button_frame, text="前に戻る", command=lambda:load_report_previous_index(report_path_list, log_text, name_label, score_combo, wb))
    next_button = ttk.Button(button_frame, text="次に進む", command=lambda:load_report_next_index(report_path_list, log_text, name_label, score_combo, wb))
    button_frame.pack()
    previous_button.pack(side="left")
    next_button.pack(side="left")

    # 備考
    biko_label = ttk.Label(window, text="保存先ファイル：このアプリが動作しているフォルダにある「reportlist2.xlsx」へファイルが出力されます。\n「前に戻る」、「次に進む」ボタンクリックで「reportlist2.xlsx」の「合計点」列に点数が書き込まれます。")
    biko_label.pack()

    # シートを指定する
    load_report_index(report_path_list, log_text, name_label)

root = tk.Tk()
root.title("manabaレポート採点支援")

# 入力元ウィジェット
src_frame = tk.Frame(root)
src_label = ttk.Label(src_frame, text="レポートフォルダ")
src_path_entry = ttk.Entry(src_frame, width=100) #フォルダのパスを表示するテキストボックス
src_browse_button = ttk.Button(src_frame, text="参照", command=lambda:browse_folder_entry(src_path_entry)) # フォルダの参照ボタン
# 配置
src_frame.pack()
src_label.pack(side="left")
src_path_entry.pack(side="left")
src_browse_button.pack(side="left")

# 実行ボタン
execute_button = ttk.Button(root, text="実行", command=lambda:load_report(src_path_entry))
execute_button.pack()

# 備考
biko_label = ttk.Label(root, text="レポートフォルダ：manabaから出力される「report-******-******」フォルダを指定する。\n「report-******-******」フォルダの中には学生のレポートが入ったフォルダが学生の数あり、reportlist.xlsxファイルがあることを想定しています。")
biko_label.pack()

root.mainloop()
