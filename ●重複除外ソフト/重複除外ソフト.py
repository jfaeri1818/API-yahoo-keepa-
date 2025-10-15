import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

def process_files(file_paths):
    try:
        combined_df = pd.DataFrame()
        total_before = 0

        for file_path in file_paths:
            df = pd.read_excel(file_path, engine="openpyxl",header=None)
            if df.empty:
                continue
            total_before += len(df)
            combined_df = pd.concat([combined_df, df], ignore_index=True)

        if combined_df.empty:
            messagebox.showwarning("警告", "有効なデータが含まれるファイルがありません。")
            return

        col_name = combined_df.columns[0]
        combined_df = combined_df.drop_duplicates(subset=[col_name], keep="first")
        total_after = len(combined_df)

        default_name = "統合結果_unique.xlsx" if len(file_paths) > 1 else os.path.splitext(os.path.basename(file_paths[0]))[0] + "_unique.xlsx"
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excelファイル", "*.xlsx")],
            initialfile=default_name
        )
        if not save_path:
            return

        combined_df.to_excel(save_path, index=False,header=None)
        messagebox.showinfo("完了", f"重複を削除しました。\n{total_before} → {total_after} 行に減少。\n\n保存先：\n{save_path}")

    except Exception as e:
        messagebox.showerror("エラー", f"処理中にエラーが発生しました：\n{e}")

def drop(event):
    raw = event.data.strip("{}")
    paths = raw.split("} {") if "} {" in raw else [raw]
    valid_paths = [p for p in paths if os.path.isfile(p) and p.endswith(".xlsx")]

    if not valid_paths:
        messagebox.showwarning("警告", "Excelファイル（.xlsx）のみ対応しています。")
        return

    process_files(valid_paths)

def select_file():
    file_paths = filedialog.askopenfilenames(
        title="Excelファイルを選択",
        filetypes=[("Excelファイル", "*.xlsx")]
    )
    if file_paths:
        process_files(list(file_paths))

# GUIウィンドウ初期化
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    root = TkinterDnD.Tk()
except Exception as e:
    messagebox.showerror("初期化エラー", f"ドラッグ＆ドロップ機能を有効化できませんでした：\n{e}")
    root = tk.Tk()

root.title("Excel重複除外・統合ツール")
root.geometry("460x320")
root.resizable(False, False)
root.configure(bg="#f4faff")

# タイトルラベル
label = tk.Label(root, text="Excelファイルをここにドロップ\n重複を除き、統合したデータを出力します。", font=("Meiryo", 11),
                 bg="#f4faff", fg="#2c82c9")
label.pack(pady=20)

# ドロップエリア
drop_area = tk.Label(root, text="ここにExcelファイルをドロップ(複数可)", width=50, height=5,
                     bg="#e2f0fb", fg="#0f3d57", relief="groove", font=("Meiryo", 10))
drop_area.pack(pady=10)

try:
    drop_area.drop_target_register(DND_FILES)
    drop_area.dnd_bind('<<Drop>>', drop)
except:
    pass

# 選択ボタン
tk.Button(root, text="ファイルを選択", command=select_file,
          bg="#2c82c9", fg="white", font=("Meiryo", 10, "bold"),
          width=20, height=1).pack(pady=15)

root.mainloop()
