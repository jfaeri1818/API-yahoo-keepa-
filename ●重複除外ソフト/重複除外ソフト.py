import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

# ============================================================
# 関数
# ============================================================
def process_file(file_path):
    try:
        df = pd.read_excel(file_path, engine="openpyxl")

        if df.empty:
            messagebox.showwarning("警告", "ファイルが空です。")
            return

        col_name = df.columns[0]
        before = len(df)
        df = df.drop_duplicates(subset=[col_name], keep="first")
        after = len(df)

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excelファイル", "*.xlsx")],
            initialfile=os.path.splitext(os.path.basename(file_path))[0] + "_unique.xlsx"
        )
        if not save_path:
            return

        df.to_excel(save_path, index=False)
        messagebox.showinfo("完了", f"重複を削除しました。\n{before} → {after} 行に減少。\n\n保存先：\n{save_path}")

    except Exception as e:
        messagebox.showerror("エラー", f"処理中にエラーが発生しました：\n{e}")


def drop(event):
    file_path = event.data.strip("{}")
    if os.path.isfile(file_path) and file_path.endswith(".xlsx"):
        process_file(file_path)
    else:
        messagebox.showwarning("警告", "Excelファイル（.xlsx）をドロップしてください。")


def select_file():
    file_path = filedialog.askopenfilename(
        title="Excelファイルを選択",
        filetypes=[("Excelファイル", "*.xlsx")]
    )
    if file_path:
        process_file(file_path)

# ============================================================
# GUI設定
# ============================================================
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    root = TkinterDnD.Tk()  # ✅ 最初からこれを使う
except Exception as e:
    messagebox.showerror("初期化エラー", f"ドラッグ＆ドロップ機能を有効化できませんでした：\n{e}")
    root = tk.Tk()  # 失敗時のみ通常Tkを使用

root.title("Excel重複除外ツール")
root.geometry("400x300")
root.resizable(False, False)

label = tk.Label(root, text="Excelファイルをここにドロップ\nまたはボタンで選択", font=("Meiryo", 11))
label.pack(pady=20)

drop_area = tk.Label(root, text="ここにファイルをドロップ", width=40, height=5, bg="#f0f0f0", relief="groove")
drop_area.pack(pady=10)

try:
    drop_area.drop_target_register(DND_FILES)
    drop_area.dnd_bind('<<Drop>>', drop)
except:
    pass

tk.Button(root, text="ファイルを選択", command=select_file).pack(pady=10)

root.mainloop()
