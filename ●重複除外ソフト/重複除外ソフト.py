import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

def process_files(file_paths):
    try:
        combined_df = pd.DataFrame()
        total_before = 0

        # ファイルを順に読み込む
        for file_path in file_paths:
            df = pd.read_excel(file_path, engine="openpyxl")
            if df.empty:
                continue
            total_before += len(df)
            combined_df = pd.concat([combined_df, df], ignore_index=True)

        if combined_df.empty:
            messagebox.showwarning("警告", "有効なデータが含まれるファイルがありません。")
            return

        # JANコード列の特定
        jan_col = None
        for col in combined_df.columns:
            if "JAN" in str(col).upper():  # 大文字小文字どちらでも対応
                jan_col = col
                break

        if not jan_col:
            messagebox.showerror("エラー", "「JANコード」列が見つかりません。")
            return

        # JANコード列で重複除去
        combined_df = combined_df.drop_duplicates(subset=[jan_col], keep="first")
        total_after = len(combined_df)

        # 出力ファイル名
        default_name = "統合結果_unique.xlsx" if len(file_paths) > 1 else os.path.splitext(os.path.basename(file_paths[0]))[0] + "_unique.xlsx"
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excelファイル", "*.xlsx")],
            initialfile=default_name
        )
        if not save_path:
            return

        combined_df.to_excel(save_path, index=False, header=True)
        messagebox.showinfo("完了", f"重複を削除しました。\n{total_before} → {total_after} 行に減少。\n\n保存先：\n{save_path}")

    except Exception as e:
        messagebox.showerror("エラー", f"処理中にエラーが発生しました：\n{e}")

def select_file():
    file_paths = filedialog.askopenfilenames(
        title="Excelファイルを選択",
        filetypes=[("Excelファイル", "*.xlsx")]
    )
    if file_paths:
        process_files(list(file_paths))

# GUI初期化
root = tk.Tk()
root.title("Excel重複除外・統合ツール（JANコード基準）")
root.geometry("460x320")
root.resizable(False, False)
root.configure(bg="#f4faff")

# タイトルラベル
label = tk.Label(root, text="Excelファイルを選択して重複を除外・統合します。", 
                 font=("Meiryo", 11), bg="#f4faff", fg="#2c82c9")
label.pack(pady=40)

# ファイル選択ボタン（デザイン維持）
tk.Button(root, 
          text="ファイルを選択", 
          command=select_file,
          bg="#2c82c9", 
          fg="white", 
          font=("Meiryo", 10, "bold"),
          width=20, 
          height=1).pack(pady=15)

root.mainloop()
