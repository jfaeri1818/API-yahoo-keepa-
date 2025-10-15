#!/usr/bin/env python
# coding: utf-8

# In[14]:


import pandas as pd
import tkinter as tk
from tkinter import messagebox, scrolledtext
from tkinterdnd2 import TkinterDnD, DND_FILES
import os

def find_jan_column(df):
    """列名からJANを自動検出"""
    for c in df.columns:
        if isinstance(c, str) and "JAN" in c.upper():
            return c
    return df.columns[0]

def determine_input_output(file1, file2):
    """中身を確認して入力・出力を自動判定（1列のみの方を入力扱い）"""
    try:
        df1 = pd.read_excel(file1, header=None)
        df2 = pd.read_excel(file2, header=None)

        # ✅ 1列しかない方を「入力ファイル」と判定
        if df1.shape[1] == 1 and df2.shape[1] > 1:
            return file1, file2
        elif df2.shape[1] == 1 and df1.shape[1] > 1:
            return file2, file1
        elif df1.shape[1] == 1 and df2.shape[1] == 1:
            # 両方1列の場合 → 行数が短い方を入力扱い
            if len(df1) <= len(df2):
                return file1, file2
            else:
                return file2, file1
        else:
            raise ValueError("1列だけのExcelを入力としてドラッグしてください。")
    except Exception as e:
        raise ValueError(f"ファイル判定中にエラーが発生しました: {e}")

def process_files():
    """2つのファイルが揃った時点で処理"""
    if len(file_list) != 2:
        return  # 2つ揃うまで待機

    try:
        input_path, output_path = determine_input_output(file_list[0], file_list[1])

        # 🟢 入力（ヘッダーなし）
        input_df = pd.read_excel(input_path, header=None)
        input_df.columns = ["JAN"]
        input_df["JAN"] = input_df["JAN"].astype(str).str.strip()

        # 🟢 出力（ヘッダーあり）
        output_df = pd.read_excel(output_path)
        jan_col_output = find_jan_column(output_df)
        output_df[jan_col_output] = output_df[jan_col_output].astype(str).str.strip()

        # 🟢 並び替え処理
        merged = pd.merge(input_df, output_df, left_on="JAN", right_on=jan_col_output, how="left")

        # ✅ 出力側のJAN列を削除
        merged.drop(columns=[jan_col_output], inplace=True)

        # 💾 保存
        save_path = os.path.join(os.path.dirname(output_path), "JAN整列結果.xlsx")
        merged.to_excel(save_path, index=False)

        messagebox.showinfo("完了", f"JAN整列が完了しました！\n\n保存先：\n{save_path}")

    except Exception as e:
        messagebox.showerror("エラー", f"処理中にエラーが発生しました：\n{e}")

def drop(event):
    """ドラッグ＆ドロップ時の処理"""
    global file_list
    dropped_files = list(root.tk.splitlist(event.data))
    for f in dropped_files:
        if f not in file_list:
            file_list.append(f)

    # ファイル名を表示
    file_text.delete(1.0, tk.END)
    file_text.insert(tk.END, "\n".join(os.path.basename(f) for f in file_list))

    # 2つ揃ったら自動実行
    if len(file_list) == 2:
        process_files()

def clear_files():
    """ドロップされたファイルを削除（リセット）"""
    global file_list
    file_list = []
    file_text.delete(1.0, tk.END)
    file_text.insert(tk.END, "📂 ファイル未選択")
    messagebox.showinfo("リセット", "ファイルリストをリセットしました。")

# 🪶 UI設定
root = TkinterDnD.Tk()
root.title("JAN整列ツール")
root.geometry("520x370")
root.configure(bg="#f7f3e9")
root.resizable(False, False)  # サイズ固定

title_label = tk.Label(
    root,
    text="JAN整列ツール\n\n"
         "① 入力（1列のExcel）と出力（一覧のExcel）を\n"
         "　ここにドラッグ＆ドロップしてください。\n"
         "② 2つのファイルが揃うと自動で整列します。",
    bg="#f7f3e9",
    fg="#333",
    font=("Meiryo", 11)
)
title_label.pack(pady=10)

drop_frame = tk.Label(
    root,
    text="⬇️ ここに2つのファイルをドロップ ⬇️",
    bg="#fffaf2",
    fg="#555",
    relief="ridge",
    width=50,
    height=5
)
drop_frame.pack(pady=5)

# ドロップ登録
drop_frame.drop_target_register(DND_FILES)
drop_frame.dnd_bind("<<Drop>>", drop)

# 📜 ファイル名表示欄（スクロール付き）
file_text = scrolledtext.ScrolledText(
    root,
    width=55,
    height=4,
    bg="#fff",
    fg="#444",
    font=("Meiryo", 9),
    wrap="none"
)
file_text.insert(tk.END, "📂 ファイル未選択")
file_text.pack(pady=5)

# 🗑 ファイル削除（リセット）ボタン
clear_button = tk.Button(
    root,
    text="🗑 ファイルをリセット",
    command=clear_files,
    bg="#f4cfa0",
    fg="#333",
    font=("Meiryo", 10, "bold"),
    relief="raised",
    width=20,
    height=1
)
clear_button.pack(pady=10)

file_list = []

root.mainloop()


# %%
