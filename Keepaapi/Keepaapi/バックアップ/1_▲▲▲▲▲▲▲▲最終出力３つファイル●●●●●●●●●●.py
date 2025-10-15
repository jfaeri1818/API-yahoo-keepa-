#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import tkinter as tk
from tkinter import messagebox, scrolledtext
from tkinterdnd2 import TkinterDnD, DND_FILES
import os

# ======================
# 🔍 JAN整列ロジック
# ======================
def find_jan_column(df):
    """列名からJANを自動検出"""
    for c in df.columns:
        if isinstance(c, str) and "JAN" in c.upper():
            return c
    return df.columns[0]

def determine_input_output(file1, file2):
    """1列しかない方を入力ファイルとして判定"""
    df1 = pd.read_excel(file1, header=None)
    df2 = pd.read_excel(file2, header=None)

    if df1.shape[1] == 1 and df2.shape[1] > 1:
        return file1, file2
    elif df2.shape[1] == 1 and df1.shape[1] > 1:
        return file2, file1
    elif df1.shape[1] == 1 and df2.shape[1] == 1:
        return (file1, file2) if len(df1) <= len(df2) else (file2, file1)
    else:
        raise ValueError("1列だけのExcelを入力としてドラッグしてください。")

# ======================
# 🧮 処理メイン
# ======================
def process_files():
    """ドラッグされた2つのExcelを処理"""
    if len(file_list) != 2:
        return

    try:
        # --- JAN整列 ---
        input_path, output_path = determine_input_output(file_list[0], file_list[1])

        # 入力ファイル（1列JAN）
        input_df = pd.read_excel(input_path, header=None)
        input_df.columns = ["JAN"]
        input_df["JAN"] = input_df["JAN"].astype(str).str.strip()

        # 出力ファイル（商品一覧）
        output_df = pd.read_excel(output_path)
        jan_col = find_jan_column(output_df)
        output_df[jan_col] = output_df[jan_col].astype(str).str.strip()

        # JAN整列処理
        merged = pd.merge(input_df, output_df, left_on="JAN", right_on=jan_col, how="left")
        merged.drop(columns=[jan_col], inplace=True)

        # 保存フォルダの作成（出力ファイルと同じ場所に「結果」フォルダ）
        folder = os.path.dirname(output_path)
        result_folder = os.path.join(folder, "結果")
        os.makedirs(result_folder, exist_ok=True)

        # JAN整列結果を保存
        jan_result_path = os.path.join(result_folder, "JAN整列結果.xlsx")
        merged.to_excel(jan_result_path, index=False)

        # --- 価格分類 ---
        df = merged.copy()
        df.columns = df.columns.str.strip()

        if "価格" not in df.columns or "備考" not in df.columns:
            raise ValueError("❌ 「価格」または「備考」列が見つかりません。")

        # 価格列を数値化（変換できない値はNaN）
        df["価格数値"] = pd.to_numeric(df["価格"], errors="coerce")

        # 価格取得成功：価格が数値（NaN以外）
        df_success = df[df["価格数値"].notna()]

        # 商品が見つからなかったもの：備考に該当文字あり
        df_not_found = df[df["備考"].astype(str).str.contains("商品が見つからない", na=False)]

        # JAN以外が空欄の行
        non_jan_cols = [c for c in df.columns if c != "JAN"]
        df_blank_except_jan = df[df[non_jan_cols].isna().all(axis=1)]

        # 価格取得失敗：残り全て（成功・見つからない・空欄以外）
        df_fail = df[
            ~df.index.isin(df_success.index)
            & ~df.index.isin(df_not_found.index)
            | df.index.isin(df_blank_except_jan.index)
        ].drop_duplicates()

        # 補助列削除
        for d in (df_success, df_fail, df_not_found):
            d.drop(columns=["価格数値"], inplace=True)

        # 結果を出力（すべて結果フォルダ内）
        df_success.to_excel(os.path.join(result_folder, "価格取得成功.xlsx"), index=False)
        df_fail.to_excel(os.path.join(result_folder, "価格取得失敗.xlsx"), index=False)
        df_not_found.to_excel(os.path.join(result_folder, "商品が見つからなかったもの.xlsx"), index=False)

        # 完了メッセージ
        messagebox.showinfo(
            "完了",
            f"処理が完了しました！\n\n"
            f"📂 保存先フォルダ：\n{result_folder}\n\n"
            "✅ 出力ファイル：\n"
            "・JAN整列結果.xlsx\n"
            "・価格取得成功.xlsx\n"
            "・価格取得失敗.xlsx\n"
            "・商品が見つからなかったもの.xlsx"
        )

    except Exception as e:
        messagebox.showerror("エラー", f"処理中にエラーが発生しました：\n{e}")

# ======================
# 🎨 GUI構築
# ======================
def drop(event):
    """ドラッグ＆ドロップ時の処理"""
    global file_list
    dropped_files = list(root.tk.splitlist(event.data))
    for f in dropped_files:
        if f not in file_list:
            file_list.append(f)
    file_text.delete(1.0, tk.END)
    file_text.insert(tk.END, "\n".join(os.path.basename(f) for f in file_list))
    if len(file_list) == 2:
        process_files()

def clear_files():
    """ドロップされたファイルリストをリセット"""
    global file_list
    file_list = []
    file_text.delete(1.0, tk.END)
    file_text.insert(tk.END, "📂 ファイル未選択")
    messagebox.showinfo("リセット", "ファイルリストをリセットしました。")

# ======================
# 💻 UI設定
# ======================
root = TkinterDnD.Tk()
root.title("結果出力ツール")
root.geometry("550x400")
root.configure(bg="#d9d9d9")
root.resizable(False, False)

title_label = tk.Label(
    root,
    text="結果出力ツール\n\n"
         "入力したファイルと出力されたファイルを\n"
         "　ここにドラッグ＆ドロップしてください。\n"
  ,
    bg="#d9d9d9", fg="#333", font=("Meiryo", 14)
)
title_label.pack(pady=10)

drop_frame = tk.Label(
    root,
    text="⬇️ ここに2つのExcelをドロップ ⬇️",
    bg="#eeeeee", fg="#444",
    relief="ridge", width=55, height=5
)
drop_frame.pack(pady=5)
drop_frame.drop_target_register(DND_FILES)
drop_frame.dnd_bind("<<Drop>>", drop)

file_text = scrolledtext.ScrolledText(
    root, width=60, height=4, bg="#f5f5f5",
    fg="#333", font=("Meiryo", 9), wrap="none"
)
file_text.insert(tk.END, "📂 ファイル未選択")
file_text.pack(pady=5)

clear_button = tk.Button(
    root, text="🗑 ファイルをリセット", command=clear_files,
    bg="#c0c0c0", fg="#222", font=("Meiryo", 10, "bold"),
    relief="raised", width=20, height=1
)
clear_button.pack(pady=10)

file_list = []

root.mainloop()


# In[ ]:




