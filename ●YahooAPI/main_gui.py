# ============================================================
# 🌸 Yahoo!商品情報取得ツール - Flower Edition（GUI）
# ============================================================

import tkinter as tk
from tkinter import ttk, messagebox
from tkinterdnd2 import TkinterDnD
import threading
from yahoo_api import run_yahoo_api   # yahooapi 内のrun yahoo関数を使えるようにする

# ============================================================
# GUI本体
# ============================================================
root = TkinterDnD.Tk()
root.title("Yahoo!商品情報取得ツール - Flower Edition")
root.resizable(False, False)
root.minsize(600, 620)        # ← ★追加：最小サイズ固定
root.maxsize(600, 620)        # ← ★追加：最大サイズ固定


# スタイル設定
style = ttk.Style()
style.theme_use("clam")
style.configure("TLabel", font=("Meiryo", 10))
style.configure("TEntry", font=("Meiryo", 10))
style.configure("TButton", font=("Meiryo", 10, "bold"), padding=4)

frame = ttk.Frame(root, padding=20)
frame.pack(fill="both", expand=True)

# 入力欄を生成するヘルパー関数
def create_input_row(parent, label, row, show=None, default=""):
    lbl = ttk.Label(parent, text=label)
    lbl.grid(row=row, column=0, sticky="e", padx=8, pady=8)
    entry = ttk.Entry(parent, width=45, show=show)
    entry.insert(0, default)
    entry.grid(row=row, column=1, padx=8, pady=8)
    return entry

# 入力欄
# api_url_entry = create_input_row(frame, "API URL：", 0, default="https://shopping.yahooapis.jp/ShoppingWebService/V3/itemSearch")
# app_id_entry = create_input_row(frame, "Client ID：", 1, show="●",default="dj00aiZpPXlkOGd5bDlUcTlWRyZzPWNvbnN1bWVyc2VjcmV0Jng9MmE-")
api_url_entry = create_input_row(frame, "API URL：", 0, show="*")
app_id_entry = create_input_row(frame, "Client ID：", 1, show="*")

#開始番号
# lbl_start = ttk.Label(frame, text="開始番号：")
# lbl_start.grid(row=2, column=0, sticky="e", padx=8, pady=8)
# start_values = ["1"] + [str(i) for i in range(10000, 1000001, 10000)]
# start_number_combo = ttk.Combobox(frame, values=start_values, width=44, state="readonly")
# start_number_combo.set("1")
# start_number_combo.grid(row=2, column=1, padx=8, pady=8)

#価格レンジ
low_price_entry = create_input_row(frame, "下限価格：", 3)
high_price_entry = create_input_row(frame, "上限価格：", 4)

# id
seller_id_entry = create_input_row(frame, "販売者ID：", 5, default="hands-net")

# ログ表示欄
log_frame = ttk.LabelFrame(frame, text="進行ログ", padding=10)
log_frame.grid(row=6, column=0, columnspan=2, sticky="nsew", pady=8)

log_text = tk.Text(log_frame, height=12, width=65, font=("Consolas", 9), bg="#f9f9f9", wrap="word")
log_text.pack(side="left", fill="both", expand=True)
scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=log_text.yview)
scrollbar.pack(side="right", fill="y")
log_text.config(yscrollcommand=scrollbar.set)

def append_log(text):
    log_text.insert(tk.END, text + "\n")
    log_text.see(tk.END)

# 実行スレッド
def start_threaded(mode):
    client_id = app_id_entry.get().strip()
    seller_id = seller_id_entry.get().strip()
    low_price = low_price_entry.get().strip()
    high_price = high_price_entry.get().strip()

    # start_num = int(start_number_combo.get().strip())
    api_url = api_url_entry.get().strip()

    if not client_id or not seller_id or not api_url:
        messagebox.showwarning("入力不足", "API URL・Client ID・販売者IDをすべて入力してください。")
        return

    append_log(f"[INFO] {mode} を開始します...\n")

    threading.Thread(
        target=run_yahoo_api,  # ← yahoo_api.py の関数を呼び出す
        args=(client_id, mode, seller_id, api_url, append_log, low_price, high_price),
        daemon=True
    ).start()

# ボタン配置
ttk.Button(frame, text="商品数を調べる", width=30, command=lambda: start_threaded("count")).grid(row=7, column=0, columnspan=2, pady=8)
ttk.Button(frame, text="商品取得を実行", width=30, command=lambda: start_threaded("normal")).grid(row=8, column=0, columnspan=2, pady=8)

root.mainloop()
