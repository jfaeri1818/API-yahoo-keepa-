import requests
import sys
import os
import pandas as pd
import time
import tkinter as tk
from tkinter import ttk, messagebox
from tkinterdnd2 import TkinterDnD
import threading

# ============================================================
# 商品取得処理
# ============================================================
def run_yahoo_api(app_id, mode, seller_id, start_number, api_url, log_callback):
    try:
        result_log = "result.txt"

        # ============================================================
        # 件数確認モード
        # ============================================================
        if mode == "count":
            log_callback(f"[INFO] 商品数を調べています（販売者ID: {seller_id}）...")

            params = {
                "appid": app_id,
                "seller_id": seller_id,
                "results": 1,
                "start": 1,
                "sort": "+price"
            }

            try:
                response = requests.get(api_url, params=params, timeout=10)
                data = response.json()
                total_available = data.get("totalResultsAvailable", 0)

                result_text = f"[RESULT] 条件に一致した全体のヒット件数: {total_available:,} 件"
                log_callback(result_text)
                with open(result_log, "w", encoding="utf-8") as f:
                    f.write(result_text)
            except Exception as e:
                err = f"[ERROR] エラー発生: {e}"
                log_callback(err)
                with open(result_log, "w", encoding="utf-8") as f:
                    f.write(err)
            return

        # ============================================================
        # 通常モード（1万件単位で取得）
        # ============================================================
        total_items = 10000           # ← 1回の実行で取得したい件数
        results_per_call = 50         # ← Yahoo! APIの1回上限
        calls = total_items // results_per_call  # 200回
        wait_sec = 0.8                # 呼び出し間隔（秒）

        output_file = f"{seller_id}_商品情報_{start_number}.xlsx"
        log_callback(f"[INFO] 商品取得を開始（開始番号: {start_number}）")
        all_rows = []

        for i in range(calls):
            start = start_number + i * results_per_call
            params = {
                "appid": app_id,
                "seller_id": seller_id,
                "results": results_per_call,
                "start": start,
                "sort": "+price",
                "condition": "new"
            }

            try:
                response = requests.get(api_url, params=params, timeout=10)
                data = response.json()
                hits = data.get("hits", [])
                total_available = data.get("totalResultsAvailable", 0)

                if not hits:
                    log_callback(f"[WARN] {i+1}回目（start={start}）に商品データなし。終了します。")
                    break

                for h in hits:
                    name = h.get("name") or ""
                    in_stock = h.get("inStock")
                    price = h.get("price") or ""
                    jan = h.get("janCode") or ""
                    all_rows.append([name, in_stock, price, jan])

                log_callback(f"[OK] {i+1}/{calls} ページ完了（start={start}〜{start + results_per_call - 1}）")
                time.sleep(wait_sec)

            except Exception as e:
                log_callback(f"[ERROR] エラー発生: {e}")
                log_callback("[WAIT] 30秒待機して再試行します...")
                time.sleep(30)

        # ============================================================
        # 保存処理
        # ============================================================
        df = pd.DataFrame(all_rows, columns=["商品名", "在庫あり", "価格", "JANコード"])
        df.to_excel(output_file, index=False)

        summary_text = (
            f"[DONE] 取得完了: {len(all_rows)}件\n"
            f"[NEXT] 次回の開始番号: {start_number + len(all_rows)}\n"
            f"[FILE] 保存先: {output_file}\n"
            f"[RESULT] 条件に一致した全体のヒット件数: {total_available:,} 件\n"
        )
        log_callback(summary_text)

        with open(result_log, "w", encoding="utf-8") as f:
            f.write(summary_text)

    except Exception as e:
        log_callback(f"[ERROR] 処理全体で例外発生: {e}")

# ============================================================
# GUIモード
# ============================================================
root = TkinterDnD.Tk()
root.title("Yahoo!商品情報取得ツール - Flower Edition")
root.geometry("600x620")
root.resizable(False, False)

style = ttk.Style()
style.theme_use("clam")
style.configure("TLabel", font=("Meiryo", 10))
style.configure("TEntry", font=("Meiryo", 10))
style.configure("TButton", font=("Meiryo", 10, "bold"), padding=4)
style.configure("TCombobox", font=("Meiryo", 10))

frame = ttk.Frame(root, padding=20)
frame.pack(fill="both", expand=True)

def create_input_row(parent, label, row, show=None, default=""):
    lbl = ttk.Label(parent, text=label)
    lbl.grid(row=row, column=0, sticky="e", padx=8, pady=8)
    entry = ttk.Entry(parent, width=45, show=show)
    entry.insert(0, default)
    entry.grid(row=row, column=1, padx=8, pady=8)
    return entry

api_url_entry = create_input_row(frame, "API URL：", 0, default="https://shopping.yahooapis.jp/ShoppingWebService/V3/itemSearch")
app_id_entry = create_input_row(frame, "Client ID：", 1, show="●")
lbl_start = ttk.Label(frame, text="開始番号：")
lbl_start.grid(row=2, column=0, sticky="e", padx=8, pady=8)
start_values = ["1"] + [str(i) for i in range(10000, 1000001, 10000)]
start_number_combo = ttk.Combobox(frame, values=start_values, width=44, state="readonly")
start_number_combo.set("1")
start_number_combo.grid(row=2, column=1, padx=8, pady=8)
seller_id_entry = create_input_row(frame, "販売者ID：", 3, default="hands-net")

separator = ttk.Separator(frame, orient="horizontal")
separator.grid(row=4, column=0, columnspan=2, sticky="ew", pady=10)

count_label = ttk.Label(frame, text="取得できる商品数：未取得", font=("Meiryo", 10, "bold"), foreground="#0078D4")
count_label.grid(row=5, column=0, columnspan=2, pady=5)

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

def start_threaded(mode):
    client_id = app_id_entry.get().strip()
    seller_id = seller_id_entry.get().strip()
    start_num = int(start_number_combo.get().strip())
    api_url = api_url_entry.get().strip()

    if not client_id or not seller_id or not api_url:
        messagebox.showwarning("入力不足", "API URL・Client ID・販売者IDをすべて入力してください。")
        return

    append_log(f"[INFO] {mode} を開始します...\n")
    threading.Thread(
        target=run_yahoo_api,
        args=(client_id, mode, seller_id, start_num, api_url, append_log),
        daemon=True
    ).start()

check_button = ttk.Button(frame, text="商品数を調べる", width=30, command=lambda: start_threaded("count"))
check_button.grid(row=7, column=0, columnspan=2, pady=8)
execute_button = ttk.Button(frame, text="商品取得を実行", width=30, command=lambda: start_threaded("normal"))
execute_button.grid(row=8, column=0, columnspan=2, pady=8)

root.mainloop()
