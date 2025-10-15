#!/usr/bin/env python
# coding: utf-8

# In[3]:


import requests
import time
import pandas as pd
import tkinter as tk
from tkinter import messagebox, filedialog
import threading
import datetime
import traceback
import ctypes
import os

# ============================================================
# 設定値
# ============================================================
DOMAIN_JP = 5
STOP_FLAG = False
MAX_TOKENS_PER_ITEM = 10
MAX_SECONDS_ALLOWED = 10       # タイムアウト：10秒
ERROR_WAIT_TIME = 1800          # エラー時の待機時間（秒）＝5分
SAVE_INTERVAL = 10             # ✅ 10件ごとに保存

# ============================================================
# スリープ防止（Windows）
# ============================================================
ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001

def prevent_sleep():
    try:
        ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS | ES_SYSTEM_REQUIRED)
    except Exception:
        pass

def allow_sleep():
    try:
        ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)
    except Exception:
        pass

# ============================================================
# Keepa API呼び出し
# ============================================================
def fetch_top_display_price(api_key: str, code: str):
    url = (
        f"https://api.keepa.com/product?key={api_key}"
        f"&domain={DOMAIN_JP}&code={code}"
        "&history=0&offers=20&onlyLiveOffers=0&buybox=1&stats=0"
    )
    start_time = time.time()
    try:
        resp = requests.get(url, timeout=MAX_SECONDS_ALLOWED)
        if resp.status_code == 429:
            return None, None, "トークン枯渇", 0
        data = resp.json()
    except Exception as e:
        return None, None, f"通信エラー: {e}", 0

    if time.time() - start_time > MAX_SECONDS_ALLOWED:
        return None, None, f"処理時間超過（{MAX_SECONDS_ALLOWED}秒）", 0
    if not data or "products" not in data:
        return None, None, "データなし", 0

    products = data["products"]
    if not products:
        return None, None, "商品が見つからない", 0

    product = products[0]
    title = product.get("title", "")
    stats = product.get("stats") or {}

    # ✅ BuyBox優先
    for key in ("buyBoxPrice", "buyBoxShippingPrice", "current_BUY_BOX_SHIPPING"):
        v = stats.get(key)
        if isinstance(v, (int, float)) and v > 0:
            return title, int(v), None, 0

    # ✅ Prime優先（なければ最初の出品）
    offers = product.get("offers") or []
    order = product.get("liveOffersOrder") or []
    ordered = [offers[i] for i in order if isinstance(i, int) and i < len(offers)]
    if not ordered and offers:
        ordered = offers

    prime_offer = next((o for o in ordered if o.get("isPrime")), None)
    chosen = prime_offer or (ordered[0] if ordered else None)

    if chosen:
        price = chosen.get("price")
        ship = chosen.get("shipping") or 0
        if price and price > 0:
            total = int(price) + int(ship)
            return title, total, None, 0

    hit_count = len(offers)
    if hit_count > 0:
        return title, None, f"価格取得失敗（{hit_count}件ヒット）", 0
    else:
        return title, None, "商品が見つからない", 0

# ============================================================
# ログ出力まとめ
# ============================================================
def flush_logs(log_box, buffer):
    if not buffer:
        return
    log_box.insert(tk.END, "".join(buffer))
    log_box.see(tk.END)
    buffer.clear()

# ============================================================
# メイン処理
# ============================================================
def start_process(api_key, filepath, log_box, start_button):
    global STOP_FLAG
    STOP_FLAG = False
    prevent_sleep()
    start_button.config(state="disabled")

    try:
        df = pd.read_excel(filepath, header=None)
    except Exception as e:
        messagebox.showerror("読込エラー", f"Excelファイルを開けませんでした。\n{e}")
        start_button.config(state="normal")
        allow_sleep()
        return

    total = len(df)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    output_file = os.path.join(desktop_path, f"結果_{timestamp}.xlsx")

    log_box.insert(tk.END, f"📘 ファイル読込完了: {filepath}\n🔢 全{total}件の処理を開始します。\n\n")
    log_box.see(tk.END)

    results = []
    log_buffer = []

    try:
        for i, row in df.iterrows():
            if STOP_FLAG:
                log_box.insert(tk.END, "🛑 強制停止を検出 → 現在の結果を保存中...\n")
                pd.DataFrame(results).to_excel(output_file, index=False)
                log_box.insert(tk.END, f"💾 中断時の結果を保存しました → {output_file}\n")
                break

            jan = str(row.iloc[0]).strip() if len(row) > 0 else ""
            if not jan or jan.lower() == "nan":
                continue

            title, price, error, _ = fetch_top_display_price(api_key, jan)

            if error:
                if "トークン枯渇" in error:
                    log_box.insert(tk.END, f"🪙 {i+1}/{total} {jan} → トークン枯渇。30分待機。\n")
                    log_box.see(tk.END)
                    time.sleep(1800)
                    continue
                elif "通信エラー" in error:
                    log_box.insert(tk.END, f"⚠️ {i+1}/{total} {jan} → 処理時間超過のためスキップ\n")
                    log_box.see(tk.END)
                    time.sleep(10)
                    continue

            results.append({
                "JANコード": jan,
                "価格": price if price is not None else "Null",
                "商品名": title or "",
                "備考": error or ""
            })

            log_buffer.append(f"🕐 {i+1}/{total} 件処理完了\n")
            if len(log_buffer) >= 1:
                flush_logs(log_box, log_buffer)

            if (i + 1) % SAVE_INTERVAL == 0:
                pd.DataFrame(results).to_excel(output_file, index=False)
                log_box.insert(tk.END, f"💾 {i+1}件完了 → 一時保存しました。\n")
                log_box.see(tk.END)

        pd.DataFrame(results).to_excel(output_file, index=False)
        log_box.insert(tk.END, f"\n🎉 完了！結果を「{output_file}」に保存しました。\n")
        messagebox.showinfo("完了", f"処理が完了しました！\n結果ファイル: {output_file}")

    except Exception as e:
        log_box.insert(tk.END, f"⚠️ エラー発生: {e}\n{traceback.format_exc()}")
        messagebox.showerror("エラー", f"処理中に問題が発生しました。\n{output_file}")

    finally:
        start_button.config(state="normal")
        allow_sleep()

# ============================================================
# GUI構築（TkinterDnDなし・高速化）
# ============================================================
def create_gui():
    root = tk.Tk()
    root.title("Keepa価格取得ツール")
    root.geometry("420x330")
    root.configure(bg="#f5f0e6")
    root.resizable(False, False)

    def on_close():
        if messagebox.askyesno("確認", "本当に終了しますか？"):
            allow_sleep()
            root.destroy()
    root.protocol("WM_DELETE_WINDOW", on_close)

    var_topmost = tk.BooleanVar(value=True)
    chk_top = tk.Checkbutton(root, text="常に前面に表示", variable=var_topmost,
                             bg="#f5f0e6", font=("Meiryo", 9),
                             command=lambda: root.attributes("-topmost", var_topmost.get()))
    chk_top.pack(anchor="e", padx=10, pady=(3, 0))
    root.attributes("-topmost", True)

    tk.Label(root, text="Keepa APIキー：", bg="#f5f0e6", font=("Meiryo", 10, "bold")).pack(anchor="w", padx=10, pady=2)
    api_entry = tk.Entry(root, width=55, show="*")
    api_entry.pack(padx=10)

    tk.Label(root, text="Excelファイル：", bg="#f5f0e6", font=("Meiryo", 10, "bold")).pack(anchor="w", padx=10, pady=2)
    file_label = tk.Label(root, text="（ファイルを選択してください）", bg="white",
                          width=55, height=1, relief="groove", font=("Meiryo", 9))
    file_label.pack(padx=10, pady=2)

    def select_file():
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filepath:
            file_label.config(text=filepath)
            file_label.filepath = filepath

    tk.Button(root, text="ファイルを選択", command=select_file, font=("Meiryo", 9), width=18).pack(pady=3)

    frame_log = tk.Frame(root, bg="#f5f0e6")
    frame_log.pack(padx=10, pady=(5, 0), fill="both", expand=True)
    log_box = tk.Text(frame_log, height=6, width=55, font=("Meiryo", 8))
    log_box.pack(side="left", fill="both", expand=True)
    scrollbar = tk.Scrollbar(frame_log, command=log_box.yview)
    scrollbar.pack(side="right", fill="y")
    log_box.config(yscrollcommand=scrollbar.set)

    frame_buttons = tk.Frame(root, bg="#f5f0e6")
    frame_buttons.pack(pady=5)

    def force_stop():
        global STOP_FLAG
        STOP_FLAG = True
        log_box.insert(tk.END, "\n🛑 強制終了ボタンが押されました。\n")
        log_box.see(tk.END)

    start_button = tk.Button(frame_buttons, text="▶ 開始", bg="#4CAF50", fg="white",
                             font=("Meiryo", 10, "bold"), width=14)
    start_button.pack(side="left", padx=15)

    tk.Button(frame_buttons, text="■ 強制終了", bg="#d9534f", fg="white",
              font=("Meiryo", 10, "bold"), width=14, command=force_stop).pack(side="right", padx=15)

    start_button.config(command=lambda: threading.Thread(
        target=start_process,
        args=(api_entry.get().strip(), getattr(file_label, "filepath", None), log_box, start_button),
        daemon=True
    ).start())

    root.mainloop()

# ============================================================
# 軽量
# ============================================================
if __name__ == "__main__":
    create_gui()


# In[ ]:




