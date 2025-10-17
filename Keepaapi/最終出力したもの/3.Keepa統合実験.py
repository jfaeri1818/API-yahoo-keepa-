#!/usr/bin/env python
# coding: utf-8

"""
フロー：
 1) Excelファイル（JANの1列リスト）を1つドロップ
 2) Keepa APIキーを入力
 3) 自動で価格取得開始（GUIログ表示・10件ごとに進捗表示）
 4) 完了後、「JAN整列結果.xlsx」「価格取得成功.xlsx」
    「価格取得失敗.xlsx」「商品が見つからなかったもの.xlsx」を出力

出力先：
  デスクトップに「結果_YYYYMMDD_HHMMSS」フォルダを自動作成し、その中へ保存
"""

import os
import time
import datetime
import threading
import traceback
import ctypes
import requests
import pandas as pd
import tkinter as tk
from tkinter import messagebox, scrolledtext
from tkinterdnd2 import TkinterDnD, DND_FILES

# =========================
# 設定
# =========================
DOMAIN_JP = 5
MAX_SECONDS_ALLOWED = 10       # リクエストタイムアウト
TOKEN_WAIT_SECONDS = 1800      # トークン枯渇時の待機（30分）
STOP_FLAG = False

# =========================
# スリープ防止（Windows）
# =========================
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

# =========================
# Keepa API
# =========================
def fetch_top_display_price(api_key: str, code: str):
    """
    価格決定ロジック（BuyBox > Prime > 先頭オファー）
    戻り値: (title, total_price_or_None, error_message_or_None, hit_count)
    """
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

    # ✅ Prime優先 → なければ先頭
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
            return title, total, None, len(offers)

    hit_count = len(offers)
    if hit_count > 0:
        return title, None, f"価格取得失敗（{hit_count}件ヒット）", hit_count
    else:
        return title, None, "商品が見つからない", hit_count

# =========================
# 便利関数
# =========================
def ensure_result_folder():
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    folder = os.path.join(desktop, f"結果_{ts}")
    os.makedirs(folder, exist_ok=True)
    return folder

def classify_and_save(df, result_folder):
    """
    df: 列(JANコード, 価格, 商品名, 備考)
    価格を数値化して分類。3ファイル＋JAN整列結果.xlsx を保存
    """
    df = df.copy()

    rename_map = {}
    for c in df.columns:
        if "JAN" in str(c).upper():
            rename_map[c] = "JANコード"
    df.rename(columns=rename_map, inplace=True)

    df["価格数値"] = pd.to_numeric(df["価格"], errors="coerce")

    df_success = df[df["価格数値"].notna()].copy()
    df_not_found = df[df["備考"].astype(str).str.contains("商品が見つからない", na=False)].copy()
    df_fail = df[~df.index.isin(df_success.index) & ~df.index.isin(df_not_found.index)].copy()

    for d in (df_success, df_fail, df_not_found):
        if "価格数値" in d.columns:
            d.drop(columns=["価格数値"], inplace=True)
    df.drop(columns=["価格数値"], inplace=True)

    df.to_excel(os.path.join(result_folder, "JAN整列結果.xlsx"), index=False)
    df_success.to_excel(os.path.join(result_folder, "価格取得成功.xlsx"), index=False)
    df_fail.to_excel(os.path.join(result_folder, "価格取得失敗.xlsx"), index=False)
    df_not_found.to_excel(os.path.join(result_folder, "商品が見つからなかったもの.xlsx"), index=False)

# =========================
# メイン処理
# =========================
def run_keepa_then_align(api_key: str, jan_file_path: str, log_box: tk.Text, start_button: tk.Button):
    global STOP_FLAG
    STOP_FLAG = False
    prevent_sleep()
    start_button.config(state="disabled")

    result_folder = ensure_result_folder()

    try:
        df_in = pd.read_excel(jan_file_path, header=None)
    except Exception as e:
        messagebox.showerror("読込エラー", f"Excelファイルを開けませんでした。\n{e}")
        start_button.config(state="normal")
        allow_sleep()
        return

    df_in = df_in.iloc[:, :1].copy()
    df_in.columns = ["JANコード"]
    df_in["JANコード"] = df_in["JANコード"].astype(str).str.strip()
    total = len(df_in)

    log_box.insert(tk.END, f"📘 JANファイル読込: {jan_file_path}\n🔢 全{total}件の処理を開始します。\n\n")
    log_box.see(tk.END)

    results = []

    try:
        for i, row in df_in.iterrows():
            if STOP_FLAG:
                log_box.insert(tk.END, "🛑 強制停止を検出 → 現在の結果を出力中...\n")
                break

            jan = str(row["JANコード"]).strip()
            if not jan or jan.lower() == "nan":
                continue

            title, price, error, hit_count = fetch_top_display_price(api_key, jan)

            if error:
                if "トークン枯渇" in error:
                    log_box.insert(tk.END, f"🪙 {i+1}/{total} {jan} → トークン枯渇。{TOKEN_WAIT_SECONDS//60}分待機。\n")
                    log_box.see(tk.END)
                    for sec in range(TOKEN_WAIT_SECONDS, 0, -1):
                        if STOP_FLAG:
                            break
                        if sec % 60 == 0:
                            log_box.insert(tk.END, f"⏳ 残り {sec//60} 分...\n")
                            log_box.see(tk.END)
                        time.sleep(1)
                    if STOP_FLAG:
                        break
                    title, price, error, hit_count = fetch_top_display_price(api_key, jan)

                elif "通信エラー" in error or "処理時間超過" in error:
                    log_box.insert(tk.END, f"⚠️ {i+1}/{total} {jan} → {error} のためスキップ\n")
                    log_box.see(tk.END)
                    time.sleep(10)

            results.append({
                "JANコード": jan,
                "価格": price if price is not None else "Null",
                "商品名": title or "",
                "備考": error or ""
            })

            log_box.insert(tk.END, f"🕐 {i+1}/{total} 件完了\n")
            log_box.see(tk.END)

        df_keepa = pd.DataFrame(results)
        classify_and_save(df_keepa, result_folder)

        messagebox.showinfo(
            "完了",
            "処理が完了しました！\n\n"
            f"📂 保存先フォルダ：\n{result_folder}\n\n"
            "✅ 出力ファイル：\n"
            "・JAN整列結果.xlsx\n"
            "・価格取得成功.xlsx\n"
            "・価格取得失敗.xlsx\n"
            "・商品が見つからなかったもの.xlsx"
        )

    except Exception as e:
        log_box.insert(tk.END, f"⚠️ エラー発生: {e}\n{traceback.format_exc()}")
        messagebox.showerror("エラー", f"処理中に問題が発生しました。\n{e}")

    finally:
        start_button.config(state="normal")
        allow_sleep()

# =========================
# GUI
# =========================
class App:
    def __init__(self):
        self.root = TkinterDnD.Tk()
        self.root.title("Keepa自動実行 → JAN整列ツール")
        self.root.geometry("560x490")
        self.root.configure(bg="#f5f0e6")
        self.root.resizable(False, False)

        title = tk.Label(
            self.root,
            text="Excelを1つドロップ → APIキー入力 → 自動で価格取得 → 自動でJAN整列",
            bg="#f5f0e6", fg="#333", font=("Meiryo", 12, "bold")
        )
        title.pack(pady=(10, 6))

        self.drop_frame = tk.Label(
            self.root,
            text="⬇️ ここにExcel（JAN1列）を1つドロップ ⬇️",
            bg="#ffffff", fg="#444",
            relief="ridge", width=62, height=4
        )
        self.drop_frame.pack(pady=6)
        self.drop_frame.drop_target_register(DND_FILES)
        self.drop_frame.dnd_bind("<<Drop>>", self.on_drop)

        self.file_label = tk.Label(
            self.root, text="📂 ファイル未選択", bg="#f5f0e6",
            fg="#333", font=("Meiryo", 10)
        )
        self.file_label.pack(pady=(2, 6))

        api_row = tk.Frame(self.root, bg="#f5f0e6")
        api_row.pack(pady=(2, 2))
        tk.Label(api_row, text="Keepa APIキー：", bg="#f5f0e6",
                 font=("Meiryo", 10, "bold")).pack(side="left")
        self.api_entry = tk.Entry(api_row, width=50, show="*")
        self.api_entry.pack(side="left", padx=(4, 0))
        self.api_entry.bind("<Return>", self.try_auto_start)

        frame_log = tk.Frame(self.root, bg="#f5f0e6")
        frame_log.pack(padx=10, pady=(6, 0), fill="both", expand=True)
        self.log_box = tk.Text(frame_log, height=13, width=70, font=("Meiryo", 9))
        self.log_box.pack(side="left", fill="both", expand=True)
        scrollbar = tk.Scrollbar(frame_log, command=self.log_box.yview)
        scrollbar.pack(side="right", fill="y")
        self.log_box.config(yscrollcommand=scrollbar.set)

        btn_row = tk.Frame(self.root, bg="#f5f0e6")
        btn_row.pack(pady=8)
        self.start_button = tk.Button(
            btn_row, text="▶ 開始（手動）", bg="#4CAF50", fg="white",
            font=("Meiryo", 10, "bold"), width=18, command=self.manual_start
        )
        self.start_button.pack(side="left", padx=10)

        self.stop_button = tk.Button(
            btn_row, text="■ 強制停止", bg="#d9534f", fg="white",
            font=("Meiryo", 10, "bold"), width=18, command=self.force_stop
        )
        self.stop_button.pack(side="left", padx=10)

        self.jan_file_path = None
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def on_close(self):
        if messagebox.askyesno("確認", "本当に終了しますか？"):
            allow_sleep()
            self.root.destroy()

    def on_drop(self, event):
        files = list(self.root.tk.splitlist(event.data))
        if not files:
            return
        self.jan_file_path = files[0]
        self.file_label.config(text=self.jan_file_path)
        self.log_box.insert(tk.END, "✅ ファイルを受け取りました。APIキーを入力してください。\n")
        self.log_box.see(tk.END)
        self.api_entry.focus_set()

    def ready_to_start(self):
        return bool(self.jan_file_path) and bool(self.api_entry.get().strip())

    def try_auto_start(self, _evt=None):
        if self.ready_to_start():
            self.log_box.insert(tk.END, "🚀 APIキー入力を検知 → 自動開始します。\n")
            self.log_box.see(tk.END)
            self.start_thread()

    def manual_start(self):
        if not self.jan_file_path:
            messagebox.showwarning("注意", "先にExcel（JAN1列）をドロップしてください。")
            return
        if not self.api_entry.get().strip():
            messagebox.showwarning("注意", "Keepa APIキーを入力してください。")
            self.api_entry.focus_set()
            return
        self.start_thread()

    def start_thread(self):
        api_key = self.api_entry.get().strip()
        threading.Thread(
            target=run_keepa_then_align,
            args=(api_key, self.jan_file_path, self.log_box, self.start_button),
            daemon=True
        ).start()

    def force_stop(self):
        global STOP_FLAG
        STOP_FLAG = True
        self.log_box.insert(tk.END, "\n🛑 強制終了ボタンが押されました。\n")
        self.log_box.see(tk.END)

    def run(self):
        self.root.mainloop()

# =========================
# 実行エントリ
# =========================
if __name__ == "__main__":
    App().run()
