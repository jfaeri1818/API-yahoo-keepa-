# ============================================================
# 🌸 Yahoo!ショッピングAPI ロジック部分（保存先選択対応版）
# ============================================================

import requests
import pandas as pd
import time
from tkinter import filedialog  # ✅ 追加：保存先を選択するために必要

def run_yahoo_api(app_id, mode, seller_id, api_url, log_callback, low_price=None, high_price=None):
    """
    Yahoo!ショッピングAPIから商品情報を取得・件数確認を行うメイン処理。
    mode: "count"（件数確認）または "normal"（商品取得）
    log_callback: GUI側から渡されるログ出力用関数
    """

    try:
        result_log = "result.txt"  # 実行結果ログファイル

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
                "sort": "+price",
                "condition": "new"
            }

            if low_price:
                params["price_from"] = int(low_price)
            if high_price:
                params["price_to"] = int(high_price)

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
            return  # 件数モード終了

        # ============================================================
        # 通常モード（商品情報取得）
        # ============================================================
        total_items = 1000
        results_per_call = 50
        calls = total_items // results_per_call
        wait_sec = 0.8

        log_callback(f"[INFO] 商品取得を開始します...")
        all_rows = []

        for i in range(calls):
            start = 1 + results_per_call * i
            params = {
                "appid": app_id,
                "seller_id": seller_id,
                "results": results_per_call,
                "start": start,
                "sort": "+price",
                "condition": "new"
            }
            if low_price:
                params["price_from"] = int(low_price)
            if high_price:
                params["price_to"] = int(high_price)

            try:
                response = requests.get(api_url, params=params, timeout=10)
                data = response.json()
                hits = data.get("hits", [])
                total_available = data.get("totalResultsAvailable", 0)

                if not hits:
                    log_callback("これ以上商品データがありません。終了します。")
                    break

                for h in hits:
                    name = h.get("name") or ""
                    in_stock = h.get("inStock")
                    price = h.get("price") or ""
                    jan = h.get("janCode") or ""
                    all_rows.append([name, in_stock, price, jan])

                log_callback(f"[OK] {i+1}/{calls} ページ完了")
                time.sleep(wait_sec)

            except Exception as e:
                log_callback(f"[ERROR] エラー発生: {e}")
                log_callback("[WAIT] 30秒待機して再試行します...")
                time.sleep(30)

        # ============================================================
        # 保存処理（保存先をユーザーが選択）
        # ============================================================
        df = pd.DataFrame(all_rows, columns=["商品名", "在庫あり", "価格", "JANコード"])

        # ✅ 保存先ダイアログ
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excelファイル", "*.xlsx")],
            initialfile=f"{seller_id}_商品情報_{low_price or 'min'}-{high_price or 'max'}.xlsx",
            title="保存先を選択してください"
        )

        if save_path:
            df.to_excel(save_path, index=False)
            summary_text = (
                f"[DONE] 取得完了: {len(all_rows)}件\n"
                f"[FILE] 保存先: {save_path}\n"
            )
        else:
            summary_text = (
                f"[CANCELLED] 保存がキャンセルされました。\n"
                f"[INFO] 取得件数: {len(all_rows)}件（未保存）\n"
            )

        log_callback(summary_text)
        with open(result_log, "w", encoding="utf-8") as f:
            f.write(summary_text)

    except Exception as e:
        log_callback(f"[ERROR] 処理全体で例外発生: {e}")
