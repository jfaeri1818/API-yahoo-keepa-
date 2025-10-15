# ============================================================
# 🌸 Yahoo!ショッピングAPI ロジック部分
# ============================================================

import requests             # ← Yahoo!APIにアクセスするための通信ライブラリ
import pandas as pd         # ← データを表形式にしてExcelに保存するために使用
import time                 # ← API呼び出しの間に少し休ませるために使用

def run_yahoo_api(app_id, mode, seller_id, api_url, log_callback, low_price=None, high_price=None):
    """
    Yahoo!ショッピングAPIから商品情報を取得・件数確認を行うメイン処理。
    mode: "count"（件数確認）または "normal"（商品取得）
    log_callback: GUI側から渡されるログ出力用関数
    """

    try:
        result_log = "result.txt"  # ← 実行結果（ログ）を保存するファイル名を設定

        # ============================================================
        # 件数確認モード（全体のヒット件数だけ調べる）
        # ============================================================
        if mode == "count":  # ← 「件数確認モード」で呼ばれた場合
            log_callback(f"[INFO] 商品数を調べています（販売者ID: {seller_id}）...")  # ← GUIログに出力

            params = {                     # ← Yahoo!APIに送る検索条件をセット
                "appid": app_id,            # ← あなたのYahoo!アプリケーションID
                "seller_id": seller_id,     # ← 出店者ID（例: hands-net）
                "results": 1,               # ← 1件だけ取得（件数調査目的なので1件でOK）
                "start": 1,                 # ← 開始位置（最初の1件）
                "sort": "+price" ,           # ← 価格の昇順でソート
                "condition": "new" ,
                         }
            if low_price:
                params["price_from"] = int(low_price)
            if high_price:
                params["price_to"] = int(high_price)
            
            

            try:
                response = requests.get(api_url, params=params, timeout=10)  # ← APIを呼び出す（10秒でタイムアウト）
                data = response.json()                                       # ← 結果をJSON形式で読み取る
                total_available = data.get("totalResultsAvailable", 0)       # ← 総ヒット件数を取得

                result_text = f"[RESULT] 条件に一致した全体のヒット件数: {total_available:,} 件"  # ← 表示用テキストを作る
                log_callback(result_text)                                     # ← GUIログに出力
                with open(result_log, "w", encoding="utf-8") as f:           # ← result.txtに保存
                    f.write(result_text)

            except Exception as e:
                err = f"[ERROR] エラー発生: {e}"                            # ← エラー内容を文字列に変換
                log_callback(err)                                            # ← GUIログに出力
                with open(result_log, "w", encoding="utf-8") as f:           # ← result.txtにも保存
                    f.write(err)
            return  # ← 件数モードはここで終了

        # ============================================================
        # 通常モード（商品情報をすべて取得）
        # ============================================================
        total_items = 1000              # ← 最大で取得したい商品件数（仮に1万件）
        results_per_call = 50            # ← 1回のAPI呼び出しで取れる上限（Yahooの仕様）
        calls = total_items // results_per_call  # ← 全部で何回呼び出すか（1万 ÷ 50 = 200回）
        wait_sec = 0.8                   # ← 呼び出し間隔（秒）を設定（Yahoo制限対策）

        output_file = f"{seller_id}_商品情報_{low_price or 'min'}-{high_price or 'max'}.xlsx"
                      # ← 出力ファイル名を定義
        log_callback(f"[INFO] 商品取得を開始）")  # ← GUIに開始メッセージを表示
        all_rows = []                    # ← 取得した商品データを入れるリスト

        # ============================================================
        # 🔁 ページごとに繰り返し取得（200回分）
        # ============================================================
        for i in range(calls):  
            start = 1 + results_per_call+i*50         
            
            params = {                                        # ← APIに送る条件を設定
                "appid": app_id,
                "seller_id": seller_id,
                "results": results_per_call,#一回で呼び出す数
                "start": start,#取得するスタート位置
                "sort": "+price",
                "condition": "new" ,
                "price_from": int(low_price), # 下限価格
                "price_to":int(high_price)    # 上限価格
                    }

            try:
                response = requests.get(api_url, params=params, timeout=10)  # ← API呼び出し
                data = response.json()                                       # ← 結果をJSON形式で受け取る
                hits = data.get("hits", [])                                  # ← 商品リストを取り出す
                total_available = data.get("totalResultsAvailable", 0)       # ← 全体件数を取得（参考）

                if not hits:                                                 # ← 商品データがなければ終了
                    log_callback("これ以上商品データなし。終了します。")
                    break

                # ============================================================
                # 商品1件ずつ取り出してリストに追加
                # ============================================================
                for h in hits:
                    name = h.get("name") or ""       # ← 商品名
                    in_stock = h.get("inStock")      # ← 在庫あり(True/False)
                    price = h.get("price") or ""     # ← 価格
                    jan = h.get("janCode") or ""     # ← JANコード
                    all_rows.append([name, in_stock, price, jan])  # ← 一行ずつ追加

                log_callback(f"[OK] {i+1}/{calls} ページ完了")  # ← 進行表示
                time.sleep(wait_sec)  # ← Yahooへの負荷軽減のため、0.8秒休む

            except Exception as e:  # ← 通信エラーが出たとき
                log_callback(f"[ERROR] エラー発生: {e}")             # ← GUIにエラーログ出力
                log_callback("[WAIT] 30秒待機して再試行します...")   # ← ユーザーに待機案内
                time.sleep(30)                                       # ← 30秒待って再開

        # ============================================================
        # 保存処理（Excelファイルに出力）
        # ============================================================
        df = pd.DataFrame(all_rows, columns=["商品名", "在庫あり", "価格", "JANコード"])  # ← 取得データを表形式に変換
        df.to_excel(output_file, index=False)  # ← Excelに書き出し

        summary_text = (  # ← 完了報告用のメッセージをまとめる
            f"[DONE] 取得完了: {len(all_rows)}件\n"
            # f"[NEXT] 次回の開始番号: {start_number + len(all_rows)}\n"
            f"[FILE] 保存先: {output_file}\n"

        )
        log_callback(summary_text)  # ← GUIに最終結果を表示

        with open(result_log, "w", encoding="utf-8") as f:  # ← result.txt にも最終結果を保存
            f.write(summary_text)

    except Exception as e:
        log_callback(f"[ERROR] 処理全体で例外発生: {e}")  # ← 予期しないエラーが発生した場合
