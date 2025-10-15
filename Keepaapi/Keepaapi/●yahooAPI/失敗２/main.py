# ============================================================
# 🌸 Yahoo!ショッピングAPIから商品情報を取得してExcelに保存するスクリプト（main.py）
# ============================================================
# このファイルは、GUIツール（Yahoo!商品情報取得ツール - Flower Edition）から呼び出されます。
# 入力された App ID・出店者ID・件数をもとに Yahoo!ショッピングAPI から商品情報を取得します。

import requests      # インターネット通信でAPIを呼び出すために使用
import pandas as pd  # 表形式データ（表・リスト）を扱うため
import time          # 一時停止（ウェイト）を入れるため
import sys           # コマンドライン引数を受け取るため
import os            # ファイル保存用にパス操作で使用

# ============================================================
# 💡 Yahoo!ショッピングAPIから商品情報を取得する関数
# ============================================================
def fetch_yahoo_items(APP_ID, SELLER_ID, RESULTS_PER_CALL, TOTAL_ITEMS):
    API_URL = "https://shopping.yahooapis.jp/ShoppingWebService/V3/itemSearch"  # Yahoo!APIのURL
    CALLS = TOTAL_ITEMS // RESULTS_PER_CALL   # APIを何回呼ぶか（総件数 ÷ 1回の取得数）
    all_rows = []   # 商品データをためていくリスト

    print("==============================================")
    print(f"🌸 {SELLER_ID} の商品情報を取得中...")
    print("==============================================")

    # ============================================================
    # 🔁 ページごとに商品を取得
    # ============================================================
    for i in range(CALLS):
        start = i * RESULTS_PER_CALL + 1  # 開始位置を計算（ページ番号）
        params = {
            "appid": APP_ID,             # あなたのYahoo! App ID
            "seller_id": SELLER_ID,      # 出店者ID（例: hands-net）
            "results": RESULTS_PER_CALL, # 1回の取得件数
            "start": start,              # 取得開始位置
            "sort": "+price",            # 価格の昇順で並べ替え
            "condition": "new"           # 新品商品のみ取得
        }

        try:
            # 📡 APIにリクエスト送信
            response = requests.get(API_URL, params=params, timeout=10)
            data = response.json()  # 結果をJSON形式で取得
            hits = data.get("hits", [])  # 商品リストを取り出す

            # 商品がない場合は終了
            if not hits:
                print(f"{i+1}ページ目: 商品なし（終了）")
                break

            # ============================================================
            # 📦 商品ごとの情報を整理して保存
            # ============================================================
            for h in hits:
                name = h.get("name") or ""       # 商品名
                in_stock = h.get("inStock")      # 在庫あり True / False
                price = h.get("price") or ""     # 価格
                jan = h.get("janCode") or ""     # JANコード

                all_rows.append([name, in_stock, price, jan])

            print(f"{i+1}ページ目完了 ✅（累計: {len(all_rows)}件）")

            # Yahoo!APIの制限を避けるため、0.5秒休む
            time.sleep(0.5)

        except Exception as e:
            # ネットワークエラーやYahoo!側の一時的な不具合
            print(f"⚠️ エラー発生: {e}")
            print("5秒後に再試行します...")
            time.sleep(5)

    # ============================================================
    # ⚠️ データが取れなかった場合
    # ============================================================
    if not all_rows:
        print("⚠️ 商品データを取得できませんでした。")
        return

    # ============================================================
    # 💾 Excelファイルとして保存
    # ============================================================
    df = pd.DataFrame(all_rows, columns=["商品名", "在庫あり", "価格", "JANコード"])

    # 保存先を実行中ディレクトリに設定
    output_path = os.path.join(os.getcwd(), f"{SELLER_ID}_商品情報.xlsx")

    try:
        df.to_excel(output_path, index=False)
        available = data.get("totalResultsAvailable", 0)
        print("----------------------------------------------")
        print(f"📊 総ヒット件数: {available:,} 件")
        print(f"💾 保存完了 → {output_path}")
        print("✅ すべての処理が完了しました。")
        print("----------------------------------------------")
    except Exception as e:
        print(f"⚠️ Excelファイル保存中にエラーが発生しました: {e}")

# ============================================================
# 🚀 コマンドライン（またはGUI）から実行した場合
# ============================================================
if __name__ == "__main__":
    # GUIやターミナルから渡された引数を確認
    if len(sys.argv) == 5:
        APP_ID = sys.argv[1]             # App ID
        SELLER_ID = sys.argv[2]          # 出店者ID
        RESULTS_PER_CALL = int(sys.argv[3])  # 1回あたりの件数
        TOTAL_ITEMS = int(sys.argv[4])      # 合計件数

        # 関数を実行
        fetch_yahoo_items(APP_ID, SELLER_ID, RESULTS_PER_CALL, TOTAL_ITEMS)
    else:
        print("使い方: python main.py <APP_ID> <SELLER_ID> <RESULTS_PER_CALL> <TOTAL_ITEMS>")
        print("例: python main.py dj00aiZp〇〇 hands-net 50 500")
