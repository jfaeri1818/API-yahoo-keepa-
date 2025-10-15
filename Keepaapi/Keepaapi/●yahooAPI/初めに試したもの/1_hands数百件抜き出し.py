import requests                              # HTTP通信（API呼び出し）に使うライブラリ
import pandas as pd                           # データを表形式で扱うためのライブラリ
import time                                   # 処理を一時停止するために使用

API_URL = "https://shopping.yahooapis.jp/ShoppingWebService/V3/itemSearch"  # APIのエンドポイントURL
APP_ID = "dj00aiZpPXlkOGd5bDlUcTlWRyZzPWNvbnN1bWVyc2VjcmV0Jng9MmE-"        # 自分のアプリケーションID（Yahoo!から発行）

# ============================================================
# 検索設定
# ============================================================

SELLER_ID = "hands-net"         # 出店者ID（東急ハンズYahoo!店）
RESULTS_PER_CALL = 1            # 1回のリクエストで取得する件数（最大50まで指定可能）
TOTAL_ITEMS = 1                 # 合計で取得したい商品件数
CALLS = TOTAL_ITEMS // RESULTS_PER_CALL   # 必要なリクエスト回数を計算

all_rows = []                   # 取得した商品データを入れるリスト

print("Yahoo!ショッピングAPIから商品情報を取得中...")  # 進行状況の表示

# ============================================================
# データ取得ループ（ページごとに商品データを取得）
# ============================================================

for i in range(CALLS):                                           # CALLSの回数分だけ繰り返す
    start = i * RESULTS_PER_CALL + 1                             # ページの開始位置を計算（1ページ目＝1）

    # APIに送るリクエストパラメータ
    params = {
        "appid": APP_ID,                                         # アプリケーションID
        "seller_id": SELLER_ID,                                  # 店舗ID
        "results": RESULTS_PER_CALL,                             # 取得件数
        "start": start,                                          # 取得開始位置
        "sort": "+price",                                        # 価格の安い順（人気順ではない）
        "condition": "new"                                       # 新品のみ取得
    }

    try:
        response = requests.get(API_URL, params=params, timeout=10)   # APIにリクエストを送信（10秒でタイムアウト）
        data = response.json()                                         # JSON形式のデータをPythonの辞書に変換
        hits = data.get("hits", [])                                   # 商品リストを取得（無ければ空リスト）

        if not hits:                                                  # 商品が取得できなかった場合
            print(f"{i+1}ページ目に商品なし。")                      # メッセージ表示
            break                                                     # ループを抜ける

        for h in hits:                                                # 各商品データを処理
            name = h.get("name") or ""                                # 商品名
            in_stock = h.get("inStock")                               # 在庫あり（True/False）
            price = h.get("price") or ""                              # 価格
            jan = h.get("janCode") or ""                              # JANコード
            all_rows.append([name, in_stock, price, jan])             # リストに1行として追加

        print(f"{i+1}ページ目完了（累計: {len(all_rows)}件）")      # 処理進行を表示

        time.sleep(0.5)                                               # Yahoo!側への負荷軽減のため0.5秒休止

    except Exception as e:                                            # エラーが発生した場合
        print(f"エラー: {e}")                                        # エラー内容を表示
        time.sleep(5)                                                 # 少し待って再試行できるように5秒停止

# ============================================================
# 取得データをExcelファイルに出力
# ============================================================

df = pd.DataFrame(all_rows, columns=["商品名", "在庫あり", "価格", "JANコード"])  # pandasで表形式データに変換
output_path = "handsnet_商品情報.xlsx"                                         # 出力ファイル名
df.to_excel(output_path, index=False)                                          # インデックスなしでExcel出力

# 検索条件に合致した全商品件数を取得
available = data.get("totalResultsAvailable")                                  # API応答から総ヒット件数を取得
print(f"全体のヒット件数: {available:,} 件")                                  # カンマ区切りで表示
print("処理が完了しました。Excelファイルを作成しました。")                    # 完了メッセージ
print(f"保存先: {output_path}")                                               # 出力ファイルのパスを表示
