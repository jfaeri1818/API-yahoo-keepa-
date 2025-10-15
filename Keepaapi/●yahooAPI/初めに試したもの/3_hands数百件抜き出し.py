import requests
import pandas as pd
import time

# ============================================================
# Yahoo!ショッピング 商品検索API（V3）
# ============================================================

API_URL = "https://shopping.yahooapis.jp/ShoppingWebService/V3/itemSearch"
APP_ID = "dj00aiZpPXlkOGd5bDlUcTlWRyZzPWNvbnN1bWVyc2VjcmV0Jng9MmE-"  # あなたのアプリケーションID

# ============================================================
# 設定値（1日1万件仕様）
# ============================================================

START_NUMBER = 1             # ★ここを手動で変更（例：翌日は10001）
TOTAL_ITEMS_PER_DAY = 1      # 取得したい数（1万件まで）
SELLER_ID = "hands-net"      # 出店者ID
RESULTS_PER_CALL = 1        # 1回で取得する件数（最大50）
CALLS = TOTAL_ITEMS_PER_DAY // RESULTS_PER_CALL  # 実行回数
WAIT_SEC = 2.0               # 呼び出し間隔（秒）
OUTPUT_FILE = f"{SELLER_ID}_商品情報_{START_NUMBER}.xlsx"

# ============================================================
# 実行開始
# ============================================================

print(f"🔍 Yahoo!ショッピングAPIから商品情報を取得開始（開始位置: {START_NUMBER}）")

all_rows = []

for i in range(CALLS):
    start = START_NUMBER + i * RESULTS_PER_CALL

    params = {
        "appid": APP_ID,
        "seller_id": SELLER_ID,
        "results": RESULTS_PER_CALL,
        "start": start,
        "sort": "+price",
        "condition": "new"
    }

    try:
        response = requests.get(API_URL, params=params, timeout=10)
        data = response.json()
        hits = data.get("hits", [])
        total_available = data.get("totalResultsAvailable", 0)

        if not hits:
            print(f"⚠️ {i+1}回目（start={start}）に商品データなし。")
            break

        for h in hits:
            name = h.get("name") or ""
            in_stock = h.get("inStock")
            price = h.get("price") or ""
            jan = h.get("janCode") or ""
            all_rows.append([name, in_stock, price, jan])

        print(f"✅ {i+1}/{CALLS} ページ完了（start={start}〜{start + RESULTS_PER_CALL - 1}）")

        time.sleep(WAIT_SEC)

    except Exception as e:
        print(f"❌ エラー発生: {e}")
        print("⏳ 30秒待機して再試行します...")
        time.sleep(30)

# ============================================================
# Excel出力
# ============================================================

df = pd.DataFrame(all_rows, columns=["商品名", "在庫あり", "価格", "JANコード"])
df.to_excel(OUTPUT_FILE, index=False)

print(f"\n🎉 1日分（約{len(all_rows)}件）の取得完了。")
print(f"🆕 次回の開始番号: {START_NUMBER + len(all_rows)}")
print(f"📁 保存先: {OUTPUT_FILE}")

# 最後に取得したリクエストから情報を表示
print(f"\n🔎 最後の取得件数: {len(hits)}件")
print(f"📊 条件に一致した全体のヒット件数: {total_available:,} 件")
