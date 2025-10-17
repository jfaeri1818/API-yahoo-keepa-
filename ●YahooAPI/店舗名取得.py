import requests
import time

API_URL = "https://shopping.yahooapis.jp/ShoppingWebService/V3/itemSearch"
APP_ID = "dj00aiZpPXlkOGd5bDlUcTlWRyZzPWNvbnN1bWVyc2VjcmV0Jng9MmE-"

store_ids = set()

for i in range(20):  # 試しに10ページ分
    start = i * 50 + 50  # ページ先頭位置
    params = {
        "appid": APP_ID,
        "query": "全商品",  # キーワードは必須
        "results": 50,             # 最大50件
        "start": start,
        "sort": "-score",          # 人気順
        "availability": "1",       # 文字列で指定（整数だと400エラーになる場合あり）
        "condition": "new"         # これも "new" / "used" の文字列指定
    }

    try:
        response = requests.get(API_URL, params=params, timeout=10)
        response.raise_for_status()
        data = response.json()

        hits = data.get("hits", [])
        if not hits:
            print(f"{i+1}ページ目にデータがありません。終了します。")
            break

        for h in hits:
            seller = h.get("seller", {})
            store = seller.get("sellerId")
            if store:
                store_ids.add(store)

        print(f"{i+1}ページ目完了。現在の店舗数: {len(store_ids)}")

        time.sleep(0.5)

    except requests.exceptions.RequestException as e:
        print(f"{i+1}ページ目でリクエストエラー: {e}")
        print(f"レスポンス: {response.text if 'response' in locals() else 'なし'}")
        break
    except Exception as e:
        print(f"予期せぬエラー: {e}")
        break

print(f"\n最終的に取得した店舗数: {len(store_ids)}")
print(store_ids)


# dj00aiZpPXlkOGd5bDlUcTlWRyZzPWNvbnN1bWVyc2VjcmV0Jng9MmE-
# https://shopping.yahooapis.jp/ShoppingWebService/V3/itemSearch