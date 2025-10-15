import tkinter as tk
from tkinter import ttk, messagebox
import subprocess

# 以下でそれぞれの値を取得して定義する
def run_main():
    app_id = app_id_entry.get().strip()
    seller_id = seller_id_entry.get().strip()
    results = results_entry.get().strip()
    total = total_entry.get().strip()

    if not all([app_id, seller_id, results, total]):
        messagebox.showerror("入力エラー", "すべての項目を入力してください。")
        return

    cmd = ["python", "main.py", app_id, seller_id, results, total]
    try:
        subprocess.run(cmd)
        messagebox.showinfo("完了", "商品情報の取得が完了しました。")
    except Exception as e:
        messagebox.showerror("実行エラー", str(e))

# 画面出力
root = tk.Tk()
root.title("Yahoo!商品情報取得ツール - Flower Edition")
root.geometry("480x360")
root.resizable(False, False)

# フレームを配置、設定
frame = ttk.Frame(root, padding=20)
frame.pack(fill="both", expand=True)

ttk.Label(frame, text="Yahoo! App ID").grid(row=0, column=0, sticky="e", pady=8)
app_id_entry = ttk.Entry(frame, width=40)# 入力欄
app_id_entry.grid(row=0, column=1)#行0　列2

ttk.Label(frame, text="出店者ID (例: hands-net)").grid(row=1, column=0, sticky="e", pady=8)
seller_id_entry = ttk.Entry(frame, width=40)# 入力欄
seller_id_entry.grid(row=1, column=1)

ttk.Label(frame, text="1回の取得件数 (最大50)").grid(row=2, column=0, sticky="e", pady=8)
results_entry = ttk.Entry(frame, width=40)
results_entry.insert(0, "10")#先頭に10と入れておく
results_entry.grid(row=2, column=1)

ttk.Label(frame, text="合計取得件数").grid(row=3, column=0, sticky="e", pady=8)
total_entry = ttk.Entry(frame, width=40)
total_entry.insert(0, "50")
total_entry.grid(row=3, column=1)

ttk.Button(frame, text="実行 ▶", command=run_main).grid(row=4, column=1, pady=20, sticky="e")

root.mainloop()
