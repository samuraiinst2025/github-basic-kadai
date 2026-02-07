# ==============================
# ライブラリ読み込み
# ==============================

import json
import fnmatch
from pathlib import Path

import pandas as pd

# display を残す（IPythonがあればdisplay、無ければprintにフォールバック）
try:
    from IPython.display import display
except ImportError:
    def display(x):
        print(x)

# Google Drive API 部品（project/src/drive_api.py）
from src.drive_api import download_file, list_files_in_folder


# ==============================
# Config 読み込み
# ==============================

PROJECT_DIR = Path(__file__).resolve().parent
CONFIG_PATH = PROJECT_DIR / "config" / "drive_config.json"

with open(CONFIG_PATH, "r", encoding="utf-8") as f:
    cfg = json.load(f)

ORDER_FOLDER_ID = cfg["order_folder_id"]
INVENTORY_FILE_ID = cfg["inventory_file_id"]
PICKUP_FILE_ID = cfg["pickup_file_id"]

# Driveから落とす一時フォルダ（必要なら gitignore 推奨）
TMP_DIR = PROJECT_DIR / "tmp"
TMP_DIR.mkdir(parents=True, exist_ok=True)


# ==============================
# ① 各店の注文情報をまとめる
# ==============================

# 今日の日付（ファイル名に使われている形式）
today = "20230524"

# Drive上の注文フォルダから一覧取得
drive_files = list_files_in_folder(ORDER_FOLDER_ID)

# 今日の日付の注文ファイルだけ取得（Drive上で名前マッチ）
pattern = f"order_*_{today}.xlsx"
order_files_meta = [f for f in drive_files if fnmatch.fnmatch(f.get("name", ""), pattern)]

print(f"対象ファイル数: {len(order_files_meta)}")

# 合計注文数を入れるDataFrame
total_df = None

# DriveからローカルへDLしてから読み込む
for meta in order_files_meta:
    file_id = meta["id"]
    name = meta["name"]

    local_path = TMP_DIR / name
    local_path = download_file(file_id, local_path)  # 戻り値は実際に保存されたPath

    df = pd.read_excel(local_path)

    if total_df is None:
        total_df = df.copy()
    else:
        total_df = total_df.add(df, fill_value=0)

# 注文ファイルが1件もなかった場合のエラー対策
if total_df is None:
    raise ValueError("本日の注文ファイルが見つかりませんでした")

print("▼ 各野菜の合計注文数")
display(total_df)


# ==============================
# ② 注文集計結果を保存（ローカル保存）
# ==============================

output_path = TMP_DIR / f"summary_order_{today}.xlsx"
total_df.to_excel(output_path, index=False)

print(f"集計結果を保存しました: {output_path}")


# ==============================
# ③ 現在の在庫状況を確認
# ==============================

# Driveから inventory をDL（Googleスプレッドシートでもxlsx化して落ちる）
inventory_local_path = download_file(INVENTORY_FILE_ID, TMP_DIR / "inventory.xlsx")

# 在庫表を読み込む
inventory_df = pd.read_excel(inventory_local_path)

# 最終行（最新在庫）
latest_inventory = inventory_df.iloc[-1]

print("▼ 最新の在庫情報（元データ）")
display(latest_inventory)

# 数値列（野菜）だけ抽出（列が存在する時だけ落とす）
drop_cols = [c for c in ["日付", "曜日"] if c in latest_inventory.index]
latest_inventory_no_date = latest_inventory.drop(drop_cols)


# ==============================
# ④ 注文反映後の在庫を計算
# ==============================

# 各野菜の合計注文数（Series化）
# numeric_only=True を付けておくと、日付など混ざってても事故りにくい
total_order_series = total_df.sum(numeric_only=True)

# 残在庫 = 最新在庫 - 注文数（インデックスを揃えてズレ防止）
total_order_series = total_order_series.reindex(latest_inventory_no_date.index).fillna(0)
remaining_inventory = latest_inventory_no_date - total_order_series

print("▼ 注文反映後の在庫数")
display(remaining_inventory)


# ==============================
# ⑤ pickup.xlsx からしきい値・発注数を取得
# ==============================

# Driveから pickup をDL
pickup_local_path = download_file(PICKUP_FILE_ID, TMP_DIR / "pickup.xlsx")

# 1列目をindexとして読み込む
pickup_df = pd.read_excel(pickup_local_path, index_col=0)

print("▼ pickup.xlsx")
display(pickup_df)

# しきい値
threshold_series = pickup_df.loc["しきい値"]

# 発注数（追加量）
order_qty_series = pickup_df.loc["追加量"]


# ==============================
# ⑥ 発注が必要な野菜を特定
# ==============================

# 残在庫がしきい値を下回っているか（インデックスを揃える）
threshold_series = threshold_series.reindex(remaining_inventory.index)
below_threshold = remaining_inventory < threshold_series

print("▼ しきい値を下回っているか")
display(below_threshold)

# 発注対象の野菜
low_stock = remaining_inventory[below_threshold]

print("▼ 発注が必要な野菜（残在庫）")
display(low_stock)


# ==============================
# ⑦ 発注対象と発注数を紐づける
# ==============================

# 発注が必要な野菜名
order_items = low_stock.index

# 発注数を取得（ズレ防止で reindex）
order_list = order_qty_series.reindex(order_items)

print("▼ 発注対象と発注数")
display(order_list)


# ==============================
# ⑧ メール用 発注内容テキスト作成
# ==============================

order_lines = []

for veg, qty in order_list.items():
    if pd.isna(qty):
        continue
    line = f"・{veg}：{int(qty)} 個"
    order_lines.append(line)

order_text = "\n".join(order_lines)

print("▼ メール用 発注内容")
print(order_text)


# ==============================
# ⑨ 発注メール本文を生成
# ==============================

mail_body = f"""
〇〇農園 御中

いつもお世話になっております。
株式会社△△の□□です。

下記内容にて野菜の発注をお願いいたします。

【発注内容】
{order_text}

納品日：明日

ご不明点等ございましたらご連絡ください。
何卒よろしくお願いいたします。

――――――――――――
株式会社△△
□□
"""

print("▼ 発注メール本文")
print(mail_body)
