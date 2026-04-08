from env import ENV
from 儲值金系統設定 import run_process


def parse_row_input(row_text: str):
    rows = set()
    parts = [p.strip() for p in row_text.split(",") if p.strip()]

    for part in parts:
        if "-" in part:
            start_str, end_str = part.split("-", 1)
            start = int(start_str.strip())
            end = int(end_str.strip())

            if start > end:
                raise ValueError(f"區間錯誤：{part}")

            rows.update(range(start, end + 1))
        else:
            rows.add(int(part))

    return sorted(rows)


print(f"目前環境：{ENV}")

sheet_name = input("請輸入工作表名稱（例如 202604）：").strip()
row_input = input("請輸入列號（例如 2,3,5-7）：").strip()

try:
    target_rows = parse_row_input(row_input)
except Exception as e:
    print(f"❌ 列號格式錯誤：{e}")
    exit()

print(f"👉 將執行列：{target_rows}")
print("====================================")

total_success = 0
total_fail = 0

for row_no in target_rows:
    print(f"\n🚀 開始執行第 {row_no} 列...")

    try:
        result = run_process(sheet_name, row_no, row_no, ENV)

        # 如果你的 run_process 有回傳統計
        if isinstance(result, dict):
            success = result.get("success_count", 0)
            fail = result.get("fail_count", 0)

            total_success += success
            total_fail += fail

            print(f"✅ 第 {row_no} 列完成：成功 {success} / 失敗 {fail}")
        else:
            print(f"✅ 第 {row_no} 列完成")

    except Exception as e:
        total_fail += 1
        print(f"❌ 第 {row_no} 列失敗：{e}")

print("\n====================================")
print("🎯 執行完成")

print(f"✅ 成功總數：{total_success}")
print(f"❌ 失敗總數：{total_fail}")
print(f"📊 總執行列數：{len(target_rows)}")
