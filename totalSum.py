# -*- coding: utf-8 -*-
"""
批量报价计算
- 支持两种批量方式：
  1) 交互式：多行输入“成本,克重(kg)”
  2) 文件：自动读取同目录下 批量输入.xlsx / 批量输入.csv
     需要列：成本, 克重(kg)；可选：Zone(5/6), 利润比例(如 0.6)
- 逻辑：
  * ≤450g 用克表第一个 >= g 的价
  * >450g 用KG表区间“起点断点”的价（不进位）
  * 头程 = 区间单价/公斤 × 实际重量(kg)
  * 尾程(￥) = 尾程USD × USD→CNY 实时汇率
  * 合计 = 成本 + 头程 + 尾程 + 面单(5) + 利润
"""

from pathlib import Path
import sys
import pandas as pd
import requests

# ===== 头程区间单价 (RMB/kg；包含上界) =====
HEAD_RATE_TABLE = [
    (0.000, 0.100, 80),
    (0.101, 0.200, 80),
    (0.201, 0.450, 85),
    (0.451, 0.700, 85),
    (0.701, 1.500, 90),
    (1.501, 2.000, 90),
    (2.001, 30.000, 90),
]

FACE_SHEET_RMB = 5.0
OUTPUT_FILE = "报价计算_批量.xlsx"

# ===== 克表（≤450g）：g上限 -> Zone5 USD, Zone6 USD =====
GRAM_PRICE_TABLE = [
    (28,  4.23, 4.35),
    (56,  4.23, 4.35),
    (85,  4.23, 4.35),
    (113, 4.23, 4.35),
    (141, 4.23, 4.35),
    (170, 4.23, 4.35),
    (198, 4.23, 4.35),
    (226, 4.23, 4.35),
    (255, 4.23, 4.35),
    (283, 5.38, 5.56),
    (311, 5.38, 5.87),
    (340, 5.38, 5.87),
    (368, 5.87, 6.07),
    (396, 6.21, 6.41),
    (425, 6.36, 6.43),
    (450, 6.80, 7.04),
]

# ===== KG表（>0.451kg，按区间起点，不进位）=====
KG_BREAKS = [0.45, 0.90, 1.36, 1.81, 2.26, 2.72, 3.17, 3.62, 4.08, 4.53, 4.98, 5.44, 5.89, 6.35, 6.80]
Z5_USD    = [7.79, 9.00, 10.03, 10.71, 11.34, 12.11, 12.78, 13.45, 14.21, 14.90, 15.79, 16.35, 16.95, 17.62, 18.52]
Z6_USD    = [8.75, 10.50, 12.16, 12.78, 13.57, 14.35, 15.12, 15.92, 16.71, 17.40, 18.21, 18.84, 19.47, 20.11, 21.03]

def usd_cny():
    for url in ("https://api.exchangerate.host/latest?base=USD&symbols=CNY",
                "https://open.er-api.com/v6/latest/USD"):
        try:
            r = requests.get(url, timeout=6).json()
            rate = float(r["rates"]["CNY"])
            if rate > 0: return rate
        except Exception:
            pass
    s = input("汇率获取失败，手动输入 USD→CNY（回车用 7.20）：").strip()
    return 7.20 if s == "" else float(s)

def head_charge(weight_kg: float) -> float:
    for lo, hi, rate in HEAD_RATE_TABLE:
        if lo <= weight_kg <= hi:
            return round(weight_kg * rate, 2)
    return round(weight_kg * HEAD_RATE_TABLE[-1][2], 2)

def match_tail_usd(weight_kg: float, zone: str) -> float:
    """≤450g 用克表第一个≥g；>450g 用KG表区间起点价（不进位）"""
    is_z5 = str(zone).strip() in ("5", "Zone-5")
    grams = round(weight_kg * 1000)

    for g_upper, z5, z6 in GRAM_PRICE_TABLE:
        if grams <= g_upper:
            return z5 if is_z5 else z6

    # 在 KG_BREAKS 里找 weight_kg 所在区间的起点
    # 遍历到第一个 > weight_kg 的断点，用它的前一个断点价；若都≤，用最后一档
    idx = 0
    for i, br in enumerate(KG_BREAKS):
        if weight_kg < br:
            idx = max(0, i - 1)
            break
        idx = i  # 如果始终不小于，就用最后一个断点
    return (Z5_USD[idx] if is_z5 else Z6_USD[idx])

def compute_row(cost_rmb: float, weight_kg: float, zone: str, profit_ratio: float, rate: float):
    head = head_charge(weight_kg)
    tail_rmb = round(match_tail_usd(weight_kg, zone) * rate, 2)
    profit = round(cost_rmb * profit_ratio, 2)
    total = round(cost_rmb + head + tail_rmb + FACE_SHEET_RMB + profit, 2)
    return {
        "成本": round(cost_rmb, 2),
        "克重(kg)": round(weight_kg, 3),
        "头程": head,
        "尾程": tail_rmb,
        "利润": profit,
        "面单": round(FACE_SHEET_RMB, 2),
        "合计": total
    }

def read_batch_inputs(default_zone: str, default_ratio: float):
    """
    若存在 批量输入.xlsx/.csv 则优先读取；否则进入交互式多行输入模式。
    文件模式：需要列 '成本','克重(kg)'; 可选 'Zone','利润比例'
    """
    xlsx = Path("批量输入.xlsx")
    csv  = Path("批量输入.csv")
    rows = []

    if xlsx.exists() or csv.exists():
        if xlsx.exists():
            df = pd.read_excel(xlsx)
        else:
            df = pd.read_csv(csv)
        for _, r in df.iterrows():
            cost = float(r["成本"])
            wkg  = float(r["克重(kg)"])
            zone = str(r.get("Zone", default_zone)).strip() or default_zone
            ratio = float(r.get("利润比例", default_ratio))
            rows.append((cost, wkg, zone, ratio))
        print(f"已读取 {len(rows)} 条记录。")
        return rows

    print("交互式输入模式：逐行输入 '成本,克重(kg)'；直接回车结束。示例： 71,0.52")
    while True:
        s = input("> ").strip()
        if not s:
            break
        try:
            parts = [x.strip() for x in s.replace("，", ",").split(",")]
            cost = float(parts[0]); wkg = float(parts[1])
            rows.append((cost, wkg, default_zone, default_ratio))
        except Exception:
            print("格式不对，请输入：成本,克重(kg)  例如：71,0.52")
    return rows

def main():
    # 全局设置：默认 Zone 与 利润比例
    default_zone = (input("默认尾程区域 5 或 6（默认5）：").strip() or "5")
    ratio_in = input("默认利润比例(默认0.6)：").strip()
    default_ratio = 0.6 if ratio_in == "" else float(ratio_in)

    rate = usd_cny()
    print(f"USD→CNY：{rate:.4f}")

    rows = read_batch_inputs(default_zone, default_ratio)
    if not rows:
        print("没有任何输入，已退出。"); sys.exit(0)

    out = []
    for cost, wkg, zone, ratio in rows:
        out.append(compute_row(cost, wkg, zone, ratio, rate))

    df = pd.DataFrame(out, columns=["成本","克重(kg)","头程","尾程","利润","面单","合计"])
    df.to_excel(OUTPUT_FILE, index=False)
    print(f"已生成：{Path(OUTPUT_FILE).resolve()}")
    print(df)

if __name__ == "__main__":
    main()