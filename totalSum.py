# -*- coding: utf-8 -*-
"""
批量报价计算（输入克重=克 g）
- 批量来源：
  1) 文件：程序目录下 批量输入.xlsx / 批量输入.csv
     必填列：成本 | 克重(g) | 利润比例   可选列：Zone(5/6)
  2) 交互式：每行输入  成本,克重(g),利润比例[,Zone]
     比例可写 0.6 / .6 / 60% / 60
- 逻辑：
  * ≤450g 用“克表”第一个 >= g 的价
  * >450g 用“KG 表”区间起点价（不进位；weight==断点 用该断点）
  * 头程 = 区间单价(元/kg) × 实际重量(kg)
  * 尾程(￥) = 尾程USD × USD→CNY 实时汇率
  * 利润 = 成本 × 利润比例（每行必填）
  * 合计 = 成本 + 头程 + 尾程 + 面单(5) + 利润
"""
from pathlib import Path
from datetime import datetime
import sys
import pandas as pd
import requests

# —— 程序目录（exe 友好）——
def base_dir() -> Path:
    return Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).parent

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
OUTPUT_NAME = "报价计算_批量.xlsx"

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

# —— 汇率 ——
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

# —— 工具：解析比例（0.6 / .6 / 60% / 60）——
def parse_ratio(x) -> float:
    if isinstance(x, str):
        t = x.strip().replace("％", "%")
        if t.endswith("%"):
            return float(t[:-1]) / 100.0
        v = float(t)
    else:
        v = float(x)
    return v / 100.0 if v > 1 else v

# —— 头程 / 尾程 ——
def head_charge(weight_kg: float) -> float:
    for lo, hi, rate in HEAD_RATE_TABLE:
        if lo <= weight_kg <= hi:
            return round(weight_kg * rate, 2)
    return round(weight_kg * HEAD_RATE_TABLE[-1][2], 2)

def match_tail_usd_from_grams(grams: int, zone: str) -> float:
    """克表优先；超出450g转kg后按KG表“区间起点价”（不进位；等于断点取该断点价）"""
    is_z5 = str(zone).strip() in ("5", "Zone-5")
    # 1) 克表（第一个 >= g）
    for g_upper, z5, z6 in GRAM_PRICE_TABLE:
        if grams <= g_upper:
            return z5 if is_z5 else z6
    # 2) KG 表
    weight_kg = grams / 1000.0
    idx = 0
    for i, br in enumerate(KG_BREAKS):
        if weight_kg >= br:
            idx = i      # 等于断点走该断点
        else:
            break
    return (Z5_USD[idx] if is_z5 else Z6_USD[idx])

def compute_row(cost_rmb: float, grams: int, zone: str, ratio: float, rate_usd_cny: float):
    weight_kg = grams / 1000.0
    head = head_charge(weight_kg)
    tail_rmb = round(match_tail_usd_from_grams(grams, zone) * rate_usd_cny, 2)
    profit = round(cost_rmb * ratio, 2)
    total = round(cost_rmb + head + tail_rmb + FACE_SHEET_RMB + profit, 2)
    return {
        "成本": round(cost_rmb, 2),
        "克重(g)": int(grams),
        "克重(kg)": round(weight_kg, 3),
        "Zone": str(zone),
        "头程": head,
        "尾程": tail_rmb,
        "利润比例": ratio,
        "利润": profit,
        "面单": round(FACE_SHEET_RMB, 2),
        "合计": total
    }

# —— 读取批量输入（优先文件）——
def read_batch_inputs(default_zone: str):
    rows = []
    p_xlsx = base_dir() / "批量输入.xlsx"
    p_csv  = base_dir() / "批量输入.csv"

    if p_xlsx.exists() or p_csv.exists():
        df = pd.read_excel(p_xlsx) if p_xlsx.exists() else pd.read_csv(p_csv, encoding="utf-8-sig")
        need = {"成本","克重(g)","利润比例"}
        if not need.issubset(set(df.columns)):
            raise SystemExit("批量文件缺少必要列：成本、克重(g)、利润比例")
        for _, r in df.iterrows():
            cost = float(r["成本"])
            grams = int(round(float(r["克重(g)"])))
            ratio = parse_ratio(r["利润比例"])
            zone = str(r.get("Zone", default_zone)).strip() or default_zone
            rows.append((cost, grams, ratio, zone))
        print(f"已从文件读取 {len(rows)} 条记录。")
        return rows

    print("交互式：每行输入 '成本,克重(g),利润比例[,Zone]'；示例： 71,520,60% 或 71,520,0.6,5")
    while True:
        s = input("> ").strip()
        if not s:
            break
        try:
            parts = [x.strip() for x in s.replace("，", ",").split(",")]
            cost = float(parts[0]); grams = int(round(float(parts[1]))); ratio = parse_ratio(parts[2])
            zone = parts[3] if len(parts) >= 4 and parts[3] else default_zone
            rows.append((cost, grams, ratio, zone))
        except Exception:
            print("格式不对，请输入：成本,克重(g),利润比例[,Zone]")
    return rows

# —— 写 Excel（若占用则改名）——
def write_output(df: pd.DataFrame):
    out = base_dir() / OUTPUT_NAME
    try:
        df.to_excel(out, index=False)
        print(f"已生成：{out}")
    except PermissionError:
        alt = base_dir() / f"报价计算_批量_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
        df.to_excel(alt, index=False)
        print(f"文件被占用，已改名保存：{alt}")

def main():
    default_zone = input("默认尾程区域 5 或 6（默认5）：").strip() or "5"
    rate = usd_cny()
    print(f"USD→CNY：{rate:.4f}")

    items = read_batch_inputs(default_zone)
    if not items:
        print("没有任何输入，已退出。"); return

    out_rows = [compute_row(cost, grams, zone, ratio, rate) for (cost, grams, ratio, zone) in items]
    df = pd.DataFrame(out_rows, columns=["成本","克重(g)","克重(kg)","Zone","头程","尾程","利润比例","利润","面单","合计"])
    write_output(df)
    print(df)

if __name__ == "__main__":
    main()
