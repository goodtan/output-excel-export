# -*- coding: utf-8 -*-
"""
批量报价计算（输入克重=克 g）
- 批量来源：
  1) 文件：程序目录下 批量输入.xlsx / 批量输入.csv
     必填列：成本 | 克重(g) | 利润比例   可选列：Zone(5/6) | 备注
  2) 交互式：每行输入  成本,克重(g),利润比例[,Zone][,备注]
     比例可写 0.6 / .6 / 60% / 60
- 逻辑：
  * ≤450g 用“克表”第一个 >= g 的价
  * >450g 用“KG 表”区间起点价（不进位；weight==断点 用该断点）
  * 头程 = 区间单价(元/kg) × 实际重量(kg)
  * 尾程(￥) = 尾程USD × USD→CNY 实时汇率
  * 利润 = 成本 × 利润比例（每行必填）
  * 合计 = 成本 + 头程 + 尾程 + 面单(5) + 利润
  * 每行生成中文描述并写入 "描述" 列，同时在控制台打印
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
            # 两个 API 的返回结构可能不同，但都通常含 rates.CNY
            if "rates" in r and "CNY" in r["rates"]:
                rate = float(r["rates"]["CNY"])
                if rate > 0: return rate
            if "rates" in r and isinstance(r["rates"], dict):
                rate = float(r["rates"].get("CNY", 0))
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
def get_head_rate_per_kg(weight_kg: float) -> float:
    """返回用于计算头程的 区间单价(元/kg)"""
    for lo, hi, rate in HEAD_RATE_TABLE:
        if lo <= weight_kg <= hi:
            return rate
    return HEAD_RATE_TABLE[-1][2]

def head_charge(weight_kg: float):
    """返回 (头程金额(四舍五入到两位), 区间单价(元/kg))"""
    rate = get_head_rate_per_kg(weight_kg)
    amt = round(weight_kg * rate, 2)
    return amt, rate

def match_tail_usd_from_grams(grams: int, zone: str) -> float:
    """克表优先；超出450g转kg后按KG表“区间起点价”（不进位；等于断点取该断点价）"""
    is_z5 = str(zone).strip() in ("5", "Zone-5", "Zone5", "Z5")
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

def compute_row(cost_rmb: float, grams: int, zone: str, ratio: float, rate_usd_cny: float, remark: str = ""):
    weight_kg = grams / 1000.0
    head_amt, head_rate = head_charge(weight_kg)
    tail_rmb = round(match_tail_usd_from_grams(grams, zone) * rate_usd_cny, 2)
    profit = round(cost_rmb * ratio, 2)
    total = round(cost_rmb + head_amt + tail_rmb + FACE_SHEET_RMB + profit, 2)

    # 生成中文描述，数值均以两位小数显示（克重用整数 g）
    desc_parts = []
    desc_parts.append(f"成衣克重 {int(grams)}g")
    desc_parts.append(f"成本 {cost_rmb:.2f}")
    desc_parts.append(f"头程 {head_amt:.2f} ({int(head_rate)} 元/kg)")
    desc_parts.append(f"面单处理费 {FACE_SHEET_RMB:.2f}")
    desc_parts.append(f"尾程 {tail_rmb:.2f}")
    desc_parts.append(f"利润 {profit:.2f}")
    desc = "，".join(desc_parts) + f"，合计 {total:.2f}。"
    if remark:
        desc += f" 备注：{remark}"

    return {
        "成本": round(cost_rmb, 2),
        "克重(g)": int(grams),
        "克重(kg)": round(weight_kg, 3),
        "Zone": str(zone),
        "头程": head_amt,
        "头程单价(元/kg)": head_rate,
        "尾程": tail_rmb,
        "利润比例": ratio,
        "利润": profit,
        "面单": round(FACE_SHEET_RMB, 2),
        "合计": total,
        "备注": remark or "",
        "描述": desc
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
        # Zone 与 备注 为可选列
        for _, r in df.iterrows():
            cost = float(r["成本"])
            grams = int(round(float(r["克重(g)"])))
            ratio = parse_ratio(r["利润比例"])
            zone = str(r.get("Zone", default_zone)).strip() or default_zone
            remark = str(r.get("备注", "") if "备注" in r.index else "").strip()
            rows.append((cost, grams, ratio, zone, remark))
        print(f"已从文件读取 {len(rows)} 条记录。")
        return rows

    print("交互式：每行输入 '成本,克重(g),利润比例[,Zone][,备注]'；示例： 71,520,60% 或 71,520,0.6,5 或 71,520,60%,5,客户A备注")
    while True:
        s = input("> ").strip()
        if not s:
            break
        try:
            parts = [x.strip() for x in s.replace("，", ",").split(",")]
            cost = float(parts[0])
            grams = int(round(float(parts[1])))
            ratio = parse_ratio(parts[2])

            zone = default_zone
            remark = ""
            if len(parts) >= 4 and parts[3]:
                # 如果第4项看起来是 zone (5/6 或 Zone5/Zone-5 等) 就当 zone，否则当备注
                p4 = parts[3]
                if p4 in ("5", "6") or p4.lower().startswith("zone") or p4.upper().startswith("Z"):
                    zone = p4
                    if len(parts) >= 5:
                        remark = parts[4]
                else:
                    # 第4项不是 zone，就当备注
                    remark = p4
                    # 若还有第5项，且第5项看起来像 zone，则把它作为 zone（容错）
                    if len(parts) >= 5 and parts[4] in ("5", "6"):
                        zone = parts[4]
            rows.append((cost, grams, ratio, zone, remark))
        except Exception:
            print("格式不对，请输入：成本,克重(g),利润比例[,Zone][,备注]")
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

    out_rows = [compute_row(cost, grams, zone, ratio, rate, remark) for (cost, grams, ratio, zone, remark) in items]
    # 指定列顺序，包含 '描述'
    cols = ["成本","克重(g)","克重(kg)","Zone","头程单价(元/kg)","头程","尾程","利润比例","利润","面单","合计","备注","描述"]
    df = pd.DataFrame(out_rows, columns=cols)
    write_output(df)

    # 控制台逐行打印描述（便于查看）
    print("\n逐行描述：")
    for i, r in df.iterrows():
        print(f"{i+1}. {r['描述']}")

    print("\n汇总表：")
    print(df.drop(columns=["描述"]))  # 控制台表格不重复显示描述列

if __name__ == "__main__":
    main()
