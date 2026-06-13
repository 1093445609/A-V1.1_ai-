import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import os
from datetime import datetime

# ==========================
# 配置
# ==========================
EXCEL_FILE = "AI_Trade_Record.xlsx"
SHEET_NAME = "阶段识别记录"


# ==========================
# 计算评分与阶段
# ==========================
def analyze_stock(data):
    score = 0
    risk_text = []
    advice = ""
    stage = ""

    close = data["close"]
    ma5 = data["ma5"]
    ma10 = data["ma10"]
    ma20 = data["ma20"]
    high20 = data["high20"]
    volume = data["volume"]
    avg_volume = data["avg_volume"]

    # 自动计算
    bias = (close - ma20) / ma20 * 100
    volume_ratio = volume / avg_volume if avg_volume > 0 else 0
    high_ratio = close / high20 * 100 if high20 > 0 else 0

    # =====================
    # 趋势评分（30）
    # =====================
    trend_score = 0
    if ma5 > ma10 > ma20:
        trend_score = 30
    elif ma10 > ma20:
        trend_score = 20
    elif close > ma20:
        trend_score = 10

    # =====================
    # 结构评分（30）
    # =====================
    structure_score = 0
    if high_ratio >= 98:
        structure_score = 25
    elif high_ratio >= 92:
        structure_score = 18
    else:
        structure_score = 10

    # =====================
    # 量能评分（20）
    # =====================
    volume_score = 0
    if 1 <= volume_ratio <= 2:
        volume_score = 20
    elif 0.8 <= volume_ratio < 1 or 2 < volume_ratio <= 2.5:
        volume_score = 15
    else:
        volume_score = 8

    # =====================
    # 风险评分（20）
    # =====================
    risk_score = 20

    if bias > 18:
        risk_score -= 10
        risk_text.append("股价远离20日均线，存在追高风险。")

    if volume_ratio > 2.5:
        risk_score -= 5
        risk_text.append("成交量异常放大，可能接近情绪高潮。")

    if close < ma20:
        risk_score -= 5
        risk_text.append("股价跌破20日均线，趋势转弱。")

    risk_score = max(risk_score, 0)

    # =====================
    # 阶段判断（核心）
    # =====================
    if (
        ma5 > ma10 > ma20
        and close > ma20
        and bias <= 8
        and 1 <= volume_ratio <= 2
    ):
        stage = "起涨阶段"
        advice = "【可买】符合右侧交易模型，可列入重点观察。"

    elif (
        ma5 > ma10 > ma20
        and close > ma20
        and 8 < bias <= 18
    ):
        stage = "加速阶段"
        advice = "【持有】已有持仓可继续持有，不建议追高。"

    elif bias > 18:
        stage = "末端阶段"
        advice = "【风险区】禁止追高，等待回调。"

    else:
        stage = "震荡整理"
        advice = "【观望】暂不符合右侧交易条件。"

    total_score = trend_score + structure_score + volume_score + risk_score

    return {
        "bias": round(bias, 2),
        "volume_ratio": round(volume_ratio, 2),
        "high_ratio": round(high_ratio, 2),
        "stage": stage,
        "advice": advice,
        "trend_score": trend_score,
        "structure_score": structure_score,
        "volume_score": volume_score,
        "risk_score": risk_score,
        "total_score": total_score,
        "risk_text": risk_text
    }


# ==========================
# 保存Excel
# ==========================
def save_to_excel(record):
    today = record["分析日期"]

    if os.path.exists(EXCEL_FILE):
        df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
    else:
        df = pd.DataFrame(columns=[
            "分析日期", "股票代码", "股票名称",
            "收盘价", "MA5", "MA10", "MA20",
            "20日最高价", "今日成交量", "5日均量",
            "乖离率(%)", "量比", "阶段",
            "综合评分", "交易建议",
            "计划买入", "备注"
        ])

    exist = df[
        (df["分析日期"] == today) &
        (df["股票代码"].astype(str) == str(record["股票代码"]))
    ]

    if not exist.empty:
        print("\n发现今天已经分析过该股票！")
        print("1：覆盖记录")
        print("2：新增记录")
        print("3：取消保存")

        choice = input("请选择(1/2/3)：")

        if choice == "1":
            df = df.drop(exist.index)
        elif choice == "3":
            print("已取消保存。")
            return

    df = pd.concat([df, pd.DataFrame([record])], ignore_index=True)

    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name=SHEET_NAME, index=False)

# ==========================
# Excel美化
# ==========================
wb = load_workbook(EXCEL_FILE)
ws = wb[SHEET_NAME]

# 开启筛选
ws.auto_filter.ref = ws.dimensions

# 冻结首行
ws.freeze_panes = "A2"

# 自动列宽
for col in ws.columns:
    max_length = 0
    column = col[0].column_letter

    for cell in col:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(str(cell.value))
        except:
            pass

    ws.column_dimensions[column].width = max_length + 4

# 表头加粗
for cell in ws[1]:
    cell.font = Font(bold=True)

wb.save(EXCEL_FILE)

print(f"\n✅ 已保存到 {EXCEL_FILE}")


# ==========================
# 主程序
# ==========================
def main():
    print("=" * 50)
    print("      A股AI右侧交易助手 V1.0")
    print("         阶段识别系统")
    print("=" * 50)

    code = input("请输入股票代码：")
    name = input("请输入股票名称：")

    close = float(input("请输入今日收盘价："))
    ma5 = float(input("请输入MA5："))
    ma10 = float(input("请输入MA10："))
    ma20 = float(input("请输入MA20："))

    high20 = float(input("请输入最近20日最高价："))
    low20 = float(input("请输入最近20日最低价（仅记录，不参与计算）："))

    volume = float(input("请输入今日成交量（万手）："))
    avg_volume = float(input("请输入最近5日平均成交量（万手）："))

    data = {
        "close": close,
        "ma5": ma5,
        "ma10": ma10,
        "ma20": ma20,
        "high20": high20,
        "low20": low20,
        "volume": volume,
        "avg_volume": avg_volume
    }

    result = analyze_stock(data)

    print("\n" + "=" * 50)
    print(f"股票：{code} {name}")
    print("=" * 50)
    print(f"乖离率：{result['bias']} %")
    print(f"量比：{result['volume_ratio']}")
    print(f"距离20日高点：{result['high_ratio']:.2f} %")
    print("-" * 50)
    print(f"【阶段识别】{result['stage']}")
    print(f"【交易建议】{result['advice']}")
    print("-" * 50)

    if result["risk_text"]:
        print("【风险提示】")
        for i, txt in enumerate(result["risk_text"], start=1):
            print(f"{i}. {txt}")
    else:
        print("【风险提示】当前风险较低。")

    print("-" * 50)
    print(f"趋势评分：{result['trend_score']}/30")
    print(f"结构评分：{result['structure_score']}/30")
    print(f"量能评分：{result['volume_score']}/20")
    print(f"风险评分：{result['risk_score']}/20")
    print(f"\n★★★★★ 综合评分：{result['total_score']}/100")
    print("=" * 50)

    save_choice = input("\n是否保存本次分析到Excel？(Y/N)：").strip().upper()

    if save_choice == "Y":
        plan_buy = input("是否计划纳入观察池或交易计划？(Y/N)：").strip().upper()
        remark = input("请输入备注（可直接回车跳过）：")

        record = {
            "分析日期": datetime.now().strftime("%Y-%m-%d"),
            "股票代码": code,
            "股票名称": name,
            "收盘价": close,
            "MA5": ma5,
            "MA10": ma10,
            "MA20": ma20,
            "20日最高价": high20,
            "今日成交量": volume,
            "5日均量": avg_volume,
            "乖离率(%)": result["bias"],
            "量比": result["volume_ratio"],
            "阶段": result["stage"],
            "综合评分": result["total_score"],
            "交易建议": result["advice"],
            "计划买入": plan_buy,
            "备注": remark
        }

        save_to_excel(record)
        # 保存DataFrame
        with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=SHEET_NAME, index=False)
        # 重新打开Excel美化
        wb = load_workbook(EXCEL_FILE)
        
    else:
        print("\n本次分析未保存。")

    print("\n感谢使用 A股AI右侧交易助手 V1.0")


# ==========================
# 程序入口
# ==========================
if __name__ == "__main__":
    main()
    except Exception as e:
        print("\n❌ 程序报错：")
        print(e)
    input("\n按回车退出...")
