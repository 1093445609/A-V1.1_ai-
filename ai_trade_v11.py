import os
import pandas as pd
from datetime import datetime

# ==========================
# Excel配置
# ==========================
FILE = "AI_Trade_Record.xlsx"
SHEET = "trade"

# ==========================
# 阶段识别
# ==========================
def get_stage(close, ma5, ma10, ma20, bias, volume_ratio):
    if ma5 > ma10 > ma20 and close > ma20 and bias <= 8:
        return "起涨阶段", "BUY"
    elif ma5 > ma10 > ma20 and bias <= 18:
        return "加速阶段", "HOLD"
    elif bias > 18:
        return "末端阶段", "SELL"
    else:
        return "震荡整理", "WATCH"


# ==========================
# 持仓风控
# ==========================
def position_risk(current, cost, ma10, ma20):
    profit = (current - cost) / cost * 100

    action = []
    risk_level = "正常"

    if profit <= -8:
        action.append("强制止损（亏损≥8%）")
        risk_level = "高风险"
        print("不允许“扛单”，这是账户保护线")
    if current < ma20:
        action.append(" 跌破MA20 → 清仓信号")
        risk_level = "高风险"
        print("右侧趋势结束，必须执行纪律")
    elif current < ma10:
        action.append(" 跌破MA10 → 建议减仓50%")
        risk_level = "中风险"

    if profit >= 10:
        action.append(" 已盈利≥10% → 可分批止盈")

    if profit >= 20:
        action.append(" 盈利≥20% → 高位风险区")

    return round(profit, 2), risk_level, action


# ==========================
# 评分系统
# ==========================
def score(ma5, ma10, ma20, close, bias, volume_ratio):
    s = 0

    if ma5 > ma10 > ma20:
        s += 30
    if close > ma20:
        s += 20
    if 0 < bias < 10:
        s += 20
    if 1 <= volume_ratio <= 2:
        s += 20
    if bias > 18:
        s -= 20

    return max(min(s, 100), 0)


# ==========================
# Excel保存
# ==========================
def save(record):
    import pandas as pd
    import os
    from openpyxl import load_workbook
    from openpyxl.styles import PatternFill, Font

    # =========================
    # 读取或创建Excel
    # =========================
    if os.path.exists(FILE):
        df = pd.read_excel(FILE)
    else:
        df = pd.DataFrame()

    df = pd.concat([df, pd.DataFrame([record])], ignore_index=True)
    df.to_excel(FILE, index=False)

    # =========================
    # Excel美化处理（核心）
    # =========================
    wb = load_workbook(FILE)
    ws = wb.active

    # -------- 1. 开启筛选 --------
    ws.auto_filter.ref = ws.dimensions

    # -------- 2. 冻结首行 --------
    ws.freeze_panes = "A2"

    # -------- 3. 自动列宽 --------
    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter

        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))

        ws.column_dimensions[col_letter].width = max_len + 3

    # -------- 4. 表头加粗 --------
    for cell in ws[1]:
        cell.font = Font(bold=True)

    # -------- 5. 阶段颜色标记 --------
    # 找到“阶段”列位置（防止列错位）
    stage_col = None
    for idx, cell in enumerate(ws[1], 1):
        if cell.value == "阶段":
            stage_col = idx
            break

    if stage_col:
        for row in ws.iter_rows(min_row=2):
            cell = row[stage_col - 1]

            if cell.value == "起涨阶段":
                fill = PatternFill("solid", fgColor="C6EFCE")  # 绿
            elif cell.value == "加速阶段":
                fill = PatternFill("solid", fgColor="FFF2CC")  # 黄
            elif cell.value == "末端阶段":
                fill = PatternFill("solid", fgColor="F8CBAD")  # 红
            else:
                fill = PatternFill("solid", fgColor="E7E6E6")  # 灰

            cell.fill = fill

    wb.save(FILE)

    print("✔ Excel已保存 + 筛选 + 排序 + 颜色标记完成")


# ==========================
# 主程序
# ==========================
def main():
    print("=" * 50)
    print("A股右侧交易系统 V1.1")
    print("只做右侧，不做预测")
    print("宁可错过，不可做错")
    print("任何交易必须服从风控，而不是情绪")
    print("=" * 50)
    
    mode = input("选择模式 1=分析 2=持仓：")

    code = input("股票代码：")
    name = input("股票名称：")

    close = float(input("收盘价："))
    ma5 = float(input("MA5："))
    ma10 = float(input("MA10："))
    ma20 = float(input("MA20："))

    high20 = float(input("20日最高价："))
    vol = float(input("今日成交量："))
    avg_vol = float(input("5日均量："))

    bias = (close - ma20) / ma20 * 100
    volume_ratio = vol / avg_vol if avg_vol != 0 else 0

    stage, signal = get_stage(close, ma5, ma10, ma20, bias, volume_ratio)
    total_score = score(ma5, ma10, ma20, close, bias, volume_ratio)

    print("\n===== 分析结果 =====")
    print("阶段：", stage)
    print("信号：", signal)
    print("评分：", total_score)
    print("乖离率：", round(bias, 2), "%")
    print("量比：", round(volume_ratio, 2))

    # ==========================
    # 持仓模式
    # ==========================
    if mode == "2":
        cost = float(input("买入成本："))

        profit, risk_level, actions = position_risk(close, cost, ma10, ma20)

        print("\n===== 持仓分析 =====")
        print("收益率：", profit, "%")
        print("风险等级：", risk_level)

        print("\n===== 风控建议 =====")
    
        for a in actions:
            print(a)
            
            

    # ==========================
    # 是否保存
    # ==========================
    save_choice = input("\n是否保存到Excel？Y/N：")
   

    if save_choice.upper() == "Y":
        remark = input("请输入备注（可直接回车跳过）：")
        
        
            
        record = {
            "时间": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "代码": code,
            "名称": name,
            "阶段": stage,
            "评分": total_score,
            "乖离率": round(bias, 2),
            "量比": round(volume_ratio, 2)
        }
        if mode == "2":
            advice = actions
            record["收益率"] = profit
            record["风险"] = risk_level
            record["交易建议"] = advice
       
        
        record["备注"] = remark
        save(record)

    print("\n运行结束")

    input("\n按回车退出（防闪退）...")


# ==========================
# 防闪退入口
# ==========================
if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("\n❌ 程序错误：", e)
        
        input("按回车退出...")
