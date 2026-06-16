from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import pandas as pd
import os
import time
from datetime import datetime, timedelta

# ===========================
# 配置
# ===========================
TOP_N = 10
KEEP_DAYS = 30

# 东方财富板块页面
BOARD_URL = "https://quote.eastmoney.com/center/hsbk.html"

# Excel文件路径
INDUSTRY_EXCEL = "industry_sectors.xlsx"
CONCEPT_EXCEL = "concept_sectors.xlsx"


# ===========================
# 初始化浏览器
# ===========================
def init_browser():
    print("正在启动浏览器（首次需下载驱动，请等待）...")
    options = Options()
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-blink-features=AutomationControlled")
  
    # 手动指定驱动路径（如果你已经装了）
    driver = webdriver.Chrome(service=Service("E:\\workSoft\\A_right_system_V1.1\\chromedriver-win64\\chromedriver-win64\\chromedriver.exe"), options=options)
    
    # driver = webdriver.Chrome(
    #    service=Service(ChromeDriverManager().install()),
    #    options=options
   # )
    
    print("浏览器启动成功！")  # 加这行确认
    driver.set_page_load_timeout(30)
    return driver



# ===========================
# 抓取板块前10数据
# ===========================
def get_top10_sectors(driver, sector_type="industry"):
    """
    sector_type: "industry" 行业板块, "concept" 概念板块
    """

    print(f"正在抓取{'行业' if sector_type == 'industry' else '概念'}板块...")

    # 根据板块类型定位对应表格
    # 行业板块标题: <div class="t">行业板块涨幅</div>
    # 概念板块标题: <div class="t">概念板块涨幅</div>
    # 表格class: quotetable_m
    
    try:
        if sector_type == "industry":
            # 定位行业板块表格（找到包含"行业板块涨幅"的div后面的表格）
            table_xpath = "//div[contains(text(),'行业板块涨幅')]/ancestor::div[@class='title2quotetable']//table[@class='quotetable_m']"
        else:
            # 定位概念板块表格
            table_xpath = "//div[contains(text(),'概念板块涨幅')]/ancestor::div[@class='title2quotetable']//table[@class='quotetable_m']"

        # 等待表格加载
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, table_xpath))
        )

        # 获取表格行
        rows = driver.find_elements(By.XPATH, f"{table_xpath}//tbody/tr")

        sectors = []
        for i, row in enumerate(rows[:TOP_N]):
            cols = row.find_elements(By.TAG_NAME, "td")

            # 表格列：排名(0)、板块名称(1)、相关链接(2)、相关链接(3)、涨跌幅(4)...
            if len(cols) >= 5:
                rank = i + 1
                sector_name = cols[1].text.strip()
                change = cols[4].text.strip()

                # 过滤无效数据
                if sector_name and change and change != '-':
                    sectors.append({
                        "日期": datetime.now().strftime("%Y-%m-%d"),
                        "排名": rank,
                        "板块名称": sector_name,
                        "涨跌幅": change
                    })

        print(f"成功抓取{len(sectors)}条{'行业' if sector_type == 'industry' else '概念'}板块数据")
        return pd.DataFrame(sectors)

    except Exception as e:
        print(f"抓取失败: {e}")
        return pd.DataFrame()


# ===========================
# 更新Excel文件
# ===========================
def update_excel(today_df, excel_file):
    today_date = datetime.now().strftime("%Y-%m-%d")
    today_sheet_name = today_date  # sheet名称为日期格式

    if os.path.exists(excel_file):
        # 读取现有Excel文件
        try:
            # 获取所有sheet名称
            excel_file_obj = pd.ExcelFile(excel_file)
            existing_sheets = excel_file_obj.sheet_names

            # 收集所有历史数据（排除Statistics sheet）
            all_data = []
            for sheet in existing_sheets:
                if sheet != "Statistics":
                    df = pd.read_excel(excel_file, sheet_name=sheet)
                    all_data.append(df)
            
            history = pd.concat(all_data, ignore_index=True) if all_data else pd.DataFrame(columns=["日期", "排名", "板块名称", "涨跌幅"])
            
        except Exception as e:
            print(f"读取现有文件失败: {e}")
            history = pd.DataFrame(columns=["日期", "排名", "板块名称", "涨跌幅"])
            existing_sheets = []
    else:
        history = pd.DataFrame(columns=["日期", "排名", "板块名称", "涨跌幅"])
        existing_sheets = []

    # 删除今天旧数据（避免重复）
    if not history.empty:
        history["日期"] = pd.to_datetime(history["日期"])
        history = history[history["日期"].dt.strftime("%Y-%m-%d") != today_date]

    # 添加今日数据
    history = pd.concat([history, today_df], ignore_index=True)

    # 保留最近30天
    if not history.empty:
        history["日期"] = pd.to_datetime(history["日期"])
        cutoff = datetime.now() - timedelta(days=KEEP_DAYS)
        history = history[history["日期"] >= cutoff]

    # 统计板块出现次数
    statistics = (
        history.groupby("板块名称")
        .agg(
            出现次数=("板块名称", "count"),
            最近一次出现=("日期", "max")
        )
        .sort_values("出现次数", ascending=False)
        .reset_index()
    )

    # 将日期格式化为字符串用于写入
    history["日期"] = history["日期"].dt.strftime("%Y-%m-%d %H:%M:%S")
    statistics["最近一次出现"] = statistics["最近一次出现"].dt.strftime("%Y-%m-%d")

    # 写入Excel（按日期分sheet）
    with pd.ExcelWriter(excel_file, engine="openpyxl") as writer:
        # 写入每日数据sheet
        history["日期"] = pd.to_datetime(history["日期"])
        dates = history["日期"].dt.strftime("%Y-%m-%d").unique()
        
        for date in dates:
            date_data = history[history["日期"].dt.strftime("%Y-%m-%d") == date]
            date_data = date_data.sort_values("排名", ascending=True)
            date_data["日期"] = date_data["日期"].dt.strftime("%Y-%m-%d %H:%M:%S")
            date_data.to_excel(writer, sheet_name=date, index=False)
        
        # 写入统计sheet
        statistics.to_excel(writer, sheet_name="Statistics", index=False)

    print(f"Excel更新完成: {excel_file}")
    print(f"今日sheet: {today_sheet_name} {'(已覆盖)' if today_sheet_name in existing_sheets else '(新建)'}")


# ===========================
# 主函数
# ===========================
def main():

    print("=" * 50)
    print("东方财富板块数据抓取")
    print("=" * 50)

    driver = init_browser()

    try:
        print("\n打开网页...")
        driver.get(BOARD_URL)
        print("网页已打开")

        # 手动验证（如果需要）
        input("请完成网页验证后按回车继续（如无需验证直接回车）...")

        # 抓取行业板块
        print("\n" + "-" * 30)
        df_industry = get_top10_sectors(driver, sector_type="industry")
        if not df_industry.empty:
            print("\n行业板块前十:")
            print(df_industry)
            update_excel(df_industry, INDUSTRY_EXCEL)
        else:
            print("未获取到行业板块数据")

        # 抓取概念板块
        print("\n" + "-" * 30)
        df_concept = get_top10_sectors(driver, sector_type="concept")
        if not df_concept.empty:
            print("\n概念板块前十:")
            print(df_concept)
            update_excel(df_concept, CONCEPT_EXCEL)
        else:
            print("未获取到概念板块数据")

        print("\n" + "=" * 50)
        print("任务完成！")
        print(f"行业板块数据: {INDUSTRY_EXCEL}")
        print(f"概念板块数据: {CONCEPT_EXCEL}")
        print("=" * 50)

    finally:
        driver.quit()


# ==========================
# 防闪退入口
# ==========================
if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("\n❌ 程序错误：", e)

        input("按回车退出...")
        driver.quit()
