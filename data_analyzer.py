import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# 設定要分析的檔案名稱 (請確保檔案已轉存為 UTF-8)
target_file = '103-105_students.csv'

# 設定 Pandas 讓終端機顯示中文時可以對齊
pd.set_option('display.unicode.ambiguous_as_wide', True)
pd.set_option('display.unicode.east_asian_width', True)
pd.set_option('display.width', 180)

def load_and_clean(path):
    """
    第一步：讀取資料並清理環境
    """
    print("\n系統訊息: 正在讀取資料檔案...")
    try:
        # 讀取 CSV
        df = pd.read_csv(path, encoding='utf-8')
        
        # 清理資料：把空格補 0，刪掉重複的行
        df_clean = df.fillna(0).drop_duplicates()
        
        print(f"完成！原始資料 {len(df)} 筆，清理後剩餘 {len(df_clean)} 筆。")
        return df_clean
    except Exception as e:
        print(f"讀取失敗，請確認檔案編碼是否為 UTF-8。原因: {e}")
        return None

def run_analysis(df):
    """
    第二步：執行各項統計計算
    """
    print("\n" + "="*60)
    print("                數據統計分析報告 (103-105學年度)")
    print("="*60)

    # 1. 統計學位比例
    level_data = df.groupby('等級別')['總計'].sum().reset_index()
    print("\n[一、學位分布統計表]")
    print("-" * 30)
    print(level_data.to_string(index=False, justify='center'))
    print("-" * 30)
    
    # 2. 趨勢分析 (使用透視表看這三年的變化)
    trend = df.pivot_table(index='學年度', values='總計', aggfunc='sum').reset_index()
    trend['增加人數'] = trend['總計'].diff().fillna(0).astype(int)
    trend['成長率(%)'] = (trend['總計'].pct_change() * 100).round(2).fillna(0)
    
    print("\n[二、年度學生人數增長趨勢]")
    print("-" * 60)
    print(trend.to_string(index=False, justify='center'))
    print("-" * 60)

    # 3. 縣市排名 (加入互動功能，讓使用者決定看前幾名)
    user_in = input("\n請輸入想查看的熱門縣市數量 (直接按 Enter 預設為 10): ")
    n = int(user_in) if user_in.isdigit() else 10
    
    city_top = df.groupby('縣市名稱')['總計'].sum().sort_values(ascending=False).head(n).reset_index()
    print(f"\n[三、學生人數前 {n} 名縣市排名]")
    print("-" * 40)
    print(city_top.to_string(index=False, justify='center'))
    print("-" * 40)

    print("\n數據分析執行完畢。")
    print("="*60)

    return level_data, city_top, trend

def generate_reports(d1, d2, d3):
    """
    第三步：輸出圖表報告與 Excel 檔
    """
    # 設定字體解決中文亂碼 (微軟正黑體)
    plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']
    plt.rcParams['axes.unicode_minus'] = False

    print("\n系統訊息: 正在生成視覺化圖表與 Excel 報表...")
    
    # 建立一個大畫布，裡面放三張小圖
    plt.figure(figsize=(20, 6))

    # 圖 1：圓餅圖
    plt.subplot(1, 3, 1)
    plt.pie(d1['總計'], labels=d1['等級別'], autopct='%1.1f%%', startangle=90)
    plt.title('學位分布比例')

    # 圖 2：長條圖
    plt.subplot(1, 3, 2)
    sns.barplot(x='總計', y='縣市名稱', data=d2)
    plt.title('縣市人數排名')

    # 圖 3：折線圖
    plt.subplot(1, 3, 3)
    plt.plot(d3['學年度'], d3['總計'], marker='o', color='blue', linewidth=2)
    plt.title('103-105 學年度總人數趨勢')

    # 存成圖片
    plt.tight_layout()
    plt.savefig('project_final_report.png')
    print("存檔成功: project_final_report.png")

    # 匯出成 Excel 檔案
    try:
        with pd.ExcelWriter('student_analysis_results.xlsx') as writer:
            d1.to_excel(writer, sheet_name='學位統計', index=False)
            d2.to_excel(writer, sheet_name='縣市排名', index=False)
            d3.to_excel(writer, sheet_name='年度趨勢', index=False)
        print("存檔成功: student_analysis_results.xlsx")
    except Exception as e:
        print(f"Excel 存檔失敗，請檢查是否安裝 openpyxl。錯誤: {e}")

    # 顯示圖表
    plt.show()

# --- 程式主入口 ---
if __name__ == "__main__":
    # 1. 載入與清理
    clean_df = load_and_clean(target_file)
    
    if clean_df is not None:
        # 2. 統計分析
        res1, res2, res3 = run_analysis(clean_df)
        
        # 3. 輸出報告
        generate_reports(res1, res2, res3)
        
        print("\n所有程序執行結束，請檢查專案資料夾下的報表檔案。")