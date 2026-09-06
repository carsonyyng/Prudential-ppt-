import streamlit as st
import pdfplumber
from playwright.sync_api import sync_playwright

def parse_insurance_pdf(uploaded_file):
    # 這裡將使用 pdfplumber 抓取座標或表格資料
    with pdfplumber.open(uploaded_file) as pdf:
        pass
        
    # 暫時使用 Mr. Hor 的資料進行模擬回傳
    return {
        "name": "Hor",
        "age": 51,
        "gender": "M",
        "annual_premium_usd": 3562.76, 
        "sum_assured_usd": 60000,      
        "val_at_86_usd": 194126        
    }

def fetch_hkmc_payout(gender, death_benefit_hkd):
    # 背景啟動無頭瀏覽器模擬操作 HKMC 網站
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto("https://www.hkmc.com.hk/chi/online_tools/policy_reverse_mortgage_programme/policy_reverse_mortgage_calculator.html")
        
        if gender == "M":
            page.click("input#gender_M") # 需替換為實際網頁元素 ID
        else:
            page.click("input#gender_F") # 需替換為實際網頁元素 ID
            
        page.fill("input#age", "65")
        page.fill("input#death_benefit", str(death_benefit_hkd))
        page.fill("input#policy_value", "0")
        
        page.click("button#calculate_btn") # 需替換為實際按鈕 ID
        
        page.wait_for_selector(".result-table") # 需替換為實際表格 Class
        monthly_payout_text = page.inner_text("tr:has-text('20年') >> td.payout-value")
        
        browser.close()
        
        return float(monthly_payout_text.replace(',', '').replace('$', '').strip())

def generate_ppt(data, discount_rate, monthly_payout, annual_payout, total_contribution, 
                 total_20_years, a_hkd, b_hkd, a_minus_b, cost_performance):
    
    # 這裡為 python-pptx 排版邏輯預留的變數對接區域：
    # [右側 RHS 區塊]
    # - 每月提取港幣: monthly_payout
    # - 全年約港幣: annual_payout
    # - 共收取現金約港幣: total_20_years
    # - 身故賠償展示 (約港幣 A-B): A 顯示為 a_hkd, B 顯示為 b_hkd
    # - 共約港幣——————給至愛親人: a_minus_b
    # [左下側 LHS 區塊]
    # - 總供款港幣: total_contribution
    # - 性價比 =: cost_performance

    ppt_file_path = f"generated_ppt_{data['name']}.pptx"
    
    # 暫時生成一個空檔案供下載測試
    with open(ppt_file_path, "w") as f:
        f.write("PPTX content goes here")
        
    return ppt_file_path

# --- Streamlit UI ---
st.title("📊 保單逆按揭分析與 PPT 產生器")

EXCHANGE_RATE = 7.85

discount_rate = st.number_input("首年折扣 (%)", min_value=0.0, max_value=100.0, step=1.0)
uploaded_file = st.file_uploader("上傳保單建議書 (PDF)", type="pdf")

if st.button("掃描數據並生成 PPT"):
    if uploaded_file is not None and discount_rate > 0:
        
        with st.spinner('🔄 正在讀取 PDF 數據...'):
            data = parse_insurance_pdf(uploaded_file)
            st.success(f"成功擷取客戶資訊: {data['name']}, {data['age']}歲, {data['gender']}")
            
        with st.spinner('🧮 正在進行逆按揭與財務公式計算...'):
            # 1. 基礎轉換
            death_benefit_hkd = data['sum_assured_usd'] * EXCHANGE_RATE
            total_contribution_hkd = data['annual_premium_usd'] * 9.9 * EXCHANGE_RATE
            
            # 2. 獲取 HKMC 每月年金 (使用 try-except 在爬蟲未配置好前用模擬數據測試)
            try:
                monthly_payout_hkd = fetch_hkmc_payout(data['gender'], death_benefit_hkd)
            except Exception as e:
                st.warning("HKMC 爬蟲選擇器尚未配置，使用模擬數據繼續流程。")
                monthly_payout_hkd = death_benefit_hkd * 0.0025
                
            # 3. 提取現金計算
            annual_payout_hkd = monthly_payout_hkd * 12
            total_20_years_hkd = annual_payout_hkd * 20  # 共收取現金約港幣
            
            # 4. 身故賠償 A-B 算式
            a_hkd = data['val_at_86_usd'] * EXCHANGE_RATE # 將 @ANB 86歲 (A)+(B) 轉為港幣
            b_hkd = annual_payout_hkd * 1.46
            a_minus_b_hkd = a_hkd - b_hkd                 # 共約港幣——————給至愛親人
            
            # 5. 最終性價比計算
            cost_performance = (total_20_years_hkd + a_minus_b_hkd) / total_contribution_hkd

        with st.spinner('📝 正在排版並生成 PPTX 檔案...'):
            ppt_path = generate_ppt(
                data=data,
                discount_rate=discount_rate,
                monthly_payout=monthly_payout_hkd,
                annual_payout=annual_payout_hkd,
                total_contribution=total_contribution_hkd,
                total_20_years=total_20_years_hkd,
                a_hkd=a_hkd,
                b_hkd=b_hkd,
                a_minus_b=a_minus_b_hkd,
                cost_performance=cost_performance
            )
            
        st.success(f"✅ 處理完成！計算得出性價比為: {cost_performance:.2f}X")
        
        with open(ppt_path, "rb") as file:
            st.download_button(
                label="下載 PPT 簡報",
                data=file,
                file_name=f"保單逆按揭分析_{data['name']}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
    else:
        st.error("請確認已填寫首年折扣並上傳 PDF 檔案。")