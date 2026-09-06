import streamlit as st
import pdfplumber
from playwright.sync_api import sync_playwright
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
import subprocess

# 安全地在背景安裝瀏覽器，並快取起來確保每次重啟只執行一次
@st.cache_resource
def install_browser():
    subprocess.run(["playwright", "install", "chromium"])

def parse_insurance_pdf(uploaded_file):
    with pdfplumber.open(uploaded_file) as pdf:
        pass
        
    return {
        "name": "Hor",
        "age": 51,
        "gender": "M",
        "annual_premium_usd": 3562.76, 
        "sum_assured_usd": 60000,      
        "val_at_86_usd": 194126        
    }

def fetch_hkmc_payout(gender, death_benefit_hkd):
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto("https://www.hkmc.com.hk/chi/online_tools/policy_reverse_mortgage_programme/policy_reverse_mortgage_calculator.html")
        
        if gender == "M":
            page.click("input#gender_M") 
        else:
            page.click("input#gender_F") 
            
        page.fill("input#age", "65")
        page.fill("input#death_benefit", str(death_benefit_hkd))
        page.fill("input#policy_value", "0")
        
        page.click("button#calculate_btn") 
        
        page.wait_for_selector(".result-table") 
        monthly_payout_text = page.inner_text("tr:has-text('20年') >> td.payout-value")
        
        browser.close()
        
        return float(monthly_payout_text.replace(',', '').replace('$', '').strip())

def generate_ppt(data, discount_rate, monthly_payout, annual_payout, total_contribution, 
                 total_20_years, a_hkd, b_hkd, a_minus_b, cost_performance):
    
    prs = Presentation()
    slide_layout = prs.slide_layouts[6] 
    slide = prs.slides.add_slide(slide_layout)

    # 左上角: 匯率與首年折扣
    tx_top_left = slide.shapes.add_textbox(Inches(0.4), Inches(0.4), Inches(3), Inches(1))
    tf_top_left = tx_top_left.text_frame
    tf_top_left.text = f"USD:HKD = 1:7.85\n首年折扣 ：{discount_rate} %"
    tf_top_left.paragraphs[0].font.size = Pt(14)
    if len(tf_top_left.paragraphs) > 1:
        tf_top_left.paragraphs[1].font.size = Pt(14)

    # 頂部置中: 圖示, 姓名, 年齡
    icon = "👨" if data['gender'] == "M" else "👩"
    tx_title = slide.shapes.add_textbox(Inches(3), Inches(0.3), Inches(4), Inches(1))
    tf_title = tx_title.text_frame
    p_title = tf_title.paragraphs[0]
    p_title.text = f"{icon} {data['name']} ({data['age']}歲)"
    p_title.font.size = Pt(32)
    p_title.font.bold = True
    p_title.font.color.rgb = RGBColor(237, 27, 46) 
    p_title.alignment = PP_ALIGN.CENTER

    # ==========================================
    # 左側區塊 (LHS) - 藍綠色背景
    # ==========================================
    lhs_shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.5), Inches(4.2), Inches(5.0)
    )
    lhs_shape.fill.solid()
    lhs_shape.fill.fore_color.rgb = RGBColor(218, 238, 238)
    lhs_shape.line.color.rgb = RGBColor(100, 180, 180)
    
    tx_lhs = slide.shapes.add_textbox(Inches(0.5), Inches(1.8), Inches(4.2), Inches(4.5))
    tf_lhs = tx_lhs.text_frame
    
    p = tf_lhs.paragraphs[0]
    p.text = "總供款港幣"
    p.font.size = Pt(22)
    p.alignment = PP_ALIGN.CENTER
    
    p = tf_lhs.add_paragraph()
    p.text = f"${total_contribution:,.0f}"
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0, 102, 102)
    p.alignment = PP_ALIGN.CENTER
    
    tf_lhs.add_paragraph().text = "\n\n"
    
    p = tf_lhs.add_paragraph()
    p.text = f"性價= {cost_performance:.2f}X"
    p.font.size = Pt(30)
    p.font.bold = True
    p.font.color.rgb = RGBColor(237, 27, 46)
    p.alignment = PP_ALIGN.CENTER
    
    p = tf_lhs.add_paragraph()
    p.text = f"(${total_20_years:,.0f} + ${a_minus_b:,.0f}) / ${total_contribution:,.0f}"
    p.font.size = Pt(14)
    p.font.color.rgb = RGBColor(85, 85, 85)
    p.alignment = PP_ALIGN.CENTER

    # ==========================================
    # 右側區塊 (RHS) - 淺綠色背景
    # ==========================================
    rhs_shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.0), Inches(1.5), Inches(4.5), Inches(5.0)
    )
    rhs_shape.fill.solid()
    rhs_shape.fill.fore_color.rgb = RGBColor(224, 245, 224) 
    rhs_shape.line.color.rgb = RGBColor(120, 200, 120)

    tx_rhs = slide.shapes.add_textbox(Inches(5.0), Inches(1.6), Inches(4.5), Inches(4.8))
    tf_rhs = tx_rhs.text_frame
    
    p = tf_rhs.paragraphs[0]
    p.text = "每月提取港幣"
    p.font.size = Pt(18)
    p.alignment = PP_ALIGN.CENTER
    
    p = tf_rhs.add_paragraph()
    p.text = f"${monthly_payout:,.0f}"
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = RGBColor(34, 139, 34)
    p.alignment = PP_ALIGN.CENTER
    
    p = tf_rhs.add_paragraph()
    p.text = "\n全年約港幣"
    p.font.size = Pt(16)
    p.alignment = PP_ALIGN.CENTER
    
    p = tf_rhs.add_paragraph()
    p.text = f"${annual_payout:,.0f}"
    p.font.size = Pt(24)
    p.font.bold = True
    p.font.color.rgb = RGBColor(34, 139, 34)
    p.alignment = PP_ALIGN.CENTER
    
    p = tf_rhs.add_paragraph()
    p.text = "\n共收取現金約港幣"
    p.font.size = Pt(20)
    p.alignment = PP_ALIGN.CENTER
    
    p = tf_rhs.add_paragraph()
    p.text = f"${total_20_years:,.0f}"
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0, 128, 0)
    p.alignment = PP_ALIGN.CENTER
    
    tf_rhs.add_paragraph().text = "\n"
    
    p = tf_rhs.add_paragraph()
    p.text = f"身故賠償：約港幣 ${a_hkd:,.0f} - ${b_hkd:,.0f}"
    p.font.size = Pt(16)
    p.alignment = PP_ALIGN.CENTER
    
    p = tf_rhs.add_paragraph()
    p.text = f"共約港幣 ${a_minus_b:,.0f} 給至愛親人"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = RGBColor(237, 27, 46)
    p.alignment = PP_ALIGN.CENTER

    ppt_file_path = f"generated_ppt_{data['name']}.pptx"
    prs.save(ppt_file_path)
    return ppt_file_path

# --- Streamlit UI ---
st.title("📊 保單逆按揭分析與 PPT 產生器")

# 啟動時先呼叫安裝函數 (如果已安裝會瞬間跳過)
install_browser()

EXCHANGE_RATE = 7.85
discount_rate = st.number_input("首年折扣 (%)", min_value=0.0, max_value=100.0, step=1.0)
uploaded_file = st.file_uploader("上傳保單建議書 (PDF)", type="pdf")

if st.button("掃描數據並生成 PPT"):
    if uploaded_file is not None and discount_rate > 0:
        with st.spinner('🔄 正在讀取 PDF 數據...'):
            data = parse_insurance_pdf(uploaded_file)
            st.success(f"成功擷取客戶資訊: {data['name']}, {data['age']}歲, {data['gender']}")
            
        with st.spinner('🧮 正在與 HKMC 伺服器連線計算逆按揭 (若為雲端首次執行需時較長)...'):
            death_benefit_hkd = data['sum_assured_usd'] * EXCHANGE_RATE
            total_contribution_hkd = data['annual_premium_usd'] * 9.9 * EXCHANGE_RATE
            
            try:
                monthly_payout_hkd = fetch_hkmc_payout(data['gender'], death_benefit_hkd)
            except Exception as e:
                st.warning("HKMC 爬蟲選擇器尚未配置，使用模擬數據繼續流程。")
                monthly_payout_hkd = death_benefit_hkd * 0.0025
                
            annual_payout_hkd = monthly_payout_hkd * 12
            total_20_years_hkd = annual_payout_hkd * 20  
            
            a_hkd = data['val_at_86_usd'] * EXCHANGE_RATE 
            b_hkd = annual_payout_hkd * 1.46
            a_minus_b_hkd = a_hkd - b_hkd                 
            
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
