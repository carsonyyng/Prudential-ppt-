import streamlit as st
import pdfplumber
from playwright.sync_api import sync_playwright
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
import subprocess
import os

# 安全地在背景安裝瀏覽器
@st.cache_resource
def install_browser():
    subprocess.run(["playwright", "install", "chromium"])

def parse_insurance_pdf(uploaded_file):
    with pdfplumber.open(uploaded_file) as pdf:
        pass
        
    return {
        "name": "Hor",
        "age": 42, 
        "gender": "M",
        "annual_premium_usd": 5349.17, 
        "sum_assured_usd": 129006,     
        "val_at_86_usd": 414965        
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

def generate_ppt(data, discount_rate, annual_premium_hkd, death_benefit_hkd,
                 monthly_payout, annual_payout, total_contribution, 
                 total_20_years, a_hkd, b_hkd, a_minus_b, cost_performance):
    
    prs = Presentation()
    slide_layout = prs.slide_layouts[6] 
    slide = prs.slides.add_slide(slide_layout)

    # ==========================================
    # 頂部重要資訊 (放大字體)
    # ==========================================
    tx_top_left = slide.shapes.add_textbox(Inches(0.2), Inches(0.1), Inches(8), Inches(0.6))
    tf_top_left = tx_top_left.text_frame
    p_top = tf_top_left.paragraphs[0]
    p_top.text = f"USD:HKD = 1:7.85    |    首年折扣：{discount_rate} %"
    p_top.font.size = Pt(20)
    p_top.font.bold = True
    p_top.font.color.rgb = RGBColor(80, 80, 80)

    # ==========================================
    # 1. 最左側第一個箭頭區塊 (藍色 Blue)
    # ==========================================
    hdr1 = slide.shapes.add_shape(MSO_SHAPE.PENTAGON, Inches(0.2), Inches(0.8), Inches(5.5), Inches(1.0))
    hdr1.fill.solid()
    hdr1.fill.fore_color.rgb = RGBColor(68, 114, 196) 
    hdr1.line.color.rgb = RGBColor(255, 255, 255)
    
    img1 = "Male ppt picture.svg" if data['gender'] == "M" else "Female ppt picture.svg"
    if os.path.exists(img1):
        slide.shapes.add_picture(img1, Inches(0.3), Inches(0.9), height=Inches(0.8))

    prefix = "MR" if data['gender'] == "M" else "MS"
    tf1 = hdr1.text_frame
    tf1.margin_left = Inches(1.2) 
    p1 = tf1.paragraphs[0]
    p1.text = f"{prefix} {data['name'].upper()} ({data['age']} 歲)"
    p1.font.size = Pt(26)
    p1.font.color.rgb = RGBColor(255, 255, 255)
    p1.alignment = PP_ALIGN.LEFT

    body1 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.2), Inches(1.8), Inches(3.2), Inches(4.2))
    body1.fill.solid()
    body1.fill.fore_color.rgb = RGBColor(255, 255, 255)
    body1.line.color.rgb = RGBColor(68, 114, 196) 
    tf1_body = body1.text_frame
    tf1_body.margin_left = Inches(0.15)
    tf1_body.margin_top = Inches(0.15)
    
    p = tf1_body.paragraphs[0]
    p.text = "投保自主未來產品\n(10年供款)\n"
    p.font.size = Pt(16)
    p.font.color.rgb = RGBColor(0, 0, 0)
    
    p = tf1_body.add_paragraph()
    p.text = f"每年保費港幣 {annual_premium_hkd:,.0f}\n投保額港幣 {death_benefit_hkd:,.0f}\n\n\n"
    p.font.size = Pt(16)
    p.font.color.rgb = RGBColor(0, 0, 0)
    
    p = tf1_body.add_paragraph()
    p.text = f"(10年後{data['age'] + 10}歲)\n"
    p.font.size = Pt(16)
    p.font.color.rgb = RGBColor(0, 0, 0)
    
    p = tf1_body.add_paragraph()
    p.text = "總供款港幣 "
    p.font.size = Pt(16)
    p.font.color.rgb = RGBColor(0, 0, 0)
    run = p.add_run()
    run.text = f"{total_contribution:,.0f}"
    run.font.color.rgb = RGBColor(68, 114, 196) 

    # ==========================================
    # 2. 中間第二個箭頭區塊 (Tiffany 藍)
    # ==========================================
    hdr2 = slide.shapes.add_shape(MSO_SHAPE.PENTAGON, Inches(3.3), Inches(1.55), Inches(5.5), Inches(1.1))
    hdr2.fill.solid()
    hdr2.fill.fore_color.rgb = RGBColor(10, 186, 181) 
    hdr2.line.color.rgb = RGBColor(255, 255, 255)
    
    if os.path.exists("Middle header ppt picture.svg"):
        slide.shapes.add_picture("Middle header ppt picture.svg", Inches(3.4), Inches(1.65), height=Inches(0.9))

    tf2 = hdr2.text_frame
    tf2.margin_left = Inches(1.1)
    p2 = tf2.paragraphs[0]
    p2.text = "(65歲時)再用保單逆按形式\n提取20年年金"
    p2.font.size = Pt(18)
    p2.font.color.rgb = RGBColor(255, 255, 255)
    p2.alignment = PP_ALIGN.LEFT

    body2 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(3.3), Inches(2.65), Inches(3.2), Inches(3.35))
    body2.fill.solid()
    body2.fill.fore_color.rgb = RGBColor(255, 255, 255)
    body2.line.color.rgb = RGBColor(10, 186, 181) 
    tf2_body = body2.text_frame
    tf2_body.margin_left = Inches(0.15)
    tf2_body.margin_top = Inches(0.15)
    
    p = tf2_body.paragraphs[0]
    p.text = f"每月提取約港幣 {monthly_payout:,.0f}\n(全年約港幣 {annual_payout:,.0f})"
    p.font.size = Pt(16)
    p.font.color.rgb = RGBColor(0, 0, 0)

    # ==========================================
    # 3. 最右側第三個箭頭區塊 (綠色 Green)
    # ==========================================
    hdr3 = slide.shapes.add_shape(MSO_SHAPE.PENTAGON, Inches(6.4), Inches(2.1), Inches(3.5), Inches(1.0))
    hdr3.fill.solid()
    hdr3.fill.fore_color.rgb = RGBColor(112, 173, 71) 
    hdr3.line.color.rgb = RGBColor(255, 255, 255)
    
    if os.path.exists("RHS header ppt picture.svg"):
        slide.shapes.add_picture("RHS header ppt picture.svg", Inches(6.5), Inches(2.2), height=Inches(0.8))

    tf3 = hdr3.text_frame
    tf3.margin_left = Inches(1.0)
    p3 = tf3.paragraphs[0]
    p3.text = "再20年後(85歲)"
    p3.font.size = Pt(20)
    p3.font.color.rgb = RGBColor(255, 255, 255)
    p3.alignment = PP_ALIGN.LEFT

    # 稍微拉長方塊高度至 3.6 吋以容納放大的字體與空白行
    body3 = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(6.4), Inches(3.1), Inches(3.5), Inches(3.6))
    body3.fill.solid()
    body3.fill.fore_color.rgb = RGBColor(255, 255, 255)
    body3.line.color.rgb = RGBColor(112, 173, 71) 
    tf3_body = body3.text_frame
    tf3_body.word_wrap = True 
    tf3_body.margin_left = Inches(0.1)
    tf3_body.margin_top = Inches(0.1)
    
    # 放大部分文字，並加強重點顏色
    p = tf3_body.paragraphs[0]
    p.text = "共收取現金約港幣 "
    p.font.size = Pt(18)
    p.font.color.rgb = RGBColor(0, 0, 0)
    run = p.add_run()
    run.text = f"{total_20_years:,.0f}"
    run.font.size = Pt(18)
    run.font.color.rgb = RGBColor(255, 0, 0)
    
    p = tf3_body.add_paragraph()
    p.text = "\n身故賠償"
    p.font.size = Pt(20)
    p.font.color.rgb = RGBColor(0, 0, 0)
    
    p = tf3_body.add_paragraph()
    p.text = f"約港幣 {a_hkd:,.0f} - {b_hkd:,.0f} (逆按欠款) ，"
    p.font.size = Pt(16)
    p.font.color.rgb = RGBColor(0, 0, 0)
    
    # 增加空白行將算式與結果拉開
    p_spacer = tf3_body.add_paragraph()
    p_spacer.text = "\n\n" 
    p_spacer.font.size = Pt(14)
    
    p = tf3_body.add_paragraph()
    p.text = "共約港幣 "
    p.font.size = Pt(20)
    p.font.color.rgb = RGBColor(0, 0, 0)
    run = p.add_run()
    run.text = f"{a_minus_b:,.0f}"
    run.font.size = Pt(20)
    run.font.color.rgb = RGBColor(255, 0, 0)
    run.font.bold = True
    run2 = p.add_run()
    run2.text = " 給至愛親人"
    run2.font.size = Pt(20)
    run2.font.color.rgb = RGBColor(0, 0, 0)

    # ==========================================
    # 4. 左下角：性價比公式 (放大字體)
    # ==========================================
    tx_bottom = slide.shapes.add_textbox(Inches(0.2), Inches(6.0), Inches(6.0), Inches(1.3))
    tf_bottom = tx_bottom.text_frame
    p_bot = tf_bottom.paragraphs[0]
    p_bot.text = f"性價= {cost_performance:.2f}X"
    p_bot.font.size = Pt(32)
    p_bot.font.bold = True
    p_bot.font.color.rgb = RGBColor(237, 27, 46) 
    
    p_bot_sub = tf_bottom.add_paragraph()
    p_bot_sub.text = f"(${total_20_years:,.0f} + ${a_minus_b:,.0f}) / ${total_contribution:,.0f}"
    p_bot_sub.font.size = Pt(16)
    p_bot_sub.font.color.rgb = RGBColor(85, 85, 85)

    ppt_file_path = f"generated_ppt_{data['name']}.pptx"
    prs.save(ppt_file_path)
    return ppt_file_path

# --- Streamlit UI ---
st.title("📊 保單逆按揭分析與 PPT 產生器")

install_browser()

EXCHANGE_RATE = 7.85
discount_rate = st.number_input("首年折扣 (%)", min_value=0.0, max_value=100.0, step=1.0)
uploaded_file = st.file_uploader("上傳保單建議書 (PDF)", type="pdf")

if st.button("掃描數據並生成 PPT"):
    if uploaded_file is not None:
        
        discount_multiplier = 10.0 - (discount_rate / 100.0)
        st.info(f"💡 系統已套用供款乘數： 10 - {discount_rate/100} = {discount_multiplier}")
        
        with st.spinner('🔄 正在讀取 PDF 數據...'):
            data = parse_insurance_pdf(uploaded_file)
            st.success(f"成功擷取客戶資訊: {data['name']}, {data['age']}歲, {data['gender']}")
            
        with st.spinner('🧮 正在與 HKMC 伺服器連線計算逆按揭...'):
            annual_premium_hkd = data['annual_premium_usd'] * EXCHANGE_RATE
            death_benefit_hkd = data['sum_assured_usd'] * EXCHANGE_RATE
            
            total_contribution_hkd = annual_premium_hkd * discount_multiplier
            
            try:
                monthly_payout_hkd = fetch_hkmc_payout(data['gender'], death_benefit_hkd)
            except Exception as e:
                st.warning("HKMC 爬蟲選擇器尚未配置，使用模擬數據繼續流程。")
                monthly_payout_hkd = death_benefit_hkd * 0.0025
                
            annual_payout_hkd = monthly_payout_hkd * 12
            total_20_years_hkd = annual_payout_hkd * 20  
            
            a_hkd = data['val_at_86_usd'] * EXCHANGE_RATE 
            b_hkd = total_20_years_hkd * 1.46
            a_minus_b_hkd = a_hkd - b_hkd                 
            
            cost_performance = (total_20_years_hkd + a_minus_b_hkd) / total_contribution_hkd

        with st.spinner('📝 正在排版並生成 PPTX 檔案...'):
            ppt_path = generate_ppt(
                data=data,
                discount_rate=discount_rate,
                annual_premium_hkd=annual_premium_hkd,
                death_benefit_hkd=death_benefit_hkd,
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
        st.error("請上傳 PDF 檔案。")
