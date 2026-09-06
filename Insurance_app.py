import streamlit as st

# 先拔除 Playwright 相關套件與下載指令，純測試 UI 是否能正常顯示
st.title("📊 保單逆按揭分析與 PPT 產生器 (UI 測試版)")

discount_rate = st.number_input("首年折扣 (%)", min_value=0.0, max_value=100.0, step=1.0)
uploaded_file = st.file_uploader("上傳保單建議書 (PDF)", type="pdf")

if st.button("掃描數據並生成 PPT"):
    st.success("UI 介面顯示成功！如果按下按鈕能看到這行字，代表您的雲端連線與檔案路徑完全正確！")
