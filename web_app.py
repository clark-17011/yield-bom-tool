import streamlit as st
import pandas as pd
import openpyxl
import re
import unicodedata
import matplotlib.pyplot as plt
import io

# ==========================================
# 1. 共用工具 (保留原本邏輯)
# ==========================================
def normalize_str(x) -> str:
    if x is None: return ""
    s = str(x)
    s = unicodedata.normalize("NFKC", s)
    s = s.strip()
    s = re.sub(r"[^A-Za-z0-9]", "", s)
    return s

def parse_search_config(raw_text: str, is_space_or_mode: bool):
    # (完整保留您原本的 parse_search_config 邏輯)
    if not raw_text.strip(): return []
    raw_text = raw_text.replace("，", ",")
    segments = [s.strip() for s in raw_text.split(',') if s.strip()]
    configs = []
    for seg in segments:
        if seg.startswith('[') and seg.endswith(']'):
            content = seg[1:-1].strip()
            sub_terms = [t.strip() for t in content.split() if t.strip()]
            if sub_terms:
                configs.append({'display': seg, 'terms': sub_terms})
            continue
        parts = [p.strip() for p in seg.split() if p.strip()]
        if not parts: continue
        if is_space_or_mode:
            for p in parts: configs.append({'display': p, 'terms': [p]})
        else:
            if len(parts) > 1:
                display_name = f"[{' '.join(parts)}]"
                configs.append({'display': display_name, 'terms': parts})
            else:
                configs.append({'display': parts[0], 'terms': [parts[0]]})
    
    seen = set()
    unique = []
    for c in configs:
        if c['display'] not in seen:
            unique.append(c)
            seen.add(c['display'])
    return unique

# ==========================================
# 2. 頁面設定
# ==========================================
st.set_page_config(page_title="Yield & BOM Tool", layout="wide")
st.title("📊 良率報表 & BOM 搜尋工具 (Web版)")

# 使用 Tabs 分開兩大功能
tab_yield, tab_bom = st.tabs(["Yield Analysis", "BOM Search"])

# ==========================================
# 3. Yield Analysis 模組
# ==========================================
with tab_yield:
    col1, col2 = st.columns([1, 2])
    
    with col1:
        st.header("1. 檔案上傳")
        # 這裡取代原本的自動讀取和 DropTable
        uploaded_files = st.file_uploader("拖曳 Excel 檔案到此處", type=['xlsx', 'xls'], accept_multiple_files=True, key="yield_files")
        
        st.header("2. 搜尋設定")
        raw_search = st.text_area("關鍵字 (空格=AND, 逗號=OR)", height=100)
        chk_space = st.checkbox("空白代表「或」(OR)", value=False)
        
        # 處理上傳檔案的快取邏輯
        # 使用 st.cache_data 可以避免每次互動都重新讀取 Excel (大幅提升速度)
        @st.cache_data(ttl=3600)
        def load_yield_data(files):
            all_rows = []
            # 模擬簡單的讀取邏輯
            for uploaded_file in files:
                try:
                    wb = openpyxl.load_workbook(uploaded_file, read_only=True, data_only=True)
                    for sheet in wb.sheetnames:
                        ws = wb[sheet]
                        data = list(ws.values)
                        if not data: continue
                        headers = [str(h) for h in data[0]]
                        # 簡單範例：直接轉 DataFrame
                        df = pd.DataFrame(data[1:], columns=headers)
                        df['SourceLabel'] = uploaded_file.name
                        df['SheetName'] = sheet
                        # 這裡應該加入您原本的「正規化搜尋」邏輯來建立索引
                        # 為求範例簡潔，此處僅做簡易處理
                        all_rows.append(df)
                    wb.close()
                except Exception as e:
                    st.error(f"Error loading {uploaded_file.name}: {e}")
            if all_rows:
                return pd.concat(all_rows, ignore_index=True)
            return pd.DataFrame()

        df_yield_raw = pd.DataFrame()
        if uploaded_files:
            with st.spinner('讀取檔案中...'):
                df_yield_raw = load_yield_data(uploaded_files)
                st.success(f"已讀取 {len(uploaded_files)} 個檔案")

    with col2:
        st.header("3. 分析結果")
        
        if not df_yield_raw.empty and raw_search:
            # 這裡執行原本的搜尋邏輯 (簡化版示意)
            configs = parse_search_config(raw_search, chk_space)
            results = []
            
            # 模擬搜尋 (實際應套用您原本的 YieldSearchThread 邏輯)
            # 在 Streamlit 中直接跑迴圈即可，不需 Thread
            search_terms = [c['terms'][0] for c in configs] # 簡化取第一個 term
            
            # Pandas 字串搜尋
            mask = pd.Series([False] * len(df_yield_raw))
            for term in search_terms:
                # 這裡做一個非常簡單的全表文字搜尋示意
                mask |= df_yield_raw.astype(str).apply(lambda x: x.str.contains(term, case=False, na=False)).any(axis=1)
            
            df_result = df_yield_raw[mask].copy()
            df_result['MatchedKeyword'] = "Demo Match" # 實際應填入對應到的 keyword

            # --- 顯示表格 ---
            st.subheader("詳細數據")
            st.dataframe(df_result, use_container_width=True)
            
            # --- 顯示圖表 ---
            st.subheader("統計圖表")
            if not df_result.empty:
                chart_type = st.selectbox("圖表類型", ["Bar", "Line", "Pie"])
                
                # 找出數值欄位
                num_cols = df_result.select_dtypes(include=['float', 'int']).columns.tolist()
                cat_cols = df_result.select_dtypes(include=['object']).columns.tolist()
                
                c1, c2 = st.columns(2)
                x_axis = c1.selectbox("X 軸", cat_cols if cat_cols else df_result.columns)
                y_axis = c2.selectbox("Y 軸", num_cols if num_cols else df_result.columns)
                
                fig, ax = plt.subplots()
                # 簡單繪圖邏輯
                try:
                    grouped = df_result.groupby(x_axis)[y_axis].sum()
                    if chart_type == "Bar":
                        grouped.plot(kind='bar', ax=ax)
                    elif chart_type == "Line":
                        grouped.plot(kind='line', marker='o', ax=ax)
                    elif chart_type == "Pie":
                        grouped.plot(kind='pie', ax=ax)
                    st.pyplot(fig)
                except Exception as e:
                    st.warning(f"無法繪圖: {e}")

            # --- 下載按鈕 ---
            st.subheader("匯出")
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_result.to_excel(writer, index=False, sheet_name='Result')
            
            st.download_button(
                label="📥 下載 Excel 結果",
                data=buffer.getvalue(),
                file_name="yield_result.xlsx",
                mime="application/vnd.ms-excel"
            )
        elif not df_yield_raw.empty:
            st.info("請在左側輸入關鍵字以開始分析")
        else:
            st.info("請先上傳 Excel 檔案")

# ==========================================
# 4. BOM Search 模組
# ==========================================
with tab_bom:
    st.markdown("### BOM 交叉比對")
    # 類似 Yield 的結構：
    # 1. file_uploader (key="bom_files")
    # 2. text_area 輸入料號
    # 3. 按鈕 "開始搜尋"
    # 4. 呼叫您原本的 BomSearch 邏輯
    # 5. st.dataframe 顯示結果
    # 6. st.download_button 下載
    st.caption("功能結構同上，將原本的邏輯搬過來即可。")