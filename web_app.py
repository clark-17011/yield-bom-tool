import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import re
import unicodedata
import io
import matplotlib.pyplot as plt
import matplotlib
import platform

# ==========================================
# 0. 基礎設定與工具 (Shared Utilities)
# ==========================================

# 設定 Matplotlib 字型以免中文亂碼
def configure_chart_font():
    system_name = platform.system()
    if system_name == "Windows":
        plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'SimHei', 'Arial']
    elif system_name == "Darwin": # macOS
        plt.rcParams['font.sans-serif'] = ['Arial Unicode MS', 'PingFang TC', 'Heiti TC']
    else:
        # Linux / Streamlit Cloud 通常是 Linux
        plt.rcParams['font.sans-serif'] = ['WenQuanYi Zen Hei', 'DejaVu Sans']
    plt.rcParams['axes.unicode_minus'] = False 

configure_chart_font()

def normalize_str(x) -> str:
    if x is None: return ""
    s = str(x)
    s = unicodedata.normalize("NFKC", s)
    s = s.strip()
    s = re.sub(r"[^A-Za-z0-9]", "", s)
    return s.lower()

def parse_search_config(raw_text: str, is_space_or_mode: bool):
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
# 1. Yield Report 邏輯
# ==========================================

@st.cache_data(ttl=3600, show_spinner=False)
def load_yield_files(uploaded_files):
    """
    讀取 Yield Report Excel 檔案，建立原始資料與搜尋索引
    """
    raw_data = [] # 存放 (label, sheet_name, header, data_rows)
    row_texts = [] # 存放 (label, sheet_name, normalized_texts_list)
    
    total_sheets = 0
    
    for file in uploaded_files:
        label = file.name # 使用檔名作為標籤
        try:
            wb = openpyxl.load_workbook(file, read_only=True, data_only=True)
            for sheet_name in wb.sheetnames:
                try:
                    ws = wb[sheet_name]
                    rows = list(ws.iter_rows(values_only=True))
                    if not rows: continue
                    header = rows[0]
                    data_rows = rows[1:]
                    if not data_rows: continue
                    
                    # 儲存原始數據
                    raw_data.append({
                        "label": label,
                        "sheet": sheet_name,
                        "header": header,
                        "rows": data_rows
                    })
                    
                    # 建立搜尋索引 (正規化字串)
                    current_sheet_texts = []
                    for row in data_rows:
                        joined = "".join([str(c) if c is not None else "" for c in row])
                        current_sheet_texts.append(normalize_str(joined))
                    
                    row_texts.append({
                        "label": label,
                        "sheet": sheet_name,
                        "texts": current_sheet_texts
                    })
                    total_sheets += 1
                except: pass
            wb.close()
        except Exception as e:
            print(f"Error loading {label}: {e}")
            
    return raw_data, row_texts, total_sheets

def execute_yield_search(raw_data, row_texts, configs):
    """
    執行 Yield 搜尋
    """
    if not configs: return pd.DataFrame(), set()

    prepared_configs = []
    for cfg in configs:
        norm_terms = [normalize_str(t) for t in cfg['terms'] if t.strip()]
        if norm_terms:
            prepared_configs.append({'display': cfg['display'], 'terms': norm_terms})

    all_rows_data = []
    found_display_names = set()
    
    # 遍歷所有已讀取的 Sheet
    # row_texts 結構: [{"label":..., "sheet":..., "texts": [...]}, ...]
    for idx, sheet_info in enumerate(row_texts):
        label = sheet_info["label"]
        sheet_name = sheet_info["sheet"]
        sheet_norm_texts = sheet_info["texts"]
        
        # 取得對應的原始資料
        # raw_data 結構與 row_texts 索引對應 (因為是順序讀取的)
        header = raw_data[idx]["header"]
        all_rows = raw_data[idx]["rows"]
        
        # 處理 Header (重複名稱問題)
        unique_header = []
        seen_counts = {}
        for col in header:
            c_str = str(col).strip() if col is not None else ""
            if not c_str: c_str = "Unnamed"
            if c_str in seen_counts:
                seen_counts[c_str] += 1
                new_name = f"{c_str}.{seen_counts[c_str]}"
            else:
                seen_counts[c_str] = 0
                new_name = c_str
            unique_header.append(new_name)
            
        # 開始搜尋該 Sheet 的每一行
        for row_idx, row_str in enumerate(sheet_norm_texts):
            for cfg in prepared_configs:
                is_match = True
                for term in cfg['terms']:
                    if term not in row_str:
                        is_match = False
                        break
                
                if is_match:
                    found_display_names.add(cfg['display'])
                    original_row = all_rows[row_idx]
                    
                    row_dict = {
                        "MatchedKeyword": cfg['display'],
                        "SourceLabel": label,
                        "SheetName": sheet_name
                    }
                    
                    for h_idx, col_name in enumerate(unique_header):
                        val = original_row[h_idx] if h_idx < len(original_row) else None
                        row_dict[col_name] = val
                    
                    all_rows_data.append(row_dict)

    if all_rows_data:
        df_result = pd.DataFrame(all_rows_data)
        # 欄位排序
        cols = list(df_result.columns)
        sys_cols = ['MatchedKeyword', 'SourceLabel', 'SheetName']
        other_cols = [c for c in cols if c not in sys_cols]
        df_result = df_result[sys_cols + other_cols]
    else:
        df_result = pd.DataFrame()

    all_targets = set(c['display'] for c in prepared_configs)
    missing = all_targets - found_display_names
    
    return df_result, missing

# ==========================================
# 2. BOM Tool 邏輯
# ==========================================

PCB_VENDOR_MAP = {"P": "PRV", "S": "SCC", "U": "旭德", "H": "AKM", "D": "科佳"}
PCB_FINISH_MAP = {"G": "化金", "N": "鎳鈀金", "P": "OSP"}

BASE_OUTPUT_ORDER = [
    "MPN", "Device Name", "ASIC (簡化 BOM)", "Sensor (簡化 BOM)", "PCB 供應商",
    "錫膏", "PCB 簡化 BOM", "金屬殼", "電容 / 電組 / 電感", "磁珠", "防水膜 / 金屬網", "Coating"
]

def unify_key(s):
    if not isinstance(s, str): return str(s)
    s = re.sub(r"[\s\(\)\/]", "", s)
    return s.lower()

def format_value(val):
    if pd.isna(val) or val is None: return ""
    val_str = str(val).replace("\n", " ").replace("\r", " ")
    return " ".join(val_str.split())

@st.cache_data(ttl=3600, show_spinner=False)
def load_bom_files(uploaded_files):
    """
    讀取 BOM Excel 檔案
    """
    raw_data = [] 
    row_texts = [] 
    
    for file in uploaded_files:
        # 自動判斷 Label
        name = file.name
        label = "CPC" if "CPC" in name.upper() else ("HELE" if "HELE" in name.upper() else name)
        
        try:
            wb = openpyxl.load_workbook(file, read_only=True, data_only=True)
            for sheet_name in wb.sheetnames:
                try:
                    ws = wb[sheet_name]
                    rows = list(ws.iter_rows(values_only=True))
                    if not rows: continue
                    header = rows[0]
                    data_rows = rows[1:]
                    if not data_rows: continue
                    
                    raw_data.append({
                        "label": label,
                        "sheet": sheet_name,
                        "header": header,
                        "rows": data_rows
                    })
                    
                    current_sheet_texts = []
                    for row in data_rows:
                        joined = "".join([str(c) if c is not None else "" for c in row])
                        current_sheet_texts.append(normalize_str(joined))
                    
                    row_texts.append({
                        "label": label,
                        "sheet": sheet_name,
                        "texts": current_sheet_texts
                    })
                except: pass
            wb.close()
        except Exception as e:
            print(f"Error loading {label}: {e}")
            
    return raw_data, row_texts

def parse_pcb_details(green_bom):
    details = []
    if pd.isna(green_bom) or not green_bom: return details
    tokens = re.split(r'[\s\n]+', str(green_bom).strip())
    for token in tokens:
        token = token.strip()
        if len(token) >= 10:
            v_code = token[-4]
            f_code = token[-2]
            vendor = PCB_VENDOR_MAP.get(v_code, None)
            finish = PCB_FINISH_MAP.get(f_code, None)
            details.append({'code': token, 'vendor': vendor, 'finish': finish})
    return details

def get_col_by_keyword(header, keyword, exclude=None):
    target_key = unify_key(keyword)
    for idx, col in enumerate(header):
        col_str = str(col)
        col_key = unify_key(col_str)
        if target_key in col_key:
            if exclude and exclude in col_str: continue
            return idx
    return None

def execute_bom_search(raw_data, row_texts, terms_raw):
    """
    執行 BOM 搜尋
    """
    export_list = []
    found_terms = set()
    
    norm_terms_map = {t: normalize_str(t) for t in terms_raw}
    
    for term in terms_raw:
        n_term = norm_terms_map[term]
        if not n_term: continue
        
        # 遍歷所有 Sheet
        for idx, sheet_info in enumerate(row_texts):
            sheet_norm_texts = sheet_info["texts"]
            matched_row_idx = -1
            
            # 尋找匹配行
            for r_idx, row_str in enumerate(sheet_norm_texts):
                if n_term in row_str:
                    matched_row_idx = r_idx
                    break
            
            if matched_row_idx != -1:
                found_terms.add(term)
                
                # 提取資料
                header = raw_data[idx]["header"]
                row = raw_data[idx]["rows"][matched_row_idx]
                label = sheet_info["label"]
                
                row_data = extract_bom_data(header, row, label, term)
                export_list.append(row_data)

    missing_terms = set(terms_raw) - found_terms
    return export_list, missing_terms

def extract_bom_data(header, row, label, term):
    row_data = {
        "Search Term": term,
        "Source File": label
    }

    def get_val(idx):
        return row[idx] if idx is not None and idx < len(row) else None

    # PCB 分析邏輯
    pcb_green_idx = get_col_by_keyword(header, "PCB BOM (Green)")
    val_green = get_val(pcb_green_idx)
    pcb_details = parse_pcb_details(val_green)
    vendors = sorted(list(set([d['vendor'] for d in pcb_details if d['vendor']])))
    
    pcb_simple_idx = get_col_by_keyword(header, "PCB 簡化 BOM")
    if not pcb_simple_idx: pcb_simple_idx = get_col_by_keyword(header, "PCB 簡化BOM")
    val_simple = format_value(get_val(pcb_simple_idx))
    
    if not vendors and val_simple:
        for code, name in PCB_VENDOR_MAP.items():
            if name in val_simple: vendors.append(name)

    vendors_str = " / ".join(vendors) if vendors else ""

    # 欄位定義
    LAYOUT_PLAN = [
        ("MPN", "normal", ["MPN"]),
        ("Device Name", "normal", ["Device Name"]),
        ("ASIC (簡化 BOM)", "merge", ["ASIC", "ASIC 簡化BOM"]),
        ("Sensor (簡化 BOM)", "merge", ["Sensor ID", "Sensor 簡化BOM"]),
        ("PCB 供應商", "value", vendors_str),
        ("錫膏", "normal", ["錫膏"]),
        ("PCB 簡化 BOM", "normal", ["PCB 簡化BOM"]),
        ("PCB List", "pcb_list", pcb_details),
        ("金屬殼", "normal", ["金屬殼 BOM (Blue)"]),
        ("電容 / 電組 / 電感", "normal", ["Indigo"]),
        ("磁珠", "normal", ["磁珠"]),
        ("防水膜 / 金屬網", "normal", ["防水膜"]),
        ("Coating", "normal", ["Coating BOM (Black)"]),
    ]

    for label_text, type_, args in LAYOUT_PLAN:
        final_val = ""
        if type_ == "normal":
            col_key = args[0]
            exclude = "簡化" if "簡化" not in label_text and "簡化" not in col_key else None
            if "Indigo" in col_key: exclude = None
            idx = get_col_by_keyword(header, col_key, exclude=exclude)
            val = format_value(get_val(idx))
            if val: final_val = val
        elif type_ == "merge":
            main_key, sub_key = args
            idx_main = get_col_by_keyword(header, main_key, exclude="簡化")
            idx_sub = get_col_by_keyword(header, sub_key)
            val_main = format_value(get_val(idx_main))
            val_sub = format_value(get_val(idx_sub))
            if val_main and val_sub:
                final_val = f"{val_main} ({val_sub})" if val_main != val_sub else val_main
            elif val_main: final_val = val_main
            elif val_sub: final_val = val_sub
        elif type_ == "value":
            final_val = str(args)
        elif type_ == "pcb_list":
            for i, item in enumerate(args, 1):
                code = item['code']
                finish = item['finish']
                display = code + (f" ({finish})" if finish else "")
                row_data[f"PCB {i}"] = display
            continue

        if final_val: row_data[label_text] = final_val
        
    return row_data

# ==========================================
# 3. Streamlit UI 主程式
# ==========================================

st.set_page_config(page_title="Yield & BOM Tool", layout="wide", page_icon="📊")

st.title("📊 良率報表 & BOM 搜尋工具")
st.caption("支援 Excel 拖曳上傳 | 多檔案搜尋 | 自動彙整")

tab1, tab2 = st.tabs(["📈 Yield Analysis", "🔍 BOM Search"])

# --- TAB 1: Yield Analysis ---
with tab1:
    col_left, col_right = st.columns([1, 2])
    
    with col_left:
        st.subheader("1. 檔案與設定")
        yield_files = st.file_uploader(
            "上傳 Yield Report (Excel)", 
            type=['xlsx', 'xls'], 
            accept_multiple_files=True,
            key="yield_uploader"
        )
        
        # 搜尋關鍵字
        raw_search = st.text_area(
            "關鍵字 (空格=AND, 逗號=OR)", 
            height=100,
            placeholder="例如: [Device A], [Device B]\n或: Device A, Device B",
            help="使用 [] 可以精確比對，例如 [Device Name]。"
        )
        chk_space = st.checkbox("空白代表「或」(OR)", value=False, help="勾選後，空白分隔的字詞會變成多個搜尋目標。")
        
        btn_search_yield = st.button("開始搜尋", type="primary", key="btn_yield")

    # 處理資料讀取 (快取)
    yield_raw_data = []
    yield_row_texts = []
    
    if yield_files:
        with st.spinner("讀取檔案中..."):
            yield_raw_data, yield_row_texts, sheet_count = load_yield_files(yield_files)
        if sheet_count > 0:
            col_left.success(f"已載入 {len(yield_files)} 個檔案，共 {sheet_count} 個 Sheet")
    
    with col_right:
        st.subheader("2. 分析結果")
        
        if btn_search_yield and yield_files:
            if not raw_search.strip():
                st.warning("請輸入關鍵字")
            else:
                configs = parse_search_config(raw_search, chk_space)
                # 執行搜尋
                with st.spinner("搜尋運算中..."):
                    df_res, missing = execute_yield_search(yield_raw_data, yield_row_texts, configs)
                
                if missing:
                    st.error(f"未找到: {', '.join(missing)}")
                
                if not df_res.empty:
                    st.success(f"找到 {len(df_res)} 筆資料")
                    
                    # 存入 Session State 以便後續繪圖使用 (避免重整消失)
                    st.session_state['yield_result'] = df_res
                else:
                    st.info("無符合資料")
                    st.session_state['yield_result'] = pd.DataFrame()

        # 顯示結果 (如果有)
        if 'yield_result' in st.session_state and not st.session_state['yield_result'].empty:
            df_display = st.session_state['yield_result']
            
            # 分頁：數據 vs 圖表
            sub_t1, sub_t2 = st.tabs(["詳細數據", "統計圖表"])
            
            with sub_t1:
                st.dataframe(df_display, use_container_width=True)
                
                # Excel 下載
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    # 簡單格式化邏輯
                    unique_sheets = df_display['SheetName'].unique()
                    for s_name in unique_sheets:
                        sub_df = df_display[df_display['SheetName'] == s_name]
                        # 移除全空欄位
                        sub_df = sub_df.dropna(axis=1, how='all')
                        safe_name = str(s_name)[:30]
                        sub_df.to_excel(writer, sheet_name=safe_name, index=False)
                
                st.download_button(
                    label="📥 下載 Excel 結果",
                    data=buffer.getvalue(),
                    file_name="yield_result.xlsx",
                    mime="application/vnd.ms-excel"
                )

            with sub_t2:
                st.markdown("#### 繪圖設定")
                c1, c2, c3, c4 = st.columns(4)
                
                # 篩選數值與類別欄位
                num_cols = df_display.select_dtypes(include=['number']).columns.tolist()
                all_cols = df_display.columns.tolist()
                
                chart_type = c1.selectbox("圖表類型", ["Bar (長條)", "Line (折線)", "Pie (圓餅)", "Scatter (散佈)"])
                x_axis = c2.selectbox("X 軸 (分組)", all_cols, index=0)
                y_axis = c3.selectbox("Y 軸 (數值)", num_cols if num_cols else all_cols, index=0)
                agg_func = c4.selectbox("計算方式", ["Sum", "Mean", "Count", "Max"])
                
                if st.button("更新圖表"):
                    try:
                        fig, ax = plt.subplots(figsize=(8, 4))
                        
                        # 簡易資料處理
                        chart_df = df_display.copy()
                        # 嘗試轉數值
                        chart_df[y_axis] = pd.to_numeric(chart_df[y_axis], errors='coerce').fillna(0)
                        
                        if agg_func == "Count":
                            data = chart_df[x_axis].value_counts()
                        else:
                            agg_map = {"Sum": "sum", "Mean": "mean", "Max": "max"}
                            data = chart_df.groupby(x_axis)[y_axis].agg(agg_map[agg_func])
                        
                        if chart_type == "Bar (長條)":
                            data.plot(kind='bar', ax=ax, color='#007AFF')
                        elif chart_type == "Line (折線)":
                            data.plot(kind='line', marker='o', ax=ax, color='#007AFF')
                        elif chart_type == "Pie (圓餅)":
                            data.plot(kind='pie', autopct='%1.1f%%', ax=ax)
                            ax.set_ylabel('')
                        elif chart_type == "Scatter (散佈)":
                            ax.scatter(chart_df[x_axis], chart_df[y_axis], color='#007AFF')

                        ax.set_title(f"{agg_func} of {y_axis} by {x_axis}")
                        plt.tight_layout()
                        st.pyplot(fig)
                    except Exception as e:
                        st.error(f"繪圖失敗: {e}")


# --- TAB 2: BOM Search ---
with tab2:
    col_b_left, col_b_right = st.columns([1, 2])
    
    with col_b_left:
        st.subheader("1. BOM 檔案設定")
        bom_files = st.file_uploader(
            "上傳 BOM 對應表 (Excel)", 
            type=['xlsx', 'xls'], 
            accept_multiple_files=True,
            key="bom_uploader"
        )
        
        st.info("系統會自動依檔名辨識 Label (HELE/CPC/其他)")
        
        bom_input = st.text_area("輸入料號 (支援 Excel 整欄貼上)", height=150)
        chk_bom_space = st.checkbox("空白分隔 (Split by space)", value=False, key="bom_space")
        
        btn_search_bom = st.button("開始比對", type="primary", key="btn_bom")

    # 載入 BOM (快取)
    bom_raw_data = []
    bom_row_texts = []
    if bom_files:
        with st.spinner("建立 BOM 索引..."):
            bom_raw_data, bom_row_texts = load_bom_files(bom_files)
        if bom_raw_data:
            col_b_left.success(f"已載入 {len(bom_files)} 個 BOM 檔")

    with col_b_right:
        st.subheader("2. 比對結果")
        
        if btn_search_bom and bom_files:
            if not bom_input.strip():
                st.warning("請輸入料號")
            else:
                # 解析輸入
                sep = r'[,\n\r\t\s，;]+' if chk_bom_space else r'[,\n\r\t，;]+'
                terms_raw = re.split(sep, bom_input)
                clean_terms = [t.strip() for t in terms_raw if t.strip()]
                clean_terms = list(dict.fromkeys(clean_terms)) # 去重
                
                st.write(f"搜尋 {len(clean_terms)} 筆料號...")
                
                with st.spinner("比對中..."):
                    res_list, missing_terms = execute_bom_search(bom_raw_data, bom_row_texts, clean_terms)
                
                if missing_terms:
                    st.error(f"⚠️ 未找到 ({len(missing_terms)}): {', '.join(missing_terms)}")
                else:
                    st.success("✅ 全部找到！")
                
                if res_list:
                    # 整理 DataFrame
                    df_bom = pd.DataFrame(res_list)
                    
                    # 動態排序與 PCB 欄位處理
                    max_pcb = 0
                    for row in res_list:
                        for k in row.keys():
                            if k.startswith("PCB ") and k[4:].isdigit():
                                max_pcb = max(max_pcb, int(k[4:]))
                    
                    final_headers = ["Search Term", "Source File"]
                    # 嘗試插入 PCB 欄位
                    base_order = list(BASE_OUTPUT_ORDER) # copy
                    try: 
                        ins_idx = base_order.index("PCB 簡化 BOM") + 1
                    except: 
                        ins_idx = len(base_order)
                        
                    pcb_cols = [f"PCB {i}" for i in range(1, max_pcb + 1)]
                    
                    final_cols = final_headers + base_order[:ins_idx] + pcb_cols + base_order[ins_idx:]
                    
                    # Reindex
                    df_bom = df_bom.reindex(columns=final_cols)
                    
                    st.dataframe(df_bom, use_container_width=True)
                    
                    # 匯出
                    buffer_bom = io.BytesIO()
                    with pd.ExcelWriter(buffer_bom, engine='openpyxl') as writer:
                        df_bom.to_excel(writer, index=False, sheet_name='Search Results')
                        
                        # 自動調整欄寬 (簡易版)
                        ws = writer.sheets['Search Results']
                        for column in ws.columns:
                            col_letter = get_column_letter(column[0].column)
                            ws.column_dimensions[col_letter].width = 20
                            
                    st.download_button(
                        label="📥 下載 BOM 結果",
                        data=buffer_bom.getvalue(),
                        file_name="BOM_Result.xlsx",
                        mime="application/vnd.ms-excel"
                    )
