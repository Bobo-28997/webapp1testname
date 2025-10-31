# =====================================
# Streamlit Web App: 模拟Project：人事用合同记录表自动审核
# (V2: 缓存优化版)
# =====================================

import streamlit as st
import pandas as pd
import time
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
from io import BytesIO

# =====================================
# 🧰 工具函数区 (不变)
# =====================================

def normalize_contract_key(series: pd.Series) -> pd.Series:
    s = series.astype(str)
    s = s.str.replace(r"\.0$", "", regex=True) 
    s = s.str.strip()
    s = s.str.upper() 
    s = s.str.replace('－', '-', regex=False)
    s = s.str.replace(r'\s+', '', regex=True)
    return s

def find_file(files_list, keyword):
    """
    注意：此版本接收 Streamlit 的 UploadedFile 对象列表
    """
    for f in files_list:
        if keyword in f.name:
            return f
    # (修改：返回 None 而不是 raise Error, 允许缓存函数处理)
    return None 

def normalize_colname(c): return str(c).strip().lower()

def find_col(df, keyword, exact=False):
    key = keyword.strip().lower()
    for col in df.columns:
        cname = normalize_colname(col)
        if (exact and cname == key) or (not exact and key in cname):
            return col
    return None

def find_sheet(xls, keyword):
    for s in xls.sheet_names:
        if keyword in s:
            return s
    raise ValueError(f"❌ 未找到包含关键词「{keyword}」的sheet: {keyword}")

def normalize_num(val):
    if pd.isna(val): return None
    s = str(val).replace(",", "").strip()
    if s in ["", "-", "nan"]: return None
    try:
        if "%" in s: return float(s.replace("%", "")) / 100
        return float(s)
    except ValueError:
        return s

def prepare_ref_df(ref_df, mapping, prefix):
    contract_col = find_col(ref_df, "合同") 
    if not contract_col:
        st.warning(f"⚠️ 在 {prefix} 参考表中未找到'合同'列，跳过此数据源。")
        return pd.DataFrame(columns=['__KEY__'])
        
    std_df = pd.DataFrame()
    std_df['__KEY__'] = normalize_contract_key(ref_df[contract_col])
    
    for main_kw, ref_kw in mapping.items():
        exact = (main_kw == "城市经理")
        ref_col_name = find_col(ref_df, ref_kw, exact=exact)
        
        if ref_col_name:
            s_ref_raw = ref_df[ref_col_name]
            if prefix == 'fk' and main_kw == '租赁期限':
                s_ref_transformed = pd.to_numeric(s_ref_raw, errors='coerce') * 12
                std_df[f'ref_{prefix}_{main_kw}'] = s_ref_transformed
            else:
                std_df[f'ref_{prefix}_{main_kw}'] = s_ref_raw
        else:
            st.warning(f"⚠️ 在 {prefix} 参考表中未找到列 (main: '{main_kw}', ref: '{ref_kw}')")

    std_df = std_df.drop_duplicates(subset=['__KEY__'], keep='first')
    return std_df

def compare_series_vec(s_main, s_ref, main_kw):
    merge_failed_mask = s_ref.isna() 
    main_is_na = pd.isna(s_main) | (s_main.astype(str).str.strip().isin(["", "nan", "None"]))
    ref_is_na = pd.isna(s_ref) | (s_ref.astype(str).str.strip().isin(["", "nan", "None"]))
    both_are_na = main_is_na & ref_is_na
    
    if any(k in main_kw for k in ["日期", "时间"]):
        d_main = pd.to_datetime(s_main, errors='coerce')
        d_ref = pd.to_datetime(s_ref, errors='coerce')
        valid_dates_mask = d_main.notna() & d_ref.notna()
        date_diff_mask = (d_main.dt.date != d_ref.dt.date)
        errors = valid_dates_mask & date_diff_mask
    else:
        s_main_norm = s_main.apply(normalize_num)
        s_ref_norm = s_ref.apply(normalize_num)
        main_is_na_norm = pd.isna(s_main_norm) | (s_main_norm.astype(str).str.strip().isin(["", "nan", "None"]))
        ref_is_na_norm = pd.isna(s_ref_norm) | (s_ref_norm.astype(str).str.strip().isin(["", "nan", "None"]))
        both_are_na_norm = main_is_na_norm & ref_is_na_norm
        is_num_main = s_main_norm.apply(lambda x: isinstance(x, (int, float)))
        is_num_ref = s_ref_norm.apply(lambda x: isinstance(x, (int, float)))
        both_are_num = is_num_main & is_num_ref
        errors = pd.Series(False, index=s_main.index)
        
        if both_are_num.any():
            num_main = s_main_norm[both_are_num].fillna(0)
            num_ref = s_ref_norm[both_are_num].fillna(0)
            diff = (num_main - num_ref).abs()
            
            if main_kw == "保证金比例":
                num_errors = (diff > 0.00500001)
            elif "租赁期限" in main_kw:
                num_errors = (diff >= 1.0) 
            else:
                num_errors = (diff > 1e-6)
            
            errors.loc[both_are_num] = num_errors

        not_num_mask = ~both_are_num
        if not_num_mask.any():
            str_main = s_main_norm[not_num_mask].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
            str_ref = s_ref_norm[not_num_mask].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
            str_errors = (str_main != str_ref)
            errors.loc[not_num_mask] = str_errors
            
        errors = errors & ~both_are_na_norm

    final_errors = errors & ~both_are_na
    lookup_failure_mask = merge_failed_mask & ~main_is_na
    final_errors = final_errors & ~lookup_failure_mask
    return final_errors

# --- VVVV (【修改点 1】: 将 reboot_app1 移到这里) VVVV ---
def reboot_app1():
    """
    一个用于“重新上传”按钮的回调函数。
    它会清除所有缓存和 session 状态，让 app 恢复到初始状态。
    """
    # 1. 清除函数缓存
    run_full_audit.clear()
    
    # 2. 定义需要从 session_state 中清除的 key
    keys_to_delete = ['audit_run_app1', 'uploader_app1'] # <--- 'uploader_app1' 是关键
    
    # 3. 循环删除
    for key in keys_to_delete:
        if key in st.session_state:
            del st.session_state[key]
    
    # (不需要 st.rerun(), on_click 会自动触发)
# --- ^^^^ (修改结束) ^^^^ ---


# =====================================
# 🧮 (修改) 单sheet检查函数 - 现在返回文件
# =====================================
def check_one_sheet(sheet_keyword, main_file, ref_dfs_std_dict, mappings_all):
    """
    (已修改)
    1. 移除 st.download_button
    2. 返回 (stats, files_dict)
    """
    start_time = time.time()
    xls_main = pd.ExcelFile(main_file)

    try:
        target_sheet = find_sheet(xls_main, sheet_keyword)
    except ValueError:
        st.warning(f"⚠️ 未找到包含「{sheet_keyword}」的sheet，跳过。")
        return (0, None, 0, set()), {} # 返回 (stats, files_dict)

    try:
        main_df = pd.read_excel(xls_main, sheet_name=target_sheet, header=1)
    except Exception as e:
        st.error(f"❌ 读取「{sheet_keyword}」时出错: {e}")
        return (0, None, 0, set()), {}
        
    if main_df.empty:
        st.warning(f"⚠️ 「{sheet_keyword}」为空，跳过。")
        return (0, None, 0, set()), {}

    contract_col_main = find_col(main_df, "合同")
    if not contract_col_main:
        st.error(f"❌ 在「{sheet_keyword}」中未找到合同列。")
        return (0, None, 0, set()), {}

    # --- 写入临时 BytesIO 而不是磁盘 ---
    temp_output_stream = BytesIO()
    with pd.ExcelWriter(temp_output_stream, engine='openpyxl') as writer:
        # (保留原始空行)
        empty_row = pd.DataFrame([[""] * len(main_df.columns)], columns=main_df.columns)
        pd.concat([empty_row, main_df], ignore_index=True).to_excel(writer, index=False, sheet_name=target_sheet)
    temp_output_stream.seek(0)
    
    wb = load_workbook(temp_output_stream)
    ws = wb[target_sheet] # 确保激活正确的 sheet
    
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    main_df['__ROW_IDX__'] = main_df.index
    main_df['__KEY__'] = normalize_contract_key(main_df[contract_col_main])
    contracts_seen = set(main_df['__KEY__'].dropna())

    merged_df = main_df.copy()
    for prefix, std_df in ref_dfs_std_dict.items():
        if not std_df.empty:
            merged_df = pd.merge(merged_df, std_df, on='__KEY__', how='left')
    
    total_errors = 0
    skip_city_manager = [0]
    errors_locations = set()
    row_has_error = pd.Series(False, index=merged_df.index) 

    progress = st.progress(0)
    status = st.empty()

    total_comparisons = sum(len(m[0]) for m in mappings_all.values())
    current_comparison = 0

    for prefix, (mapping, std_df) in mappings_all.items():
        if std_df.empty:
            current_comparison += len(mapping)
            continue
            
        for main_kw, ref_kw in mapping.items():
            current_comparison += 1
            status.text(f"检查「{sheet_keyword}」: {prefix} - {main_kw}...")
            
            exact = (main_kw == "城市经理")
            main_col = find_col(main_df, main_kw, exact=exact)
            ref_col = f'ref_{prefix}_{main_kw}'

            if not main_col or ref_col not in merged_df.columns:
                continue 

            s_main = merged_df[main_col]
            s_ref = merged_df[ref_col]

            skip_mask = pd.Series(False, index=merged_df.index)
            if main_kw == "城市经理":
                na_strings = ["", "-", "nan", "none", "null"]
                skip_mask = pd.isna(s_ref) | s_ref.astype(str).str.strip().isin(na_strings)
                skip_city_manager[0] += skip_mask.sum()
            
            errors_mask = compare_series_vec(s_main, s_ref, main_kw)
            final_errors_mask = errors_mask & ~skip_mask
            
            if final_errors_mask.any():
                total_errors += final_errors_mask.sum()
                row_has_error |= final_errors_mask
                bad_indices = merged_df[final_errors_mask]['__ROW_IDX__']
                for idx in bad_indices:
                    errors_locations.add((idx, main_col))
            
            progress.progress(current_comparison / total_comparisons)

    status.text(f"「{sheet_keyword}」比对完成，正在生成标注文件...")

    original_cols_list = list(main_df.drop(columns=['__ROW_IDX__', '__KEY__']).columns)
    col_name_to_idx = {name: i + 1 for i, name in enumerate(original_cols_list)}

    for (row_idx, col_name) in errors_locations:
        if col_name in col_name_to_idx:
            ws.cell(row_idx + 3, col_name_to_idx[col_name]).fill = red_fill

    if contract_col_main in col_name_to_idx:
        contract_col_excel_idx = col_name_to_idx[contract_col_main]
        error_row_indices = merged_df[row_has_error]['__ROW_IDX__']
        for row_idx in error_row_indices:
            ws.cell(row_idx + 3, contract_col_excel_idx).fill = yellow_fill

    # --- (10. 修改为 return 文件) ---
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    
    files_to_save = {
        "full_report": (f"月重卡_{sheet_keyword}_审核标注版.xlsx", output),
        "error_report": (None, None)
    }
    output_errors_only = None

    # --- (11. 修改为 return 文件) ---
    if row_has_error.any():
        try:
            df_errors_only = merged_df.loc[row_has_error, original_cols_list].copy()
            original_indices_with_error = merged_df.loc[row_has_error, '__ROW_IDX__']
            original_idx_to_new_excel_row = {
                original_idx: new_row_num 
                for new_row_num, original_idx in enumerate(original_indices_with_error, start=2)
            }
            wb_errors = Workbook()
            ws_errors = wb_errors.active
            for r in dataframe_to_rows(df_errors_only, index=False, header=True):
                ws_errors.append(r)
            for (original_row_idx, col_name) in errors_locations:
                if original_row_idx in original_idx_to_new_excel_row:
                    new_row = original_idx_to_new_excel_row[original_row_idx]
                    if col_name in col_name_to_idx:
                        new_col = col_name_to_idx[col_name]
                        ws_errors.cell(row=new_row, column=new_col).fill = red_fill
            
            output_errors_only = BytesIO()
            wb_errors.save(output_errors_only)
            output_errors_only.seek(0)
            
            files_to_save["error_report"] = (f"月重卡_{sheet_keyword}_仅错误行_标红.xlsx", output_errors_only)
            
        except Exception as e:
            st.error(f"❌ 生成“仅错误行”文件时出错: {e}")
            
    elapsed = time.time() - start_time
    st.success(f"✅ {sheet_keyword} 检查完成，共 {total_errors} 处错误，用时 {elapsed:.2f} 秒。")
    
    stats = (total_errors, elapsed, skip_city_manager[0], contracts_seen)
    return stats, files_to_save

# =====================================
# 🕵️ (新) 漏填检查函数
# =====================================
def run_leaky_check(zd_df, contract_col_zd, contracts_seen_all_sheets):
    """
    执行漏填检查并返回 BytesIO 文件。
    """
    st.info("ℹ️ 正在执行漏填检查...")
    files_to_save = {}
    
    field_contracts = zd_df[contract_col_zd].dropna().astype(str).str.strip()
    col_car_manager = find_col(zd_df, "是否车管家", exact=True)
    col_bonus_type = find_col(zd_df, "提成类型", exact=True)

    missing_contracts_mask = (~field_contracts.isin(contracts_seen_all_sheets))

    if col_car_manager:
        missing_contracts_mask &= ~(zd_df[col_car_manager].astype(str).str.strip().str.lower() == "是")
    if col_bonus_type:
        missing_contracts_mask &= ~(
            zd_df[col_bonus_type].astype(str).str.strip().isin(["联合租赁", "驻店"])
        )

    zd_df_missing = zd_df.copy()
    zd_df_missing["漏填检查"] = ""
    zd_df_missing.loc[missing_contracts_mask, "漏填检查"] = "❗ 漏填"
    漏填合同数 = zd_df_missing["漏填检查"].eq("❗ 漏填").sum()
    st.warning(f"⚠️ 共发现 {漏填合同数} 个合同在记录表中未出现（已排除车管家、联合租赁、驻店）")

    # --- 导出文件逻辑 ---
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    # 文件1：全字段表
    wb = Workbook()
    ws = wb.active
    for r in dataframe_to_rows(zd_df_missing, index=False, header=True):
        ws.append(r)
    
    check_col_idx = -1
    for c_idx, cell in enumerate(ws[1], 1): # 遍历表头
        if cell.value == "漏填检查":
            check_col_idx = c_idx
            break
            
    if check_col_idx > 0:
        for row in ws.iter_rows(min_row=2, min_col=check_col_idx, max_col=check_col_idx):
            cell = row[0]
            if cell.value == "❗ 漏填":
                cell.fill = yellow_fill

    output_all = BytesIO()
    wb.save(output_all)
    output_all.seek(0)
    files_to_save["leaky_full"] = ("字段表_漏填标注版.xlsx", output_all)

    # 文件2：仅漏填
    zd_df_only_missing = zd_df_missing[zd_df_missing["漏填检查"] == "❗ 漏填"].copy()
    if not zd_df_only_missing.empty:
        wb2 = Workbook()
        ws2 = wb2.active
        for r in dataframe_to_rows(zd_df_only_missing, index=False, header=True):
            ws2.append(r)
        
        check_col_idx_2 = -1
        for c_idx, cell in enumerate(ws2[1], 1):
            if cell.value == "漏填检查":
                check_col_idx_2 = c_idx
                break
        if check_col_idx_2 > 0:
            for row in ws2.iter_rows(min_row=2, min_col=check_col_idx_2, max_col=check_col_idx_2):
                if row[0].value == "❗ 漏填":
                    row[0].fill = yellow_fill

        out2 = BytesIO()
        wb2.save(out2)
        out2.seek(0)
        files_to_save["leaky_only"] = ("字段表_仅漏填.xlsx", out2)
    else:
        files_to_save["leaky_only"] = (None, None)

    return 漏填合同数, files_to_save


# =====================================
# 🚀 (新) 缓存的审核主函数
# =====================================
@st.cache_data(show_spinner="正在执行审核，请稍候...")
def run_full_audit(_uploaded_files):
    """
    执行所有文件读取、预处理和检查，并返回所有结果。
    此函数被缓存，只有在 _uploaded_files 更改时才会重新运行。
    """
    
    # --- 1. 📖 文件读取 ---
    main_file = find_file(_uploaded_files, "月重卡")
    fk_file = find_file(_uploaded_files, "放款明细")
    zd_file = find_file(_uploaded_files, "字段")
    ec_file = find_file(_uploaded_files, "二次明细")
    
    if not all([main_file, fk_file, zd_file, ec_file]):
        raise FileNotFoundError("未能找到所有必需的文件（月重卡、放款明细、字段、二次明细）。")

    st.info("ℹ️ 正在读取并预处理参考文件...")
    fk_df = pd.read_excel(pd.ExcelFile(fk_file), sheet_name=find_sheet(pd.ExcelFile(fk_file), "威田"))
    zd_df = pd.read_excel(pd.ExcelFile(zd_file), sheet_name=find_sheet(pd.ExcelFile(zd_file), "重卡"))
    ec_df = pd.read_excel(ec_file)

    contract_col_zd = find_col(zd_df, "合同") # 仅为漏填检查

    # --- 2. 🗺️ 映射表 ---
    mapping_fk = {
        "授信方": "授信方",
        "租赁本金": "租赁本金", 
        "租赁期限": "租赁期限",
        "挂车台数": "挂车数量",
        "起租收益率": "XIRR"
    }
    mapping_zd = {"保证金比例": "保证金比例_2", "项目提报人": "提报", "起租时间": "起租日_商", "客户经理": "客户经理_资产", "所属省区": "区域", "主车台数": "主车台数", "城市经理": "城市经理"}
    mapping_ec = {"二次时间": "出本流程时间"}

    mappings_all = {
        'fk': (mapping_fk, None), # DF 稍后填充
        'zd': (mapping_zd, None),
        'ec': (mapping_ec, None),
    }

    # --- 3. 🚀 预处理 ---
    fk_std = prepare_ref_df(fk_df, mapping_fk, 'fk')
    zd_std = prepare_ref_df(zd_df, mapping_zd, 'zd')
    ec_std = prepare_ref_df(ec_df, mapping_ec, 'ec')

    ref_dfs_std_dict = {
        'fk': fk_std,
        'zd': zd_std,
        'ec': ec_std
    }
    
    mappings_all['fk'] = (mapping_fk, fk_std)
    mappings_all['zd'] = (mapping_zd, zd_std)
    mappings_all['ec'] = (mapping_ec, ec_std)
    st.success("✅ 参考数据预处理完成。")
    
    # --- 4. 🧾 多sheet循环 ---
    sheet_keywords = ["二次", "部分担保", "随州", "驻店客户"]
    total_all = elapsed_all = skip_total = 0
    contracts_seen_all_sheets = set()
    
    all_generated_files = [] # 存储所有 (文件名, BytesIO) 元组

    for kw in sheet_keywords:
        stats, files_dict = check_one_sheet(kw, main_file, ref_dfs_std_dict, mappings_all)
        (count, used, skipped, seen) = stats
        
        all_generated_files.append(files_dict["full_report"])
        if files_dict["error_report"][0] is not None:
            all_generated_files.append(files_dict["error_report"])
        
        total_all += count
        elapsed_all += used or 0
        skip_total += skipped
        contracts_seen_all_sheets.update(seen)

    # --- 5. 🕵️ 漏填检查 ---
    漏填合同数, leaky_files_dict = run_leaky_check(
        zd_df, contract_col_zd, contracts_seen_all_sheets
    )
    
    all_generated_files.append(leaky_files_dict["leaky_full"])
    if leaky_files_dict["leaky_only"][0] is not None:
        all_generated_files.append(leaky_files_dict["leaky_only"])

    # --- 6. 返回所有结果 ---
    stats_summary = {
        "total_all": total_all,
        "elapsed_all": elapsed_all,
        "漏填合同数": 漏填合同数
    }
    
    return all_generated_files, stats_summary

# =====================================
# 🏁 应用标题与说明 (重构版)
# =====================================
st.title("📊 模拟人事用薪资计算表自动审核系统-1重卡")
st.image("image/app1(1).png")

# =====================================
# 📂 上传文件区：要求上传 4 个 xlsx 文件 (重构版)
# =====================================
uploaded_files = st.file_uploader(
    "请上传文件名中包含以下字段的文件：月重卡、放款明细、字段、二次明细。最后誊写，需检的表为文件名包含‘月重卡’字段的表。",
    type="xlsx",
    accept_multiple_files=True,
    key="uploader_app1" # <--- VVVV 新增此 key VVVV
)

# =====================================
# 🚀 (新) 主执行逻辑 (V2 - 缓存修复版)
# =====================================

def reboot_app1():
    """
    一个用于“重新上传”按钮的回调函数。
    它会清除所有缓存和 session 状态，让 app 恢复到初始状态。
    """
    # 1. 清除函数缓存
    run_full_audit.clear()
    
    # 2. 定义需要从 session_state 中清除的 key
    keys_to_delete = ['audit_run_app1', 'uploader_app1'] # <--- 'uploader_app1' 是关键
    
    # 3. 循环删除
    for key in keys_to_delete:
        if key in st.session_state:
            del st.session_state[key]
    
    # (不需要 st.rerun(), on_click 会自动触发)

# -------------------------------------
# 主程序开始
# -------------------------------------

if not uploaded_files or len(uploaded_files) < 4:
    st.warning("⚠️ 请上传所有 4 个文件后继续")
    # (健壮性：如果用户清空了文件, 也重置审核状态)
    if 'audit_run_app1' in st.session_state:
        del st.session_state.audit_run_app1
    st.stop()
else:
    st.success("✅ 文件上传完成")
    
    # (新) 创建两列用于放置按钮
    col1, col2 = st.columns(2)
    
    with col1:
        # (新) “开始审核”按钮
        if st.button("🚀 开始审核", type="primary", use_container_width=True):
            # 将运行状态存入 session state
            st.session_state.audit_run_app1 = True 
            # (点击按钮会自动 rerun)
    
    with col2:
        # --- VVVV (【核心修改】按钮逻辑) VVVV ---
        # (新) “重新上传 (Reboot)” 按钮
        st.button("🔄 重新上传 (Reboot)", on_click=reboot_app1, use_container_width=True)
        # --- ^^^^ (修改结束) ^^^^ ---

    # (新) 只有在 "开始审核" 被点击后才执行
    if 'audit_run_app1' in st.session_state and st.session_state.audit_run_app1:
        try:
            # 1. (新) 调用缓存的审核函数
            all_files, stats = run_full_audit(uploaded_files)

            # 2. (新) 显示统计摘要
            st.success(f"🎯 全部审核完成，共 {stats['total_all']} 处错误，总耗时 {stats['elapsed_all']:.2f} 秒。")
            st.warning(f"⚠️ 共发现 {stats['漏填合同数']} 个合同在记录表中未出现（已排除车管家、联合租赁、驻店）")
            
            # 3. (新) 显示所有下载按钮
            st.divider()
            st.subheader("📤 下载审核结果文件")
            
            # (新) 将下载按钮放入两列
            cols_dl = st.columns(2) # (使用新变量名, 避免与 col1/col2 混淆)
            col_idx = 0
            
            for (filename, data) in all_files:
                if filename and data: # 确保文件名和数据都存在
                    with cols_dl[col_idx % 2]:
                        st.download_button(
                            label=f"📥 下载 {filename}",
                            data=data,
                            file_name=filename,
                            key=f"download_btn_{filename}" # 使用唯一key
                        )
                    col_idx += 1
            
            st.success("✅ 所有检查、标注与导出完成！")
            
        except FileNotFoundError as e:
            st.error(f"❌ 文件查找失败: {e}")
            st.info("请确保您上传了所有必需的文件（月重卡、放款明细、字段、二次明细）。")
            st.session_state.audit_run_app1 = False # 出错时重置状态
        except ValueError as e:
            st.error(f"❌ Sheet 查找失败: {e}")
            st.info(f"请确保您的Excel文件包含必需的 sheet（例如 '威田', '重卡'）。错误详情: {e}")
            st.session_state.audit_run_app1 = False
        except Exception as e:
            st.error(f"❌ 审核过程中发生未知错误: {e}")
            st.exception(e)
            st.session_state.audit_run_app1 = False
