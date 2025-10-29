# =====================================
# Streamlit Web App: 合同记录表自动审核（向量化 + 四Sheet + 漏填检查 + 驻店客户版）
# =====================================

import streamlit as st
import pandas as pd
import time
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill
from io import BytesIO

# =====================================
# 🏁 应用标题
# =====================================
st.title("📊 模拟实际运用环境Project：人事用合同记录表自动审核系统（向量化 + 四Sheet + 漏填检查 + 驻店客户版）")

# =====================================
# 📂 上传文件
# =====================================
uploaded_files = st.file_uploader(
    "请上传以下文件：记录表、放款明细、字段、二次明细、重卡数据",
    type="xlsx",
    accept_multiple_files=True
)
if not uploaded_files or len(uploaded_files) < 5:
    st.warning("⚠️ 请上传所有 5 个文件后继续")
    st.stop()
else:
    st.success("✅ 文件上传完成")

# =====================================
# 🧰 工具函数
# =====================================
def find_file(files_list, keyword):
    for f in files_list:
        if keyword in f.name:
            return f
    raise FileNotFoundError(f"❌ 未找到包含关键词「{keyword}」的文件")

def normalize_colname(c):
    return str(c).strip().lower()

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
    raise ValueError(f"❌ 未找到包含关键词「{keyword}」的sheet")

def normalize_num(val):
    if pd.isna(val): return None
    s = str(val).replace(",", "").strip()
    if s in ["", "-", "nan"]: return None
    try:
        if "%" in s: return float(s.replace("%", "")) / 100
        return float(s)
    except ValueError:
        return s

def same_date_ymd(a, b):
    try:
        da = pd.to_datetime(a, errors='coerce')
        db = pd.to_datetime(b, errors='coerce')
        if pd.isna(da) or pd.isna(db): return False
        return (da.year, da.month, da.day) == (db.year, db.month, db.day)
    except Exception:
        return False

def read_excel_clean(file, sheet_name=None, header=0):
    df = pd.read_excel(file, sheet_name=sheet_name, header=header)
    df.columns = [str(c).strip() for c in df.columns]
    return df

# =====================================
# 📖 文件读取
# =====================================
main_file = find_file(uploaded_files, "记录表")
fk_file = find_file(uploaded_files, "放款明细")
zd_file = find_file(uploaded_files, "字段")
ec_file = find_file(uploaded_files, "二次明细")
zk_file = find_file(uploaded_files, "重卡数据")

xls_main = pd.ExcelFile(main_file)
sheet_keywords = ["二次", "部分担保", "随州", "驻店客户"]

fk_df = read_excel_clean(fk_file, sheet_name=find_sheet(pd.ExcelFile(fk_file), "本司"))
zd_df = read_excel_clean(zd_file, sheet_name=find_sheet(pd.ExcelFile(zd_file), "重卡"))
ec_df = read_excel_clean(ec_file)
zk_df = read_excel_clean(zk_file)

# =====================================
# 📌 合同列
# =====================================
contract_col_fk = find_col(fk_df, "合同", exact=True)
contract_col_zd = find_col(zd_df, "合同", exact=True)
contract_col_ec = find_col(ec_df, "合同", exact=True)
contract_col_zk = find_col(zk_df, "合同", exact=True)

# =====================================
# 🔗 字段映射
# =====================================
mapping_fk = {"授信方":"授信", "租赁本金":"本金", "租赁期限月":"租赁期限月",
              "客户经理":"客户经理", "起租收益率":"收益率", "主车台数":"主车台数", "挂车台数":"挂车台数"}
mapping_zd = {"保证金比例":"保证金比例_2", "项目提报人":"提报", "起租时间":"起租日_商",
              "租赁期限月":"总期数_商_资产", "起租收益率":"XIRR_商_起租", "所属省区":"区域", "城市经理":"城市经理"}
mapping_ec = {"二次时间":"出本流程时间"}
mapping_zk = {"授信方":"授信方"}

# =====================================
# ⚡ 向量化比对函数（支持 exact 客户经理/城市经理 + 日期 + 保证金比例容差）
# =====================================
def compare_fields_vectorized(main_df, ref_df, main_contract_col, ref_contract_col, mapping_dict, tolerance_dict=None):
    tolerance_dict = tolerance_dict or {}
    main_df_clean = main_df.copy()
    main_df_clean[main_contract_col] = main_df_clean[main_contract_col].astype(str).str.strip()
    ref_df_clean = ref_df.copy()
    ref_df_clean[ref_contract_col] = ref_df_clean[ref_contract_col].astype(str).str.strip()

    ref_cols_needed = [ref_contract_col] + list(mapping_dict.values())
    missing_cols = [c for c in ref_cols_needed if c not in ref_df_clean.columns]
    if missing_cols:
        st.error(f"❌ 参考表缺少列: {missing_cols}")
        mask_empty = pd.DataFrame(False, index=main_df.index, columns=mapping_dict.keys())
        return main_df_clean.copy(), mask_empty

    ref_sub = ref_df_clean[ref_cols_needed]
    merged = main_df_clean.merge(ref_sub, how="left", left_on=main_contract_col, right_on=ref_contract_col, suffixes=("", "_ref"))
    mask = pd.DataFrame(False, index=merged.index, columns=mapping_dict.keys())

    for main_col, ref_col in mapping_dict.items():
        if main_col not in merged.columns: continue
        main_vals = merged[main_col]
        ref_vals = merged[f"{ref_col}_ref"]
        is_date_col = any(k in main_col for k in ["日期","时间"]) or any(k in ref_col for k in ["日期","时间"])
        tol = tolerance_dict.get(main_col, 0)
        exact_match = main_col in ["客户经理","城市经理"]

        # 向量化比较
        if is_date_col:
            main_dt = pd.to_datetime(main_vals, errors='coerce').dt.normalize()
            ref_dt = pd.to_datetime(ref_vals, errors='coerce').dt.normalize()
            mask[main_col] = ~(main_dt.eq(ref_dt))
        else:
            main_num = main_vals.apply(normalize_num)
            ref_num = ref_vals.apply(normalize_num)
            num_mask = (main_num.notna() & ref_num.notna()) & ((main_num - ref_num).abs() > tol)
            text_mask = (~main_num.eq(ref_num)) & (~num_mask)
            if exact_match:
                text_mask = ~main_vals.astype(str).str.strip().eq(ref_vals.astype(str).str.strip())
            nan_mask = (main_num.isna() & ref_num.notna()) | (main_num.notna() & ref_num.isna())
            mask[main_col] = num_mask | text_mask | nan_mask

    return merged, mask

# =====================================
# 🧮 单 sheet 检查
# =====================================
def check_one_sheet(sheet_keyword):
    start_time = time.time()
    main_df = read_excel_clean(main_file, sheet_name=find_sheet(xls_main, sheet_keyword), header=1)
    output_path = f"记录表_{sheet_keyword}_审核标注版.xlsx"
    empty_row = pd.DataFrame([[""]*len(main_df.columns)], columns=main_df.columns)
    pd.concat([empty_row, main_df], ignore_index=True).to_excel(output_path, index=False)
    wb = load_workbook(output_path)
    ws = wb.active
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    global contract_col_main
    contract_col_main = find_col(main_df, "合同", exact=True)
    if not contract_col_main:
        st.error(f"❌ 在「{sheet_keyword}」中未找到合同列。")
        return 0, None, 0, set()

    total_errors = 0
    contracts_seen = set()
    progress = st.progress(0)
    status = st.empty()

    merged_fk, mask_fk = compare_fields_vectorized(main_df, fk_df, contract_col_main, contract_col_fk, mapping_fk)
    merged_zd, mask_zd = compare_fields_vectorized(main_df, zd_df, contract_col_main, contract_col_zd, mapping_zd, tolerance_dict={"保证金比例":0.005})
    merged_ec, mask_ec = compare_fields_vectorized(main_df, ec_df, contract_col_main, contract_col_ec, mapping_ec)
    merged_zk, mask_zk = compare_fields_vectorized(main_df, zk_df, contract_col_main, contract_col_zk, mapping_zk)

    mask_all = pd.concat([mask_fk, mask_zd, mask_ec, mask_zk], axis=1)
    mask_any = mask_all.any(axis=1)

    for r_idx, row in main_df.iterrows():
        contracts_seen.add(str(row[contract_col_main]).strip())
        for col in mask_all.columns:
            if mask_all.at[r_idx,col]:
                c_idx = list(main_df.columns).index(col)+1
                ws.cell(r_idx+3,c_idx).fill = red_fill
        if mask_any.at[r_idx]:
            c_contract = list(main_df.columns).index(contract_col_main)+1
            ws.cell(r_idx+3,c_contract).fill = yellow_fill
        if (r_idx+1) % 10 == 0:
            status.text(f"检查「{sheet_keyword}」... {r_idx+1}/{len(main_df)}")
        progress.progress((r_idx+1)/len(main_df))

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    st.download_button(f"📥 下载 {sheet_keyword} 审核标注版", output, f"记录表_{sheet_keyword}_审核标注版.xlsx")
    total_errors = mask_any.sum()
    elapsed = time.time()-start_time
    st.success(f"✅ {sheet_keyword} 检查完成，共 {total_errors} 处错误，用时 {elapsed:.2f} 秒。")
    return total_errors, elapsed, 0, contracts_seen

# =====================================
# 🧾 多 sheet 循环
# =====================================
total_all = elapsed_all = skip_total = 0
contracts_seen_all_sheets = set()
for kw in sheet_keywords:
    count, used, skipped, seen = check_one_sheet(kw)
    total_all += count
    elapsed_all += used or 0
    skip_total += skipped
    contracts_seen_all_sheets.update(seen)
st.success(f"🎯 全部审核完成，共 {total_all} 处错误，总耗时 {elapsed_all:.2f} 秒。")

# =====================================
# 🕵️ 漏填检查
# =====================================
field_contracts = zd_df[contract_col_zd].dropna().astype(str).str.strip()
col_car_manager = find_col(zd_df, "是否车管家", exact=True)
col_bonus_type = find_col(zd_df, "提成类型", exact=True)
missing_contracts_mask = (~field_contracts.isin(contracts_seen_all_sheets))
if col_car_manager:
    missing_contracts_mask &= ~(zd_df[col_car_manager].astype(str).str.strip().str.lower() == "是")
if col_bonus_type:
    missing_contracts_mask &= ~(zd_df[col_bonus_type].astype(str).str.strip().isin(["联合租赁","驻店"]))

zd_df_missing = zd_df.copy()
zd_df_missing["漏填检查"] = ""
zd_df_missing.loc[missing_contracts_mask, "漏填检查"] = "❗ 漏填"
漏填合同数 = zd_df_missing["漏填检查"].eq("❗ 漏填").sum()
st.warning(f"⚠️ 共发现 {漏填合同数} 个合同在记录表中未出现（已排除车管家、联合租赁、驻店）")

# =====================================
# 📤 导出字段表
# =====================================
yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
wb_all = Workbook()
ws_all = wb_all.active
for c_idx, c in enumerate(zd_df_missing.columns,1): ws_all.cell(1,c_idx,c)
for r_idx, row in enumerate(zd_df_missing.itertuples(index=False),2):
    for c_idx, v in enumerate(row,1):
        ws_all.cell(r_idx,c_idx,v)
        if zd_df_missing.columns[c_idx-1]=="漏填检查" and v=="❗ 漏填":
            ws_all.cell(r_idx,c_idx).fill = yellow_fill
output_all = BytesIO()
wb_all.save(output_all)
output_all.seek(0)
st.download_button("📥 下载字段表漏填标注版", output_all,"字段表_漏填标注版.xlsx")

zd_df_only_missing = zd_df_missing[zd_df_missing["漏填检查"]=="❗ 漏填"].copy()
if not zd_df_only_missing.empty:
    wb2 = Workbook()
    ws2 = wb2.active
    for c_idx, c in enumerate(zd_df_only_missing.columns,1): ws2.cell(1,c_idx,c)
    for r_idx,row in enumerate(zd_df_only_missing.itertuples(index=False),2):
        for c_idx,v in enumerate(row,1):
            ws2.cell(r_idx,c_idx,v)
            if zd_df_only_missing.columns[c_idx-1]=="漏填检查" and v=="❗ 漏填":
                ws2.cell(r_idx,c_idx).fill = yellow_fill
    out2 = BytesIO()
    wb2.save(out2)
    out2.seek(0)
    st.download_button("📥 下载仅漏填字段表", out2,"字段表_仅漏填.xlsx")

st.success("✅ 所有检查、标注与导出完成！")
