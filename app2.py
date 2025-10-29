# =====================================
# Streamlit Web App: 合同记录表自动审核（优化版 - 向量化比对 + 四输出表 + 漏填检查）
# =====================================

import streamlit as st
import pandas as pd
import time
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from io import BytesIO

# =====================================
# 🏁 应用标题
# =====================================
st.title("📊 模拟Project：人事用合同记录表自动审核系统（四Sheet + 向量化比对 + 漏填检查优化版）")

# =====================================
# 📂 上传文件区
# =====================================
uploaded_files = st.file_uploader(
    "请上传以下文件：记录表、放款明细、字段、二次明细、重卡数据",
    type="xlsx",
    accept_multiple_files=True
)

if not uploaded_files or len(uploaded_files) < 5:
    st.warning("⚠️ 请上传全部 5 个文件后继续")
    st.stop()
else:
    st.success("✅ 文件上传完成")

# =====================================
# 🔧 工具函数
# =====================================
def find_file(files_list, keyword):
    for f in files_list:
        if keyword in f.name:
            return f
    raise FileNotFoundError(f"❌ 未找到包含关键词「{keyword}」的文件")

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

# =====================================
# 🧮 快速对照函数（向量化版本）
# =====================================
def check_one_sheet_fast(sheet_keyword):
    start_time = time.time()
    xls_main = pd.ExcelFile(main_file)

    try:
        target_sheet = find_sheet(xls_main, sheet_keyword)
    except ValueError:
        st.warning(f"⚠️ 未找到包含「{sheet_keyword}」的sheet，跳过。")
        return 0, None, 0, set()

    main_df = pd.read_excel(xls_main, sheet_name=target_sheet, header=1)
    contract_col_main = find_col(main_df, "合同")
    if not contract_col_main:
        st.error(f"❌ 在「{sheet_keyword}」中未找到合同列。")
        return 0, None, 0, set()

    main_df[contract_col_main] = main_df[contract_col_main].astype(str).str.strip()
    contracts_seen = set(main_df[contract_col_main].dropna().unique())

    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    wb = Workbook()
    ws = wb.active
    for c_idx, c in enumerate(main_df.columns, 1): ws.cell(1, c_idx, c)

    # === 初始化错误矩阵 ===
    error_flags = pd.DataFrame(False, index=main_df.index, columns=main_df.columns)
    skip_city_manager = 0

    def batch_compare(mapping_dict, ref_df, ref_contract_col):
        nonlocal skip_city_manager
        ref_df[ref_contract_col] = ref_df[ref_contract_col].astype(str).str.strip()
        merged = main_df.merge(
            ref_df[[ref_contract_col] + list(ref_df.columns)],
            left_on=contract_col_main,
            right_on=ref_contract_col,
            how="left",
            suffixes=("", "_ref")
        )
        for main_kw, ref_kw in mapping_dict.items():
            mc = find_col(main_df, main_kw)
            rc = find_col(merged, ref_kw + "_ref") or (ref_kw + "_ref")
            if mc and rc in merged.columns:
                main_vals = merged[mc].apply(normalize_num)
                ref_vals = merged[rc].apply(normalize_num)
                mismatched = []
                for a, b in zip(main_vals, ref_vals):
                    # 城市经理跳过
                    if main_kw == "城市经理":
                        if b in [None, "", "-", "nan", "none", "null"]:
                            skip_city_manager += 1
                            mismatched.append(False)
                            continue
                    if (a is None and b is None):
                        mismatched.append(False)
                    elif any(k in main_kw for k in ["日期", "时间"]) or any(k in ref_kw for k in ["日期", "时间"]):
                        mismatched.append(not same_date_ymd(a, b))
                    elif isinstance(a, (int, float)) and isinstance(b, (int, float)):
                        diff = abs(a - b)
                        if main_kw == "保证金比例" and ref_kw == "保证金比例_2":
                            mismatched.append(diff > 0.005)
                        else:
                            mismatched.append(diff > 1e-6)
                    else:
                        mismatched.append(str(a).strip().lower().replace(".0", "") != str(b).strip().lower().replace(".0", ""))
                error_flags[mc] |= mismatched

    # === 执行四类批量对照 ===
    batch_compare(mapping_fk, fk_df, contract_col_fk)
    batch_compare(mapping_zd, zd_df, contract_col_zd)
    batch_compare(mapping_ec, ec_df, contract_col_ec)
    batch_compare(mapping_zk, zk_df, contract_col_zk)

    # === 输出结果 ===
    total_errors = error_flags.sum().sum()
    for i, (_, row) in enumerate(main_df.iterrows(), 2):
        row_has_error = False
        for j, col in enumerate(main_df.columns, 1):
            v = row[col]
            ws.cell(i, j, v)
            if error_flags.loc[i - 2, col]:
                ws.cell(i, j).fill = red_fill
                row_has_error = True
        if row_has_error:
            ws.cell(i, list(main_df.columns).index(contract_col_main) + 1).fill = yellow_fill

    output = BytesIO()
    wb.save(output)
    output.seek(0)
    st.download_button(f"📥 下载 {sheet_keyword} 审核标注版", output, f"记录表_{sheet_keyword}_审核标注版.xlsx")

    elapsed = time.time() - start_time
    st.success(f"✅ {sheet_keyword} 检查完成，共 {total_errors} 处错误，用时 {elapsed:.2f} 秒。")
    return total_errors, elapsed, skip_city_manager, contracts_seen

# =====================================
# 📖 文件读取区
# =====================================
main_file = find_file(uploaded_files, "记录表")
fk_file = find_file(uploaded_files, "放款明细")
zd_file = find_file(uploaded_files, "字段")
ec_file = find_file(uploaded_files, "二次明细")
zk_file = find_file(uploaded_files, "重卡数据")

fk_df = pd.read_excel(pd.ExcelFile(fk_file), sheet_name=find_sheet(pd.ExcelFile(fk_file), "本司"))
zd_df = pd.read_excel(pd.ExcelFile(zd_file), sheet_name=find_sheet(pd.ExcelFile(zd_file), "重卡"))
ec_df = pd.read_excel(ec_file)
zk_df = pd.read_excel(zk_file)

contract_col_fk = find_col(fk_df, "合同")
contract_col_zd = find_col(zd_df, "合同")
contract_col_ec = find_col(ec_df, "合同")
contract_col_zk = find_col(zk_df, "合同")

# 字段映射
mapping_fk = {"授信方": "授信", "租赁本金": "本金", "租赁期限月": "租赁期限月", "客户经理": "客户经理", "起租收益率": "收益率", "主车台数": "主车台数", "挂车台数": "挂车台数"}
mapping_zd = {"保证金比例": "保证金比例_2", "项目提报人": "提报", "起租时间": "起租日_商", "租赁期限月": "总期数_商_资产", "所属省区": "区域", "城市经理": "城市经理"}
mapping_ec = {"二次时间": "出本流程时间"}
mapping_zk = {"授信方": "授信方"}

# =====================================
# 🧾 多sheet循环检查
# =====================================
sheet_keywords = ["二次", "部分担保", "随州", "驻店客户"]
total_all = elapsed_all = skip_total = 0
contracts_seen_all_sheets = set()

for kw in sheet_keywords:
    count, used, skipped, seen = check_one_sheet_fast(kw)
    total_all += count
    elapsed_all += used or 0
    skip_total += skipped
    contracts_seen_all_sheets.update(seen)

st.success(f"🎯 全部检查完成，共 {total_all} 处错误，总耗时 {elapsed_all:.2f} 秒。")
st.info(f"跳过字段表中空城市经理合同数量：{skip_total}")

# =====================================
# 🕵️ 漏填检查
# =====================================
field_contracts = zd_df[contract_col_zd].dropna().astype(str).str.strip()
col_car_manager = find_col(zd_df, "是否车管家", exact=True)
col_bonus_type = find_col(zd_df, "提成类型", exact=True)

missing_mask = ~field_contracts.isin(contracts_seen_all_sheets)
if col_car_manager:
    missing_mask &= ~(zd_df[col_car_manager].astype(str).str.strip().str.lower() == "是")
if col_bonus_type:
    missing_mask &= ~zd_df[col_bonus_type].astype(str).str.strip().isin(["联合租赁", "驻店"])

zd_df_missing = zd_df.copy()
zd_df_missing["漏填检查"] = ""
zd_df_missing.loc[missing_mask, "漏填检查"] = "❗ 漏填"
漏填合同数 = zd_df_missing["漏填检查"].eq("❗ 漏填").sum()
st.warning(f"⚠️ 共发现 {漏填合同数} 个合同在记录表中未出现（已排除车管家、联合租赁、驻店）")

# 导出字段表（含漏填标注）
yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
wb_all = Workbook()
ws_all = wb_all.active
for c_idx, c in enumerate(zd_df_missing.columns, 1): ws_all.cell(1, c_idx, c)
for r_idx, row in enumerate(zd_df_missing.itertuples(index=False), 2):
    for c_idx, v in enumerate(row, 1):
        ws_all.cell(r_idx, c_idx, v)
        if zd_df_missing.columns[c_idx-1] == "漏填检查" and v == "❗ 漏填":
            ws_all.cell(r_idx, c_idx).fill = yellow_fill
out_all = BytesIO()
wb_all.save(out_all)
out_all.seek(0)
st.download_button("📥 下载字段表漏填标注版", out_all, "字段表_漏填标注版.xlsx")

# 仅漏填
zd_df_only_missing = zd_df_missing[zd_df_missing["漏填检查"] == "❗ 漏填"].copy()
if not zd_df_only_missing.empty:
    wb2 = Workbook()
    ws2 = wb2.active
    for c_idx, c in enumerate(zd_df_only_missing.columns, 1): ws2.cell(1, c_idx, c)
    for r_idx, row in enumerate(zd_df_only_missing.itertuples(index=False), 2):
        for c_idx, v in enumerate(row, 1):
            ws2.cell(r_idx, c_idx, v)
            if zd_df_only_missing.columns[c_idx-1] == "漏填检查" and v == "❗ 漏填":
                ws2.cell(r_idx, c_idx).fill = yellow_fill
    out2 = BytesIO()
    wb2.save(out2)
    out2.seek(0)
    st.download_button("📥 下载仅漏填字段表", out2, "字段表_仅漏填.xlsx")

st.success("✅ 所有检查、标注与导出完成！")

