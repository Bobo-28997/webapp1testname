# =====================================
# Streamlit Web App: 模拟Project——人事用合同记录表自动审核系统
# （四输出表 + 漏填检查 + 驻店客户 + 向量化优化版）
# =====================================

import streamlit as st
import pandas as pd
import time
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill
from io import BytesIO

# =====================================
# 🏁 应用标题与说明
# =====================================
st.title("📊 模拟Project：人事用合同记录表自动审核系统（四Sheet + 向量化优化 + 漏填检查 + 驻店客户）")

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

def same_date_ymd(a, b):
    try:
        da = pd.to_datetime(a, errors="coerce")
        db = pd.to_datetime(b, errors="coerce")
        if pd.isna(da) or pd.isna(db): return False
        return (da.year, da.month, da.day) == (db.year, db.month, db.day)
    except Exception:
        return False

# =====================================
# 📖 读取文件
# =====================================
main_file = find_file(uploaded_files, "记录表")
fk_file = find_file(uploaded_files, "放款明细")
zd_file = find_file(uploaded_files, "字段")
ec_file = find_file(uploaded_files, "二次明细")
zk_file = find_file(uploaded_files, "重卡数据")

fk_xls = pd.ExcelFile(fk_file)
zd_xls = pd.ExcelFile(zd_file)

fk_df = pd.read_excel(fk_xls, sheet_name=find_sheet(fk_xls, "本司"))
zd_df = pd.read_excel(zd_xls, sheet_name=find_sheet(zd_xls, "重卡"))
ec_df = pd.read_excel(ec_file)
zk_df = pd.read_excel(zk_file)

# 自动去重列名，避免 merge 报错
for df in [fk_df, zd_df, ec_df, zk_df]:
    # 自动去重列名
    def dedup_columns(columns):
        seen = {}
        new_cols = []
        for c in columns:
            if c not in seen:
                seen[c] = 0
                new_cols.append(c)
            else:
                seen[c] += 1
                new_cols.append(f"{c}.{seen[c]}")
        return new_cols

    df.columns = dedup_columns(df.columns)


# =====================================
# 🧩 字段映射
# =====================================
mapping_fk = {"授信方": "授信", "租赁本金": "本金", "租赁期限月": "租赁期限月", "客户经理": "客户经理", "起租收益率": "收益率", "主车台数": "主车台数", "挂车台数": "挂车台数"}
mapping_zd = {"保证金比例": "保证金比例_2", "项目提报人": "提报", "起租时间": "起租日_商", "租赁期限月": "总期数_商_资产", "所属省区": "区域", "城市经理": "城市经理"}
mapping_ec = {"二次时间": "出本流程时间"}
mapping_zk = {"授信方": "授信方"}

# =====================================
# ⚙️ 向量化对比函数
# =====================================
def check_one_sheet_fast(sheet_keyword):
    start_time = time.time()
    xls_main = pd.ExcelFile(main_file)

    try:
        sheet_name = find_sheet(xls_main, sheet_keyword)
    except ValueError:
        st.warning(f"⚠️ 未找到包含「{sheet_keyword}」的sheet，跳过。")
        return 0, 0, 0, set()

    main_df = pd.read_excel(xls_main, sheet_name=sheet_name, header=1)
    contract_col_main = find_col(main_df, "合同")
    if not contract_col_main:
        st.error(f"❌ 在 {sheet_keyword} 未找到合同号列")
        return 0, 0, 0, set()

    # 输出文件预备
    wb = Workbook()
    ws = wb.active
    for i, c in enumerate(main_df.columns, 1): ws.cell(1, i, c)

    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    total_errors = 0
    skip_city_manager = 0
    contracts_seen = set(main_df[contract_col_main].dropna().astype(str).str.strip())

    # 核心向量化比对函数
    def batch_compare(mapping_dict, ref_df, ref_contract_col):
        nonlocal total_errors, skip_city_manager

        ref_df = ref_df.loc[:, ~ref_df.columns.duplicated()].copy()
        cols_keep = [ref_contract_col] + [v for v in mapping_dict.values() if v in ref_df.columns]
        ref_df = ref_df[cols_keep]

        merged = main_df.merge(
            ref_df,
            left_on=contract_col_main,
            right_on=ref_contract_col,
            how="left",
            suffixes=("", "_ref")
        )

        for main_kw, ref_kw in mapping_dict.items():
            if main_kw not in main_df.columns or ref_kw not in merged.columns:
                continue

            a = merged[main_kw]
            b = merged[ref_kw]

            # 城市经理字段跳过空值情况
            if main_kw == "城市经理":
                skip_city_manager += b.isna().sum()
                b = b.fillna(a)  # 避免误判

            # 日期字段：精确到年月日
            if any(k in main_kw for k in ["日期", "时间"]) or any(k in ref_kw for k in ["日期", "时间"]):
                mismatch = ~a.combine(b, same_date_ymd)
            else:
                # 尝试数值比较
                a_num = pd.to_numeric(a, errors="coerce")
                b_num = pd.to_numeric(b, errors="coerce")
                both_num = a_num.notna() & b_num.notna()
                mismatch = (
                    (both_num & (abs(a_num - b_num) > 1e-6))
                    | (~both_num & (a.astype(str).str.strip() != b.astype(str).str.strip()))
                )

            total_errors += mismatch.sum()
            # 标红单元格
            for idx in merged[mismatch].index:
                ws.cell(idx + 2, list(main_df.columns).index(main_kw) + 1).fill = red_fill

    # 执行批量对比
    contract_col_fk = find_col(fk_df, "合同")
    contract_col_zd = find_col(zd_df, "合同")
    contract_col_ec = find_col(ec_df, "合同")
    contract_col_zk = find_col(zk_df, "合同")

    batch_compare(mapping_fk, fk_df, contract_col_fk)
    batch_compare(mapping_zd, zd_df, contract_col_zd)
    batch_compare(mapping_ec, ec_df, contract_col_ec)
    batch_compare(mapping_zk, zk_df, contract_col_zk)

    # 合同列黄标
    cidx = list(main_df.columns).index(contract_col_main) + 1
    for r in range(len(main_df)):
        if any(ws.cell(r + 2, c).fill == red_fill for c in range(1, len(main_df.columns) + 1)):
            ws.cell(r + 2, cidx).fill = yellow_fill

    # 导出
    out = BytesIO()
    wb.save(out)
    out.seek(0)
    st.download_button(f"📥 下载 {sheet_keyword} 审核结果", out, f"记录表_{sheet_keyword}_审核标注版.xlsx")

    elapsed = time.time() - start_time
    st.success(f"✅ {sheet_keyword} 检查完成，共 {total_errors} 处错误，用时 {elapsed:.2f} 秒。")
    return total_errors, elapsed, skip_city_manager, contracts_seen


# =====================================
# 🔄 执行四sheet检查
# =====================================
sheet_keywords = ["二次", "部分担保", "随州", "驻店客户"]
total_all = elapsed_all = skip_total = 0
contracts_seen_all_sheets = set()

for kw in sheet_keywords:
    count, used, skipped, seen = check_one_sheet_fast(kw)
    total_all += count
    elapsed_all += used
    skip_total += skipped
    contracts_seen_all_sheets.update(seen)

st.success(f"🎯 全部检查完成，共 {total_all} 处错误，耗时 {elapsed_all:.2f} 秒。")

# =====================================
# 🕵️ 漏填检查（含跳过条件）
# =====================================
contract_col_zd = find_col(zd_df, "合同")
col_car_manager = find_col(zd_df, "是否车管家", exact=True)
col_bonus_type = find_col(zd_df, "提成类型", exact=True)

field_contracts = zd_df[contract_col_zd].dropna().astype(str).str.strip()
missing_mask = ~field_contracts.isin(contracts_seen_all_sheets)

if col_car_manager:
    missing_mask &= ~(zd_df[col_car_manager].astype(str).str.strip() == "是")
if col_bonus_type:
    missing_mask &= ~zd_df[col_bonus_type].astype(str).str.strip().isin(["联合租赁", "驻店"])

zd_df_missing = zd_df.copy()
zd_df_missing["漏填检查"] = ""
zd_df_missing.loc[missing_mask, "漏填检查"] = "❗ 漏填"

漏填数 = zd_df_missing["漏填检查"].eq("❗ 漏填").sum()
st.warning(f"⚠️ 共发现 {漏填数} 个漏填合同（已排除车管家、联合租赁、驻店）")

# =====================================
# 📤 导出字段表
# =====================================
yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

def export_excel(df, filename):
    wb = Workbook()
    ws = wb.active
    for c_idx, c in enumerate(df.columns, 1): ws.cell(1, c_idx, c)
    for r_idx, row in enumerate(df.itertuples(index=False), 2):
        for c_idx, v in enumerate(row, 1):
            ws.cell(r_idx, c_idx, v)
            if df.columns[c_idx - 1] == "漏填检查" and v == "❗ 漏填":
                ws.cell(r_idx, c_idx).fill = yellow_fill
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    st.download_button(f"📥 下载 {filename}", bio, f"{filename}.xlsx")

export_excel(zd_df_missing, "字段表_漏填标注版")
zd_df_only_missing = zd_df_missing[zd_df_missing["漏填检查"] == "❗ 漏填"]
if not zd_df_only_missing.empty:
    export_excel(zd_df_only_missing, "字段表_仅漏填")

st.success("✅ 所有检查、标注与导出完成！")

