# =====================================
# Streamlit Web App: 模拟Project——人事用合同记录表自动审核系统
# （四Sheet + 向量化优化 + 漏填检查 + 驻店客户）
# =====================================

import streamlit as st
import pandas as pd
import time
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from io import BytesIO

st.title("📊 模拟Project：人事用合同记录表自动审核系统（四Sheet + 向量化优化 + 漏填检查 + 驻店客户）")

# -------- 上传文件 ----------
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

# -------- 工具：查找文件 / 列 / sheet ----------
def find_file(files_list, keyword):
    for f in files_list:
        if keyword in f.name:
            return f
    raise FileNotFoundError(f"❌ 未找到包含关键词「{keyword}」的文件")

def normalize_colname(c):
    return str(c).strip().lower()

def find_col(df, keyword, exact=False):
    if df is None:
        return None
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

# 小的帮助：把合同列统一为 string 且 strip
def normalize_contract_col(df, col):
    if col is None:
        return df
    df[col] = df[col].astype(str).where(~df[col].isna(), "").str.strip()
    return df

# 日期按年月日判断
def ymd_series(s):
    return pd.to_datetime(s, errors="coerce").dt.normalize()

# -------- 读取文件并准备 dataframes ----------
main_file = find_file(uploaded_files, "记录表")
fk_file   = find_file(uploaded_files, "放款明细")
zd_file   = find_file(uploaded_files, "字段")
ec_file   = find_file(uploaded_files, "二次明细")
zk_file   = find_file(uploaded_files, "重卡数据")

fk_xls = pd.ExcelFile(fk_file)
zd_xls = pd.ExcelFile(zd_file)

# 读取参照表（用模糊匹配 sheet 名）
fk_df = pd.read_excel(fk_xls, sheet_name=find_sheet(fk_xls, "本司"))
zd_df = pd.read_excel(zd_xls, sheet_name=find_sheet(zd_xls, "重卡"))
ec_df = pd.read_excel(ec_file)
zk_df = pd.read_excel(zk_file)

# 防止参照表列名重复导致 merge 报错：对每个参照表保留第一组同名列
def dedup_keep_first(df):
    # 如果存在重复列名，保留第一次出现，删除重复后的同名列
    cols = df.columns.tolist()
    seen = set()
    keep_cols = []
    for c in cols:
        if c not in seen:
            keep_cols.append(c)
            seen.add(c)
    return df.loc[:, keep_cols]

fk_df = dedup_keep_first(fk_df)
zd_df = dedup_keep_first(zd_df)
ec_df = dedup_keep_first(ec_df)
zk_df = dedup_keep_first(zk_df)

# -------- 映射配置（保持与你之前相同） ----------
mapping_fk = {
    "授信方": "授信",
    "租赁本金": "本金",
    "租赁期限月": "租赁期限月",
    "客户经理": "客户经理",
    "起租收益率": "收益率",
    "主车台数": "主车台数",
    "挂车台数": "挂车台数"
}
mapping_zd = {
    "保证金比例": "保证金比例_2",
    "项目提报人": "提报",
    "起租时间": "起租日_商",
    "租赁期限月": "总期数_商_资产",
    "所属省区": "区域",
    "城市经理": "城市经理"
}
mapping_ec = {"二次时间": "出本流程时间"}
mapping_zk = {"授信方": "授信方"}

# -------- 识别参照表合同列（会在每次运行前normalize合同列） ----------
contract_col_fk = find_col(fk_df, "合同")
contract_col_zd = find_col(zd_df, "合同")
contract_col_ec = find_col(ec_df, "合同")
contract_col_zk = find_col(zk_df, "合同")

# 把参照表的合同列统一为字符串并 strip（方便 later merge/isin）
if contract_col_fk:
    fk_df = normalize_contract_col(fk_df, contract_col_fk)
if contract_col_zd:
    zd_df = normalize_contract_col(zd_df, contract_col_zd)
if contract_col_ec:
    ec_df = normalize_contract_col(ec_df, contract_col_ec)
if contract_col_zk:
    zk_df = normalize_contract_col(zk_df, contract_col_zk)

# -------- 批量/向量化比对辅助函数 ----------
def safe_get_ref_col_name(merged_df, ref_kw):
    """merge 后可能出现 ref_kw 或 ref_kw + '_ref'，优先返回存在的那个"""
    if ref_kw in merged_df.columns:
        return ref_kw
    if f"{ref_kw}_ref" in merged_df.columns:
        return f"{ref_kw}_ref"
    # as fallback, try any column that endswith ref_kw (rare)
    for c in merged_df.columns:
        if c.endswith(ref_kw):
            return c
    return None

def write_workbook_with_marks(main_df, red_marks_by_col, filename, contract_col_main, yellow_fill, red_fill):
    """
    main_df: 原始主表（未插入空行）
    red_marks_by_col: dict main_col -> set(orig_idx) 要标红的行索引（这些索引对应 main_df.index）
    生成 workbook 并返回 BytesIO
    """
    wb = Workbook()
    ws = wb.active

    # 写表头
    for c_idx, c in enumerate(main_df.columns, start=1):
        ws.cell(1, c_idx, c)

    # 写数据行
    for r_idx, (_, row) in enumerate(main_df.iterrows(), start=2):
        for c_idx, c in enumerate(main_df.columns, start=1):
            ws.cell(r_idx, c_idx, row[c])

    # 标红指定单元格（red_marks_by_col）
    for col_name, idx_set in red_marks_by_col.items():
        if col_name not in main_df.columns:
            continue
        col_idx = list(main_df.columns).index(col_name) + 1
        for orig_idx in idx_set:
            # 找到在 excel 中对应的行号（header + 1blank行未要求，这里直接写 main_df）
            excel_row = list(main_df.index).index(orig_idx) + 2
            ws.cell(excel_row, col_idx).fill = red_fill

    # 标黄色合同列（整行有任何红就黄）
    if contract_col_main in main_df.columns:
        contract_col_idx = list(main_df.columns).index(contract_col_main) + 1
        for r_idx, orig_idx in enumerate(main_df.index, start=2):
            # check if any red in that row: scan red_marks_by_col
            has_red = False
            for col_name, idx_set in red_marks_by_col.items():
                if orig_idx in idx_set:
                    has_red = True
                    break
            if has_red:
                ws.cell(r_idx, contract_col_idx).fill = yellow_fill

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio

# -------- 主检查函数（向量化） ----------
def check_one_sheet_fast(sheet_keyword):
    """对一个 sheet 做向量化检查，返回 (errors_count, elapsed_seconds, skip_city_manager_count, contracts_seen_set, excel_bytesio)"""
    start_time = time.time()
    xls_main = pd.ExcelFile(main_file)
    try:
        sheet_name = find_sheet(xls_main, sheet_keyword)
    except ValueError:
        st.warning(f"⚠️ 未找到包含「{sheet_keyword}」的sheet，跳过。")
        return 0, 0, 0, set(), None

    main_df = pd.read_excel(xls_main, sheet_name=sheet_name, header=1)
    # 统一合同列并保留原始 index 便于回写
    contract_col_main = find_col(main_df, "合同")
    if not contract_col_main:
        st.error(f"❌ 在 {sheet_keyword} 未找到合同号列")
        return 0, 0, 0, set(), None
    main_df = normalize_contract_col(main_df, contract_col_main)
    main_df["_orig_idx"] = main_df.index  # 保存原始索引

    # red marks accumulator: main_col -> set(orig_idx)
    red_marks = {col: set() for col in main_df.columns}

    total_errors = 0
    skip_city_manager = 0
    contracts_seen = set(main_df[contract_col_main].dropna().astype(str).str.strip())

    # 进度显示（粗略）
    progress = st.progress(0)
    status = st.empty()

    # 内部：执行一次 mapping 的批量比较
    def batch_compare(mapping_dict, ref_df, ref_contract_col):
        nonlocal total_errors, skip_city_manager

        if ref_df is None or ref_contract_col is None:
            return

        # 只保留参照表需要的列：合同列 + 映射中的参照列（存在则保留）
        needed_ref_cols = []
        for v in mapping_dict.values():
            if v in ref_df.columns:
                needed_ref_cols.append(v)
        needed_ref_cols = [ref_contract_col] + needed_ref_cols
        ref_sub = ref_df.loc[:, [c for c in needed_ref_cols if c in ref_df.columns]].copy()

        # 规范参照表合同列为字符串
        if ref_contract_col in ref_sub.columns:
            ref_sub[ref_contract_col] = ref_sub[ref_contract_col].astype(str).where(~ref_sub[ref_contract_col].isna(), "").str.strip()

        # 为避免 merge 时列名冲突，先把 ref_sub 的列与 main_df 的列名比较；我们会在 merge 后使用 safe_get_ref_col_name 查找
        # 把 main_df 的合同和 ref_sub 的合同都作为 key 做左连接
        merged = main_df.merge(
            ref_sub,
            left_on=contract_col_main,
            right_on=ref_contract_col,
            how="left",
            suffixes=("", "_ref")
        )

        # merged 中保留原始索引列 _orig_idx
        for main_kw, ref_kw in mapping_dict.items():
            # 如果主表不包含 main_kw，就跳过（比如用户主表列名不完全一致）
            if main_kw not in merged.columns:
                continue

            # 找到参照列在 merged 中的名字（可能是 ref_kw 或 ref_kw + '_ref'）
            ref_col_in_merged = safe_get_ref_col_name(merged, ref_kw)
            if ref_col_in_merged is None:
                # 参照表不含该列，跳过
                continue

            a = merged[main_kw]            # 来自主表的值（Series）
            b = merged[ref_col_in_merged]  # 来自参照表的值（Series）

            # 城市经理：参照表为空的直接跳过并计数（视为未填写 -> 不判错）
            if main_kw == "城市经理":
                # count rows where b is blank/NaN/empty string
                b_is_blank = b.isna() | (b.astype(str).str.strip() == "") | (b.astype(str).str.strip().str.lower().isin(["nan", "none", "null", "-"]))
                skip_city_manager += int(b_is_blank.sum())
                # for comparison, fill those b blanks with the a value so they won't be flagged below
                b = b.mask(b_is_blank, a)

            # 判断是否为日期字段（主或参照字段名包含“日期”或“时间”）
            is_date_field = any(k in main_kw for k in ["日期", "时间"]) or any(k in ref_kw for k in ["日期", "时间"])

            if is_date_field:
                # 转为日期，然后比较年月日是否一致
                a_dt = pd.to_datetime(a, errors="coerce")
                b_dt = pd.to_datetime(b, errors="coerce")
                a_na = a_dt.isna()
                b_na = b_dt.isna()
                both_na = a_na & b_na
                mismatch = ~( (a_dt.dt.year == b_dt.dt.year) & (a_dt.dt.month == b_dt.dt.month) & (a_dt.dt.day == b_dt.dt.day) )
                # 当两者都为 NaT 时，不视为 mismatch
                mismatch = mismatch & (~both_na)
            else:
                # 尝试数值比较（向量化）
                a_num = pd.to_numeric(a, errors="coerce")
                b_num = pd.to_numeric(b, errors="coerce")
                a_na = a.isna()
                b_na = b.isna()
                both_na = a_na & b_na

                both_num = a_num.notna() & b_num.notna()
                # numeric mismatch
                # 针对保证金比例使用容差 0.005；注意 mapping_zd 的那一对
                if main_kw == "保证金比例" and ref_kw == "保证金比例_2":
                    tol = 0.005
                else:
                    tol = 1e-6

                numeric_mismatch = both_num & ( (a_num - b_num).abs() > tol )

                # 非数值比较（字符串比较）：在此之前把 NaN 视为空字符串，但如果两者均为空，应视为相等（因此用 both_na 屏蔽）
                a_str = a.astype(str).where(~a_na, "").str.strip().str.lower()
                b_str = b.astype(str).where(~b_na, "").str.strip().str.lower()
                nonnum_mismatch = (~both_num) & (~both_na) & (a_str != b_str)

                mismatch = numeric_mismatch | nonnum_mismatch

            # 把 mismatch 转为需要标红的主表原始 index 集合
            mismatch_idx = merged.loc[mismatch, "_orig_idx"].tolist()
            if mismatch_idx:
                total_errors += len(mismatch_idx)
                # accumulate
                red_marks.setdefault(main_kw, set()).update(mismatch_idx)

    # 逐个 mapping 批量比较（向量化）
    contract_col_fk_local = contract_col_fk
    contract_col_zd_local = contract_col_zd
    contract_col_ec_local = contract_col_ec
    contract_col_zk_local = contract_col_zk

    # 执行批量比较（每次都会做一次 merge）
    batch_compare(mapping_fk, fk_df, contract_col_fk_local)
    # 映射字段表
    batch_compare(mapping_zd, zd_df, contract_col_zd_local)
    # 二次
    batch_compare(mapping_ec, ec_df, contract_col_ec_local)
    # 重卡
    batch_compare(mapping_zk, zk_df, contract_col_zk_local)

    # 写出 Excel（把 main_df 原样写入，然后应用 red_marks）
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    # 生成 workbook bytes
    bio = write_workbook_with_marks(main_df.drop(columns=["_orig_idx"]), red_marks, f"记录表_{sheet_keyword}_审核标注版.xlsx", contract_col_main, yellow_fill, red_fill)

    elapsed = time.time() - start_time
    st.success(f"✅ {sheet_keyword} 检查完成，共 {total_errors} 处错误，用时 {elapsed:.2f} 秒。")
    return total_errors, elapsed, skip_city_manager, contracts_seen, bio

# -------- 循环执行四个 sheet ----------
sheet_keywords = ["二次", "部分担保", "随州", "驻店客户"]
total_all = elapsed_all = skip_total = 0
contracts_seen_all_sheets = set()

# 为每个 sheet 提供下载按钮（向量化版本生成的 bytesio）
sheet_bios = {}

for kw in sheet_keywords:
    count, used, skipped, seen, bio = check_one_sheet_fast(kw)
    total_all += count
    elapsed_all += used
    skip_total += skipped
    contracts_seen_all_sheets.update(seen)
    sheet_bios[kw] = bio
    # 显示并提供下载（如果 bio 为 None 则跳过）
    if bio is not None:
        st.download_button(label=f"📥 下载 {kw} 审核标注版", data=bio, file_name=f"记录表_{kw}_审核标注版.xlsx")

st.success(f"🎯 全部检查完成，共 {total_all} 处错误，总耗时 {elapsed_all:.2f} 秒。")
st.info(f"📍 跳过字段表中空城市经理的合同数量总数（估算）：{skip_total}")

# -------- 字段表 漏填 检查（跳过条件：车管家=是；提成类型=联合租赁/驻店） ----------
contract_col_zd = find_col(zd_df, "合同")
col_car_manager = find_col(zd_df, "是否车管家", exact=True)
col_bonus_type = find_col(zd_df, "提成类型", exact=True)

field_contracts = zd_df[contract_col_zd].dropna().astype(str).str.strip()
missing_mask = ~field_contracts.isin(contracts_seen_all_sheets)

# 跳过车管家=是
if col_car_manager:
    missing_mask &= ~(zd_df[col_car_manager].astype(str).str.strip().str.lower() == "是")
# 跳过提成类型联合租赁/驻店
if col_bonus_type:
    missing_mask &= ~(zd_df[col_bonus_type].astype(str).str.strip().isin(["联合租赁", "驻店"]))

zd_df_missing = zd_df.copy()
zd_df_missing["漏填检查"] = ""
zd_df_missing.loc[missing_mask, "漏填检查"] = "❗ 漏填"

漏填数 = zd_df_missing["漏填检查"].eq("❗ 漏填").sum()
st.warning(f"⚠️ 共发现 {漏填数} 个漏填合同（已排除车管家、联合租赁、驻店）")

# 输出字段表（带黄色标注）
def export_excel_with_yellow(df, filename):
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    wb = Workbook()
    ws = wb.active
    # headers
    for c_idx, c in enumerate(df.columns, start=1):
        ws.cell(1, c_idx, c)
    # rows
    for r_idx, row in enumerate(df.itertuples(index=False), start=2):
        for c_idx, v in enumerate(row, start=1):
            ws.cell(r_idx, c_idx, v)
            if df.columns[c_idx-1] == "漏填检查" and v == "❗ 漏填":
                ws.cell(r_idx, c_idx).fill = yellow_fill
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    st.download_button(label=f"📥 下载 {filename}", data=bio, file_name=f"{filename}.xlsx")

export_excel_with_yellow(zd_df_missing, "字段表_漏填标注版")
zd_df_only_missing = zd_df_missing[zd_df_missing["漏填检查"] == "❗ 漏填"].copy()
if not zd_df_only_missing.empty:
    export_excel_with_yellow(zd_df_only_missing, "字段表_仅漏填")

st.success("✅ 所有检查、标注与导出完成！")

