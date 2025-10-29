# =====================================
# Streamlit Web App: 人事用合同记录表自动审核（四Sheet + 漏填检查 + 驻店客户版）向量化优化版
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
st.title("📊 人事合同记录表自动审核系统（四Sheet + 漏填检查 + 驻店客户版）- 向量化优化")

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
# 🟢 向量化比对函数
# =====================================
def compare_fields_vectorized(main_df, ref_df, main_contract_col, ref_contract_col, mapping_dict, tolerance_dict={}):
    """
    向量化比对
    - main_df: 记录表sheet
    - ref_df: 参考表
    - mapping_dict: {主表列: 参考表列}
    - tolerance_dict: 对数值列设置容差
    返回：merged DataFrame, mask 错误标记布尔矩阵
    """
    merged = main_df.merge(
        ref_df[[ref_contract_col] + list(mapping_dict.values())],
        left_on=main_contract_col, right_on=ref_contract_col,
        how='left', suffixes=('', '_ref')
    )
    
    mask = pd.DataFrame(False, index=merged.index, columns=mapping_dict.keys())

    for main_col, ref_col in mapping_dict.items():
        a = merged[main_col]
        b = merged[f"{ref_col}"]

        # 统一数值
        a_num = pd.to_numeric(a.astype(str).str.replace("%","").str.replace(",",""), errors='coerce')
        b_num = pd.to_numeric(b.astype(str).str.replace("%","").str.replace(",",""), errors='coerce')

        # 日期字段
        date_mask = a.astype(str).str.contains("日期|时间") | b.astype(str).str.contains("日期|时间")
        for idx in merged.index[date_mask]:
            if not same_date_ymd(a[idx], b[idx]):
                mask.loc[idx, main_col] = True

        # 保证金比例容差
        tol = tolerance_dict.get(main_col, 1e-6)
        num_mask = (~date_mask) & (a_num.notna() & b_num.notna())
        mask.loc[num_mask, main_col] = (abs(a_num[num_mask] - b_num[num_mask]) > tol)

        # 文本对比（包括NaN处理）
        text_mask = (~date_mask) & (a_num.isna() | b_num.isna())
        mask.loc[text_mask, main_col] = a.astype(str)[text_mask].str.strip().str.lower() != b.astype(str)[text_mask].str.strip().str.lower()

    return merged, mask

# =====================================
# 📖 文件读取
# =====================================
main_file = find_file(uploaded_files, "记录表")
fk_file = find_file(uploaded_files, "放款明细")
zd_file = find_file(uploaded_files, "字段")
ec_file = find_file(uploaded_files, "二次明细")
zk_file = find_file(uploaded_files, "重卡数据")

fk_df = pd.read_excel(fk_file, sheet_name=find_sheet(pd.ExcelFile(fk_file), "本司"))
zd_df = pd.read_excel(zd_file, sheet_name=find_sheet(pd.ExcelFile(zd_file), "重卡"))
ec_df = pd.read_excel(ec_file)
zk_df = pd.read_excel(zk_file)

contract_col_fk = find_col(fk_df, "合同")
contract_col_zd = find_col(zd_df, "合同")
contract_col_ec = find_col(ec_df, "合同")
contract_col_zk = find_col(zk_df, "合同")

mapping_fk = {"授信方": "授信","租赁本金":"本金","租赁期限月":"租赁期限月","客户经理":"客户经理","起租收益率":"收益率","主车台数":"主车台数","挂车台数":"挂车台数"}
mapping_zd = {"保证金比例":"保证金比例_2","项目提报人":"提报","起租时间":"起租日_商","租赁期限月":"总期数_商_资产","起租收益率":"XIRR_商_起租","所属省区":"区域","城市经理":"城市经理"}
mapping_ec = {"二次时间":"出本流程时间"}
mapping_zk = {"授信方":"授信方"}

# =====================================
# 🧾 多sheet循环
# =====================================
sheet_keywords = ["二次", "部分担保", "随州", "驻店客户"]
total_all = elapsed_all = 0
contracts_seen_all_sheets = set()

red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

for sheet_kw in sheet_keywords:
    start_time = time.time()
    xls_main = pd.ExcelFile(main_file)
    try:
        target_sheet = find_sheet(xls_main, sheet_kw)
    except ValueError:
        st.warning(f"⚠️ 未找到sheet「{sheet_kw}」, 跳过")
        continue
    main_df = pd.read_excel(xls_main, sheet_name=target_sheet, header=1)
    contract_col_main = find_col(main_df, "合同")
    if not contract_col_main:
        st.error(f"❌ sheet「{sheet_kw}」未找到合同列")
        continue

    # 放款明细对照
    merged_fk, mask_fk = compare_fields_vectorized(main_df, fk_df, contract_col_main, contract_col_fk, mapping_fk)
    # 字段表对照
    merged_zd, mask_zd = compare_fields_vectorized(main_df, zd_df, contract_col_main, contract_col_zd, mapping_zd, tolerance_dict={"保证金比例":0.005})
    # 二次明细对照
    merged_ec, mask_ec = compare_fields_vectorized(main_df, ec_df, contract_col_main, contract_col_ec, mapping_ec)
    # 重卡数据对照
    merged_zk, mask_zk = compare_fields_vectorized(main_df, zk_df, contract_col_main, contract_col_zk, mapping_zk)

    # 合并所有mask
    mask_total = pd.concat([mask_fk, mask_zd, mask_ec, mask_zk], axis=1).any(axis=1)
    total_errors = mask_total.sum()
    total_all += total_errors
    elapsed_all += time.time() - start_time

    # 导出Excel并标红
    wb = Workbook()
    ws = wb.active
    for c_idx, c in enumerate(main_df.columns, 1): ws.cell(1, c_idx, c)
    for r_idx, row in enumerate(main_df.itertuples(index=False), 2):
        for c_idx, v in enumerate(row, 1):
            ws.cell(r_idx, c_idx, v)
            # 若该行任何列在mask_total为True，则标黄合同号
            if mask_total.get(r_idx-2, False) and c_idx == list(main_df.columns).index(contract_col_main)+1:
                ws.cell(r_idx, c_idx).fill = yellow_fill
    # 标红字段
    for df_mask in [mask_fk, mask_zd, mask_ec, mask_zk]:
        for col in df_mask.columns:
            col_idx = list(main_df.columns).index(col)
            for row_idx in df_mask.index[df_mask[col]]:
                ws.cell(row_idx+2, col_idx+1).fill = red_fill

    # 保存到BytesIO
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    st.download_button(f"📥 下载 {sheet_kw} 审核标注版", output, f"{sheet_kw}_审核标注版.xlsx")
    st.success(f"✅ sheet「{sheet_kw}」检查完成，错误数 {total_errors}")

st.success(f"🎯 全部sheet完成，总错误数 {total_all}, 总耗时 {elapsed_all:.2f} 秒")
