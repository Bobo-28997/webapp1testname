# =====================================
# Streamlit Web App: 合同记录表自动审核（向量化 + 四Sheet + 漏填检查 + 驻店客户版）
# 修正版：修复列读取、_ref 回退、exact 匹配、索引对齐
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
    """在 DataFrame 中根据关键字（或 exact 精确匹配）查找列名（返回真实列名）"""
    if df is None or len(df.columns) == 0:
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

def normalize_num(val):
    if pd.isna(val): return None
    s = str(val).replace(",", "").strip()
    if s in ["", "-", "nan"]: return None
    try:
        if "%" in s:
            return float(s.replace("%", "")) / 100
        return float(s)
    except Exception:
        return None

def same_date_ymd(a, b):
    try:
        da = pd.to_datetime(a, errors='coerce')
        db = pd.to_datetime(b, errors='coerce')
        if pd.isna(da) or pd.isna(db): return False
        return (da.year, da.month, da.day) == (db.year, db.month, db.day)
    except Exception:
        return False

def read_excel_clean(file, sheet_name=None, header=0):
    """
    更稳健的读取：
      - 如果给定 sheet_name 且存在则读取它
      - 否则读取第一个 sheet
      - header 可指定（记录表 header=1）
      - 列名去前后空格
      - 发生异常返回空 DataFrame（并在 Streamlit 中显示错误）
    """
    try:
        xl = pd.ExcelFile(file)
        if sheet_name is None:
            sheet = xl.sheet_names[0]
        else:
            # 尝试模糊匹配 sheet_name（包含关系）
            sheet = None
            for s in xl.sheet_names:
                if sheet_name in s:
                    sheet = s
                    break
            if sheet is None:
                # 回退到第一个 sheet，但提示
                sheet = xl.sheet_names[0]
                st.info(f"⚠️ 指定 sheet 名「{sheet_name}」未找到，使用第一个 sheet：{sheet}")
        df = pd.read_excel(xl, sheet_name=sheet, header=header)
        # 规范化列名（去前后空格）
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception as e:
        st.error(f"❌ 读取 Excel 失败：{e}")
        return pd.DataFrame()

# =====================================
# 📖 文件定位与读取（记录表 header=1，参考表 header=0）
# =====================================
main_file = find_file(uploaded_files, "记录表")
fk_file = find_file(uploaded_files, "放款明细")
zd_file = find_file(uploaded_files, "字段")
ec_file = find_file(uploaded_files, "二次明细")
zk_file = find_file(uploaded_files, "重卡数据")

xls_main = pd.ExcelFile(main_file)
sheet_keywords = ["二次", "部分担保", "随州", "驻店客户"]

# 参考表：第一行为表头
fk_df = read_excel_clean(fk_file, sheet_name=find_sheet(pd.ExcelFile(fk_file), "本司"), header=0)
zd_df = read_excel_clean(zd_file, sheet_name=find_sheet(pd.ExcelFile(zd_file), "重卡"), header=0)
ec_df = read_excel_clean(ec_file, header=0)
zk_df = read_excel_clean(zk_file, header=0)

# =====================================
# 📌 合同列（参考表：使用 exact=True 更稳健）
# =====================================
contract_col_fk = find_col(fk_df, "合同", exact=True)
contract_col_zd = find_col(zd_df, "合同", exact=True)
contract_col_ec = find_col(ec_df, "合同", exact=True)
contract_col_zk = find_col(zk_df, "合同", exact=True)

# =====================================
# 🔗 字段映射（主表名 -> 参考表名）
# =====================================
mapping_fk = {"授信方":"授信", "租赁本金":"本金", "租赁期限月":"租赁期限月",
              "客户经理":"客户经理", "起租收益率":"收益率", "主车台数":"主车台数", "挂车台数":"挂车台数"}

mapping_zd = {"保证金比例":"保证金比例_2", "项目提报人":"提报", "起租时间":"起租日_商",
              "租赁期限月":"总期数_商_资产", "起租收益率":"XIRR_商_起租", "所属省区":"区域", "城市经理":"城市经理"}

mapping_ec = {"二次时间":"出本流程时间"}
mapping_zk = {"授信方":"授信方"}

# =====================================
# ⚡ 向量化比对函数（更稳健）
# =====================================
def compare_fields_vectorized(main_df, ref_df, main_contract_col, ref_contract_col, mapping_dict, tolerance_dict=None):
    """
    向量化比对（稳健版）：
    - 先通过 find_col 找到主/参考表真实列名（支持 exact 指定）
    - merge 后安全读取参考列（尝试 ref_col_ref -> 回退到 ref_col）
    - 返回 merged（用于 debug / 后续扩展）与 mask（列名为主表真实列名）
    """
    tolerance_dict = tolerance_dict or {}

    # 保证 main_df 的索引连续 0..n-1，便于 later 使用位置索引
    main_df = main_df.reset_index(drop=True)
    main_df_clean = main_df.copy()
    main_df_clean[main_contract_col] = main_df_clean[main_contract_col].astype(str).str.strip()

    ref_df_clean = ref_df.copy()
    ref_df_clean[ref_contract_col] = ref_df_clean[ref_contract_col].astype(str).str.strip()

    # 准备 ref 子表（若缺列则提示并返回空 mask）
    ref_cols_needed = [ref_contract_col] + list(mapping_dict.values())
    missing_cols = [c for c in ref_cols_needed if c not in ref_df_clean.columns]
    if missing_cols:
        st.error(f"❌ 参考表缺少列: {missing_cols}")
        mask_empty = pd.DataFrame(False, index=main_df.index, columns=[
            # 把 mask 列名设为主表真实列名（尽量找出）
            (find_col(main_df, mk, exact=(mk in ["客户经理","城市经理"])) or mk)
            for mk in mapping_dict.keys()
        ])
        return main_df_clean.copy(), mask_empty

    ref_sub = ref_df_clean[ref_cols_needed]

    # merge（左连接）
    merged = main_df_clean.merge(ref_sub, how="left",
                                 left_on=main_contract_col, right_on=ref_contract_col,
                                 suffixes=("", "_ref"))

    # mask 列名使用主表真实列名（避免 mapping 键与实际列名不一致造成混乱）
    actual_main_cols = []
    for mk in mapping_dict.keys():
        actual = find_col(main_df, mk, exact=(mk in ["客户经理","城市经理"]))
        if actual is None:
            # 回退使用 mapping key 本身（可能主表缺列）
            actual = mk
        actual_main_cols.append(actual)

    mask = pd.DataFrame(False, index=merged.index, columns=actual_main_cols)

    # 逐字段做向量化比较（不使用 Python 层面的逐行循环）
    for mk, rv in mapping_dict.items():
        main_col = find_col(main_df, mk, exact=(mk in ["客户经理","城市经理"]))
        ref_col = find_col(ref_df, rv, exact=(rv in ["客户经理","城市经理"]))
        if main_col is None:
            # 主表没有此列，跳过（已在 mask 中保留为 False）
            continue
        # 确定参考列在 merged 中的列名：先尝试带 _ref 的版本（当列名冲突时 pandas 会添加后缀）
        ref_col_in_merged = f"{ref_col}_ref" if f"{ref_col}_ref" in merged.columns else ref_col

        # 若参考列不在 merged（极少见），将整列标为 False 并继续
        if ref_col_in_merged not in merged.columns:
            continue

        main_vals = merged[main_col]
        ref_vals = merged[ref_col_in_merged]

        is_date_col = any(k in mk for k in ["日期", "时间"]) or any(k in rv for k in ["日期", "时间"])
        tol = tolerance_dict.get(mk, 0)
        exact_match = mk in ["客户经理", "城市经理"]

        # 日期比较（向量化）
        if is_date_col:
            main_dt = pd.to_datetime(main_vals, errors='coerce').dt.normalize()
            ref_dt = pd.to_datetime(ref_vals, errors='coerce').dt.normalize()
            # 注意：NaT 比较会返回 False，所以我们把 NaT-NaT 视为相等
            date_mismatch = ~(main_dt.eq(ref_dt) | (main_dt.isna() & ref_dt.isna()))
            mask.loc[date_mismatch.index, main_col] = date_mismatch.fillna(False)
            continue

        # 非日期：先尝试数值比较（使用 normalize_num 向量化）
        main_num = main_vals.apply(normalize_num)
        ref_num = ref_vals.apply(normalize_num)

        # 数值都存在时按容差比较
        both_num = main_num.notna() & ref_num.notna()
        num_mismatch = pd.Series(False, index=merged.index)
        if both_num.any():
            # 注意： main_num/ref_num 是混合类型的 Series（可能包含 None），先转换为 float where possible
            try:
                # compute difference for numeric positions
                diff = (main_num - ref_num).abs()
                num_mismatch = both_num & (diff > tol)
            except Exception:
                # 若减法失败（非数值），保持 False
                num_mismatch = both_num & False

        # 文本/其他比较（含 exact 逻辑）
        # 当数值比较不适用（任一为 Na），按字符串比较（忽略大小写与尾部 .0）
        text_mask = pd.Series(False, index=merged.index)
        non_num_positions = ~(both_num)
        if non_num_positions.any():
            a_str = main_vals.astype(str).fillna("").str.strip().str.lower().str.replace(".0", "")
            b_str = ref_vals.astype(str).fillna("").str.strip().str.lower().str.replace(".0", "")
            if exact_match:
                text_mask = non_num_positions & (~a_str.eq(b_str))
            else:
                text_mask = non_num_positions & (~a_str.eq(b_str))

        # Na/NotNa 不等也视为 mismatch
        nan_mismatch = (main_num.isna() ^ ref_num.isna())

        mask_col_result = num_mismatch | text_mask | nan_mismatch
        mask.loc[mask_col_result.index, main_col] = mask_col_result.fillna(False)

    return merged, mask

# =====================================
# 🧮 单 sheet 检查函数（使用向量化比对）
# =====================================
def check_one_sheet(sheet_keyword):
    start_time = time.time()
    # 记录表 sheet header=1（第二行为列名）
    main_df = read_excel_clean(main_file, sheet_name=find_sheet(xls_main, sheet_keyword), header=1)
    # reset index 0..n-1 以便 mask/merged 对齐
    main_df = main_df.reset_index(drop=True)

    # 准备输出 Excel（保留空行以与原版一致：数据从 excel 行 3 开始）
    output_path = f"记录表_{sheet_keyword}_审核标注版.xlsx"
    empty_row = pd.DataFrame([[""] * len(main_df.columns)], columns=main_df.columns)
    pd.concat([empty_row, main_df], ignore_index=True).to_excel(output_path, index=False)
    wb = load_workbook(output_path)
    ws = wb.active

    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    # 合同列（强制 exact 匹配）
    global contract_col_main
    contract_col_main = find_col(main_df, "合同", exact=False)
    if not contract_col_main:
        st.error(f"❌ 在「{sheet_keyword}」中未找到合同列。")
        return 0, None, 0, set()

    contracts_seen = set()
    progress = st.progress(0)
    status = st.empty()

    # 向量化比对四张参考表
    merged_fk, mask_fk = compare_fields_vectorized(main_df, fk_df, contract_col_main, contract_col_fk, mapping_fk, tolerance_dict={})
    merged_zd, mask_zd = compare_fields_vectorized(main_df, zd_df, contract_col_main, contract_col_zd, mapping_zd, tolerance_dict={"保证金比例": 0.005})
    merged_ec, mask_ec = compare_fields_vectorized(main_df, ec_df, contract_col_main, contract_col_ec, mapping_ec, tolerance_dict={})
    merged_zk, mask_zk = compare_fields_vectorized(main_df, zk_df, contract_col_main, contract_col_zk, mapping_zk, tolerance_dict={})

    # 合并所有 mask（列名是主表实际列名）
    mask_all = pd.concat([mask_fk, mask_zd, mask_ec, mask_zk], axis=1).fillna(False)
    mask_any = mask_all.any(axis=1)

    # 标红 / 标黄（使用 main_df 的位置索引 r_idx 对应 excel 行 r_idx+3）
    for r_idx, row in main_df.iterrows():
        contracts_seen.add(str(row[contract_col_main]).strip() if not pd.isna(row[contract_col_main]) else "")
        # 标红：逐列检查 mask_all（mask_all 列名是主表实际列名）
        for col in mask_all.columns:
            try:
                if mask_all.at[r_idx, col]:
                    c_idx = list(main_df.columns).index(col) + 1
                    ws.cell(r_idx + 3, c_idx).fill = red_fill
            except Exception:
                # 若某些 mask 列不在主表列中（极少），跳过
                continue
        # 标黄合同号（若该行任一列出错）
        try:
            if mask_any.at[r_idx]:
                c_contract = list(main_df.columns).index(contract_col_main) + 1
                ws.cell(r_idx + 3, c_contract).fill = yellow_fill
        except Exception:
            pass

        # 进度显示（每 10 行刷新）
        if (r_idx + 1) % 10 == 0:
            status.text(f"检查「{sheet_keyword}」... {r_idx+1}/{len(main_df)}")
        progress.progress((r_idx + 1) / max(1, len(main_df)))

    # 导出下载
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    st.download_button(
        label=f"📥 下载 {sheet_keyword} 审核标注版",
        data=output,
        file_name=f"记录表_{sheet_keyword}_审核标注版.xlsx"
    )

    total_errors = int(mask_any.sum())
    elapsed = time.time() - start_time
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
# 🕵️ 漏填检查（保留原逻辑：排除“是否车管家=是”与“提成类型=联合租赁/驻店”）
# =====================================
field_contracts = zd_df[contract_col_zd].dropna().astype(str).str.strip()
col_car_manager = find_col(zd_df, "是否车管家", exact=True)
col_bonus_type = find_col(zd_df, "提成类型", exact=True)

missing_contracts_mask = (~field_contracts.isin(contracts_seen_all_sheets))
if col_car_manager:
    missing_contracts_mask &= ~(zd_df[col_car_manager].astype(str).str.strip().str.lower() == "是")
if col_bonus_type:
    missing_contracts_mask &= ~(
        zd_df[col_bonus_type].astype(str).str.strip().isin(["联合租赁","驻店"])
    )

zd_df_missing = zd_df.copy()
zd_df_missing["漏填检查"] = ""
zd_df_missing.loc[missing_contracts_mask, "漏填检查"] = "❗ 漏填"
漏填合同数 = zd_df_missing["漏填检查"].eq("❗ 漏填").sum()
st.warning(f"⚠️ 共发现 {漏填合同数} 个合同在记录表中未出现（已排除车管家、联合租赁、驻店）")

# =====================================
# 📤 导出字段表（含漏填标注 + 仅漏填）
# =====================================
yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
wb_all = Workbook()
ws_all = wb_all.active
for c_idx, c in enumerate(zd_df_missing.columns, 1):
    ws_all.cell(1, c_idx, c)
for r_idx, row in enumerate(zd_df_missing.itertuples(index=False), 2):
    for c_idx, v in enumerate(row, 1):
        ws_all.cell(r_idx, c_idx, v)
        if zd_df_missing.columns[c_idx-1] == "漏填检查" and v == "❗ 漏填":
            ws_all.cell(r_idx, c_idx).fill = yellow_fill
output_all = BytesIO()
wb_all.save(output_all)
output_all.seek(0)
st.download_button("📥 下载字段表漏填标注版", output_all, "字段表_漏填标注版.xlsx")

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
