import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side, Font
import io
import os
import re

# ─────────────────────────────────────────────────────────────
# 페이지 설정
# ─────────────────────────────────────────────────────────────
st.set_page_config(page_title="독성정보 자동 추출 시스템", layout="wide")
st.title("🧪 화학물질 독성정보 자동 추출 서비스")
st.info("내부식별자를 입력하면 DB에서 독성정보를 추출하여 엑셀 파일을 생성합니다.")

# ─────────────────────────────────────────────────────────────
# GitHub 파일명
# ─────────────────────────────────────────────────────────────
DB_FILENAME = "유해성미확인물질 12종 DB.xlsx"
TPL_SINGLE  = "개별물질 추출 템플릿.xlsx"
TPL_MULTI   = "다중물질 추출 템플릿.xlsx"


# ─────────────────────────────────────────────────────────────
# 공통 유틸
# ─────────────────────────────────────────────────────────────

def write_safe(ws, row, col, value):
    cell = ws.cell(row=row, column=col)
    for merged in ws.merged_cells.ranges:
        if cell.coordinate in merged:
            ws.cell(row=merged.min_row, column=merged.min_col).value = value
            return
    cell.value = value


# ─────────────────────────────────────────────────────────────
# VEGA 우선순위 로직
# ─────────────────────────────────────────────────────────────
VEGA_PRIORITY_ORDER = [
    "EXPERIMENTAL value",
    "GOOD reliability",
    "MODERATE reliability",
    "LOW reliability",
]

def vega_domain_rank(domain_val):
    """Domain status → 우선순위 점수 (높을수록 우선)"""
    d = str(domain_val)
    for i, label in enumerate(VEGA_PRIORITY_ORDER):
        if label.lower() in d.lower():
            return len(VEGA_PRIORITY_ORDER) - i  # EXPERIMENTAL=4, GOOD=3, MODERATE=2, LOW=1
    return 0

def vega_domain_score(domain_val):
    """같은 등급 내 괄호 안 숫자 추출 (클수록 우선)"""
    match = re.search(r'\(([0-9.]+)\)', str(domain_val))
    return float(match.group(1)) if match else 0.0

def get_best_vega(df):
    """VEGA 데이터에서 우선순위에 따라 최적 1행 반환"""
    if df.empty:
        return None
    temp = df.copy()
    temp['_rank']  = temp['Domain status'].apply(vega_domain_rank)
    temp['_score'] = temp['Domain status'].apply(vega_domain_score)
    temp = temp.sort_values(by=['_rank', '_score'], ascending=[False, False])
    return temp.iloc[0]


# ─────────────────────────────────────────────────────────────
# 단일 추출 로직
# ─────────────────────────────────────────────────────────────

def apply_priority(df, cat, method, exp_species=None):
    if len(df) <= 1:
        return df.iloc[0]
    temp = df.copy()
    if method == "실험값":
        if cat == "급성경구독성":
            temp['p1'] = (temp['Endpoint(표준)'] == 'LD50').astype(int)
            temp['p2'] = (temp['시험종(표준)'] == 'Rat').astype(int)
            temp['p3'] = temp['시험지침'].astype(str).str.contains('401', na=False).astype(int)
            temp = temp.sort_values(['p1','p2','p3','Result'], ascending=[False,False,False,True])
        elif cat == "급성흡입독성":
            temp['p1'] = (temp['Endpoint(표준)'] == 'LC50').astype(int)
            temp['p2'] = (temp['시험종(표준)'] == 'Rat').astype(int)
            temp['p3'] = (temp['Duration(표준)'] == '4 h').astype(int)
            temp['p4'] = temp['시험지침'].astype(str).str.contains('403', na=False).astype(int)
            temp = temp.sort_values(['p1','p2','p3','p4','Result'], ascending=[False,False,False,False,True])
        elif cat == "어류급성독성":
            temp['p1'] = (temp['Endpoint(표준)'] == 'LC50').astype(int)
            temp['p2'] = temp['시험종(표준)'].isin(['Fathead minnow','Zebrafish','Rainbow trout']).astype(int)
            temp['p3'] = (temp['Duration(표준)'] == '96 h').astype(int)
            temp['p4'] = temp['시험지침'].astype(str).str.contains('203', na=False).astype(int)
            temp = temp.sort_values(['p1','p2','p3','p4','Result'], ascending=[False,False,False,False,True])
        elif cat == "물벼룩급성독성":
            temp['p1'] = (temp['Endpoint(표준)'] == 'EC50').astype(int)
            temp['p2'] = (temp['시험종(표준)'] == 'Daphnia magna').astype(int)
            temp['p3'] = (temp['Duration(표준)'] == '48 h').astype(int)
            temp['p4'] = temp['시험지침'].astype(str).str.contains('202', na=False).astype(int)
            temp = temp.sort_values(['p1','p2','p3','p4','Result'], ascending=[False,False,False,False,True])
        elif cat == "담수조류생장저해":
            temp['p1'] = (temp['Endpoint(표준)'] == 'EC50').astype(int)
            temp['p2'] = temp['시험종(표준)'].isin(['P. subcapitata','D. subspicatus']).astype(int)
            temp['p3'] = (temp['Duration(표준)'] == '72 h').astype(int)
            temp['p4'] = temp['시험지침'].astype(str).str.contains('201', na=False).astype(int)
            temp = temp.sort_values(['p1','p2','p3','p4','Result'], ascending=[False,False,False,False,True])
    elif method == "QSAR":
        model_map = {
            "급성경구독성":      "Acute toxicity in Rat, Oral - Danish QSAR DB ACDLabs model (v1.0)",
            "담수조류생장저해":  "Pseudokirchneriella subcapitata 72h EC50 - Danish QSAR DB battery model (v1.0)",
            "물벼룩급성독성":    "Daphnia magna 48h EC50 - Danish QSAR DB battery model (v1.0)",
            "복귀돌연변이":      "Ames test in S. typhimurium (in vitro) - Danish QSAR DB battery model (v1.0)",
            "소핵시험":          "Micronucleus Test in Mouse Erythrocytes - Danish QSAR DB battery model (v1.0)",
            "어류급성독성":      "Fathead minnow 96h LC50 - Danish QSAR DB battery model (v1.0)",
            "피부부식성/자극성": "BfR skin irritation/corrosion (v1.0)"
        }
        if cat == "포유류 배양세포를 이용한 염색체이상":
            mname = ("Chromosome Aberrations in Chinese Hamster Ovary (CHO) Cells - Danish QSAR DB battery model (v1.0)"
                     if exp_species == "CHO Cells"
                     else "Chromosome Aberrations in Chinese Hamster Lung (CHL) Cells - Danish QSAR DB battery model (v1.0)")
            temp['p_q'] = (temp['모델 종류 및 버전'] == mname).astype(int)
        elif cat in model_map:
            temp['p_q'] = (temp['모델 종류 및 버전'] == model_map[cat]).astype(int)
        else:
            temp['p_q'] = 0
        temp = temp.sort_values('p_q', ascending=False)
    return temp.iloc[0]


def format_val_single(row, cat, method):
    res = str(row['Result'])
    if method == "QSAR" and str(row.get('Domain status','')) == "Out of domain":
        res += " (Out of domain)"
    val_items = ["급성경구독성","급성흡입독성","어류급성독성","물벼룩급성독성","담수조류생장저해"]
    if cat in val_items:
        return f"{row['Endpoint(표준)']} = {res} {row['단위']} ({row['시험종(표준)']})"
    return res


def extract_single(target_id, df_mat, df_tox, wb):
    """
    개별물질 템플릿 열 매핑 (VEGA 추가 반영)
    col 4~8  : 실험값 (ECHA, US, Pub, K-reach, 환경부)
    col 9    : Read-across (TB)
    col 10   : QSAR Toolbox QSAR
    col 11   : Danish QSAR
    col 12   : VEGA (QSAR)        ← NEW
    col 13   : Epi suite
    col 14   : HAZMAP
    col 15   : Protox 3.0
    col 16   : Vega (AI-based)    ← NEW
    col 17   : Cheminfomatics
    """
    ws = wb.active
    categories = [
        "급성경구독성","급성흡입독성","피부부식성/자극성","복귀돌연변이",
        "포유류 배양세포를 이용한 염색체이상","소핵시험",
        "어류급성독성","물벼룩급성독성","담수조류생장저해","이분해성"
    ]
    exp_srcs = ["ECHA CHEM","US DashBoard","Pubchem","K-reach","환경부유해성심사결과"]
    ai_srcs  = ["HAZMAP","Protox 3.0"]   # col 14, 15  (Vega AI=16, Chemi=17 별도처리)

    mat_row = df_mat[df_mat['내부식별자'] == target_id]
    if mat_row.empty:
        raise ValueError(f"'{target_id}' 물질정보를 DB에서 찾을 수 없습니다.")
    t = mat_row.iloc[0]
    write_safe(ws, 7, 3, target_id)
    write_safe(ws, 7, 4, str(t['CAS']))
    write_safe(ws, 7, 5, str(t['물질명']))
    write_safe(ws, 7, 6, str(t['분자식']))
    write_safe(ws, 7, 7, str(t['분자량']))

    for r_idx, cat in enumerate(categories):
        df_cat = df_tox[(df_tox['내부식별자'] == target_id) & (df_tox['유해성항목'] == cat)]
        exp_species_found = None
        base_row = 12 + r_idx

        # 실험값 (col 4~8)
        for c_idx, src in enumerate(exp_srcs):
            df_src = df_cat[(df_cat['결과도출방법'] == '실험값') & (df_cat['출처'] == src)]
            if not df_src.empty:
                best = apply_priority(df_src, cat, "실험값")
                ws.cell(row=base_row, column=4+c_idx).value = format_val_single(best, cat, "실험값")
                if cat == "포유류 배양세포를 이용한 염색체이상":
                    exp_species_found = best['시험종(표준)']

        # Read-across (col 9)
        df_ra = df_cat[(df_cat['결과도출방법'] == 'Read-across') & (df_cat['출처'] == 'QSAR Toolbox v.4.8')]
        if not df_ra.empty:
            ws.cell(row=base_row, column=9).value = format_val_single(df_ra.iloc[0], cat, "Read-across")

        # QSAR Toolbox QSAR (col 10)
        df_tb = df_cat[(df_cat['결과도출방법'] == 'QSAR') & (df_cat['출처'] == 'QSAR Toolbox v.4.8')]
        if not df_tb.empty:
            ws.cell(row=base_row, column=10).value = format_val_single(df_tb.iloc[0], cat, "QSAR")

        # Danish QSAR (col 11)
        df_danish = df_cat[(df_cat['결과도출방법'] == 'QSAR') & (df_cat['출처'] == 'Danish QSAR')]
        if not df_danish.empty:
            best = apply_priority(df_danish, cat, "QSAR", exp_species_found)
            ws.cell(row=base_row, column=11).value = format_val_single(best, cat, "QSAR")

        # VEGA - QSAR (col 12)
        df_vega_qsar = df_cat[(df_cat['출처'] == 'VEGA') & (df_cat['결과도출방법'] == 'QSAR')]
        if not df_vega_qsar.empty:
            best = get_best_vega(df_vega_qsar)
            if best is not None:
                ws.cell(row=base_row, column=12).value = format_val_single(best, cat, "QSAR")

        # Epi suite (col 13)
        df_epi = df_cat[(df_cat['결과도출방법'] == 'QSAR') & (df_cat['출처'] == 'Epi suite')]
        if not df_epi.empty:
            ws.cell(row=base_row, column=13).value = format_val_single(df_epi.iloc[0], cat, "QSAR")

        # HAZMAP, Protox (col 14, 15)
        for c_idx, src in enumerate(ai_srcs):
            df_src = df_cat[(df_cat['결과도출방법'] == 'AI-based QSAR') & (df_cat['출처'] == src)]
            if not df_src.empty:
                ws.cell(row=base_row, column=14+c_idx).value = format_val_single(df_src.iloc[0], cat, "AI-based QSAR")

        # VEGA - AI-based QSAR (col 16)
        df_vega_ai = df_cat[(df_cat['출처'] == 'VEGA') & (df_cat['결과도출방법'] == 'AI-based QSAR')]
        if not df_vega_ai.empty:
            best = get_best_vega(df_vega_ai)
            if best is not None:
                ws.cell(row=base_row, column=16).value = format_val_single(best, cat, "AI-based QSAR")

        # Cheminfomatics (col 17)
        df_chemi = df_cat[(df_cat['결과도출방법'] == 'AI-based QSAR') & (df_cat['출처'] == 'Cheminfomatics')]
        if not df_chemi.empty:
            ws.cell(row=base_row, column=17).value = format_val_single(df_chemi.iloc[0], cat, "AI-based QSAR")

    # 스타일
    thin = Border(left=Side(style='thin'), right=Side(style='thin'),
                  top=Side(style='thin'),  bottom=Side(style='thin'))
    for rng in [ws['C7:G7'], ws['B11:Q21']]:
        for row in rng:
            for cell in row:
                cell.border    = thin
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                cell.font      = Font(name='맑은 고딕', size=9)
    col_widths = {'B':12,'C':15,'D':22,'E':25,'F':12,'G':12,'H':22,
                  'I':18,'J':20,'K':20,'L':20,'M':20,'N':15,'O':15,'P':15,'Q':15}
    for col, w in col_widths.items():
        ws.column_dimensions[col].width = w
    for i in range(12, 22):
        ws.row_dimensions[i].height = 45


# ─────────────────────────────────────────────────────────────
# 다중 추출 로직
# ─────────────────────────────────────────────────────────────

def format_biodeg(row):
    if row['출처'] in ['환경부유해성심사결과','K-reach'] or \
       (row['결과도출방법'] == 'QSAR' and row['출처'] == 'Epi suite'):
        return str(row['Result'])
    try:
        val = float(row['Result'])
        ep  = str(row['Endpoint']).lower()
        threshold = 70 if 'doc' in ep else 60
        status = "positive(이분해성)" if val >= threshold else "negative(난분해성)"
        return f"{status} - {row['Endpoint']} = {row['Result']} {row['단위']}"
    except:
        return str(row['Result'])


def format_standard(row, cat):
    res  = str(row['Result'])
    ep   = row['Endpoint']    if pd.notna(row.get('Endpoint'))    else (row.get('Endpoint(표준)','Unknown') or 'Unknown')
    sp   = row['시험종(표준)'] if pd.notna(row.get('시험종(표준)')) else (row.get('시험종','Unknown')          or 'Unknown')
    unit = row['단위']         if pd.notna(row.get('단위'))         else ""
    if "(Out of domain)" not in res and \
       pd.notna(row.get('Domain status')) and str(row.get('Domain status')) == "Out of domain":
        res += " (Out of domain)"
    val_cats = ["급성경구독성","급성흡입독성","어류급성독성","물벼룩급성독성","담수조류생장저해"]
    if cat in val_cats:
        return f"{ep} = {res} {unit} ({sp})"
    return res


def get_best_row_multi(df, cat, src_key):
    if df.empty:
        return None
    temp = df.copy()
    temp['result_num'] = pd.to_numeric(temp['Result'], errors='coerce').fillna(999999)
    if "Cheminfomatics" in src_key:
        cons = temp[temp['모델 종류 및 버전'].astype(str).str.contains('Consensus', case=False, na=False)]
        return cons.iloc[0] if not cons.empty else temp.iloc[0]
    if cat == '이분해성':
        def gl_score(v):
            v = str(v).upper()
            return 2 if 'OECD' in v else (1 if v not in ['-','','NAN'] else 0)
        temp['gl_score'] = temp['시험지침'].apply(gl_score)
        temp = temp.sort_values(by=['gl_score','result_num'], ascending=[False,False])
        return temp.iloc[0]
    if cat in ["급성경구독성","급성흡입독성","어류급성독성","물벼룩급성독성","담수조류생장저해"]:
        target_ep = "LD50" if "경구" in cat else ("LC50" if "어류" in cat or "흡입" in cat else "EC50")
        temp['ep_score'] = (
            temp['Endpoint'].astype(str).str.contains(target_ep, case=False, na=False) |
            temp['Endpoint(표준)'].astype(str).str.contains(target_ep, case=False, na=False)
        ).astype(int) * 10
        t_sp = ("Rat"            if "경구" in cat or "흡입" in cat else
                "Fathead minnow" if "어류" in cat else
                "Daphnia magna"  if "물벼룩" in cat else "P. subcapitata")
        temp['sp_score'] = temp['시험종(표준)'].astype(str).str.contains(t_sp, case=False, na=False).astype(int) * 5
        temp['total_score'] = temp['ep_score'] + temp['sp_score']
        temp = temp.sort_values(by=['total_score','result_num'], ascending=[False,True])
        return temp.iloc[0]
    return temp.iloc[0]


def filter_skin_exp(df):
    temp = df[df['Result'].astype(str).str.lower().isin(['positive','negative'])]
    if not temp.empty:
        rabbit = temp[temp['시험종(표준)'].astype(str).str.contains('Rabbit', case=False, na=False)]
        return rabbit.iloc[0] if not rabbit.empty else temp.iloc[0]
    return None


def get_final_value_multi(best, cat, src_key):
    if cat == '이분해성':
        return format_biodeg(best)
    elif cat == '피부부식성/자극성' and "QSAR" in str(best.get('결과도출방법','')):
        val = str(best['Result'])
        if "(Out of domain)" not in val and str(best.get('Domain status')) == "Out of domain":
            val += " (Out of domain)"
        return val
    else:
        return format_standard(best, cat)


# 다중 템플릿 레이아웃 (VEGA 추가 반영)
BLOCK_HEADER_ROWS = [2, 15]
INFO_OFFSETS = {'내부식별자':1, 'CAS No.':3, '물질명':5, '분자식':7, '분자량':9}
INFO_COL = 2
CAT_OFFSETS = {
    '급성경구독성':2, '급성흡입독성':3, '피부부식성/자극성':4,
    '복귀돌연변이':5, '포유류 배양세포를 이용한 염색체이상':6, '소핵시험':7,
    '어류급성독성':8, '물벼룩급성독성':9, '담수조류생장저해':10, '이분해성':11,
}

# 다중 템플릿 열 매핑
# F=6(ECHA), G=7(US), H=8(Pub), I=9(K-reach), J=10(환경부)
# K=11(TB_RA), L=12(TB_QSAR), M=13(Danish), N=14(VEGA_QSAR), O=15(Epi)
# P=16(HAZMAP), Q=17(Protox), R=18(Vega_AI), S=19(Chemi)
MULTI_SRC_COLS = {
    'ECHA CHEM':6, 'US DashBoard':7, 'Pubchem':8, 'K-reach':9, '환경부유해성심사결과':10,
    'QSAR_RA':11, 'QSAR_QSAR':12, 'Danish QSAR':13,
    'VEGA_QSAR':14,
    'Epi suite':15,
    'HAZMAP':16, 'Protox 3.0':17,
    'VEGA_AI':18,
    'Cheminfomatics':19,
}


def extract_multi(tid1, tid2, df_mat, df_tox, wb):
    target_ids = [tid1, tid2]
    ws = wb.active
    ws.title = f"{tid1} 및 {tid2}"

    for hdr in BLOCK_HEADER_ROWS:
        for offset in INFO_OFFSETS.values():
            ws.cell(row=hdr + offset, column=INFO_COL).value = None
        for offset in CAT_OFFSETS.values():
            for col in range(6, 20):
                ws.cell(row=hdr + offset, column=col).value = None

    for tid, hdr_row in zip(target_ids, BLOCK_HEADER_ROWS):
        mat_row = df_mat[df_mat['내부식별자'] == tid]
        if mat_row.empty:
            raise ValueError(f"'{tid}' 물질정보를 DB에서 찾을 수 없습니다.")
        t = mat_row.iloc[0]
        info_vals = {
            '내부식별자': tid, 'CAS No.': str(t['CAS']),
            '물질명': str(t['물질명']), '분자식': str(t['분자식']),
            '분자량': f"{t['분자량']} g/mol",
        }
        for label, offset in INFO_OFFSETS.items():
            write_safe(ws, hdr_row + offset, INFO_COL, info_vals[label])

        df_sub = df_tox[df_tox['내부식별자'] == tid]

        for cat, cat_offset in CAT_OFFSETS.items():
            data_row = hdr_row + cat_offset
            df_cat   = df_sub[df_sub['유해성항목'] == cat]

            for src_key, col_idx in MULTI_SRC_COLS.items():

                # VEGA 별도 처리
                if src_key == 'VEGA_QSAR':
                    df_src = df_cat[(df_cat['출처'] == 'VEGA') & (df_cat['결과도출방법'] == 'QSAR')]
                    if not df_src.empty:
                        best = get_best_vega(df_src)
                        if best is not None:
                            write_safe(ws, data_row, col_idx, get_final_value_multi(best, cat, src_key))
                    continue

                if src_key == 'VEGA_AI':
                    df_src = df_cat[(df_cat['출처'] == 'VEGA') & (df_cat['결과도출방법'] == 'AI-based QSAR')]
                    if not df_src.empty:
                        best = get_best_vega(df_src)
                        if best is not None:
                            write_safe(ws, data_row, col_idx, get_final_value_multi(best, cat, src_key))
                    continue

                # QSAR Toolbox 분리
                if src_key == 'QSAR_RA':
                    df_src = df_cat[
                        df_cat['출처'].astype(str).str.contains('QSAR Toolbox', case=False, na=False) &
                        df_cat['결과도출방법'].astype(str).str.contains('Read across', case=False, na=False)
                    ]
                elif src_key == 'QSAR_QSAR':
                    df_src = df_cat[
                        df_cat['출처'].astype(str).str.contains('QSAR Toolbox', case=False, na=False) &
                        ~df_cat['결과도출방법'].astype(str).str.contains('Read across', case=False, na=False)
                    ]
                else:
                    df_src = df_cat[df_cat['출처'].astype(str).str.contains(src_key, case=False, na=False)]

                if cat == '이분해성' and src_key not in ['Epi suite','환경부유해성심사결과','K-reach']:
                    df_src = df_src[df_src['Endpoint'].notna()]

                if df_src.empty:
                    continue

                if cat == '피부부식성/자극성' and src_key not in ['QSAR_RA','QSAR_QSAR','Danish QSAR']:
                    best = filter_skin_exp(df_src)
                else:
                    best = get_best_row_multi(df_src, cat, src_key)

                if best is not None:
                    write_safe(ws, data_row, col_idx, get_final_value_multi(best, cat, src_key))

    # 스타일
    thin = Border(left=Side(style='thin'), right=Side(style='thin'),
                  top=Side(style='thin'),  bottom=Side(style='thin'))
    for hdr in BLOCK_HEADER_ROWS:
        for r in range(hdr + 2, hdr + 12):
            for c in range(6, 20):
                cell = ws.cell(row=r, column=c)
                cell.border    = thin
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                cell.font      = Font(name='맑은 고딕', size=9)


# ─────────────────────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────────────────────

if not os.path.exists(DB_FILENAME):
    st.error(f"DB 파일을 찾을 수 없습니다: **{DB_FILENAME}**")
    st.stop()

mode = st.radio("📋 추출 모드 선택", ["단일 물질 추출", "다중 물질 추출 (2개)"], horizontal=True)
st.divider()

if mode == "단일 물질 추출":
    if not os.path.exists(TPL_SINGLE):
        st.error(f"템플릿 파일 없음: **{TPL_SINGLE}**")
        st.stop()
    target_id = st.text_input("🔍 내부식별자 입력 (예: B-3)", value="B-3")
    if st.button("🚀 추출 및 엑셀 다운로드", key="btn_single"):
        with st.spinner("데이터 추출 중..."):
            try:
                df_mat = pd.read_excel(DB_FILENAME, sheet_name='물질정보')
                df_tox = pd.read_excel(DB_FILENAME, sheet_name='유해성정보')
                wb     = load_workbook(TPL_SINGLE)
                extract_single(target_id.strip(), df_mat, df_tox, wb)
                buf = io.BytesIO()
                wb.save(buf)
                st.success(f"✅ **{target_id}** 추출 완료!")
                st.download_button(
                    label="📥 결과 엑셀 다운로드",
                    data=buf.getvalue(),
                    file_name=f"추출결과_{target_id}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"오류 발생: {e}")

else:
    if not os.path.exists(TPL_MULTI):
        st.error(f"템플릿 파일 없음: **{TPL_MULTI}**")
        st.stop()
    col1, col2 = st.columns(2)
    with col1:
        tid1 = st.text_input("🔍 첫 번째 내부식별자 (예: B-1)", value="B-1")
    with col2:
        tid2 = st.text_input("🔍 두 번째 내부식별자 (예: B-3)", value="B-3")
    if st.button("🚀 추출 및 엑셀 다운로드", key="btn_multi"):
        if not tid1.strip() or not tid2.strip():
            st.warning("두 개의 내부식별자를 모두 입력해주세요.")
        elif tid1.strip() == tid2.strip():
            st.warning("서로 다른 내부식별자를 입력해주세요.")
        else:
            with st.spinner("데이터 추출 중..."):
                try:
                    df_mat = pd.read_excel(DB_FILENAME, sheet_name='물질정보')
                    df_tox = pd.read_excel(DB_FILENAME, sheet_name='유해성정보')
                    wb     = load_workbook(TPL_MULTI)
                    extract_multi(tid1.strip(), tid2.strip(), df_mat, df_tox, wb)
                    buf = io.BytesIO()
                    wb.save(buf)
                    st.success(f"✅ **{tid1}** + **{tid2}** 추출 완료!")
                    st.download_button(
                        label="📥 결과 엑셀 다운로드",
                        data=buf.getvalue(),
                        file_name=f"추출결과_{tid1}_{tid2}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"오류 발생: {e}")
