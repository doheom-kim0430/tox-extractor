import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
import io
import os

# 1. 페이지 설정
st.set_page_config(page_title="독성정보 자동 추출 시스템", layout="wide")
st.title("🧪 화학물질 독성정보 자동 추출 서비스")
st.info("내부식별자만 입력하면 서버에 내장된 DB에서 정보를 추출하여 엑셀을 생성합니다.")

# --- 파일 설정 (GitHub 업로드 파일명과 반드시 일치해야 함) ---
DB_FILENAME = "유해성미확인물질 12종 DB.xlsx"
TPL_FILENAME = "개별물질 추출 템플릿.xlsx"

# --- [기능 유지] 우선순위 로직 함수 ---
def apply_priority(df, cat, method, exp_species=None):
    if len(df) <= 1:
        return df.iloc[0]
    
    temp = df.copy()
    if method == "실험값":
        if cat == "급성경구독성":
            temp['p1'] = (temp['Endpoint(표준)'] == 'LD50').astype(int)
            temp['p2'] = (temp['시험종(표준)'] == 'Rat').astype(int)
            temp['p3'] = temp['시험지침'].astype(str).str.contains('401', na=False).astype(int)
            temp = temp.sort_values(['p1', 'p2', 'p3', 'Result'], ascending=[False, False, False, True])
        elif cat == "급성흡입독성":
            temp['p1'] = (temp['Endpoint(표준)'] == 'LC50').astype(int)
            temp['p2'] = (temp['시험종(표준)'] == 'Rat').astype(int)
            temp['p3'] = (temp['Duration(표준)'] == '4 h').astype(int)
            temp['p4'] = temp['시험지침'].astype(str).str.contains('403', na=False).astype(int)
            temp = temp.sort_values(['p1', 'p2', 'p3', 'p4', 'Result'], ascending=[False, False, False, False, True])
        elif cat == "어류급성독성":
            temp['p1'] = (temp['Endpoint(표준)'] == 'LC50').astype(int)
            temp['p2'] = temp['시험종(표준)'].isin(['Fathead minnow', 'Zebrafish', 'Rainbow trout']).astype(int)
            temp['p3'] = (temp['Duration(표준)'] == '96 h').astype(int)
            temp['p4'] = temp['시험지침'].astype(str).str.contains('203', na=False).astype(int)
            temp = temp.sort_values(['p1', 'p2', 'p3', 'p4', 'Result'], ascending=[False, False, False, False, True])
        elif cat == "물벼룩급성독성":
            temp['p1'] = (temp['Endpoint(표준)'] == 'EC50').astype(int)
            temp['p2'] = (temp['시험종(표준)'] == 'Daphnia magna').astype(int)
            temp['p3'] = (temp['Duration(표준)'] == '48 h').astype(int)
            temp['p4'] = temp['시험지침'].astype(str).str.contains('202', na=False).astype(int)
            temp = temp.sort_values(['p1', 'p2', 'p3', 'p4', 'Result'], ascending=[False, False, False, False, True])
        elif cat == "담수조류생장저해":
            temp['p1'] = (temp['Endpoint(표준)'] == 'EC50').astype(int)
            temp['p2'] = temp['시험종(표준)'].isin(['P. subcapitata', 'D. subspicatus']).astype(int)
            temp['p3'] = (temp['Duration(표준)'] == '72 h').astype(int)
            temp['p4'] = temp['시험지침'].astype(str).str.contains('201', na=False).astype(int)
            temp = temp.sort_values(['p1', 'p2', 'p3', 'p4', 'Result'], ascending=[False, False, False, False, True])
        elif cat in ["복귀돌연변이", "포유류 배양세포를 이용한 염색체이상", "소핵시험"]:
            temp = temp.head(1)
        else:
            return df.iloc[0]
            
    elif method == "QSAR":
        model_map = {
            "급성경구독성": "Acute toxicity in Rat, Oral - Danish QSAR DB ACDLabs model (v1.0)",
            "담수조류생장저해": "Pseudokirchneriella subcapitata 72h EC50 - Danish QSAR DB battery model (v1.0)",
            "물벼룩급성독성": "Daphnia magna 48h EC50 - Danish QSAR DB battery model (v1.0)",
            "복귀돌연변이": "Ames test in S. typhimurium (in vitro) - Danish QSAR DB battery model (v1.0)",
            "소핵시험": "Micronucleus Test in Mouse Erythrocytes - Danish QSAR DB battery model (v1.0)",
            "어류급성독성": "Fathead minnow 96h LC50 - Danish QSAR DB battery model (v1.0)",
            "피부부식성/자극성": "BfR skin irritation/corrosion (v1.0)"
        }
        if cat == "포유류 배양세포를 이용한 염색체이상":
            model_name = "Chromosome Aberrations in Chinese Hamster Ovary (CHO) Cells - Danish QSAR DB battery model (v1.0)" if exp_species == "CHO Cells" else "Chromosome Aberrations in Chinese Hamster Lung (CHL) Cells - Danish QSAR DB battery model (v1.0)"
            temp['p_q'] = (temp['모델 종류 및 버전'] == model_name).astype(int)
        elif cat in model_map:
            temp['p_q'] = (temp['모델 종류 및 버전'] == model_map[cat]).astype(int)
        else:
            temp['p_q'] = 0
        temp = temp.sort_values('p_q', ascending=False)
        
    return temp.iloc[0]

# --- [기능 유지] 데이터 포맷팅 함수 ---
def format_val(row, cat, method):
    res = str(row['Result'])
    if method == "QSAR" and str(row['Domain status']) == "Out of domain":
        res += " (Out of domain)"
    val_items = ["급성경구독성", "급성흡입독성", "어류급성독성", "물벼룩급성독성", "담수조류생장저해"]
    if cat in val_items:
        return f"{row['Endpoint(표준)']} = {res} {row['단위']} ({row['시험종(표준)']})"
    return res

# --- 메인 실행 UI ---
target_id = st.text_input("🔍 추출할 내부식별자 입력 (예: B-3)", value="B-3")

if st.button("🚀 데이터 추출 및 엑셀 다운로드"):
    # 파일 존재 확인
    if not os.path.exists(DB_FILENAME) or not os.path.exists(TPL_FILENAME):
        st.error(f"파일을 찾을 수 없습니다: {DB_FILENAME} 또는 {TPL_FILENAME}이 GitHub에 있는지 확인하세요.")
    else:
        try:
            # 데이터 로드
            df_mat = pd.read_excel(DB_FILENAME, sheet_name='물질정보')
            df_tox = pd.read_excel(DB_FILENAME, sheet_name='유해성정보')
            wb = load_workbook(TPL_FILENAME)
            ws = wb.active
            
            # 1. 물질정보 기입 (C7, D7:G7)
            mat_row = df_mat[df_mat['내부식별자'] == target_id].iloc[0]
            ws['C7'] = target_id
            ws['D7'] = mat_row['CAS']
            ws['E7'] = mat_row['물질명']
            ws['F7'] = mat_row['분자식']
            ws['G7'] = mat_row['분자량']
            
            # 2. 유해성정보 루프
            categories = ["급성경구독성", "급성흡입독성", "피부부식성/자극성", "복귀돌연변이", 
                          "포유류 배양세포를 이용한 염색체이상", "소핵시험", "어류급성독성", 
                          "물벼룩급성독성", "담수조류생장저해", "이분해성"]
            exp_srcs = ["ECHA CHEM", "US DashBoard", "Pubchem", "K-reach", "환경부유해성심사결과"]
            qsar_srcs = ["QSAR Toolbox v.4.8", "Danish QSAR", "Epi suite"]
            ai_srcs = ["HAZMAP", "Protox 3.0", "Vega", "Cheminfomatics"] # 오타 수정 완료

            for r_idx, cat in enumerate(categories):
                df_cat = df_tox[(df_tox['내부식별자'] == target_id) & (df_tox['유해성항목'] == cat)]
                exp_species_found = None
                
                for c_idx, src in enumerate(exp_srcs):
                    df_src = df_cat[(df_cat['결과도출방법'] == '실험값') & (df_cat['출처'] == src)]
                    if not df_src.empty:
                        best = apply_priority(df_src, cat, "실험값")
                        ws.cell(row=12+r_idx, column=4+c_idx).value = format_val(best, cat, "실험값")
                        if cat == "포유류 배양세포를 이용한 염색체이상": exp_species_found = best['시험종(표준)']

                df_ra = df_cat[(df_cat['결과도출방법'] == 'Read-across') & (df_cat['출처'] == 'QSAR Toolbox v.4.8')]
                if not df_ra.empty:
                    ws.cell(row=12+r_idx, column=9).value = format_val(df_ra.iloc[0], cat, "Read-across")

                for c_idx, src in enumerate(qsar_srcs):
                    df_src = df_cat[(df_cat['결과도출방법'] == 'QSAR') & (df_cat['출처'] == src)]
                    if not df_src.empty:
                        best = apply_priority(df_src, cat, "QSAR", exp_species_found)
                        ws.cell(row=12+r_idx, column=10+c_idx).value = format_val(best, cat, "QSAR")

                for c_idx, src in enumerate(ai_srcs):
                    df_src = df_cat[(df_cat['결과도출방법'] == 'AI-based QSAR') & (df_cat['출처'] == src)]
                    if not df_src.empty:
                        ws.cell(row=12+r_idx, column=13+c_idx).value = format_val(df_src.iloc[0], cat, "AI-based QSAR")

            # --- [기능 유지] 시각적 개선 (스타일링) ---
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            for rng in [ws['C7:G7'], ws['B11:P21']]:
                for row in rng:
                    for cell in row:
                        cell.border = thin_border
                        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                        cell.font = Font(name='맑은 고딕', size=9)
            
            col_widths = {'B': 12, 'C': 15, 'D': 22, 'E': 25, 'F': 12, 'G': 12, 'H': 22, 'I': 18, 'J': 20, 'K': 20, 'L': 20, 'M': 15, 'N': 15, 'O': 15, 'P': 15}
            for col, width in col_widths.items(): ws.column_dimensions[col].width = width
            for i in range(12, 22): ws.row_dimensions[i].height = 45 

            # 결과 다운로드
            output = io.BytesIO()
            wb.save(output)
            st.success(f"'{target_id}' 데이터 추출 완료!")
            st.download_button(label="📥 결과 엑셀 다운로드", data=output.getvalue(), file_name=f"추출결과_{target_id}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"오류 발생: {e}")
























