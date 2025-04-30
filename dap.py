# 라이브러리 설치 및 임포트
import sys
import subprocess

def install_if_missing(package):
    try:
        __import__(package)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])

for package in ['pandas', 'openpyxl', 'xlrd', 'xlsxwriter', 'ipywidgets']:
    install_if_missing(package)

import pandas as pd
import re
import io
from xlsxwriter.utility import xl_col_to_name
from IPython.display import display, HTML, clear_output
import ipywidgets as widgets
from google.colab import files


# 매핑딕셔너리 자리 표시
mapping_dict = {
    "지정후원금(지역)": ("지정후원금(지역)", 33),
    "지정후원금(기타)": ("지정후원금(기타)", 34),
    "전년도이월금(지정후원금기타)": ("전년도이월금(지정후원금기타)", 82),
    "전년도이월금(법인후원금_지역기타)": ("전년도이월금(법인후원금_지역기타)", 66),
    "전년도이월금(법인후원금_지역)": ("전년도이월금(법인후원금_지역)", 65),
    "전년도이월금(후원금)": ("전년도 이월금(후원금)", 81),
    "전년도이월금(보조금)": ("전년도 이월금(보조금)", 71),
    "전년도이월금(법인후원금기타)": ("전년도 이월금(법인후원금기타)", 62),
    "전년도이월금(법인후원금)": ("전년도 이월금(법인)", 61),
    "사업수입": ("사업수입", 42),
    "법인전입금(후원금)지역(기타)": ("법인전입금(후원금)지역(기타)", 26),
    "법인전입금(후원금)지역": ("법인전입금(후원금)지역", 25),
    "법인전입금(후원금)기타": ("법인전입금(후원금) 기타", 22),
    "기타예금이자(후원금)": ("기타예금이자(후원금)", 46),
    "기타예금이자": ("기타예금이자", 45),
    "과년도수입": ("없음", 99),
    "후원물품(지부)": ("수증물품", 99),
    "후원물품(본부)": ("수증물품", 99),
    "지정후원금(본부)": ("지정후원금", 32),
    "지정후원금": ("지정후원금", 32),
    "전년도이월액(보육료)": ("전년도 이월금(기타)", 91),
    "전년도이월금(사업수입및기타)": ("전년도 이월금(기타)", 91),
    "전년도이월금(사업수입)": ("전년도 이월금(기타)", 91),
    "전년도이월금(기타)": ("전년도 이월금(기타)", 91),
    "이자수입(후원금)": ("이자수입(본부)", 47),
    "이자수입(퇴직연금운용수익)": ("이자수입(본부)", 47),
    "이자수입(본부)": ("이자수입(본부)", 47),
    "인건비(시도보조금)": ("시도보조금", 13),
    "운영비(시도보조금)": ("시도보조금", 13),
    "시도보조금": ("시도보조금", 13),
    "사업비(시도보조금)": ("시도보조금", 13),
    "인건비(시군구보조금)": ("시군구보조금", 14),
    "운영비(시군구보조금)": ("시군구보조금", 14),
    "시군구보조금": ("시군구보조금", 14),
    "사업비(시군구보조금)": ("시군구보조금", 14),
    "특별활동비(어린이집)": ("사용자부담금", 43),
    "연장보육료": ("사용자부담금", 43),
    "사용자부담금(본부결산용)": ("사용자부담금", 43),
    "사용자부담금": ("사용자부담금", 43),
    "부모부담 보육료(어린이집)": ("사용자부담금", 43),
    "기타필요경비(어린이집)": ("사용자부담금", 43),
    "시설후원금": ("비지정후원금", 31),
    "비지정후원금": ("비지정후원금", 31),
    "직장어린이집위탁전입금": ("보조금", 16),
    "정부지원 보육료(어린이집)": ("보조금", 16),
    "자본보조금": ("보조금", 16),
    "인건비 보조금": ("보조금", 16),
    "보조금": ("보조금", 16),
    "정기후원금": ("법인전입금(후원금)", 21),
    "일시후원금(본부)": ("법인전입금(후원금)", 21),
    "법인전입금(후원금)": ("법인전입금(후원금)", 21),
    "적립금처분수입(어린이집)": ("잡수입", 41),
    "잡수익": ("잡수입", 41),
    "기타잡수입": ("잡수입", 41),
    "기타지원금": ("기타보조금", 15),
    "기타 보조금": ("기타보조금", 15),
    "인건비(국고보조금)": ("국고보조금", 12),
    "운영비(국고보조금)": ("국고보조금", 12),
    "사업비(국고보조금)": ("국고보조금", 12),
    "국고보조금": ("국고보조금", 12),
    "잡수입": ("잡수입", 41),
    "기타보조금": ("기타보조금", 15),
    "전년도이월금(법인)": ("전년도 이월금(법인)", 61),
    "": ("매핑없음", 999),
    "수증물품": ("수증물품", 99),
    "이이월금": ("이이월금", 999),
    "합계": ("합계", 999),
    "소계": ("소계", "소계"),
}

# normalize_key 함수
def normalize_key(x):
    return str(x).replace(" ", "").replace("\u200b", "").strip()


def clean_se_detail_column(df):
    df["세세목"] = df["세세목"].fillna("").astype(str).str.strip()
    df.loc[df["세세목"] == "", "세세목"] = "-"
    return df


# process_and_format 함수
def process_and_format(df):
    cols = ['예산단위', '정책사업', '단위사업', '세부사업', '세항', '비용/자본구분', '목', '세목', '세세목', '재원']
    code_re = re.compile(r'^\d+$|^[A-Z]+\d+$')

    def is_code(s): return bool(code_re.match(s.strip()))
    def is_subtotal(s): return ''.join(s.split()) == '소계'

    for col in cols:
        prev = None
        filled = []
        for raw in df[col].fillna('').astype(str):
            txt = raw.strip()
            if is_subtotal(txt):
                val = '소계'
            elif is_code(txt):
                val = prev
            else:
                val = txt
                prev = txt
            filled.append(val)
        df[col] = filled


    # ✅ 소계 처리
    df.loc[df['세목'] == '소계', '세세목'] = '소계'


    return df




# 📌 피벗 결과 정렬 함수 추가

def sort_pivot_df(df, with_se_detail):
    """
    📌 피벗 결과 정렬 함수
    - 비용/자본구분 > 세항 > 목 > 세목 > (세세목) > 재원순서 순서대로 정렬
    - 세목과 세세목은 '소계'가 가장 마지막에 오게 처리
    """
    sort_cols = ["비용/자본구분", "세항", "목", "세목"]
    if with_se_detail:
        sort_cols.append("세세목")
    sort_cols.append("재원순서")

    def sort_key(val):
        val = str(val).strip()
        return (1, "") if val == "소계" else (0, val)

    df = df.copy()
    for col in ["세목", "세세목"] if with_se_detail else ["세목"]:
        if col in df.columns:
            df[col] = df[col].map(lambda x: (1, '') if str(x).strip() == '소계' else (0, str(x)))
    df = df.sort_values(by=sort_cols, ascending=True)
    for col in ["세목", "세세목"] if with_se_detail else ["세목"]:
        if col in df.columns:
            df[col] = df[col].map(lambda x: '소계' if x == (1, '') else x[1])
    return df



# compose_jaewon_detail (지출 매핑 함수)
def compose_jaewon_detail(df):
    """
    📌 지출 데이터: '재원' 기준으로 재원순서를 매핑한다.
    📌 매핑이 안 되는 재원(빈 값, 매핑없음)은 삭제한다.
    """

    # 재원 정리
    df["재원"] = df["재원"].astype(str).str.strip()

    # 🔥 매핑 가능한 재원만 남긴다
    df = df[df["재원"].apply(lambda x: normalize_key(x) in mapping_dict.keys())]

    # 재원순서 매핑
    df["재원순서"] = df["재원"].map(lambda x: mapping_dict.get(normalize_key(x), ("매핑없음", 999))[1])

    # 🔥 "매핑없음" 들어간 행 제거 (이중 확인)
    df = df[df["재원순서"] != 999]

    # 재원 다음에 재원순서 위치
    cols = list(df.columns)
    cols = [col for col in cols if col != "재원순서"]
    idx = cols.index("재원")
    cols = cols[:idx+1] + ["재원순서"] + cols[idx+1:]

    return df[cols]


def inject_hyeonyeak_columns(df_summary, df_original):
    """
    📌 df_summary: 피벗된 결과 (5,6번 시트)
    📌 df_original: 원본 수입/지출 (df_expense / df_income)
    → "목, 세목, 재원, 세세목" 기준으로 현액/연예산 집계 후 삽입
    """
    key_cols = ["목", "세목", "재원", "세세목"]
    month_cols = [col for col in df_original.columns if re.match(r"^\d{1,2}월$", str(col))]

    # 🔥 groupby 전 숫자 필드로 변환
    df_original = df_original.copy()
    df_original[month_cols] = df_original[month_cols].apply(pd.to_numeric, errors='coerce').fillna(0)

    # 1. 현액/연예산 계산
    numeric_summary = (
        df_original
        .assign(세세목=lambda df: df["세세목"].fillna("").astype(str).str.strip().replace("", "-"))
        .groupby(key_cols, dropna=False)[month_cols]
        .sum()
        .reset_index()
    )
    numeric_summary["현액(A)"] = numeric_summary[month_cols].sum(axis=1)
    numeric_summary["연예산(계획)"] = numeric_summary["현액(A)"]

    # 2. 병합 전 기존 칼럼 제거
    df_summary = df_summary.drop(columns=["현액(A)", "연예산(계획)"], errors="ignore")

    # 3. 병합
    merged = pd.merge(
        df_summary,
        numeric_summary[key_cols + ["현액(A)", "연예산(계획)"]],
        on=key_cols,
        how="left"
    )

    # 4. 칼럼 순서: 재원순서 다음으로 현액/연예산 이동
    cols = list(merged.columns)
    if "재원순서" in cols:
        idx = cols.index("재원순서")
        new_cols = cols[:idx+1] + ["현액(A)", "연예산(계획)"] + [c for c in cols if c not in key_cols + ["현액(A)", "연예산(계획)"]]
        merged = merged[new_cols]

    return merged





# map_semo_column (수입 매핑 함수)
# 📌 수입 파일 매핑 함수 수정
def map_semo_column(df):
    """
    📌 수입 데이터는:
    (1) 재원 컬럼은 있지만 값이 비어있다.
    (2) 비어 있는 재원만 세목 기준으로 매핑한다.
    (3) 재원 기준으로 재원순서를 매핑한다.
    (4) 세목이 '소계'면 재원순서도 '소계'로 넣는다.
    """

    # 재원 컬럼 없는 경우 대비 (혹시라도)
    if "재원" not in df.columns:
        df["재원"] = None

    if "재원순서" not in df.columns:
        df["재원순서"] = None

    # (1) 재원이 비어있는 곳만 세목으로 매핑
    mask_rewon_empty = df["재원"].isnull() | (df["재원"].astype(str).str.strip() == "")
    df.loc[mask_rewon_empty, "재원"] = df["세목"].map(lambda x: mapping_dict.get(normalize_key(x), ("매핑없음", 999))[0])

    # (2) 소계 처리 (세목이 소계면 재원순서도 소계)
    mask_subtotal = df["세목"].astype(str).replace(" ", "").eq("소계")
    df.loc[mask_subtotal, "재원순서"] = "소계"

    # (3) 그 외는 재원 기준 재원순서 매핑
    mask_normal = df["재원순서"].isnull()
    df.loc[mask_normal, "재원순서"] = df.loc[mask_normal, "재원"].map(lambda x: mapping_dict.get(normalize_key(x), ("매핑없음", 999))[1])

    # (4) 재원 다음에 재원순서 위치 정렬
    cols = list(df.columns)
    cols = [col for col in cols if col != "재원순서"]
    idx = cols.index("재원")
    cols = cols[:idx+1] + ["재원순서"] + cols[idx+1:]

    return df[cols]




# 📌 소계/합계 음영 포맷 함수
def apply_subtotal_format(writer, df, sheet_name):
    wb = writer.book
    ws = writer.sheets[sheet_name]

    # 📌 소계 포맷 (기존 유지)
    subtotal_fmt = wb.add_format({
        'bg_color': '#D9D9D9'
    })

    # ✅ 합계 포맷 (요청대로: 테두리 + 음영 + 볼드)
    total_fmt = wb.add_format({
        'bold': True,
        'bg_color': '#D9D9D9',
        'border': 1,
        'align': 'center',
        'valign': 'vcenter'
    })

    if sheet_name in ["1_지출", "2_수입", "3_지출_과목", "4_수입_과목", "5_지출_과목(세세목)", "6_수입_과목(세세목)"]:
        if "세목" not in df.columns:
            return

        for row_idx, val in enumerate(df["세목"], start=1):
            if str(val).strip() == '소계':
                for col_idx in range(len(df.columns)):
                    cell_value = df.iat[row_idx-1, col_idx]
                    if pd.isna(cell_value):
                        ws.write_blank(row_idx, col_idx, None, subtotal_fmt)
                    else:
                        ws.write(row_idx, col_idx, cell_value, subtotal_fmt)

            elif str(val).strip() == '합계':
                for col_idx in range(len(df.columns)):
                    cell_value = df.iat[row_idx-1, col_idx]
                    if pd.isna(cell_value):
                        ws.write_blank(row_idx, col_idx, None, total_fmt)
                    else:
                        ws.write(row_idx, col_idx, cell_value, total_fmt)

    else:
        if "재원" not in df.columns:
            return

        for row_idx, val in enumerate(df["재원"], start=1):
            if str(val).strip() == '합계':
                for col_idx in range(len(df.columns)):
                    cell_value = df.iat[row_idx-1, col_idx]
                    if pd.isna(cell_value):
                        ws.write_blank(row_idx, col_idx, None, total_fmt)
                    else:
                        ws.write(row_idx, col_idx, cell_value, total_fmt)




# 3,4번 시트용: 세세목 무시
def create_subject_summary_excluding_se_detail(df):
    base_cols = ["비용/자본구분", "세항", "목", "세목", "재원", "재원순서"]

    if not all(col in df.columns for col in base_cols):
        raise ValueError(f"❌ 필수 컬럼 {base_cols}이(가) 없습니다.")

    idx = df.columns.get_loc("재원순서")
    numeric_cols = df.columns[idx+1:]  # 재원순서 다음부터 숫자

    df_temp = df.copy()
    df_temp[numeric_cols] = df_temp[numeric_cols].apply(pd.to_numeric, errors='coerce')

    summary = df_temp.groupby(base_cols, dropna=False)[numeric_cols].sum().reset_index()
    return summary

def generate_sheet_5_6(df_summary_input):
    """
    📌 세세목 포함 피벗 → 정제 → 정렬 → 합계 추가까지 수행
    """
    # 1. 피벗
    df_summary = create_subject_summary_including_se_detail(df_summary_input)

    # 2. 세세목 정제
    df_summary["세세목"] = df_summary["세세목"].fillna("").astype(str).str.strip()
    df_summary.loc[df_summary["세세목"] == "", "세세목"] = "-"

    # 3. 정렬
    df_summary = sort_pivot_df(df_summary, with_se_detail=True)


    return df_summary



def create_subject_summary_including_se_detail(df):
    base_cols = ["비용/자본구분", "세항", "목", "세목", "세세목", "재원", "재원순서"]
    numeric_cols = ["현액(A)", "연예산(계획)", "집행금액(B)"] + [f"{i}월" for i in range(1, 13)]
    numeric_cols = [col for col in numeric_cols if col in df.columns]

    df_temp = df.copy()
    df_temp[numeric_cols] = df_temp[numeric_cols].apply(pd.to_numeric, errors='coerce').fillna(0)

    summary = df_temp.groupby(base_cols, dropna=False)[numeric_cols].sum().reset_index()
    return summary





# 📌 7,8번 시트용: 소계 제외 pivot 생성
def create_biyo_rewon_summary(df):
    """
    📌 비용/자본구분, 재원, 재원순서 기준으로 groupby 후 숫자 합계
    - 재원순서 오른쪽은 모두 숫자로 간주
    - 소계 행은 삭제한다
    """
    base_cols = ["비용/자본구분", "재원", "재원순서"]

    if not all(col in df.columns for col in base_cols):
        raise ValueError(f"❌ 필수 컬럼 {base_cols}이(가) 없습니다.")

    idx = df.columns.get_loc("재원순서")
    numeric_cols = df.columns[idx+1:]  # 재원순서 다음부터 숫자 컬럼

    df_temp = df.copy()
    df_temp[numeric_cols] = df_temp[numeric_cols].apply(pd.to_numeric, errors='coerce')

    # 📌 소계 행 제거
    df_temp = df_temp[df_temp["재원순서"].astype(str).str.strip() != "소계"]

    # 📌 groupby
    summary = df_temp.groupby(base_cols, dropna=False)[numeric_cols].sum().reset_index()

    # 📌 재원순서를 str로 변환 후 정렬 (필요할 경우)
    summary["재원순서"] = summary["재원순서"].astype(str)
    summary = summary.sort_values(by=["재원순서"], ascending=True)

    return summary



    # 📌 7,8번 시트에 합계 행 추가하는 함수
def add_total_row(df):
    """
    📌 비용/자본구분/재원/재원순서에 '합계' 넣고
    📌 숫자 컬럼은 열 단위 sum
    """
    key_cols = ["비용/자본구분", "재원", "재원순서"]

    # 🔥 숫자 컬럼 찾기
    idx = df.columns.get_loc("재원순서")
    numeric_cols = df.columns[idx+1:]

    # 🔥 합계 row 만들기
    total_row = {col: "합계" for col in key_cols}
    for col in numeric_cols:
        total_row[col] = df[col].sum()

    # 🔥 df에 추가
    df_with_total = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)

    return df_with_total



# 📌 9번 시트용: 재원잔액(수입-지출) 계산 (재원순서 다음부터 숫자컬럼 자동 인식)
def create_rewon_balance(exp_df, inc_df):
    """
    📌 7번(지출)과 8번(수입) 시트를 재원/재원순서 기준으로 병합 후
    - 수입 금액 - 지출 금액 계산
    - 재원순서 다음 컬럼부터 숫자컬럼 자동 처리
    """
    key_cols = ["재원", "재원순서"]

    # 🔥 merge
    merged = pd.merge(
        inc_df, exp_df,
        on=key_cols,
        how="outer",
        suffixes=("_수입", "_지출")
    )

    # 🔥 없는 값은 0
    merged = merged.fillna(0)

    # 🔥 재원순서 다음부터 모든 컬럼
    idx = merged.columns.get_loc("재원순서")
    all_numeric_cols = merged.columns[idx+1:]

    # 🔥 숫자 컬럼 강제 변환
    merged[all_numeric_cols] = merged[all_numeric_cols].apply(pd.to_numeric, errors="coerce").fillna(0)

    # 🔥 수입-지출 계산 (컬럼 쌍 맞춰서)
    balance_df = merged[key_cols].copy()

    for col in all_numeric_cols:
        if col.endswith("_수입"):
            base_col = col.replace("_수입", "")
            exp_col = f"{base_col}_지출"

            if exp_col in merged.columns:
                balance_df[base_col] = merged[col] - merged[exp_col]
            else:
                # 혹시 지출 없는 경우는 그냥 수입 값 그대로
                balance_df[base_col] = merged[col]

    # 📌 재원순서 정렬
    balance_df["재원순서"] = balance_df["재원순서"].astype(str)
    balance_df = balance_df.sort_values(by=["재원순서"], ascending=True)

    return balance_df



def add_half_total_row(df):
    """
    ✅ '세목'이 '소계'인 행만 합산 → '합계' 행 추가
    """
    if "재원순서" not in df.columns or "세목" not in df.columns:
        return df

    cols = list(df.columns)
    idx = cols.index("재원순서")
    numeric_cols = cols[idx+1:]
    text_cols = cols[:idx+1]

    # 🔍 소계 필터링
    mask = df["세목"].astype(str).str.strip() == "소계"
    df_sub = df[mask]

    if df_sub.empty:
        return df  # 소계 없으면 그냥 반환

    # 🔢 숫자 컬럼 합계
    df_numeric = df_sub[numeric_cols].apply(pd.to_numeric, errors='coerce').fillna(0)
    sum_row = df_numeric.sum().to_dict()

    # 📌 텍스트 컬럼은 "합계"
    total_row = {col: "합계" for col in text_cols}
    total_row.update(sum_row)

    return pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)




# 📌 10번 시트: 누적 잔액 생성 함수
def create_cumulative_balance(df):
    """
    📌 9번 시트 기반 누적잔액 시트 생성 (10번 시트)
    - 1월~12월 컬럼을 누적합으로 변환
    """
    df = df.copy()
    idx = df.columns.get_loc("재원순서")
    numeric_cols = df.columns[idx+1:]
    df[numeric_cols] = df[numeric_cols].cumsum(axis=1)
    return df





# 최종 실행 run_final_report 함수
def run_final_report(b):
    with output:
        clear_output()
        print("🚦 보고서 생성 시작...")

        if not upload_expense.value:
            print("❌ 지출 파일을 먼저 업로드하세요.")
            return
        if not upload_income.value:
            print("❌ 수입 파일을 먼저 업로드하세요.")
            return

        # ✅ 1번 시트 처리 순서 변경: 숫자 코드 먼저 제거 → 그 후 재원 복사
        expense_filename = list(upload_expense.value.keys())[0]
        expense_content = upload_expense.value[expense_filename]['content']
        df_expense = pd.read_excel(io.BytesIO(expense_content))

        # ✅ 코드형 재원일 경우, 해당 라인의 현액/연예산 초기화
        def clear_wrong_amounts(df):
            df = df.copy()
            for col in ["현액(A)", "연예산(계획)"]:
                if col not in df.columns:
                    df[col] = None
            def is_code(val):
                val = str(val).strip()
                return bool(re.match(r'^\d+$|^[A-Z]+\d+$', val))
            mask = df["재원"].astype(str).map(is_code)
            df.loc[mask, ["현액(A)", "연예산(계획)"]] = 0
            return df

        df_expense = clear_wrong_amounts(df_expense)     # 🔥 숫자 코드 → 현액/연예산 제거
        df_expense = process_and_format(df_expense)      # 🔄 그 다음 윗줄 복사
        df_expense = compose_jaewon_detail(df_expense)   # ✅ 재원순서 추가

        # ✅ [추가] 세세목이 비어있는 경우 BLANK로 대체 (시트1 처리)
        df_expense["세세목"] = df_expense["세세목"].fillna("-").astype(str).str.strip()



        # 2번 시트
        income_filename = list(upload_income.value.keys())[0]
        income_content = upload_income.value[income_filename]['content']
        df_income = pd.read_excel(io.BytesIO(income_content))
        df_income = process_and_format(df_income)
        df_income = map_semo_column(df_income)

        # 매핑없음 제거
        df_expense = df_expense[~((df_expense["재원"] == "매핑없음") & (df_expense["재원순서"] == 999))]
        df_income = df_income[~((df_income["재원"] == "매핑없음") & (df_income["재원순서"] == 999))]

        # 🔥 (2) 진짜 빈 행 삭제
        df_expense = df_expense.dropna(how='all')  # 완전 빈 줄 삭제
        # 소계/합계 외에 완전히 빈 세목만 제거 (텍스트형이면서 공백만 있는 것도 걸러냄)
        df_expense = df_expense[df_expense["세목"].astype(str).str.strip().notnull() &
                                (df_expense["세목"].astype(str).str.strip() != "")]

        # ✅ 세세목 비어있는 값은 먼저 "BLANK"로 전처리 (pivot 전에 반드시 수행)
        df_expense["세세목"] = df_expense["세세목"].fillna("").astype(str).str.strip()
        df_expense.loc[df_expense["세세목"] == "", "세세목"] = "-"

        df_income["세세목"] = df_income["세세목"].fillna("").astype(str).str.strip()
        df_income.loc[df_income["세세목"] == "", "세세목"] = "-"

        # ✅ 5,6번 시트 생성 함수 호출
        exp_sub_full = generate_sheet_5_6(df_expense)
        inc_sub_full = generate_sheet_5_6(df_income)


        # ✅ 3,4번 시트 생성 (세세목 무시한 피벗)
        exp_sub = create_subject_summary_excluding_se_detail(df_expense)
        inc_sub = create_subject_summary_excluding_se_detail(df_income)

        # ✅ 정렬
        exp_sub = sort_pivot_df(exp_sub, with_se_detail=False)
        inc_sub = sort_pivot_df(inc_sub, with_se_detail=False)





        # 📌 7번, 8번 시트용 pivot 만들기
        exp_biyo_rewon = create_biyo_rewon_summary(df_expense)
        inc_biyo_rewon = create_biyo_rewon_summary(df_income)

        # 📌 7,8번 시트에 합계 행 추가
        exp_biyo_rewon = add_total_row(exp_biyo_rewon)
        inc_biyo_rewon = add_total_row(inc_biyo_rewon)




        # 📌 9번 시트 생성
        rewon_balance = create_rewon_balance(exp_biyo_rewon, inc_biyo_rewon)



        # 📌 10번 시트: 누적잔액 (9번 기반 누적합)
        rewon_balance_cumsum = create_cumulative_balance(rewon_balance)





        # 📌 여기 추가: 피벗 결과 정렬
        exp_sub = sort_pivot_df(exp_sub, with_se_detail=False)        # 3번 시트용
        inc_sub = sort_pivot_df(inc_sub, with_se_detail=False)        # 4번 시트용
        exp_sub_full = sort_pivot_df(exp_sub_full, with_se_detail=True)  # 5번 시트용
        inc_sub_full = sort_pivot_df(inc_sub_full, with_se_detail=True)  # 6번 시트용



        # 📌 3,4,7,8,9번 시트 저장 전에 세세목 컬럼 삭제
        for df in [exp_sub, inc_sub, exp_biyo_rewon, inc_biyo_rewon, rewon_balance]:
            if "세세목" in df.columns:
                df.drop(columns=["세세목"], inplace=True)


        # 정렬
        exp_sub_full = sort_pivot_df(exp_sub_full, with_se_detail=True)
        inc_sub_full = sort_pivot_df(inc_sub_full, with_se_detail=True)


        # 저장
        output_file = "통합_지출수입_보고서.xlsx"
        with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
            # 원본 시트
            df_expense.to_excel(writer, index=False, sheet_name="1_지출")
            df_income.to_excel(writer, index=False, sheet_name="2_수입")

            # pivot 저장
            exp_sub.to_excel(writer, index=False, sheet_name="3_지출_과목")
            inc_sub.to_excel(writer, index=False, sheet_name="4_수입_과목")
            exp_sub_full.to_excel(writer, index=False, sheet_name="5_지출_과목(세세목)")
            inc_sub_full.to_excel(writer, index=False, sheet_name="6_수입_과목(세세목)")

            # 🔥 재원 피벗 저장
            exp_biyo_rewon.to_excel(writer, index=False, sheet_name="7_지출_재원")
            inc_biyo_rewon.to_excel(writer, index=False, sheet_name="8_수입_재원")

            # 🔥 9번시트: 수입 - 지출
            rewon_balance.to_excel(writer, index=False, sheet_name="9_재원잔액(수입-지출)")

            # 📌 피벗 저장 직후에 다시 워크북 객체 얻기
            wb = writer.book

            # 📌 header 포맷 만들기
            header_fmt = wb.add_format({
                'bold': True,
                'bg_color': '#D9D9D9',
                'align': 'center',
                'valign': 'vcenter'
            })

            # 📌 시트 헤더 및 줄 색칠
            sheet_names = [
                '1_지출', '2_수입',
                '3_지출_과목', '4_수입_과목',
                '5_지출_과목(세세목)', '6_수입_과목(세세목)',
                '7_지출_재원', '8_수입_재원', '9_재원잔액(수입-지출)',
                '10_누적잔액'  # ✅ 추가됨
            ]
            df_list = [
                df_expense, df_income,
                exp_sub, inc_sub, exp_sub_full, inc_sub_full,
                exp_biyo_rewon, inc_biyo_rewon, rewon_balance,
                rewon_balance_cumsum  # ✅ 추가됨
                ]


            for sheet_name, df_sheet in zip(sheet_names, df_list):
                # 👉 합계는 1~6번 시트에서만 마지막에 추가
                if sheet_name in ["1_지출", "2_수입", "3_지출_과목", "4_수입_과목", "5_지출_과목(세세목)", "6_수입_과목(세세목)"]:
                    df_sheet = add_half_total_row(df_sheet)

                df_sheet.to_excel(writer, index=False, sheet_name=sheet_name)

                ws = writer.sheets[sheet_name]
                ws.set_row(0, None, header_fmt)  # 헤더 색칠
                apply_subtotal_format(writer, df_sheet, sheet_name)  # 소계/합계 줄 색칠


        # 📌 저장 후 Colab에서 자동 다운로드
        from google.colab import files
        files.download(output_file)


# 📌 파일 업로드 UI 세팅
upload_expense = widgets.FileUpload(description="📤 지출 업로드", accept='.xls,.xlsx', multiple=False)
upload_income = widgets.FileUpload(description="📥 수입 업로드", accept='.xls,.xlsx', multiple=False)
run_button = widgets.Button(description="🚀 보고서 생성", button_style="success")
output = widgets.Output()

display(HTML("<h4>1️⃣ 지출 파일 업로드</h4>"))
display(upload_expense)
display(HTML("<h4>2️⃣ 수입 파일 업로드</h4>"))
display(upload_income)
display(run_button)
display(output)

# 📌 버튼 실행 연결
run_button.on_click(run_final_report)
