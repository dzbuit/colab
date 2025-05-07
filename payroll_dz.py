# payroll_dz.py

def run_all():
    # 📘 설명서 출력
    print("""
📢  **디지털 기획운영부 직접 제작**

1️⃣ 왼쪽 상단 ⏵ **[시작 버튼]**을 클릭해주세요.

2️⃣ 아래쪽에 **E-HR** 에서 내려받은  
📂 **"급여대장"** 파일을 첨부합니다.

📄 **폼드(묶음형) 보고서**를  
🔲 속성별 **구조화된 보고서**로 변환해드립니다.

---
✅ **완벽하지 않을 수 있습니다.**
    """)

    import pandas as pd, io
    from google.colab import files

    def process_final_extract_all_in_one(df_raw, drop_garbage_rows=True):
        df_raw.iloc[:, 1] = df_raw.iloc[:, 1].combine_first(df_raw.iloc[:, 0])
        df_clean = df_raw.apply(lambda col: col.map(lambda x: str(x).replace(" ", "") if pd.notna(x) else x))
        col_name = df_clean.columns[1]
        name_indices = df_clean[df_clean[col_name].str.contains("성명", na=False)].index.tolist()
        corp_indices = df_clean[df_clean[col_name].str.contains("법인부담계", na=False)].index.tolist()

        sets = []
        i = 0
        while i + 1 < len(name_indices):
            first, second = name_indices[i], name_indices[i + 1]
            third_candidates = [x for x in corp_indices if x > second]
            if third_candidates:
                third = third_candidates[0]
                sets.append((first, second, third))
            i += 2
        if i + 1 == len(name_indices) - 1:
            sets.append((name_indices[i], name_indices[i + 1], df_raw.shape[0] - 1))

        merged_blocks = []
        for first, second, third in sets:
            block = df_raw.iloc[first:third + 1].reset_index(drop=True)
            지급항목 = block.iloc[:, 1].reset_index(drop=True).rename("지급항목")
            name_data = block.iloc[:, 2:].reset_index(drop=True)
            full_block = pd.concat([지급항목, name_data], axis=1)
            merged_blocks.append(full_block)

        final_merged = merged_blocks[0]
        for block in merged_blocks[1:]:  # ✅ 여기!
            final_merged = pd.concat([final_merged, block.iloc[:, 1:]], axis=1)

        data_block = final_merged.iloc[1:6, :]
        is_real_data = data_block.apply(
            lambda col: col.astype(str).str.strip().replace("계", "").replace("nan", "").ne("").any(),
            axis=0
        )
        cleaned = final_merged.loc[:, is_real_data]

        header_keywords = ['성명', '소속', '직책', '직위', '직급', '호봉', '최종승호월']
        seen_header = False
        rows_to_keep = []
        for i in range(len(cleaned)):
            current_value = str(cleaned["지급항목"].iloc[i]).replace(" ", "")
            if current_value == "성명":
                next_vals = cleaned["지급항목"].iloc[i:i+7].astype(str).str.replace(" ", "").tolist()
                if all(keyword in next_vals for keyword in header_keywords):
                    if seen_header:
                        continue
                    else:
                        seen_header = True
            rows_to_keep.append(i)

        final_filtered = cleaned.loc[rows_to_keep].reset_index(drop=True)

        if drop_garbage_rows:
            final_filtered = final_filtered.drop(index=range(46, 55), errors="ignore").reset_index(drop=True)

        return final_filtered

    # 1️⃣ 파일 업로드
    uploaded = files.upload()
    file_name = next(iter(uploaded))
    df_raw = pd.read_excel(io.BytesIO(uploaded[file_name]))

    # 2️⃣ 병합 처리
    df_final = process_final_extract_all_in_one(df_raw)

    # 3️⃣ 병합본 저장
    merged_file = "급여대장_페이롤DZ_병합본.xlsx"
    df_final.to_excel(merged_file, index=False)

    # 4️⃣ 병합본 불러와 첫 행을 header로 재설정
    df_temp = pd.read_excel(merged_file)
    new_header = df_temp.iloc[0]
    df_cleaned = df_temp[1:].copy()
    df_cleaned.columns = new_header
    df_cleaned.reset_index(drop=True, inplace=True)

    # 5️⃣ 최종본 저장
    final_file = "급여대장_페이롤DZ_최종본.xlsx"
    df_cleaned.to_excel(final_file, index=False)

    # 6️⃣ 다운로드
    files.download(final_file)
