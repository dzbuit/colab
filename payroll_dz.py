# payroll_dz.py

def run_all():
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
        for block in merged_blocks[1:]:
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

        # ✅ 마지막 줄 추가: 첫 행 제거
        final_filtered = final_filtered.drop(index=0).reset_index(drop=True)

        return final_filtered

    uploaded = files.upload()
    file_name = next(iter(uploaded))
    df_raw = pd.read_excel(io.BytesIO(uploaded[file_name]))
    df_final = process_final_extract_all_in_one(df_raw)
    output_file = "급여대장_페이롤DZ.xlsx"
    df_final.to_excel(output_file, index=False)
    files.download(output_file)
