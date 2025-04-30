# merge_xls.py
import os, glob, shutil, pandas as pd, pyzipper
from datetime import datetime
from google.colab import files
from IPython.display import clear_output

def merge_xls_from_zip(marker_mode='text', marker_value='★시작★', row_idx=0):
    clear_output()
    print("📦 기존 작업 정리 중...")
    folder = "unzipped_excels"
    if os.path.exists(folder):
        shutil.rmtree(folder)
    for f in glob.glob("엑셀머지_*.xlsx") + glob.glob("*.zip"):
        try: os.remove(f)
        except: pass

    print("📤 ZIP 파일 업로드:")
    uploaded = files.upload()
    zip_name = next(iter(uploaded.keys()))

    os.makedirs(folder, exist_ok=True)
    with pyzipper.AESZipFile(zip_name, 'r') as zip_ref:
        zip_ref.setpassword(b"")
        for name in zip_ref.namelist():
            try: decoded = name.encode('cp437').decode('utf-8')
            except: decoded = name.encode('cp437').decode('euc-kr')
            with zip_ref.open(name) as src, open(os.path.join(folder, decoded), "wb") as dst:
                shutil.copyfileobj(src, dst)

    files_to_merge = [f for f in os.listdir(folder) if f.endswith('.xlsx')]
    if not files_to_merge:
        print("❌ 엑셀 파일 없음."); return

    merged, actual_start = [], None
    for idx, file in enumerate(files_to_merge):
        df_raw = pd.read_excel(os.path.join(folder, file), header=None)
        if idx == 0:
            if marker_mode == 'text':
                found = df_raw[df_raw.iloc[:, 0].astype(str).str.strip() == marker_value.strip()]
                actual_start = found.index[0] + 1 if not found.empty else None
            else:
                actual_start = row_idx
        try:
            df = df_raw.iloc[actual_start:].copy()
            df.columns = df.iloc[0]
            df = df[1:]
            df.insert(0, 'source_file', file)
            merged.append(df)
        except Exception as e:
            print(f"⚠️ {file} 병합 실패: {e}")

    if not merged:
        print("⚠️ 병합할 데이터 없음."); return

    result = pd.concat(merged, ignore_index=True)
    outname = f"엑셀머지_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
    result.to_excel(outname, index=False)
    print(f"\n✅ 병합 완료 → {outname}")
    files.download(outname)
