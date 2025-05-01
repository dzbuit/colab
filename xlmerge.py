# xlmerge.py

import os, glob, shutil
import pandas as pd
import pyzipper
from datetime import datetime
from google.colab import files
import ipywidgets as widgets
from IPython.display import display, clear_output

def clean_workspace():
    folder = "unzipped_excels"
    if os.path.exists(folder):
        shutil.rmtree(folder)
    for f in glob.glob("엑셀머지_*.xlsx") + glob.glob("*.zip"):
        try:
            os.remove(f)
        except:
            pass

def unzip_files(zip_name, extract_folder="unzipped_excels"):
    os.makedirs(extract_folder, exist_ok=True)
    with pyzipper.AESZipFile(zip_name, 'r') as zip_ref:
        zip_ref.setpassword(b"")
        for name in zip_ref.namelist():
            try:
                decoded = name.encode('cp437').decode('utf-8')
            except:
                decoded = name.encode('cp437').decode('euc-kr')
            with zip_ref.open(name) as src, open(os.path.join(extract_folder, decoded), "wb") as dst:
                shutil.copyfileobj(src, dst)
    return [f for f in os.listdir(extract_folder) if f.endswith('.xlsx')]

def merge_excels(files_to_merge, extract_folder, mode='text', marker='★시작★', row_idx=0):
    merged = []
    actual_start = None

    for idx, file in enumerate(files_to_merge):
        path = os.path.join(extract_folder, file)
        df_raw = pd.read_excel(path, header=None)

        if idx == 0:
            if mode == 'text':
                found = df_raw[df_raw.iloc[:, 0].astype(str).str.strip() == marker.strip()]
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
        return None

    result = pd.concat(merged, ignore_index=True)
    outname = f"엑셀머지_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
    result.to_excel(outname, index=False)
    return outname

def run_merge():
    mode_radio = widgets.RadioButtons(options=[('기준 텍스트로 병합', 'text'), ('행 번호로 병합', 'row')])
    marker_input = widgets.Text(value='★시작★')
    row_input = widgets.IntText(value=0)
    input_box = widgets.VBox([marker_input])
    confirm_button = widgets.Button(description='병합 실행', button_style='success')
    output_box = widgets.Output()

    def update_input(mode): input_box.children = [marker_input] if mode == 'text' else [row_input]
    mode_radio.observe(lambda ch: update_input(ch['new']) if ch['name'] == 'value' else None, names='value')

    def on_confirm(b):
        clear_output(wait=True)
        display(output_box)
        with output_box:
            clean_workspace()
            uploaded = files.upload()
            zip_name = list(uploaded.keys())[0]
            folder = "unzipped_excels"
            files_to_merge = unzip_files(zip_name, folder)
            if not files_to_merge:
                print("❌ 병합할 엑셀 없음.")
                return
            mode = mode_radio.value
            marker = marker_input.value
            row_idx = row_input.value
            result_file = merge_excels(files_to_merge, folder, mode, marker, row_idx)
            if result_file:
                print(f"✅ 병합 완료 → {result_file}")
                files.download(result_file)
            else:
                print("⚠️ 병합 실패")

    confirm_button.on_click(on_confirm)
    display(mode_radio, input_box, confirm_button)
