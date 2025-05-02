# ocrmint.py - Colab 전용 OCRmint 모듈 (배경 보존 + 선명도 + 실시간 합계)

import pytesseract, io, base64, re
from PIL import Image, ImageEnhance, ImageFilter
import ipywidgets as widgets
from IPython.display import display, clear_output, HTML
from google.colab import output

# 전처리: 배경 보존 + 대비 강화 + 선명도 향상
def preprocess_image_strongest(img):
    gray = img.convert("L")
    contrast = ImageEnhance.Contrast(gray).enhance(1.5)
    sharpened = contrast.filter(ImageFilter.UnsharpMask(radius=2, percent=150))
    return sharpened

# 숫자 추출 + 콤마 포맷
def extract_numbers_with_commas(image_bytes):
    image = preprocess_image_strongest(Image.open(io.BytesIO(image_bytes)))
    config = '--psm 6 -c tessedit_char_whitelist=0123456789,'
    raw_text = pytesseract.image_to_string(image, config=config)
    rows, total = [], 0
    for line in raw_text.strip().split("\n"):
        tokens = re.findall(r'[-+]?[0-9A-Za-z,.\s:;()]*', line)
        nums = []
        for token in tokens:
            num_str = re.sub(r'[^0-9\-+]', '', token.upper().replace("O", "0").replace("I", "1").replace("L", "1"))
            try:
                val = int(num_str)
                nums.append(f"{val:,}")
                total += val
            except:
                continue
        if nums:
            rows.append("\t".join(nums))
    return "\n".join(rows), total

# 합계 실시간 반영
def update_sum(change):
    new_text = change['new']
    total = 0
    for line in new_text.splitlines():
        tokens = re.findall(r'-?\d[\d,]*', line)
        for t in tokens:
            try:
                total += int(t.replace(",", ""))
            except:
                pass
    info_label.value = f"<b>📊 합계: {total:,}</b>"

# 붙여넣기 → OCR 수행
def on_image_paste(data_url):
    global image_bytes_global
    header, encoded = data_url.split(",", 1)
    image_bytes_global = base64.b64decode(encoded)
    image_preview.value = image_bytes_global

    result_text, total = extract_numbers_with_commas(image_bytes_global)
    line_count = result_text.count('\n') + 1
    box_height = max(300, line_count * 40)
    text_widget.layout.height = f"{box_height}px"
    text_widget.value = result_text
    info_label.value = f"<b>📊 합계: {total:,}</b>"

    with output_area:
        clear_output()
        display(text_widget)
        print("✅ 복사하거나 수정하면 합계가 자동 갱신됩니다.")

# ✅ 메인 실행 함수 (Colab 전용)
def run_ocrmint():
    global text_widget, info_label, output_area, image_preview

    text_widget = widgets.Textarea(
        layout=widgets.Layout(width='100%'),
        style={'overflow': 'hidden'},
        placeholder="Ctrl+A → Ctrl+C 또는 수정하세요"
    )
    text_widget.observe(update_sum, names='value')

    image_preview = widgets.Image(format='png', layout=widgets.Layout(width='150px', margin='0 0 0 20px'))
    image_slider = widgets.IntSlider(value=150, min=50, max=600, step=10, description='🖼️ 크기', continuous_update=True)
    image_slider.observe(lambda ch: image_preview.layout.__setattr__('width', f"{ch['new']}px"), names='value')

    info_label = widgets.HTML("<b>📊 합계: -</b>")
    output_area = widgets.Output()

    right_panel = widgets.VBox([image_slider, image_preview])
    ui_layout = widgets.HBox([widgets.VBox([info_label, output_area]), right_panel])

    display(HTML('''
    <h4>📋 이미지 붙여넣기 (Ctrl+V)</h4>
    <div style="width:66%; max-width:650px;">
      <textarea id="paste-box" style="width:100%; height:150px; border:2px dashed #888;" placeholder="Ctrl+V로 이미지 붙여넣기"></textarea>
    </div>
    <script>
    document.getElementById('paste-box').addEventListener('paste', function(event) {
      const items = event.clipboardData.items;
      for (let i = 0; i < items.length; i++) {
        if (items[i].type.indexOf("image") !== -1) {
          const blob = items[i].getAsFile();
          const reader = new FileReader();
          reader.onload = function(e) {
            google.colab.kernel.invokeFunction('notebook.onImagePaste', [e.target.result], {});
          };
          reader.readAsDataURL(blob);
        }
      }
    });
    </script>
    '''))

    output.register_callback('notebook.onImagePaste', lambda data_url: on_image_paste(data_url))
    display(ui_layout)
