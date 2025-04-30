
from transformers import AutoTokenizer, AutoModelForTokenClassification, pipeline
import re, os


from IPython.display import display, Markdown, Javascript
import ipywidgets as widgets


# ✅ 색상 팔레트 순환 리스트
TAG_COLORS = ["#007bff", "#28a745", "#ffc107", "#17a2b8", "#6610f2", "#e83e8c", "#6f42c1"]

def render_mapping_colored(mapping):
    html_lines = []
    for idx, (tag, value) in enumerate(mapping.items()):
        color = TAG_COLORS[int(tag[1:4]) % len(TAG_COLORS)]  # N100 → 100 → mod 색상
        html_line = f"<b style='color:{color}'>{tag}</b> → <span>{value}</span><br>"
        html_lines.append(html_line)
    return "".join(html_lines)

def launch_masking_ui():
    """
    ipywidgets 기반 민감정보 마스킹 UI
    """

    # === 입력 상자 ===
    input_box = widgets.Textarea(
        value='',
        placeholder='예: 홍길동 교수님이 서울대학교에서 강의하셨습니다.\n이메일은 gil@example.com입니다.',
        description='입력:',
        layout=widgets.Layout(width='100%', height='250px')
    )

    # === 실행 버튼 ===
    run_button = widgets.Button(description='마스킹 실행', button_style='success')

    # === 출력 상자 ===
    output_box = widgets.Textarea(
        value='',
        description='마스킹 결과:',
        layout=widgets.Layout(width='100%', height='250px')
    )

    # === 매핑 상자 (컬러 HTML로 표시)
    mapping_box = widgets.HTML(
        value='',
        placeholder='매핑 정보 출력 영역',
        layout=widgets.Layout(width='100%', height='150px', overflow='auto', border='1px solid lightgray')
    )

    # === 복사 버튼들 ===
    copy_output_btn = widgets.Button(description="📋 결과 복사", button_style='info')
    copy_mapping_btn = widgets.Button(description="📋 매핑 복사", button_style='info')

    # === 인라인 메시지 표시
    info_label = widgets.Label("")

    # === 복사 로직 (텍스트만 클립보드에 복사)
    def copy_to_clipboard(value, target):
        js = f"""
        navigator.clipboard.writeText(`{value}`);
        """
        display(Javascript(js))
        info_label.value = f"{target} 복사되었습니다."

    # 📌 복사 이벤트 연결 (텍스트만 복사)
    copy_output_btn.on_click(lambda _: copy_to_clipboard(output_box.value, "결과"))
    copy_mapping_btn.on_click(lambda _: copy_to_clipboard(mapping_box.plain_text, "매핑"))

    # === 실행 로직
    def on_run_clicked(_):
        result, mapping, mapping_text = full_masking_pipeline(input_box.value)
        output_box.value = result
        mapping_box.value = render_mapping_colored(mapping)  # HTML 컬러 표시용
        mapping_box.plain_text = mapping_text                # 복사용 텍스트 저장
        info_label.value = "마스킹이 완료되었습니다."

    # 📌 버튼 클릭 이벤트 연결
    run_button.on_click(on_run_clicked)

    # === UI 렌더링
    display(Markdown("### 디지털기획운영부 민감정보 마스킹기"))
    display(input_box)
    display(run_button)
    display(output_box)
    display(copy_output_btn)
    display(mapping_box)
    display(copy_mapping_btn)
    display(info_label)


# ✅ 모델 로딩 (Drive 마운트 없이 기본 HuggingFace 캐시 사용)
model_name = "Leo97/KoELECTRA-small-v3-modu-ner"
# ✅ 모델 로딩 (캐시 디렉토리 기본값 사용)
tokenizer = AutoTokenizer.from_pretrained(model_name)
model = AutoModelForTokenClassification.from_pretrained(model_name)

ner_pipeline = pipeline("ner", model=model, tokenizer=tokenizer, aggregation_strategy="simple")



 # ✅ 확장된 호칭 리스트
COMMON_SUFFIXES = [
    # 📁 가정/관계 기반
    '어머니', '아버지', '엄마', '아빠', '형', '누나', '언니', '오빠', '동생',
    '딸', '아들', '조카', '사촌', '이모', '고모', '삼촌', '숙모', '외삼촌',
    '할머니', '할아버지', '외할머니', '외할아버지', '장모', '장인', '며느리', '사위',
    '부인', '와이프', '신랑', '올케', '형수', '제수씨', '매형', '처제', '시누이',
    # 📁 사회/교육/직업 호칭
    '학생', '초등학생', '중학생', '고등학생', '수험생', '학부모', '선생', '선생님', '교사',
    '교감', '교장', '담임', '반장', '조교수', '교수', '연구원', '강사', '박사', '석사', '학사',
    '보호자', '피해자', '아동', '주민', '당사자', '대상자', '담당자',
    # 📁 직장/조직 직급
    '대표', '이사', '전무', '상무', '부장', '차장', '과장', '대리', '사원', '팀장', '본부장',
    '센터장', '소장', '실장', '총무', '직원', '매니저', '지점장', '사무장',
    # 📁 의료/기타
    '의사', '간호사', '간병인', '기사님', '어르신', '님', '씨'
]


# ✅ 문장 기반 패턴 정의 (NER 보완용)
EXPLICIT_NAME_PATTERNS = [
    r"([가-힣]{2,4})[ ]?(씨|님|선생님|교수님)?(이|가)?\s?(말씀|전달|연락|강의|설명|작성|보고|발언)\w*",
    r"([가-힣]{2,4})[ ]?의\s?(계좌|전화번호|주소|이메일|비밀번호|계정)\w*",
    r"([가-힣]{2,4})[ ]?입니다[.]?",
    r"예금주[는\s]*([가-힣]{2,4})",
    r"계정[의\s]*주인[은\s]*([가-힣]{2,4})",
    r"작성자[는\s]*([가-힣]{2,4})",
]


# ✅ 실전급 조사 리스트
COMMON_JOSA = [
    # ✅ 기본 조사
    '이', '가', '을', '를', '은', '는', '의', '도',
    # ✅ 처소/방향/대상
    '에', '에서', '에게', '께서', '으로', '로', '부터', '까지', '한테',
    # ✅ 강조/대조/비교
    '보다', '보다도', '마저', '조차', '조차도', '까지도', '밖에', '만큼', '만큼은',
    '이라도', '이든지', '이나마', '이건', '이란', '이라서', '이지만',
    # ✅ 연결형 조사
    '이며', '이나', '이거나', '이니까', '이라면', '처럼', '대로', '하고', '그리고', '와', '과',
    # ✅ 보조/종결형 어미
    '이기도', '이었던', '이었지만', '이어서', '이었다면', '인', '일', '임', '이란', '이라는',
    # ✅ 특수형 조사/조합형
    '같은', '같아서', '까지는', '뿐만 아니라', '와는', '와도', '하고도', '으로서', '으로써'
]

# ✅ 이름 추출 (NER + 문장패턴 기반 하이브리드)
def extract_names(text):
    ner_results = ner_pipeline(text)

    # ✅ 사전 정의된 "사람이 아닌 표현" 필터링 목록
    FORBIDDEN_NAME_TERMS = {
        "보호자", "담당자", "대상자", "피해자", "의사", "간호사", "학생",
        "주민", "아동", "매니저", "대표", "교수", "강사", "직원", "총무", "사무장"
    }

    # ✅ NER 결과 기반 이름 후보 수집 (금지어 제외)
    names = [
        e["word"].replace("##", "").strip()
        for e in ner_results
        if e["entity_group"] == "PS"
        and e["word"].strip() not in FORBIDDEN_NAME_TERMS
    ]

    # ✅ 문장 패턴 기반 후보 보완
    names_from_pattern = extract_names_by_pattern(text)
    for name in names_from_pattern:
        if name not in names and name not in FORBIDDEN_NAME_TERMS:
            names.append(name)

    return names



def expand_variation_patterns(text: str, mapping: dict) -> str:
    for tag, base in mapping.items():
        prefix = r'[\s\(\["\']*'
        suffix = f"(?:{'|'.join(COMMON_SUFFIXES)})?"
        josa = f"(?:{'|'.join(COMMON_JOSA)})?"
        pattern = re.compile(rf'(?<![\w가-힣]){re.escape(base)}{suffix}{josa}(?![\w가-힣])', re.IGNORECASE)

        # ✅ 조사, 호칭 포함 시 base만 태깅하고 나머지는 그대로
        def replacer(m):
            matched = m.group(0)
            if base in matched:
                return matched.replace(base, tag)
            return matched

        text = pattern.sub(replacer, text)

    return text


def boost_mapping_from_context(text: str, mapping: dict) -> dict:
    updated = {}
    for tag, base in mapping.items():
        idx = text.find(tag)
        if idx == -1:
            updated[tag] = base
            continue

        window = text[max(0, idx - 100): idx + 100]

        # 조사와 호칭은 탐색에만 쓰고, 저장은 base만
        suffix = f"(?:{'|'.join(COMMON_SUFFIXES)})?"
        josa = f"(?:{'|'.join(COMMON_JOSA)})?"
        pattern = re.compile(rf'({re.escape(base)}{suffix}{josa})', re.IGNORECASE)

        match = pattern.search(window)
        matched = match.group(1) if match else base

        # ✅ 단서는 조사 포함된 표현이더라도, 저장은 항상 이름 본체
        if base in matched:
            updated[tag] = base
        else:
            updated[tag] = matched  # fallback (희박한 경우)
    return updated




# ✅ 문장 단서 기반 이름 추출 함수
def extract_names_by_pattern(text):
    pattern_names = set()
    for pattern in EXPLICIT_NAME_PATTERNS:
        matches = re.findall(pattern, text)
        for match in matches:
            if isinstance(match, tuple):
                pattern_names.add(match[0])
            else:
                pattern_names.add(match)
    return list(pattern_names)


# ✅ 이름 + 호칭/조사 포함 마스킹 (동일인물 표현 묶어서 계층태깅)
def mask_names_with_suffix_josa(original_text, names, start_counter=100):
    masked_text = original_text
    mapping = {}
    counter = start_counter

    # ✅ 수동 매핑 기반 이름 동의어 (예: "아윤" → "조아윤")
    name_alias_dict = {
        "아윤": "조아윤",
        # 필요 시 추가 가능
    }

    # ✅ 대표이름으로 그룹핑
    name_groups = {}
    for name in names:
        base_name = name_alias_dict.get(name, name)
        if base_name not in name_groups:
            name_groups[base_name] = []
        name_groups[base_name].append(name)

    # ✅ 각 대표이름별 태깅 진행
    for base_name, name_variants in name_groups.items():
        base_tag = f"N{counter:03d}"
        mapping[base_tag] = base_name
        suffix_pattern = f"(?:{'|'.join(COMMON_SUFFIXES)})?"
        josa_pattern = f"(?:{'|'.join(COMMON_JOSA)})?"
        pattern = re.compile(
            rf'(?<![\w가-힣])({"|".join(map(re.escape, name_variants))}{suffix_pattern}{josa_pattern})\b(?![\w가-힣])'
        )

        sub_counter = 1
        for match in pattern.finditer(masked_text):
            full_match = match.group(1)

            # ✅ 1. 대표이름은 마지막에 처리
            if full_match == base_name:
                continue

            tag = f"{base_tag}-{sub_counter}"

            # ✅ 2. 이미 태그가 들어간 조각이면 무시 (중복 방지)
            if any(existing_tag in full_match for existing_tag in mapping.keys()):
                continue

            # ✅ 조사나 호칭만 잘린 경우 무시 (예: '이며'만 단독 등장)
            if len(full_match.strip()) <= len(base_name):
                 continue


            # ✅ 4. 중복 매핑이나 텍스트에 이미 등장한 태그 무시
            if tag in masked_text or full_match in mapping.values():
                continue

            # ✅ 5. 정상적인 경우만 태깅
            mapping[tag] = base_name

            # ✅ 6. 단어 경계 기반 안전 치환
            safe_pattern = re.compile(rf'(?<![\w가-힣]){re.escape(full_match)}(?![\w가-힣])')
            masked_text = safe_pattern.sub(
                lambda m: m.group(0).replace(base_name, tag),
                masked_text,
                count=1
            )

            sub_counter += 1

        # ✅ 대표이름 단독 매칭은 마지막에 처리
        pattern_base = re.compile(rf'(?<![\w가-힣])({re.escape(base_name)})(?![\w가-힣])')
        masked_text = pattern_base.sub(base_tag, masked_text)

        counter += 1


    return masked_text, mapping



def to_chosung(text: str) -> str:
    CHOSUNG_LIST = [chr(i) for i in range(0x1100, 0x1113)]
    result = ""
    for ch in text:
        if '가' <= ch <= '힣':
            code = ord(ch) - ord('가')
            cho = code // 588
            result += CHOSUNG_LIST[cho]
        else:
            result += ch
    return result

def sanitize_sensitive_info(text):
    # 📧 이메일
    text = re.sub(r"[\w\.-]+@[\w\.-]+", r"******@****", text)
    # 🔐 주민등록번호
    text = re.sub(r"(\d{6})[-](\d)\d{6}", r"*******-\2*****", text)
    # 📞 전화번호
    text = re.sub(r"(\d{3})-(\d{4})-(\d{4})", r"\1-****-\3", text)
    # 🏠 주소
    text = re.sub(r"(\d{1,3})번지", r"***번지", text)
    text = re.sub(r"(\d{1,3})동", r"***동", text)
    text = re.sub(r"(\d{1,4})호", r"****호", text)
    text = re.sub(r"([가-힣]+(대로|로|길))\s?(\d+)(호|번길|가)?", r"\1 ***", text)
    text = re.sub(r"([가-힣]{1,10})(은행|동|로|길)\s?([\d\-]{4,})", lambda m: m.group(1) + m.group(2) + " " + re.sub(r"\d", "*", m.group(3)), text)
    # 💳 숫자형 코드들
    text = re.sub(r"(\d{2,6})[-]?(\d{2,6})[-]?(\d{2,6})", lambda m: f"{m.group(1)[:2]}{'*'*(len(m.group(1))-2)}{'*'*len(m.group(2))}{m.group(3)[-4:]}", text)
    text = re.sub(r"(\d{4})[- ]?(\d{4})[- ]?(\d{4})[- ]?(\d{4})", lambda m: f"{m.group(1)}-****-****-{m.group(4)}", text)
    # 🌐 IP
    text = re.sub(r"(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})", lambda m: f"{m.group(1)}.{m.group(2)}.*.*", text)
    # 📅 생년월일
    text = re.sub(r"(\d{4})년 (\d{1,2})월 (\d{1,2})일", r"19**년 \2월 *일", text)
    # 🏢 기관명
    text = re.sub(r"\b(good neighbors|굿네이버스|사회복지법인 굿네이버스|gn)\b", "우리법인", text, flags=re.IGNORECASE)
    # 🏫 학교명 마스킹
    text = re.sub(r"([가-힣]{2,20})(초등학교|중학교|고등학교|대학교)", lambda m: to_chosung(m.group(1)) + m.group(2), text)
    # 🏢 학과 마스킹
    text = re.sub(r"([가-힣]{2,20})학과", lambda m: to_chosung(m.group(1)) + "학과", text)
    # 🎓 학년/반 마스킹
    text = re.sub(r"(\d)학년(\s?(\d)반)?", "*학년 *반", text)
    return text

# ✅ 후처리: 조사/어미가 붙은 경우에도 이름만 정확히 치환
def final_name_remask_exact_only(original_text, masked_text, mapping_dict):
    for tag, name in mapping_dict.items():
        # ✅ 이름 뒤에 조사/호칭/어미가 붙은 경우까지 포함해 치환
        pattern = re.compile(rf'(?<![\w가-힣])({re.escape(name)})(?=[가-힣]{0,4}|\W|$)')

        # ✅ 이름만 태깅하고, 뒤 조사/호칭은 유지
        masked_text = pattern.sub(lambda m: m.group(0).replace(name, tag), masked_text)

    return masked_text




# ✅ 민감정보 전체 마스킹 파이프라인
def full_masking_pipeline(text):
    names = extract_names(text)
    # 1. 이름 + 호칭/조사 포함 정규식 기반 태깅
    masked_text, mapping = mask_names_with_suffix_josa(text, names)
    # 2. 확장 표현 내부 재치환 (김철수교수님께서 → N100께서)
    masked_text = expand_variation_patterns(masked_text, mapping)
    # 3. 매핑 정보도 보정 (태그 주변에서 가장 긴 표현으로 재매핑)
    mapping = boost_mapping_from_context(masked_text, mapping)
    # 4. 민감정보 정규식 기반 마스킹
    sanitized = sanitize_sensitive_info(masked_text)
    # 5. 혹시 다시 이름이 원문에 남았을 경우 보정
    sanitized = final_name_remask_exact_only(text, sanitized, mapping)

    # 6. ✅ 텍스트에 없는 태그는 매핑에서 제거
    mapping = {k: v for k, v in mapping.items() if k in sanitized}

    # 7. ✅ 컬러 출력용으로는 dict 자체를 넘겨야 함
    mapping_table = "\n".join([f"{k} → {v}" for k, v in mapping.items()])

    return sanitized, mapping, mapping_table  # ← dict + 텍스트 둘 다 리턴


def masking_tool(text, option=False):
    names = extract_names(text)
    masked_text, mapping = mask_names_with_suffix_josa(text, names)
    masked_text = expand_variation_patterns(masked_text, mapping)
    mapping = boost_mapping_from_context(masked_text, mapping)
    sanitized = sanitize_sensitive_info(masked_text)
    sanitized = final_name_remask_exact_only(sanitized, mapping)

    # ✅ 매핑 정보 한 줄로 정리
    mapping_line = " / ".join([f"{k} {v}" for k, v in mapping.items()])

    # ✅ 출력은 output 박스에 전부 표시
    full_output = f"{sanitized}\n\n{mapping_line}"
    return {
        "결과": full_output,
        "정보": ""
    }

