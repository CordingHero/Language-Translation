!pip install deep-translator openpyxl tqdm

from google.colab import files
from openpyxl import load_workbook
from deep_translator import GoogleTranslator, supported_languages
from tqdm import tqdm
import time
import random
import re

# ==========================================
# 설정
# ==========================================
BATCH_SIZE = 50
MAX_CHAR_LIMIT = 4500
# ==========================================

# ----------------------------------------------------
# 변수 보호를 위한 마스킹/언마스킹 함수
# ----------------------------------------------------
def mask_variables(text):
    if not isinstance(text, str):
        return text, []
    pattern = r'\{.*?\}'
    variables = re.findall(pattern, text)
    masked_text = text
    for i, var in enumerate(variables):
        placeholder = f"__VAR_{i}__"
        masked_text = masked_text.replace(var, placeholder, 1)
    return masked_text, variables

def unmask_variables(text, variables):
    if not variables:
        return text
    restored_text = text
    for i, var in enumerate(variables):
        placeholder = f"__VAR_{i}__"
        if placeholder in restored_text:
            restored_text = restored_text.replace(placeholder, var)
        else:
            pattern = f"__\s*VAR\s*_\s*{i}\s*__"
            restored_text = re.sub(pattern, var, restored_text)
    return restored_text

# ----------------------------------------------------
# 언어 코드 파싱 함수 (자동 감지용)
# ----------------------------------------------------
def parse_lang_code(cell_value):
    """
    셀 값(예: id-ID, ar-SA, fr-CH)에서 deep-translator가 이해하는
    ISO 639-1 코드(예: id, ar, fr)를 추출합니다.
    """
    if not cell_value or not isinstance(cell_value, str):
        return None
    
    val = cell_value.strip()
    
    # 'Key', '설명' 등은 건너뜀 (알파벳이 아니거나 길이가 맞지 않는 경우 필터링)
    # 1. 구분자(- 또는 _)로 분리 후 첫 번째 파트 가져오기
    part = re.split(r'[-_]', val)[0]
    
    # 2. 길이가 2자리이고 알파벳인 경우만 유효한 언어 코드로 인정 (ex: en, ko, id, ar)
    if len(part) == 2 and part.isalpha():
        return part.lower()
    
    return None

# ----------------------------------------------------
# 메인 로직
# ----------------------------------------------------
print("엑셀 파일을 업로드해주세요.")
uploaded = files.upload()

if not uploaded:
    print("파일이 업로드되지 않았습니다.")
else:
    file_path = list(uploaded.keys())[0]
    print(f"업로드된 파일: {file_path}")

    try:
        wb = load_workbook(file_path)

        target_sheets = [
            "Mobile App Web", "Mobile AppWeb",
            "Admin", "Admin-에스피텍", "Admin-유비온",
            "Cms"
        ]

        # 배치 번역 처리 함수
        def process_batch_translation(translator, text_list):
            results = []
            if not text_list:
                return results

            for i in range(0, len(text_list), BATCH_SIZE):
                batch = text_list[i : i + BATCH_SIZE]
                masked_batch = []
                batch_vars = []

                for text in batch:
                    m_text, vars_list = mask_variables(text)
                    masked_batch.append(m_text)
                    batch_vars.append(vars_list)

                translated_batch = []
                try:
                    time.sleep(random.uniform(0.5, 1.5))
                    translated_batch = translator.translate_batch(masked_batch)
                except Exception as e:
                    print(f"\n⚠️ 배치 번역 실패 (재시도 중...): {e}")
                    for idx, text in enumerate(masked_batch):
                        try:
                            time.sleep(1)
                            t_text = translator.translate(text)
                            translated_batch.append(t_text)
                        except:
                            translated_batch.append(text)

                final_batch = []
                for j, trans_text in enumerate(translated_batch):
                    if trans_text:
                        restored = unmask_variables(trans_text, batch_vars[j])
                        final_batch.append(restored)
                    else:
                        final_batch.append(trans_text)
                results.extend(final_batch)

            return results

        # 시트 순회
        for sheet_name in target_sheets:
            if sheet_name not in wb.sheetnames: continue

            print(f"\n📂 [{sheet_name}] 데이터 스캔 및 헤더 분석 중... 🚀")
            ws = wb[sheet_name]

            # 1. 헤더 행 및 소스(en-US) 위치 찾기
            header_row = None
            source_col = None
            
            # 처음 15행까지 스캔하여 'en-US'가 있는 위치를 찾음
            for r in range(1, min(16, ws.max_row + 1)):
                for c in range(1, ws.max_column + 1):
                    val = str(ws.cell(row=r, column=c).value).strip()
                    if val == "en-US":
                        header_row = r
                        source_col = c
                        break
                if header_row: break

            if not header_row or not source_col:
                print(f"   ⚠️ 'en-US' 컬럼을 찾을 수 없어 [{sheet_name}] 시트를 건너뜁니다.")
                continue

            # 2. 헤더 행을 분석하여 타겟 언어 컬럼들 자동 매핑
            # 예: id-ID -> id, en-IN -> en, ar-SA -> ar
            target_cols = {} # { col_idx: 'lang_code' }
            
            for c in range(1, ws.max_column + 1):
                if c == source_col: continue # 원본 컬럼은 스킵

                header_val = ws.cell(row=header_row, column=c).value
                lang_code = parse_lang_code(header_val)

                if lang_code:
                    target_cols[c] = lang_code
            
            print(f"   ℹ️ 감지된 언어: {list(set(target_cols.values()))}")
            if not target_cols:
                print("   ⚠️ 번역할 대상 언어 컬럼(예: id-ID, ar-SA)을 찾지 못했습니다.")
                continue

            # 3. 작업 목록 생성
            tasks_by_lang = {lang: [] for lang in set(target_cols.values())}
            total_skip_count = 0
            total_add_count = 0

            for row in range(header_row + 1, ws.max_row + 1):
                en_val = ws.cell(row=row, column=source_col).value

                if en_val and str(en_val).strip():
                    en_text = str(en_val).strip()

                    for col_idx, lang_code in target_cols.items():
                        target_cell = ws.cell(row=row, column=col_idx)
                        cell_val = target_cell.value

                        # 이미 값이 있으면 스킵
                        if cell_val is not None and str(cell_val).strip() != "":
                            total_skip_count += 1
                            continue

                        tasks_by_lang[lang_code].append({
                            'row': row,
                            'col': col_idx,
                            'text': en_text,
                            'header_origin': ws.cell(row=header_row, column=col_idx).value # 로깅용
                        })
                        total_add_count += 1

            print(f"   ℹ️ 스캔 결과: {total_skip_count}개 셀 건너뜀, {total_add_count}개 셀 작업 예정")

            if total_add_count == 0:
                print("   ✅ 모든 작업이 이미 완료되어 있습니다.")
                continue

            # 4. 언어별 번역 또는 복사 수행
            for lang_code, tasks in tasks_by_lang.items():
                if not tasks: continue

                # [중요] 타겟 언어코드가 'en'인 경우 (en-IN, en-PH 등) -> 번역 없이 원문 복사
                if lang_code == 'en':
                    print(f"   👉 [English Variant] 영어 변형({lang_code})은 원문 복사 중... ({len(tasks)}개)")
                    for task in tasks:
                        ws.cell(row=task['row'], column=task['col']).value = task['text']
                    continue

                # 그 외 언어는 번역 진행
                print(f"   👉 [{lang_code}] 번역 진행 중... ({len(tasks)}개)")

                translator = GoogleTranslator(source='en', target=lang_code)
                texts_to_translate = [t['text'] for t in tasks]
                translated_texts = []

                # tqdm 진행바 표시
                with tqdm(total=len(texts_to_translate), desc=f"   Translating to {lang_code}") as pbar:
                    for i in range(0, len(texts_to_translate), BATCH_SIZE):
                        batch_texts = texts_to_translate[i : i + BATCH_SIZE]
                        batch_results = process_batch_translation(translator, batch_texts)
                        translated_texts.extend(batch_results)
                        pbar.update(len(batch_texts))

                # 결과 엑셀에 쓰기
                for i, task in enumerate(tasks):
                    if i < len(translated_texts):
                        ws.cell(row=task['row'], column=task['col']).value = translated_texts[i]

            print(f"   ✨ [{sheet_name}] 작업 완료")

        output_path = "NextS_AutoDetected_Updated.xlsx"
        wb.save(output_path)
        print(f"\n🎉 모든 작업 완료! 저장됨: {output_path}")
        files.download(output_path)

    except Exception as e:
        print(f"오류 발생: {e}")
        # 오류 발생 시에도 현재까지 작업한 내용은 저장 시도
        try:
            wb.save("Backup_Error.xlsx")
            files.download("Backup_Error.xlsx")
        except:
            pass
