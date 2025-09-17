import traceback
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException
import time
import pyautogui
import utility.config as config
from operation.generator import login, recipe_copy, remove_steps, extract_from_excel, find_start_index, process_trashcan, process_material, update_material
from utility.utils import wait_and_click, zoom_out, param_click, input_text, wait_and_send_keys


# --- 공용: 로딩 마스크 대기용 함수 정의 ---
def wait_for_mask_to_disappear(driver, timeout=30):
    try:
        WebDriverWait(driver, timeout).until(
            EC.invisibility_of_element_located((By.CLASS_NAME, "x-mask"))
        )
    except TimeoutException:
        print("Timeout waiting for loading mask to disappear.")


# --- 엑셀 파싱 제너레이터 ---
def process_excel_data(file_path, sheet_name=None):
    """
    yield 기반 엑셀 파서.
    - sheet_name 지정 시: 해당 시트만 1회 yield
    - sheet_name 미지정 시: 모든 시트를 순회하며 yield
    ※ 드라이버는 만들지 않습니다.
    """
    xls = pd.ExcelFile(file_path)
    target_sheets = [sheet_name] if sheet_name is not None else xls.sheet_names

    for s in target_sheets:
        df = pd.read_excel(xls, sheet_name=s, header=None)
        df.columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']

        # recipe의 2.1 / 3.1 / 4.1 / 5.1 / 6.1 / 7.1 스텝 정의
        values_2_1 = [f'2.1.{i}' for i in range(1, 11)]
        values_3_1 = [f'3.1.{i}' for i in range(1, 11)]
        values_4_1 = [f'4.1.{i}' for i in range(1, 11)]
        values_5_1 = [f'5.1.{i}' for i in range(1, 11)]
        values_6_1 = [f'6.1.{i}' for i in range(1, 11)]
        values_7_1 = [f'7.1.{i}' for i in range(1, 11)]

        # extract_from_excel 함수 이용하여 엑셀에서 Recipe Name, Sample, Sample_liquid 추출
        extracted = extract_from_excel(file_path, s)
        Recipe = extracted.get('Recipe Name')
        Sample = extracted.get('Sample')
        Sample_liquid = extracted.get('Sample_liquid')
        #
        yield {
            'sheet_name': s,
            'df': df,
            'values_2_1': values_2_1,
            'values_3_1': values_3_1,
            'values_4_1': values_4_1,
            'values_5_1': values_5_1,
            'values_6_1': values_6_1,
            'values_7_1': values_7_1,
            # extract_from_excel 에서 파징 (Recipe, Sample, Sample_liquid, choice, choice_detail)
            'Recipe': Recipe,
            'Sample': Sample,
            'Sample_liquid': Sample_liquid,
            'choice': extracted.get('choice'),
            'choice_detail': extracted.get('choice_detail'),
            'excel_data': extracted
        }

# observation
# Material 기반으로 전처리/분석 단계를 채우는 동작을 수행
# 엑셀 시트에서 "Material 1" 이라는 문자열이 들어 있는 행의 Index 탐색
# > Material의 E 열이 'Y' 이면 2.1.X 단계에 자동 입력 /  Material의 F 열이 'Y' 이면 3.1.X 단계에 자동 입력
#  >> Material 행에서 G열이 'Y' (고체) 일 경우 삭제 대상 파라미터 인덱스 목록 (3,4) 삭제함
#  >> Material 행에서 H열이 'Y' (액체) 일 경우 삭제 대상 파라미터 인덱스 목록  (1,2,8) 을 삭제함

def observation(driver, df):
    start_index = find_start_index(df)
    if start_index is None:
        print("'Material 1'을 찾을 수 없습니다.")
        return

    # 변수 선언
    saved_values = []
    value_index_2_1 = 0
    value_index_3_1 = 0
    preprocessing = 1
    analyze = 21

    # 파라미터 값 생성
    values_2_1 = [f'2.1.{i}' for i in range(1, 21)]
    values_3_1 = [f'3.{i}' for i in range(1, 21)]

    # Material 조건에 따라 기입
    for index in range(start_index, len(df)):
        if pd.notna(df.at[index, 'A']) and 'Material' in df.at[index, 'A']:
            input_value = df.at[index, 'B']
            saved_values.append(input_value)

            if df.at[index, 'E'] == 'Y' and value_index_2_1 < len(values_2_1):
                try:
                    process_material(driver, df, index, input_value, values_2_1[value_index_2_1])
                    update_material(driver, preprocessing, input_value)
                    preprocessing += 1
                    value_index_2_1 += 1
                except Exception as e:
                    print(f"Error occurred for value {values_2_1[value_index_2_1]}: {e}")

            if df.at[index, 'F'] == 'Y' and value_index_3_1 < len(values_3_1):
                try:
                    process_material(driver, df, index, input_value, values_3_1[value_index_3_1])
                    update_material(driver, analyze, input_value)
                    analyze += 1
                    value_index_3_1 += 1
                except Exception as e:
                    print(f"Error occurred for value {values_3_1[value_index_3_1]}: {e}")

            if value_index_2_1 >= len(values_2_1) and value_index_3_1 >= len(values_3_1):
                break
        else:
            break

    print("Saved Values:", saved_values)


# update_sample
# 1) Recipe의 'Samples' 탭으로 이동 > 'Sample Name' 인 xpath를 확인 > sample Name을 변경
# 2) Process 탭 클릭 > 오른쪽의 Parameter 탭 클릭 > 'Expand All' 버튼 > 'Properties View' 버튼 클릭
# 3) sample name + ['_제조번호', '_제조일자', '_채취일자', '_시험 시작일자'] 로 Parameter 이름을 변경
# 4) 1번스텝 클릭 > sample_name + [_전처리 확인여부] 로 parameter 이름 변경
# 5) 1.1 스텝 클릭 > sample_name + [_샘플이름', ' SA_칭량'] 으로 parameter 이름 변경
# 6) Sample 정보가 액체인경우 > 필요 없는 process 삭제
# 7) 4번스텝 클릭 > sample_name + [_샘플이름, _결과,  시험 종료일자] 로 parameter 이름 변경


def update_sample(driver, Sample, Sample_liquid):
    # Sample 기입
    wait_and_click(driver, 20, By.XPATH, "//a[text()='Samples']")
    sample_element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
        (By.XPATH, '//div[contains(@class, "x-grid-cell-inner") and text()="Sample"]')))
    actions = ActionChains(driver)
    actions.double_click(sample_element).perform()
    time.sleep(2)
    
    for _ in range(11):
        actions.send_keys(Keys.BACKSPACE).perform()
        time.sleep(0.1)
    actions.send_keys(Sample).perform()
    time.sleep(2)
    
    wait_and_click(driver, 20, By.XPATH, "//a[text()='Process']")
    param_click(driver)
    # 우측 파라미터 정보에 파라미터 탭, Expand All 버튼, Properties View 버튼 차례로 클릭
    input_text(driver, Sample, ['_제조번호', '_제조일자', '_채취일자', '_시험 시작일자'])
    
    wait_and_click(driver, 20, By.XPATH, "//div[normalize-space(text())='1']")
    param_click(driver)
    input_text(driver, Sample, ['_전처리 확인여부'])
    
    wait_and_click(driver, 20, By.XPATH, "//div[normalize-space(text())='1.1']")
    param_click(driver)
    input_text(driver, Sample, ['_샘플이름', ' SA_칭량'])
    
    if Sample_liquid == 'Y':
        process_trashcan(driver, [1])  # 액체 Y인 경우 필요없는 Process 삭제
    
    wait_and_click(driver, 20, By.XPATH, "//div[normalize-space(text())='4']")
    param_click(driver)
    input_text(driver, Sample, ['_샘플이름', ' _결과', ' 시험 종료일자'])



# update_sample density
# 1) Recipe의 'Samples' 탭으로 이동 > 'Sample Name' 인 xpath를 확인 > sample Name을 변경
# 2) Process 탭 클릭 > 오른쪽의 Parameter 탭 클릭 > 'Expand All' 버튼 > 'Properties View' 버튼 클릭
# 3) sample name + ['_제조번호', '_제조일자', '_채취일자', '_비중 시험 시작시간', '_비중_MIN', '_비중_MAX'] 로 Parameter 이름을 변경
# 4) 2번스텝 클릭 > sample_name + ['_비중 시험법 선택', ' _비중'] 으로 parameter 이름 변경

def update_sample_density(driver, Sample, Sample_liquid):
    # Sample 기입
    wait_and_click(driver, 20, By.XPATH, "//a[text()='Samples']")
    sample_element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
        (By.XPATH, '//div[contains(@class, "x-grid-cell-inner") and text()="Sample"]')))
    actions = ActionChains(driver)
    actions.double_click(sample_element).perform()
    time.sleep(2)
    
    for _ in range(11):
        actions.send_keys(Keys.BACKSPACE).perform()
        time.sleep(0.1)
    actions.send_keys(Sample).perform()
    time.sleep(2)
    
    wait_and_click(driver, 20, By.XPATH, "//a[text()='Process']")
    param_click(driver)
    # 우측 파라미터 정보에 파라미터 탭, Expand All 버튼, Properties View 버튼 차례로 클릭
    input_text(driver, Sample, ['_제조번호', '_제조일자', '_채취일자', '_비중 시험 시작시간', '_비중_MIN', '_비중_MAX'])
    
    wait_and_click(driver, 20, By.XPATH, "//div[normalize-space(text())='2']")
    param_click(driver)
    input_text(driver, Sample, ['_비중 시험법 선택', ' _비중'])



# add_paramdesc(driver, df, 4)
# 1) 엑셀의 A열에서 'Param Dsc.1" 이 등장하는 행을 시작점으로 잡음, Param Dsc로 잡힌 데이터 만큼 실행
#  >> 못 찾으면 패스
# 2) 시작점이 잡히면 A열에 ‘Param Dsc’가 포함된 모든 행 을 순회하여 B열 값을 항목명으로 UI에 Param Dsc를 추가
# 3) Add > Filter에서 B열 값을 입력
# 4) Add Selected 클릭
# 5) 추가한 Parameter에서 'Process Result' 클릭
# 6) 'Apply' 클릭

def add_paramdsc(driver, df, value):
    # Param Dsc. 관련 파라미터 입력 처리
    wait_and_click(driver, 20, By.XPATH, f"//div[normalize-space(text())='{value}']")
    wait_and_click(driver, 20, By.XPATH, "//span[normalize-space(text())='Details']")
    param_click(driver)

    start_index = None
    for index, cell_value in df['A'].items():
        if pd.notna(cell_value) and 'Param Dsc. 1' in cell_value:
            start_index = index
            break

    if start_index is not None:
        saved_values = []
        for index in range(start_index, len(df)):
            if pd.notna(df.at[index, 'A']) and 'Param Dsc' in df.at[index, 'A']:
                input_value = df.at[index, 'B']
                saved_values.append(input_value)

                wait_and_click(driver, 20, By.XPATH, "//span[starts-with(@id, 'splitbutton-') and normalize-space(text())='Add']")
                filter_label = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//span[normalize-space(text())='Filter:']")))
                label_id = filter_label.get_attribute("id")
                input_id = label_id.replace("labelTextEl", "inputEl")
                wait_and_send_keys(driver, 20, By.ID, input_id, input_value)
                wait_and_click(driver, 20, By.XPATH, "//div[contains(@class, 'x-form-spinner-up')]")
                wait_and_click(driver, 20, By.XPATH, "//span[normalize-space(text())='Add Selected']")

                img_elements = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "img.x-tool-img.x-tool-tag")))
                element = img_elements[2]
                element.click()
                time.sleep(2)

                input_element = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, "//input[starts-with(@id, 'parametertagcombo-') and contains(@id, '-inputEl')]")))
                actions = ActionChains(driver)
                actions.double_click(input_element).perform()
                time.sleep(2)
                actions.send_keys('Process Result').perform()
                time.sleep(2)
                actions.send_keys(Keys.RETURN).perform()
                time.sleep(2)
                wait_and_click(driver, 20, By.XPATH, "//span[@class='x-btn-inner x-btn-inner-default-small' and text()='Apply']")


# physicochemistry
# Material 기반으로 전처리/분석 단계를 채우는 동작을 수행
# 엑셀 시트에서 "Material 1" 이라는 문자열이 들어 있는 행의 Index 탐색
#  > E 열이 'Y' 이면 2.1.X 단계에 자동 입력 /  F 열이 'Y' 이면 3.1.X 단계에 자동 입력
#   >> Material 행에서 G열이 'Y' (고체) 일 경우 삭제 대상 파라미터 인덱스 목록 (3,4) 삭제함
#   >> Material 행에서 H열이 'Y' (액체) 일 경우 삭제 대상 파라미터 인덱스 목록  (1,2,8) 을 삭제함

def physicochemistry(driver, df):
    # 전처리 및 분석 단계 파라미터 자동 기입
    start_index = find_start_index(df)
    if start_index is None:
        print("'Material 1'을 찾을 수 없습니다.")
        return

    saved_values = []
    value_index_2_1 = 0
    value_index_3_1 = 0
    preprocessing = 1
    analyze = 21
    values_2_1 = [f'2.1.{i}' for i in range(1, 21)]
    values_3_1 = [f'3.{i}' for i in range(1, 21)]

    for index in range(start_index, len(df)):
        if pd.notna(df.at[index, 'A']) and 'Material' in df.at[index, 'A']:
            input_value = df.at[index, 'B']
            saved_values.append(input_value)

            if df.at[index, 'E'] == 'Y' and value_index_2_1 < len(values_2_1):
                try:
                    process_material(driver, df, index, input_value, values_2_1[value_index_2_1])
                    update_material(driver, preprocessing, input_value)
                    preprocessing += 1
                    value_index_2_1 += 1
                except Exception as e:
                    print(f"Error occurred for value {values_2_1[value_index_2_1]}: {e}")

            if df.at[index, 'F'] == 'Y' and value_index_3_1 < len(values_3_1):
                try:
                    process_material(driver, df, index, input_value, values_3_1[value_index_3_1])
                    update_material(driver, analyze, input_value)
                    analyze += 1
                    value_index_3_1 += 1
                except Exception as e:
                    print(f"Error occurred for value {values_3_1[value_index_3_1]}: {e}")

            if value_index_2_1 >= len(values_2_1) and value_index_3_1 >= len(values_3_1):
                break
        else:
            break

    print("Saved Values:", saved_values)


# instrument
# 1) Recipe 편집 초기 화면에서 첫 번째 필드 에 Recipe Name 입력
# 2) sleep 후 Zoom out (3)  실행
# 3) 엑셀의 각 행을 순회하여 A열의 Material이고 D열 값 (표준품 여부)'Y' 인 경우를 standard_values에 추가
#  >> 시험에 사용하는 표준물질 이름 리스트

# 4) 나중에 삭제할 파라미터 수 계산
# 5) 7.2 라벨 영역을 클릭 > param_click 실행 > parameters 이동 > expand all > standard values 개수만큼 textarea에 '[물질명]_SST 판정' 으로 parameter 이름이력
# 6) 7.3.1 라벨 클릭 > param_click 실행 > parameters 이동 > expand all > standard values 개수만큼 textarea에 '[물질명]_ST_Area' / '[물질명]_ST_R.T.' 입력 


def instrument(driver, df, Recipe):
    # Process Tree Name 기입
    tabpanel_divs = driver.find_elements(By.XPATH, "//div[starts-with(@id, 'tabpanel-') and contains(@id, '-body')]")
    for tabpanel_div in tabpanel_divs:
        entity_textfields = tabpanel_div.find_elements(By.XPATH, ".//input[starts-with(@id, 'entitytextfield-') and contains(@id, '-inputEl')]")
        if entity_textfields:
            first_input = entity_textfields[0]
            time.sleep(2)
            first_input.clear()
            time.sleep(2)
            first_input.send_keys(Recipe)
            time.sleep(2)
    zoom_out(3)
    time.sleep(5)

    # standard 대상 수집
    standard_values = []
    for _, row in df.iterrows():
        if pd.notna(row['A']) and 'Material' in str(row['A']) and row['D'] == 'Y':
            standard_values.append(row['B'])
    delete = len(standard_values)

    # 7.2 섹션
    wait_and_click(driver, 20, By.XPATH, "//div[text()='7.2']")
    param_click(driver)

    # ★ 여기서 다시 textarea 수집 (스코프/갱신 이슈 방지)
    tabpanel_divs = driver.find_elements(By.XPATH, "//div[starts-with(@id, 'tabpanel-') and contains(@id, '-body')]")
    textarea_fields = []
    for tp in tabpanel_divs:
        cand = tp.find_elements(By.XPATH, ".//textarea[starts-with(@id, 'textarea-') and contains(@id, '-inputEl')]")
        if cand:
            textarea_fields = cand
            break

    for i, value in enumerate(standard_values):
        if i < len(textarea_fields):
            textarea_field = textarea_fields[i]
            textarea_field.clear()
            textarea_field.send_keys(f"{value}_SST 판정")
            time.sleep(2)
        else:
            print(f"Warning: Not enough textarea fields for all standard values. Stopped at index {i}")
            break

    # 필요 없는 파라미터 삭제
    img_elements = WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.CSS_SELECTOR, "img.x-tool-img.x-tool-trashcan"))
    )
    if len(img_elements) >= 10:
        for i in range(delete, 10):
            element = img_elements[i]
            time.sleep(2)
            element.click()
            time.sleep(2)
            wait_and_click(driver, 20, By.XPATH, "//span[normalize-space(text())='예']")
    else:
        print("There are less than 9 elements matching the criteria.")

    # 7.3.1 섹션
    wait_and_click(driver, 20, By.XPATH, "//div[text()='7.3.1']")
    param_click(driver)

    # 다시 textarea 수집
    tabpanel_divs = driver.find_elements(By.XPATH, "//div[starts-with(@id, 'tabpanel-') and contains(@id, '-body')]")
    textarea_fields = []
    for tp in tabpanel_divs:
        cand = tp.find_elements(By.XPATH, ".//textarea[starts-with(@id, 'textarea-') and contains(@id, '-inputEl')]")
        if cand:
            textarea_fields = cand
            break

    start_index = 1
    for value in standard_values:
        if start_index < len(textarea_fields):
            textarea_fields[start_index].clear()
            textarea_fields[start_index].send_keys(f"{value}_ST Area")
            time.sleep(2)
        else:
            print(f"Warning: Not enough textarea fields for {value}_ST Area")
            break

        if start_index + 1 < len(textarea_fields):
            textarea_fields[start_index + 1].clear()
            textarea_fields[start_index + 1].send_keys(f"{value}_SA_R.T.")
            time.sleep(2)
        else:
            print(f"Warning: Not enough textarea fields for {value}_SA_R.T.")
            break

        start_index += 2

    # 필요 없는 파라미터 삭제
    img_elements = WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.CSS_SELECTOR, "img.x-tool-img.x-tool-trashcan"))
    )
    if len(img_elements) >= 21:
        for i in range(delete * 2 + 1, 21):
            element = img_elements[i]
            time.sleep(2)
            element.click()
            time.sleep(2)
            wait_and_click(driver, 20, By.XPATH, "//span[normalize-space(text())='예']")
    else:
        print("There are less than 9 elements matching the criteria.")


# sample_instrument
# 1) Recipe에서 Samples 탭 > xpath 텍스트가 'Sample' 인 것을 찾아 지우고 Excel의 Sample Name을 입력 > Process 탭 클릭
# 2) Process 클릭하자마자 보이는 0번스텝의 Parameter 이름을 'Sample_Name'+['_제조번호', '_제조일자', '_채취일자', '_시험 시작일자'] 로 변경한다.
# 3) 1.1 스텝으로 이동해서 Parameters 탭 > Expand All > Parameter 이름을 'Sample_Name'+['_샘플이름', ' SA_칭량'] 으로 수정한다.
#  >> 이후 Sample이 액상인지를 보고, 액상이면 현재 스텝에서 2번째 인덱스인 파라미터를 삭제한다.
# 4) 7.4 스텝으로 이동해서 Parameters 탭 > Expand All > 'Sample_Name' + [' _확인시험', ' 시험 종료일자'] 로 수정한다.

def sample_instrument(driver, Sample, Sample_liquid):
    # Sample 기입
    wait_and_click(driver, 20, By.XPATH, "//a[text()='Samples']")
    sample_element = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, '//div[contains(@class, "x-grid-cell-inner") and text()="Sample"]'))
    )

    actions = ActionChains(driver)
    actions.double_click(sample_element).perform()
    time.sleep(2)

    for _ in range(11):
        actions.send_keys(Keys.BACKSPACE).perform()
        time.sleep(0.1)

    actions.send_keys(Sample).perform()
    time.sleep(2)
    wait_and_click(driver, 20, By.XPATH, "//a[text()='Process']")

    # 파라미터 Name 기입(상단)
    tabpanel_divs = driver.find_elements(By.XPATH, "//div[starts-with(@id, 'tabpanel-') and contains(@id, '-body')]")
    for tp in tabpanel_divs:
        param_click(driver)
        textarea_fields = tp.find_elements(By.XPATH, ".//textarea[starts-with(@id, 'textarea-') and contains(@id, '-inputEl')]")
        for i, textarea_field in enumerate(textarea_fields[:4]):
            textarea_field.clear()
            textarea_field.send_keys(Sample + ['_제조번호', '_제조일자', '_채취일자', '_시험 시작일자'][i])
            time.sleep(2)

    # 1.1 섹션
    wait_and_click(driver, 20, By.XPATH, "//div[text()='1.1']")
    tabpanel_divs = driver.find_elements(By.XPATH, "//div[starts_with(@id, 'tabpanel-') and contains(@id, '-body')]".replace("starts_with", "starts-with"))
    for tp in tabpanel_divs:
        param_click(driver)
        textarea_fields = tp.find_elements(By.XPATH, ".//textarea[starts-with(@id, 'textarea-') and contains(@id, '-inputEl')]")
        for i, textarea_field in enumerate(textarea_fields[:2]):
            textarea_field.clear()
            textarea_field.send_keys(Sample + ['_샘플이름', ' SA_칭량'][i])
            time.sleep(2)

    # 조건 별 파라미터 삭제
    if str(Sample_liquid).strip().upper() == 'Y':
        img_elements = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, "img.x-tool-img.x-tool-trashcan"))
        )
        if len(img_elements) > 1:
            img_elements[1].click()
            time.sleep(2)
            wait_and_click(driver, 20, By.XPATH, "//span[normalize-space(text())='예']")

    # 7.4 섹션
    wait_and_click(driver, 20, By.XPATH, "//div[text()='7.4']")
    tabpanel_divs = driver.find_elements(By.XPATH, "//div[starts-with(@id, 'tabpanel-') and contains(@id, '-body')]")
    for tp in tabpanel_divs:
        param_click(driver)
        textarea_fields = tp.find_elements(By.XPATH, ".//textarea[starts-with(@id, 'textarea-') and contains(@id, '-inputEl')]")
        for i, textarea_field in enumerate(textarea_fields[:2]):
            textarea_field.clear()
            textarea_field.send_keys(Sample + [' _확인시험', ' 시험 종료일자'][i])
            time.sleep(2)


# add_paramdsc_instrument
# ** 함수 실행은 sample_instrument 동작이 끝난 직후 바로 실행되므로, 7.4 스텝에서 동작됨.

# 1) Excel 파일에서 Param Dsc 시작 위치 찾기
# 2) 해당 범위 A열에서 Param Dsc 문자열을 찾아, B열 값을 input_value로 저장
# 3) Compose 화면에서 add 버튼 클릭 > Excel의 B열 값대로 필터링 > Parameter 추가
# 4) 추가한 Parameter에 대해 Process Result 태그 설정 > Apply 버튼 클릭

def add_paramdsc_instrument(driver, df, ID, PW, Recipe):
    # Param DSC 시작 인덱스 찾기
    start_index = None
    for index, value in df['A'].items():
        if pd.notna(value) and 'Param Dsc. 1' in str(value):
            start_index = index
            break

    if start_index is None:
        return

    saved_values = []
    # value_index_7_4 = 0  # (사용하지 않으면 제거 가능)

    for index in range(start_index, len(df)):
        if pd.notna(df.at[index, 'A']) and 'Param Dsc' in str(df.at[index, 'A']):
            input_value = df.at[index, 'B']
            saved_values.append(input_value)

            wait_and_click(driver, 20, By.XPATH, "//span[starts-with(@id, 'splitbutton-') and normalize-space(text())='Add']")

            filter_label = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//span[normalize-space(text())='Filter:']"))
            )
            label_id = filter_label.get_attribute("id")
            input_id = label_id.replace("labelTextEl", "inputEl")

            input_field = driver.find_element(By.ID, input_id)
            input_field.clear()
            input_field.send_keys(input_value)
            time.sleep(2)
            wait_and_click(driver, 20, By.XPATH, "//div[contains(@class, 'x-form-spinner-up')]")
            wait_and_click(driver, 20, By.XPATH, "//span[normalize-space(text())='Add Selected']")

            # Plan function 추가
            img_elements = WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "img.x-tool-img.x-tool-tag"))
            )
            element = img_elements[0]
            element.click()
            time.sleep(2)

            input_element = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.XPATH, "//input[starts-with(@id, 'parametertagcombo-') and contains(@id, '-inputEl')]"))
            )
            actions = ActionChains(driver)
            actions.double_click(input_element).perform()
            time.sleep(2)
            actions.send_keys('Process Result').perform()
            time.sleep(2)
            actions.send_keys(Keys.RETURN).perform()
            time.sleep(2)
            wait_and_click(driver, 20, By.XPATH, "//span[@class='x-btn-inner x-btn-inner-default-small' and text()='Apply']")

            # Planned 입력 스킵 플래그
            SKIP_PLANNED_INPUT = True
            if not SKIP_PLANNED_INPUT:
                wait_and_click(driver, 20, By.XPATH, "//span[normalize-space(text())='Expand All']")
                # ... (필요 시 기존 Planned 입력/Save 로직 복원)
                wait_for_mask_to_disappear(driver)
                time.sleep(2)


# process_materials_instrument_hplc
# 기기분석_HPLC 레시피에서 동작하는 함수 (기기매핑 포함)

# 1) Excel 행을 Material 별로 읽어서 최대 10개까지 각 섹션 단계 (2.1.x ~ 7.1.x) 에 매핑 준비
# 2) A열에서 'Material 1' 을 포함하는 첫 행을 찾아 시작점 지정, 각 material마다 B열 값을 input value로 가져옴
# 3) Material 영역의 Flag에 따라 섹션이 결정됨
#  > Detail 탭 열기
#  > Text 필드 지우고 Material 값 이름을 스텝 이름으로 입력
#  > Parameters 탭 > Expand ALL 버튼 클릭
#  > Parameter 이름 수정 > input value + 각 name 으로 parameter 이름 수정
#  > 분석기기 선택 + Parameter Reading 입력 + 입력 후 browser 재시작
#     >> 2.1.x/3.1.x/4.1.x/5.1.x/6.1.x/7.1.x 모든 섹션에서 기기분석_HPLC 는 이부분을 유지
#  > Sample 탭에서 해당 Material의 이름으로 변경
# 기존의 장비 매핑 동작에 generator.py에서 parsing 한 'equipment_primary' 를 분석장비 목록으로 받아와 실행됨.
def process_materials_instrument_hplc(driver, df, ID, PW, Recipe, Sample, equipment_primary=None):

    # 파라미터 값 생성
    values_2_1 = [f'2.1.{i}' for i in range(1, 11)]
    values_3_1 = [f'3.1.{i}' for i in range(1, 11)]
    values_4_1 = [f'4.1.{i}' for i in range(1, 11)]
    values_5_1 = [f'5.1.{i}' for i in range(1, 11)]
    values_6_1 = [f'6.1.{i}' for i in range(1, 11)]
    values_7_1 = [f'7.1.{i}' for i in range(1, 11)]


    # 시작할 행 인덱스
    start_index = None

    # 'Material 1' 값을 찾기
    for index, value in df['A'].items():
        if pd.notna(value) and 'Material 1' in value:
            start_index = index
            break
    #변수 선언
    if start_index is not None:
        saved_values = []
        value_index_2_1 = 0 
        value_index_3_1 = 0  
        value_index_4_1 = 0
        value_index_5_1 = 0
        value_index_6_1 = 0
        value_index_7_1 = 0
        standard = 1
        preprocessing = 11
        analyze = 21
        mobile = 31

        #Material 조건에 따라 기입
        all_materials_processed = False
        for index in range(start_index, len(df)):
            if pd.notna(df.at[index, 'A']) and 'Material' in df.at[index, 'A']:
                input_value = df.at[index, 'B']
                saved_values.append(input_value)

                if df.at[index, 'D'] == 'Y' and value_index_2_1 < len(values_2_1):
                    try:
                        process_material(driver, df, index, input_value, values_2_1[value_index_2_1])

                        wait_and_click(driver, 20, By.XPATH, f"//div[normalize-space(text())='{values_7_1[value_index_7_1]}']") 
                        wait_and_click(driver, 20, By.XPATH, "//span[normalize-space(text())='Details']") 

                        tabpanel_divs = driver.find_elements(By.XPATH, "//div[starts-with(@id, 'tabpanel-') and contains(@id, '-body')]")
                        
                        for tabpanel_div in tabpanel_divs:
                            entity_textfields = tabpanel_div.find_elements(By.XPATH, ".//input[starts-with(@id, 'entitytextfield-') and contains(@id, '-inputEl')]")
                            if entity_textfields:
                                first_input = entity_textfields[0]
                                time.sleep(2)
                                # 텍스트 지우기
                                first_input.clear()
                                time.sleep(2)
                                first_input.send_keys(input_value)
                                time.sleep(2)
                                
                                wait_and_click(driver, 20, By.XPATH, "//span[normalize-space(text())='Parameters']") 
                                wait_and_click(driver, 20, By.XPATH, "//span[normalize-space(text())='Expand All']")
                                wait_and_click(driver, 20, By.XPATH, "//span[contains(@class, 'x-btn-icon-el') and normalize-space(text())=' ']")

                                textarea_fields = tabpanel_div.find_elements(By.XPATH, ".//textarea[starts-with(@id, 'textarea-') and contains(@id, '-inputEl')]")
                                for i, textarea_field in enumerate(textarea_fields[:37]):
                                    textarea_field.clear()
                                    text = input_value + [' ST-1_Area', ' ST-2_Area', ' ST-3_Area', ' ST-4_Area', ' ST-5_Area', ' ST-6_Area', 
                                                            ' ST-1_R.T.', ' ST-2_R.T.', ' ST-3_R.T.', ' ST-4_R.T.', ' ST-5_R.T.', ' ST-6_R.T.',
                                                            ' ST-1_Asymmetry', 'ST-2_Asymmetry', 'ST-3_Asymmetry', 'ST-4_Asymmetry', 'ST-5_Asymmetry', 'ST-6_Asymmetry',
                                                            ' ST-1_Theoretical plates', ' ST-2_Theoretical plates', ' ST-3_Theoretical plates', ' ST-4_Theoretical plates', ' ST-5_Theoretical plates', ' ST-6_Theoretical plates',
                                                            ' ST-1_Resolution', ' ST-2_Resolution', ' ST-3_Resolution', ' ST-4_Resolution', ' ST-5_Resolution', ' ST-6_Resolution',
                                                            ' ST_Average Area', ' ST_Average R.T.', ' ST_Average Asymmetry', ' ST_Average Theoretical plates', ' ST_Average Resolution', ' ST_R.T. RSD(%)',' ST_Area RSD(%)' ][i]
                                    textarea_field.send_keys(text)
                                    time.sleep(2)

                                # --- 분석장비 매핑 (equipment_primary만 사용할 수 있게 수정) ---
                                wait_and_click(driver, 20, By.XPATH,
                                    "//span[contains(@class, 'x-btn-icon-el') and contains(@style, 'parameters_datacollection_light.png')]")

                                if not equipment_primary:
                                    print("[경고] equipment_primary 값이 없어 장비 매핑을 건너뜁니다.")
                                else:
                                    print("분석장비(primary):", equipment_primary)

                                    # 매핑 횟수는 기존과 동일하게 30회 유지
                                    for i in range(30):
                                        try:
                                            element = WebDriverWait(driver, 10).until(
                                                EC.element_to_be_clickable(
                                                    (By.XPATH, f"(//input[@placeholder='Select Recipe Equipment'])[{i+1}]")
                                                )
                                            )
                                            driver.execute_script("arguments[0].click();", element)
                                            time.sleep(1)
                                            element.send_keys(equipment_primary)  # 항상 동일한 값 입력
                                            element.send_keys(Keys.RETURN)
                                            time.sleep(2)
                                        except (TimeoutException, StaleElementReferenceException) as e:
                                            print(f"Error interacting with element {i+1}: {e}")
                                            continue

                                    print("Completed 30 iterations")


                                texts = ['Area ST-1', 'Area ST-2', 'Area ST-3', 'Area ST-4', 'Area ST-5', 'Area ST-6', 
                                        'Retention Time Min ST-1', 'Retention Time Min ST-2', 'Retention Time Min ST-3', 'Retention Time Min ST-4', 'Retention Time Min ST-5', 'Retention Time Min ST-6',
                                        'Asymmetry ST-1', 'Asymmetry ST-2', 'Asymmetry ST-3', 'Asymmetry ST-4', 'Asymmetry ST-5', 'Asymmetry ST-6',
                                        'Theoretical Plates ST-1', 'Theoretical Plates ST-2', 'Theoretical Plates ST-3', 'Theoretical Plates ST-4', 'Theoretical Plates ST-5', 'Theoretical Plates ST-6',
                                        'Resolution ST-1', 'Resolution ST-2', 'Resolution ST-3', 'Resolution ST-4', 'Resolution ST-5', 'Resolution ST-6']

                                for i in range(30):
                                    try:
                                        text = texts[i % len(texts)]
                                        element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, f"(//input[@placeholder='Select Reading'])[{i+1}]")))
                                        driver.execute_script("arguments[0].click();", element)
                                        time.sleep(1)
                                        element.send_keys(text)
                                        element.send_keys(Keys.RETURN)
                                        time.sleep(2)
                                        
                                    except (TimeoutException, StaleElementReferenceException) as e:
                                        print(f"Error interacting with element {i+1}: {e}")
                                        continue

                                print("Completed 30 iterations")  
                            
                            update_material(driver, standard, input_value)
                            standard += 1
                            value_index_2_1 += 1
                            value_index_7_1 += 1
                            
                            #브라우저 다시 시작
                            save_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[normalize-space(text())='Save']")))
                            save_button.click()
                            wait_for_mask_to_disappear(driver, 60)
                            time.sleep(5)
                            driver.quit()
                            time.sleep(5)

                            #크롬 열기
                            service = Service()
                            options = webdriver.ChromeOptions()
                            driver = webdriver.Chrome(service=service, options=options)    
                            driver.maximize_window()

                            #Compose 접속하기
                            driver.get("https://hubdev.daphgmp.dongasocio.com:9953/foundation/hub/")

                            #로그인 하기

                            if config.IS_DA_DEV == "Y":
                                username_xpath = "//input[@placeholder='이메일 또는 사용자 이름']"
                                password_xpath = "//input[@placeholder='암호']"
                            else:
                                username_xpath = "//input[@placeholder='Username']"
                                password_xpath = "//input[@placeholder='Password']"
                                
                            username = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, username_xpath)))
                            username.click()
                            username.send_keys(ID)
                            
                            password = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, password_xpath)))
                            password.click()
                            password.send_keys(PW)
                            password.send_keys(Keys.RETURN)

                            time.sleep(2)

                            #compose로 이동
                            driver.get("https://cncdev.daphgmp.dongasocio.com:9963/compose/")
                            time.sleep(5)

                            #Std Recipe 찾기
                            button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//div[@title='Recipe Filters']")))
                            button.click()                                                                 

                            button = driver.find_element(By.NAME, "formRecipeName")
                            button.send_keys(Recipe)          
                            button.send_keys(Keys.RETURN)
                            time.sleep(5)

                            button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[@title='Open the recipe in Compose.' and text()='{}']".format(Recipe))))
                            button.click()                                                                  
                            time.sleep(2)

                            Process_link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//a[text()='Process']")))
                            Process_link.click()
                            time.sleep(2)

                            expand_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//span[contains(@id, 'button-') and contains(text(), 'Expand')]")))
                            expand_button.click()
                            time.sleep(2)

                            # Ctrl 키를 누른 상태에서 - 키를 여러 번 눌러서 줌 아웃을 수행
                            for _ in range(6):  # 필요에 따라 반복 횟수 조절
                                # Ctrl와 - 키 조합하여 줌 레벨 조정
                                pyautogui.keyDown('ctrl')
                                pyautogui.press('-')
                                pyautogui.keyUp('ctrl')
                                time.sleep(0.1)  # 각 조작 사이에 약간의 지연을 추가
                                               
                    except Exception as e:
                        print(f"Error occurred for value {values_2_1[value_index_2_1]}: {e}")

                if df.at[index, 'D'] == 'N' and df.at[index, 'E'] == 'Y' and value_index_3_1 < len(values_3_1):
                    try:
                        process_material(driver, df, index, input_value, values_3_1[value_index_3_1])                    
                        update_material(driver, preprocessing, input_value)

                        preprocessing += 1
                        value_index_3_1 += 1  

                    except Exception as e:
                        print(f"Error occurred for value {values_3_1[value_index_3_1]}: {e}")

                if df.at[index, 'D'] == 'N' and df.at[index, 'F'] == 'Y' and value_index_4_1 < len(values_4_1):
                    try:
                        process_material(driver, df, index, input_value, values_4_1[value_index_4_1])
                        update_material(driver, analyze, input_value)

                        analyze += 1
                        value_index_4_1 += 1  

                    except Exception as e:
                        print(f"Error occurred for value {values_4_1[value_index_4_1]}: {e}")   

                if df.at[index, 'D'] == 'N' and df.at[index, 'I'] == 'Y' and value_index_5_1 < len(values_5_1):
                    try:

                        process_material(driver, df, index, input_value, values_5_1[value_index_5_1])                    
                        update_material(driver, mobile, input_value)

                        mobile += 1

                        value_index_5_1 += 1 

                    except Exception as e:
                        print(f"Error occurred for value {values_5_1[value_index_5_1]}: {e}")                

                if df.at[index, 'J'] == 'Y':
                    try:
                        wait_and_click(driver, 20, By.XPATH, f"//div[normalize-space(text())='{values_6_1[value_index_6_1]}']")                    
                        wait_and_click(driver, 20, By.XPATH, "//span[normalize-space(text())='Details']")

                        tabpanel_divs = driver.find_elements(By.XPATH, "//div[starts-with(@id, 'tabpanel-') and contains(@id, '-body')]")
                            
                        for tabpanel_div in tabpanel_divs:
                            entity_textfields = tabpanel_div.find_elements(By.XPATH, ".//input[starts-with(@id, 'entitytextfield-') and contains(@id, '-inputEl')]")
                            if entity_textfields:
                                first_input = entity_textfields[0]
                                time.sleep(2)
                                first_input.clear()
                                time.sleep(2)
                                first_input.send_keys(input_value)
                                time.sleep(2)

                                param_click(driver)
                                
                                textarea_fields = tabpanel_div.find_elements(By.XPATH, ".//textarea[starts-with(@id, 'textarea-') and contains(@id, '-inputEl')]")
                                for i, textarea_field in enumerate(textarea_fields[:11]): 
                                    textarea_field.clear()
                                    text = Sample + ['_컬럼_관리번호'][i]
                                    textarea_field.send_keys(text)
                                    time.sleep(2)

                            wait_and_click(driver, 20, By.XPATH, "//a[text()='Process']")

                            value_index_6_1 += 1 
                            print(f"Successfully processed {values_6_1[value_index_6_1 - 1]}")

                    except Exception as e:
                        print(f"Error occurred for value {values_6_1[value_index_6_1]}: {e}")                

            else:
                all_materials_processed = True
                break

            if all_materials_processed:
                break

        print("Saved Values:", saved_values)
    else:
        print("'Material 1'을 찾을 수 없습니다.")
    return driver



# process_materials_instrument_gc
# 기기분석_GC 레시피에서 동작하는 함수 (기기매핑 동작 없음)

# 1) Excel 행을 Material 별로 읽어서 최대 10개까지 각 섹션 단계 (2.1.x ~ 7.1.x) 에 매핑 준비
# 2) A열에서 'Material 1' 을 포함하는 첫 행을 찾아 시작점 지정, 각 material마다 B열 값을 input value로 가져옴
# 3) Material 영역의 Flag에 따라 섹션이 결정됨
#  > Detail 탭 열기
#  > Text 필드 지우고 Material 값 이름을 스텝 이름으로 입력
#  > Parameters 탭 > Expand ALL 버튼 클릭
#  > Parameter 이름 수정 > input value + 각 name 으로 parameter 이름 수정
#     >> (2.1.x ~ 7.1.x) 까지 이후 동작 진행
#  > Sample 탭에서 해당 Material의 이름으로 변경

def process_materials_instrument_gc(driver, df, ID, PW, Recipe, Sample, equipment_primary=None):

    # 파라미터 값 생성
    values_2_1 = [f'2.1.{i}' for i in range(1, 11)]
    values_3_1 = [f'3.1.{i}' for i in range(1, 11)]
    values_4_1 = [f'4.1.{i}' for i in range(1, 11)]
    values_5_1 = [f'5.1.{i}' for i in range(1, 11)]
    values_6_1 = [f'6.1.{i}' for i in range(1, 11)]
    values_7_1 = [f'7.1.{i}' for i in range(1, 11)]


    # 시작할 행 인덱스
    start_index = None

    # 'Material 1' 값을 찾기
    for index, value in df['A'].items():
        if pd.notna(value) and 'Material 1' in value:
            start_index = index
            break
    #변수 선언
    if start_index is not None:
        saved_values = []
        value_index_2_1 = 0 
        value_index_3_1 = 0  
        value_index_4_1 = 0
        value_index_5_1 = 0
        value_index_6_1 = 0
        value_index_7_1 = 0
        standard = 1
        preprocessing = 11
        analyze = 21
        mobile = 31

        #Material 조건에 따라 기입
        all_materials_processed = False
        for index in range(start_index, len(df)):
            if pd.notna(df.at[index, 'A']) and 'Material' in df.at[index, 'A']:
                input_value = df.at[index, 'B']
                saved_values.append(input_value)

                if df.at[index, 'D'] == 'Y' and value_index_2_1 < len(values_2_1):
                    try:
                        process_material(driver, df, index, input_value, values_2_1[value_index_2_1])

                        wait_and_click(driver, 20, By.XPATH, f"//div[normalize-space(text())='{values_7_1[value_index_7_1]}']") 
                        wait_and_click(driver, 20, By.XPATH, "//span[normalize-space(text())='Details']") 

                        tabpanel_divs = driver.find_elements(By.XPATH, "//div[starts-with(@id, 'tabpanel-') and contains(@id, '-body')]")
                        
                        for tabpanel_div in tabpanel_divs:
                            entity_textfields = tabpanel_div.find_elements(By.XPATH, ".//input[starts-with(@id, 'entitytextfield-') and contains(@id, '-inputEl')]")
                            if entity_textfields:
                                first_input = entity_textfields[0]
                                time.sleep(2)
                                # 텍스트 지우기
                                first_input.clear()
                                time.sleep(2)
                                first_input.send_keys(input_value)
                                time.sleep(2)
                                
                                wait_and_click(driver, 20, By.XPATH, "//span[normalize-space(text())='Parameters']") 
                                wait_and_click(driver, 20, By.XPATH, "//span[normalize-space(text())='Expand All']")
                                wait_and_click(driver, 20, By.XPATH, "//span[contains(@class, 'x-btn-icon-el') and normalize-space(text())=' ']")

                                textarea_fields = tabpanel_div.find_elements(By.XPATH, ".//textarea[starts-with(@id, 'textarea-') and contains(@id, '-inputEl')]")
                                for i, textarea_field in enumerate(textarea_fields[:37]):
                                    textarea_field.clear()
                                    text = input_value + [' Sample ID', ' Name',
                                        ' RT-1', ' RT-2',' RT-3',' RT-4',' RT-5',' RT-6',
                                        ' Area-1', ' Area-2',' Area-3',' Area-4',' Area-5',' Area-6',
                                        ' Tailing-1', ' Tailing-2', ' Tailing-3', ' Tailing-4', ' Tailing-5', ' Tailing-6',
                                        ' Signal to noise-1', ' Signal to noise-2', ' Signal to noise-3', ' Signal to noise-4', ' Signal to noise-5', ' Signal to noise-6',
                                        ' HalfWidthEP-1', ' HalfWidthEP-2', ' HalfWidthEP-3', ' HalfWidthEP-4', ' HalfWidthEP-5', ' HalfWidthEP-6' ][i]
                                    textarea_field.send_keys(text)
                                    time.sleep(2)
                            update_material(driver, standard, input_value)
                            standard += 1
                            value_index_2_1 += 1
                            value_index_7_1 += 1              
                                               
                    except Exception as e:
                        print(f"Error occurred for value {values_2_1[value_index_2_1]}: {e}")

                if df.at[index, 'D'] == 'N' and df.at[index, 'E'] == 'Y' and value_index_3_1 < len(values_3_1):
                    try:
                        process_material(driver, df, index, input_value, values_3_1[value_index_3_1])                    
                        update_material(driver, preprocessing, input_value)

                        preprocessing += 1
                        value_index_3_1 += 1  

                    except Exception as e:
                        print(f"Error occurred for value {values_3_1[value_index_3_1]}: {e}")

                if df.at[index, 'D'] == 'N' and df.at[index, 'F'] == 'Y' and value_index_4_1 < len(values_4_1):
                    try:
                        process_material(driver, df, index, input_value, values_4_1[value_index_4_1])
                        update_material(driver, analyze, input_value)

                        analyze += 1
                        value_index_4_1 += 1  

                    except Exception as e:
                        print(f"Error occurred for value {values_4_1[value_index_4_1]}: {e}")   

                if df.at[index, 'D'] == 'N' and df.at[index, 'I'] == 'Y' and value_index_5_1 < len(values_5_1):
                    try:

                        process_material(driver, df, index, input_value, values_5_1[value_index_5_1])                    
                        update_material(driver, mobile, input_value)

                        mobile += 1

                        value_index_5_1 += 1 

                    except Exception as e:
                        print(f"Error occurred for value {values_5_1[value_index_5_1]}: {e}")                

                if df.at[index, 'J'] == 'Y':
                    try:
                        wait_and_click(driver, 20, By.XPATH, f"//div[normalize-space(text())='{values_6_1[value_index_6_1]}']")                    
                        wait_and_click(driver, 20, By.XPATH, "//span[normalize-space(text())='Details']")

                        tabpanel_divs = driver.find_elements(By.XPATH, "//div[starts-with(@id, 'tabpanel-') and contains(@id, '-body')]")
                            
                        for tabpanel_div in tabpanel_divs:
                            entity_textfields = tabpanel_div.find_elements(By.XPATH, ".//input[starts-with(@id, 'entitytextfield-') and contains(@id, '-inputEl')]")
                            if entity_textfields:
                                first_input = entity_textfields[0]
                                time.sleep(2)
                                first_input.clear()
                                time.sleep(2)
                                first_input.send_keys(input_value)
                                time.sleep(2)

                                param_click(driver)
                                
                                textarea_fields = tabpanel_div.find_elements(By.XPATH, ".//textarea[starts-with(@id, 'textarea-') and contains(@id, '-inputEl')]")
                                for i, textarea_field in enumerate(textarea_fields[:11]): 
                                    textarea_field.clear()
                                    text = Sample + ['_컬럼_관리번호'][i]
                                    textarea_field.send_keys(text)
                                    time.sleep(2)

                            wait_and_click(driver, 20, By.XPATH, "//a[text()='Process']")

                            value_index_6_1 += 1 
                            print(f"Successfully processed {values_6_1[value_index_6_1 - 1]}")

                    except Exception as e:
                        print(f"Error occurred for value {values_6_1[value_index_6_1]}: {e}")                

            else:
                all_materials_processed = True
                break

            if all_materials_processed:
                break

        print("Saved Values:", saved_values)
    else:
        print("'Material 1'을 찾을 수 없습니다.")
    return driver

# run_recipe_rebuild
# process_excel_data() 에서 시트별 파징 결과를 yield 로 받음
# 각 시트별로 WebDriver 생성 → login → 조건에 따른 함수 실행 → remove_steps → 드라이버 종료
# 시트별로 조건에 맞게 Recipe 생성, 종료 후 다음 시트로 이동하여 동일하게 동작
# choice는 Excel의 '시험항목', choice_detail은 '시험분류' 값
# choice와 choice_detail 값에 따라 함수 실행 분기

def run_recipe_rebuild(sheet_name, file_path, ID, PW):
    """
    주어진 시트명(sheet_name)에 대해 기기분석 시험 자동화 실행
    - yield 기반 파서를 for 루프로 소비 (지정 시트면 1회만 돌고 종료)
    """
    for data in process_excel_data(file_path, sheet_name):
        df = data['df']
        Recipe = data['Recipe']
        Sample = data['Sample']
        Sample_liquid = data['Sample_liquid']
        excel_data = data['excel_data']

        #Excel에서 파징한 '시험항목' 값
        choice = data['choice']
        #Excel에서 파징한 '시험분류' 값
        choice_detail = data['choice_detail']

        # --- 드라이버는 공통으로 생성 ---
        service = Service()
        options = webdriver.ChromeOptions()
        options.add_experimental_option("detach", True)
        driver = webdriver.Chrome(service=service, options=options)
        driver.maximize_window()

        try:
            # 공통: login
            login(driver, ID, PW)

            if choice == "이화학":
                if choice_detail == "성상":
                    recipe_copy(driver, "이화학_성상", Recipe, excel_data)
                    observation(driver, df)
                    update_sample(driver, Sample, Sample_liquid)

                elif choice_detail == "기타":
                    recipe_copy(driver, "이화학_기타", Recipe, excel_data)
                    observation(driver, df)
                    update_sample(driver, Sample, Sample_liquid)
                    add_paramdsc(driver, df, 4)

                elif choice_detail == "비중":
                    recipe_copy(driver, "이화학_비중", Recipe, excel_data)
                    update_sample_density(driver, Sample, Sample_liquid)

                else:
                    raise ValueError(f"[recipe_copy] 지원하지 않는 choice_detail: {choice_detail}")

            elif choice == "기기분석":
                if choice_detail == "HPLC":
                    recipe_copy(driver, "기기분석_HPLC", Recipe, excel_data)
                    instrument(driver, df, Recipe)
                    sample_instrument(driver, Sample, Sample_liquid)
                    add_paramdsc_instrument(driver, df, ID, PW, Recipe)
                    driver = process_materials_instrument_hplc(
                        driver, df, ID, PW, Recipe, Sample,
                        equipment_primary=excel_data.get("Equipment_primary")
                    )

                elif choice_detail == "GC":
                    recipe_copy(driver, "기기분석_GC", Recipe, excel_data)
                    instrument(driver, df, Recipe)
                    sample_instrument(driver, Sample, Sample_liquid)
                    add_paramdsc_instrument(driver, df, ID, PW, Recipe)
                    driver = process_materials_instrument_gc(
                        driver, df, ID, PW, Recipe, Sample,
                        equipment_primary=excel_data.get("Equipment_primary")
                    )

                else:
                    raise ValueError(f"[recipe_copy] 지원하지 않는 choice_detail: {choice_detail}")

            else:
                print(f"[recipe_copy] '{sheet_name}' 시트: 조건 불일치 → 실행하지 않음 "
                    f"(choice={choice}, choice_detail={choice_detail})")

            # 공통: 마지막 단계에서 이름에 'name' 이 포함된 스텝 제거
            # 스텝 이름에 'name' 이 포함되었다는 말은 아무것도 설정되지 않은 스텝을 의미함.
            remove_steps(driver)

        finally:
            driver.quit()

        break  # 지정 시트명만 처리하고 종료