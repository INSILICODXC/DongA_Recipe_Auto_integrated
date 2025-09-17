import pandas as pd
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException
import time
from utility.utils import *
import utility.config as config
import pyautogui

# 엑셀 시트를 모두 순회하며 파징 값에 이상이 있는지를 체크하는 전용 함수를 정의함 (ssgwak)
# precheck 할 때, 각 시트에서 extract_from_excel을 호출하여 필수값 누락 여부를 검사함.
# 현재 required_map에 정의된 필수값이 모두 채워져 있는지 검사
# required_map에 정의된 값중 하나라도 없으면 시트이름과 누락값을 메시지로 출력함.
def precheck_all_sheets(excel_path: str):
    """
    전체 시트를 돌며 extract_from_excel을 실행,
    필수값(required_map) 누락 여부와 Material 고체/액체 여부를 사전에 검사한다.
    """
    xls = pd.ExcelFile(excel_path)
    errors = []

    for sheet_name in xls.sheet_names:
        if sheet_name.lower() == config.IGNORE_SHEET_NAME.lower():
            print(f"[스킵] '{sheet_name}' 시트는 참조용이므로 pre-check에서 제외")
            continue

        try:
            data = extract_from_excel(excel_path, sheet_name)

            # ----------------------------
            # 1) 필수 키 확인
            # ----------------------------
            required_keys = [
                "choice",
                "Recipe Name",
                "Method Category",
                "Recipe Location",
                "Sample",
                "Sample_liquid",
            ]
            missing = [
                k for k in required_keys
                if not data.get(k) or str(data[k]).strip().lower() in ["", "nan"]
            ]

            # 기기분석일 경우 장비 정보도 필수
            if str(data["choice"]).strip() == "기기분석":
                if not data.get("Equipment_list"):
                    missing.append("Equipment_list")
                if not data.get("Equipment_primary"):
                    missing.append("Equipment_primary")

            if missing:
                errors.append((sheet_name, f"필수값 누락: {', '.join(missing)}"))
            
            
            # ----------------------------
            # 2) Material의 고체/액체 컬럼 검증
            # ----------------------------
            try:
                df = data["df"]  # extract_from_excel에서 반환된 원본 DataFrame

                # 공백("")은 None으로 변환하여 진짜 빈값으로 처리
                norm_all = df.applymap(lambda x: str(x).strip() if str(x).strip() != "" else None)

                colA = norm_all.iloc[:, 0]
                material_row = colA[colA.eq("MATERIAL")].index.min()

                if pd.isna(material_row):
                    errors.append((sheet_name, "MATERIAL 행을 찾지 못했습니다."))
                else:
                    # MATERIAL 행에서 고체/액체 컬럼 찾기
                    headers = norm_all.iloc[material_row]
                    try:
                        solid_col = headers[headers == "고체"].index[0]
                        liquid_col = headers[headers == "액체"].index[0]
                    except IndexError:
                        # 컬럼 자체를 못 찾은 경우
                        errors.append((sheet_name, "고체/액체 컬럼 정의를 MATERIAL 행에서 찾지 못했습니다."))
                        continue

                    # 데이터 블록은 MATERIAL 행 바로 아래부터 시작
                    material_block = norm_all.iloc[material_row+1:, :]

                    # "Material + 숫자" 패턴인 행만 검사
                    material_rows = material_block[
                        material_block.iloc[:, 0].str.match(r"^Material\s*\d+$", na=False)
                    ]

                    for idx, row in material_rows.iterrows():
                        row_name = str(row.iloc[0]).strip()
                        try:
                            solid_val = (row.iloc[solid_col] or "").upper()
                            liquid_val = (row.iloc[liquid_col] or "").upper()
                        except Exception:
                            errors.append((sheet_name, f"{row_name} 행: 고체/액체 컬럼 정의를 찾지 못했습니다."))
                            continue

                        # 핵심 조건: 고체/액체 둘 다 Y가 아니면 무조건 에러
                        if solid_val != "Y" and liquid_val != "Y":
                            errors.append((sheet_name, f"{row_name} 행: 고체/액체 중 하나는 Y 여야 합니다."))

            except Exception as e:
                errors.append((sheet_name, f"Material 검증 중 오류: {e}"))

        except Exception as e:
            errors.append((sheet_name, f"파싱 오류: {e}"))

    # ----------------------------
    # 결과 출력
    # ----------------------------
    if errors:
        print("=== Pre-check 오류 발생 ===")
        for sheet, err in errors:
            print(f"[{sheet}] {err}")
        raise RuntimeError("Pre-check 실패 → 실행 중단")
    else:
        print("=== Pre-check 완료: 모든 시트 OK ===")



# 엑셀에서 Key-Value 영역, Sample 테이블 영역, Equipment 테이블 영역에서 값 추출 -> dict 형태로 return (다른 곳에서 Key 값을 가지고 사용 가능)
# 변경된 Execl 양식에서 Parsing할 값을 먼저 읽고, 이후에 각 모듈에서 해당 값을 사용
def extract_from_excel(excel_path: str, sheet=0) -> dict:

    # drop-down 시트는 무시
    if isinstance(sheet, str) and sheet.lower() == config.IGNORE_SHEET_NAME.lower():
        print(f"[스킵] '{sheet}' 시트는 참조용이므로 실행하지 않음")
        return {}

    # drop-down 제외한 전체 시트 로드
    df = pd.read_excel(excel_path, sheet_name=sheet, header=None, keep_default_na=False)

    # -----------------------------
    # 1) KV 영역 (Sample Name 직전까지)
    # -----------------------------
    # 'Sample Name' 이 나오는 첫 행 찾기
    # 행 단위로 함수 적용
    # 행의 모든 값을 문자열로 변환 후 공백 제거, "Sample Name"과 비교, 하나라도 True가 있으면 True 반환
    try:
        marker_row = df[df.apply(lambda r: r.astype(str).str.strip().eq("Sample Name").any(), axis=1)].index[0]
    except IndexError:
        raise ValueError("'Sample Name' 라벨을 KV 경계로 찾지 못했습니다. (KV 영역 경계 없음)")
    key_value_area = df.iloc[:marker_row]

    # 비교용 정규화본 (applymap deprecated → map 사용)
    # 비교 안정성을 위해 모든 값을 문자열 + 공백 제거 (Key-value 비교 전용 DataFrame)
    # 이 데이터 영역 (norm)은 Key-Value로 되어 있는 데이터 영역만 해당됨.
    norm = key_value_area.astype(str).map(lambda x: x.strip())

    # key-value 영역에서 값 추출 (Key의 오른쪽을 value로 추출)
    def value_right_of(key: str) -> str:
        """KV 테이블에서 key 오른쪽(1칸) 값"""
        pos = norm[norm.eq(key.strip())].stack().index[0]   # (row, col) 위치
        return str(key_value_area.iat[pos[0], pos[1] + 1]).strip()
    
    # key-value 항목 중 시험항목에는 detail이라는 컬럼이 하나 더 있음, detail에서 값을 뽑기 위함
    def value_offset_of(key: str, offset: int):
        """KV 테이블에서 key 기준 오른쪽 offset칸 값 (Details 등)"""
        r, c = norm[norm.eq(key.strip())].stack().index[0]
        target_c = c + offset
        if target_c >= key_value_area.shape[1]:
            return None
        val = str(key_value_area.iat[r, target_c]).strip()
        return val if val and val.lower() != "nan" else None

    # key-value에서 뽑은 각 컬럼을 key로 하여, value로 뽑고 이를 딕셔너리로 변환
    # key-value에서 '시험항목', 'Recipe Name', 'Method Category', 'Recipe Location' 의 value 추출
    choice           = value_right_of("시험항목")
    # 추후 이화학, 기기분석 확장에 details 값을 사용
    choice_detail   = value_right_of("시험분류")
    recipe_name      = value_right_of("Recipe Name")
    method_category  = value_right_of("Method Category")
    recipe_location  = value_right_of("Recipe Location")

    # -----------------------------
    # 2) 테이블 영역 (시트 전체 정규화본)
    # -----------------------------
    norm_all = df.astype(str).map(lambda x: x.strip())

    # (공통) A열에서 경계 행
    colA = norm_all.iloc[:, 0]
    samplename_row = colA[colA.eq("Sample Name")].index.min()
    #엑셀양식 변경 전
    #equipment_row = colA[colA.eq("Equipment")].index.min()

    #엑셀양식 변경 후
    equipment_row = colA[colA.eq("Equipment Name")].index.min()
    material_row = colA[colA.eq("MATERIAL")].index.min()

    # 2-1) Sample (Sample Name 바로 아래 1칸)
    # "Sample Name" 라벨이 들어있는 컬럼 번호(int)
    try:
        sample_col = norm_all.columns[(norm_all == "Sample Name").any()].tolist()[0]
        # "Sample Name"이라는 텍스트가 적혀 있는 행 번호
        sample_row = norm_all.index[norm_all[sample_col] == "Sample Name"].tolist()[0]
    except IndexError:
        raise ValueError("'Sample Name' 라벨을 표에서 찾지 못했습니다. (Sample 섹션 없음)")
    # 'Sample Name' 이 있는 행을 찾고, 거기에 +1을해줘야 그 밑에서 값을 가져올수 있음 
    sample_row_idx = sample_row + 1

    # '액체' 컬럼 찾기: 우선 헤더 행에서, 없으면 전체에서
    # 'Sample Name' 라벨이 있는 sample_row 행 번호 기반으로 '액체' 컬럼 찾기
    liquid_cols = norm_all.columns[norm_all.loc[sample_row].eq("액체")].tolist()
    if not liquid_cols:
        liquid_cols = norm_all.columns[(norm_all == "액체").any()].tolist()
        if not liquid_cols:
            raise ValueError("액체 컬럼을 찾을 수 없습니다.")
    # 엑셀 전체에서 '액체' 컬럼 찾기
    liquid_col = liquid_cols[0]

    # Sample Name, 액체 컬럼 값 추출
    # 같은 라인에 위치하지만, 실제로 Sample Name이 있는 위치와 액체가 있는 위치는 다름
    # df.iat[row,col] 방식으로 해서 sample name 값이 있는 위치를 지정
    # "Sample Name" 아래 줄, Sample Name 컬럼, strip() 으로 양쪽 공백 제거
    try:
        sample        = str(df.iat[sample_row_idx, sample_col]).strip()
        # df.iat[row,col] 방식으로 해서 sample name 값이 있는 위치->'액체' 컬럼이 있는 위치의 값을 지정
        # Sample 행에서 sample_row_idx (sample 값이 있는 줄의 위치) liquid_col ('액체' 컬럼이 있는 컬럼 번호) 기반으로, "액체" 컬럼에 해당하는 값을 추출
        sample_liquid = str(df.iat[sample_row_idx, liquid_col]).strip()
    except Exception:
        raise ValueError("Sample/Sample_liquid 값을 읽는 중 인덱스 오류가 발생했습니다.")

 # 2-2) Equipment (Equipment ~ Material 직전, 분석장비=='Y')
 # Equipment 섹션은 Equipment 에서 'Material' 이라고 써 있는 바로 전까지의 구간
    # 기존에는 무조건 에러를 던졌으나, 아래에서 choice가 '기기분석'일 때만 필수로 검증.
    eq_valid = (not pd.isna(equipment_row)) and (not pd.isna(material_row)) and (equipment_row < material_row)
    equipment_list = []
    equipment_primary = None
    if eq_valid:
        #equip_col = norm_all.columns[norm_all.loc[equipment_row].eq("Equipment")].tolist()[0]
        equip_col = norm_all.columns[norm_all.loc[equipment_row].eq("Equipment Name")].tolist()[0]
        flag_col  = norm_all.columns[norm_all.loc[equipment_row].eq("분석장비")].tolist()[0]

        block = norm_all.iloc[equipment_row + 1:material_row, :]
        # '분석장비' 열이 Y인 행만 필터
        mask_y = block.iloc[:, flag_col].str.upper().eq("Y")

        # Equipment 컬럼 위치에서, B열(index=1)의 값을 가져오도록
        #equipment_list = block.loc[mask_y, 1].tolist()

        # Equipment 컬럼 위치에서, A열(index=0)의 값을 가져오도록 (엑셀양식 변경)
        equipment_list = block.loc[mask_y, 0].tolist()
        equipment_primary = equipment_list[0] if equipment_list else None

    # -----------------------------
    # 3) 필수값 검증 (요청 사양)
    # -----------------------------
    def _empty(v):
        return (v is None) or (str(v).strip() == "") or (str(v).strip().lower() == "nan")

    # choice, choice_details, recipe_name, method_category, recipe_location, sample, sample_liquid 는 choice가 무엇이든 간에 필수
    required_map = {
        "choice": choice,
        "choice_detail": choice_detail,
        "Recipe Name": recipe_name,
        "Method Category": method_category,
        "Recipe Location": recipe_location,
        "Sample": sample,
        "Sample_liquid": sample_liquid,
    }
    missing = [k for k, v in required_map.items() if _empty(v)]
    if missing:
        raise ValueError(f"필수값 누락: {', '.join(missing)}")

    # choice가 '기기분석' 일 때만 equipment_list / equipment_primary 미존재 시 에러
    if str(choice).strip() == "기기분석":
        if not eq_valid:
            raise ValueError("기기분석: 'Equipment ~ MATERIAL' 구간을 찾지 못했습니다.")
        if not equipment_list:
            raise ValueError("기기분석: '분석장비=Y' 장비 목록이 비었습니다.")
        if _empty(equipment_primary):
            raise ValueError("기기분석: equipment_primary가 정의되지 않았습니다.")

    return {
        # Key-value 영역에서 추출하는 Data -> 나중에 딕셔너리로 return
        
        # Excel에서 선택하는 '시험항목' 의 Value 값
        "choice": choice,
        # Excel에서 선택하는 '시험분류' 의 Value 값
        "choice_detail": choice_detail,
        "Recipe Name": recipe_name,
        "Method Category": method_category,
        "Recipe Location": recipe_location,

        # Sample table 영역에서 추출하는 Data
        "Sample": sample,
        "Sample_liquid": sample_liquid,
        # Equipment table 영역에서 추출하는 Data
        "Equipment_list": equipment_list,
        "Equipment_primary": equipment_primary,

        # 필요 시 디버깅용
        "_marker_rows": {
            "SampleName": int(samplename_row) if not pd.isna(samplename_row) else None, 
            "Equipment": int(equipment_row) if not pd.isna(equipment_row) else None, 
            "Material": int(material_row) if not pd.isna(material_row) else None
        },
        "df": df,
    }



# ID,PW 입력
def get_id(id_prompt="ID:", pw_prompt="비밀번호:"):

    try:
        id_input = input(id_prompt + " ")
        if not id_input.strip():
            return None
        
        pw_input = input(pw_prompt + " ")
        if not pw_input.strip():
            return None
        
        return (id_input, pw_input)
    except KeyboardInterrupt:
        print("\n입력이 취소되었습니다.")
        return None

# username, password로 로그인    
def login(driver, username, password):
    driver.get(config.URL_HUB)
    time.sleep(5)

    if config.IS_DA_DEV == "Y":
        # 동아 개발환경
        wait_and_send_keys(driver, 20, By.XPATH, "//input[@placeholder='이메일 또는 사용자 이름']", username)
        wait_and_send_keys(driver, 20, By.XPATH, "//input[@placeholder='암호']", password)
        wait_and_click(driver, 20, By.XPATH, "//input[@value='로그인']")
    else:
        # 인실리코 내부 개발환경
        wait_and_send_keys(driver, 20, By.XPATH, "//input[@placeholder='Username']", username)
        wait_and_send_keys(driver, 20, By.XPATH, "//input[@placeholder='Password']", password)
        wait_and_click(driver, 20, By.XPATH, "//input[@value='SIGN IN']")


# recipe_name을 찾고 복사 한 후 product_name 으로 변경
# recipe_name은 원본 레시피 이름, product_name은 복사 후 새 레시피 이름
def recipe_copy(driver, recipe_name, product_name, excel_data):

    driver.get(config.URL_COMPOSE)
    time.sleep(5)

    wait_and_click(driver, 20, By.XPATH, "//div[@title='Recipe Filters']")                                                                 
    wait_and_send_keys(driver, 20, By.NAME, "formRecipeName", recipe_name)
    wait_and_click(driver, 20, By.XPATH, f"//i[@title='Clone {recipe_name}']")

    # 'Recipe Type' 콤보박스 클릭
    wait_and_click(driver, 20, By.NAME, "recipeType")

    # Recipe Clone 할 때 Recipe Type은 'Master' / Recipe Workflow는 'Activity Driven' 으로 고정
    wait_and_click(driver, 20, By.XPATH, "//li[normalize-space(text())='Master']")
    
    # 'recipe workflow' 콤보박스 선택
    wait_and_click(driver, 20, By.NAME, "recipeWorkflow")  
    wait_and_click(driver, 20, By.XPATH, "//li[normalize-space(text())='Activity Driven']")

    # 'method category' 콤보박스 선택
    wait_and_click(driver, 20, By.NAME, "methodCategory")
    wait_and_click(driver, 20, By.XPATH, f"//li[normalize-space(text())='{excel_data['Method Category']}']")

    wait_and_click(driver, 20, By.XPATH, "//span[normalize-space(text())='Recipe Library']")
    actions_down(driver, 9)
    wait_and_click(driver, 20, By.XPATH, f"//span[normalize-space(text())='{excel_data['Recipe Location']}']")
    wait_and_click(driver, 20, By.XPATH, "//span[normalize-space(text())='OK']")
    #time.sleep(20)
    
    # 동아 개발환경에서 OK 클릭 후 다음 동작시까지 오래 걸리는 경우 발생, 30초로 늘림
    time.sleep(30)


    wait_and_send_keys(driver, 20, By.XPATH, "//input[@type='text' and @role='textbox' and contains(@class, 'x-form-field') and @data-ref='inputEl' and not(@readonly)]", product_name)
    wait_and_click(driver, 20, By.XPATH, "//span[normalize-space(text())='Save']")
    wait_and_click(driver, 20, By.XPATH, "//a[text()='Process']")
    wait_and_click(driver, 20, By.XPATH, "//span[contains(@id, 'button-') and contains(text(), 'Expand')]")
    input_param(driver, 20, product_name)
    zoom_out(3)

# 엑셀에서 Material 1 찾기
def find_start_index(df):
    for index, value in df['A'].items():
        if pd.notna(value) and 'Material 1' in value:
            return index
    return None

# num 횟수 만큼 파라미터 삭제 버튼 반복 클릭
def process_trashcan(driver, num):
    img_elements = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "img.x-tool-img.x-tool-trashcan"))) # 휴지동 아이콘 클릭
    if len(img_elements) >= 9:
        for i in num:
            element = img_elements[i]
            element.click()
            time.sleep(2)
            wait_and_click(driver, 20, By.XPATH, "//span[normalize-space(text())='예']") #휴지통 아이콘 클릭 후 예 버튼 클릭
    else:
        print("There are less than 9 elements matching the criteria.")

# material_number의 material을 찾아 input_value로 변경
def update_material(driver, material_number, input_value):
    #XPATH에 Material 1, Material 2 등으로 되어있는 패턴을 찾아 f-string으로 처리
    wait_and_click(driver, 20, By.XPATH, "//a[text()='Materials']")
    #문자열 앞에 f를 붙여서 변수값을 넣을 수 있게 함 (변수 값을 사용하려면 모든 경우에서 아래와 같이 f-string으로 처리)
    #f를 쓰면 변수 값이 치환되며, f를 쓰지 않으면 단순 문자열로 인식
    material_xpath = f"//div[contains(@class, 'x-grid-cell-inner') and text()='Material {material_number}']"
    # driver가 조작하는 브라우저에서 최대 10초 동안 특정 조건이 만족되기를 기다리고, until 은 조건이 만족되면 해당 요소를 반환홤
    # 10초 안에 조건이 만족되지 않으면 TimeoutException 반환을 하고, XPATH에 material_xpath가 보이고 클릭 가능한 상태일 때 동작
    material_element = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, material_xpath)))
    ## ActionChains - 드라이버에 연결된 액션 체인 객체 생성, 마우스/키보드 같은 고급 사용자 입력 이벤트를 다룰 수 있게 해주는 클래스
    actions = ActionChains(driver)
    # material_element라는 WebElement에 대해 더블클릭 동작을 정의, .perform() 으로 실제 동작을 실행함
    actions.double_click(material_element).perform()
    # 2초 동안 대기 (time.sleep 없이 바로 다음 동작을 진행할 시 Web에 요소가 미처 업로드 되지 못하면 문제 발생, 타임아웃 설정으로 다음 동작 전 대기시간 설정)
    time.sleep(2)
    # BACKSPACE 키를 11번 눌러 기존 값을 지움, Keys.BACKSPACE는 Selenium에서 키보드의 BACKSPACE 키를 나타냄
    # .perform() 으로 실제 동작을 실행함
    actions.send_keys(Keys.BACKSPACE * 11).perform()
    actions.send_keys(input_value).perform()
    time.sleep(2)
    wait_and_click(driver, 20, By.XPATH, "//a[text()='Process']")

# 파라미터 수정 후 엑셀 조건에 따라 필요없는 파라미터 삭제
def process_material(driver, df, index, input_value, value):
    wait_and_click(driver, 20, By.XPATH, f"//div[normalize-space(text())='{value}']")
    wait_and_click(driver, 20, By.XPATH, "//span[normalize-space(text())='Details']")
    input_param(driver, 20, input_value)
    wait_and_click(driver, 20, By.XPATH, "//span[normalize-space(text())='Parameters']")
    wait_and_click(driver, 20, By.XPATH, "//span[normalize-space(text())='Expand All']")
    wait_and_click(driver, 20, By.XPATH, "//span[contains(@class, 'x-btn-icon-el') and normalize-space(text())=' ']")
    input_text(driver, input_value, ['_Barcode', '_칭량', '_Quantity', '_사용량', '_Quantity2', '_사용기한', '_Purity', '_Lot No.', '_계산식 사용 칭량값']) # 반복해서 입력할 리스트

    if df.at[index, 'G'] == 'Y':
        process_trashcan(driver, [3, 4])

    if df.at[index, 'H'] == 'Y':
        process_trashcan(driver, [1, 2, 8])


# 기존의 remove_steps 함수에서, Remove 버튼 클릭이 안되는 이슈로 인해 다른 패턴의 버튼 클릭을 시도하는 로직 추가
def remove_steps(driver):

    while True:
        try:
            name_elements = WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.XPATH, "//span[contains(text(), 'name')]"))
            )
            if len(name_elements) == 0:
                break

            for index, element in enumerate(name_elements):
                try:
                    print(f"Clicking element {index + 1} with id: {element.get_attribute('id')}")
                    element.click()
                    time.sleep(2)

                    try:
                        # 1차: 기본 클릭 (기존 그대로 Remove 버튼을 찾는 패턴)
                        wait_and_click(driver, 10, By.XPATH, "//span[contains(@id, 'button-') and contains(text(), 'Remove')]")
                    except TimeoutException:
                        print("[Remove] 기본 클릭 실패")
                        try:
                           # 2차 : 기존과 다르게, web에서 보이고 클릭할 수 있는 것만 대상으로 필터링 해서 클릭하는 패턴, 현재는 기기분석에서 해당사항이 발견되어 사용중)
                            wait_and_click_visible(
                                driver, 10, By.XPATH,
                                "//div[contains(@class,'x-toolbar') and not(contains(@class,'x-hidden')) and not(contains(@class,'x-hide-offsets'))]"
                                "//a[contains(@class,'x-btn') and not(contains(@class,'x-btn-disabled')) "
                                " and .//span[@data-ref='btnInnerEl' and normalize-space(.)='Remove']]"
                            )

                        except TimeoutException:
                            print("[Remove] last 클릭 실패")
                            try:
                                #Remove 버튼 클릭 시 모든 경우에 대해서 실패했을 경우, find_cause 함수를 호출하여 원인 파악 (버튼 속성 등이 출력되어서 해당 결과를 기반으로 디버깅)
                                find_cause(driver)
                            except Exception as e:
                                print(f"[find_cause] 오류: {e}")

                    # 확인 팝업 '예'
                    try:
                        wait_and_click(driver, 20, By.XPATH, "//span[normalize-space(text())='예']")
                    except TimeoutException:
                        print("확인 팝업 '예' 버튼을 찾지 못했습니다.")

                except StaleElementReferenceException:
                    print("Element is stale, retrying...")
                    continue

        except TimeoutException:
            print("No elements found containing 'name' within timeout period.")
            break

    # 마지막 저장
    wait_and_click(driver, 20, By.XPATH, "//span[normalize-space(text())='Save']")


# 버튼 클릭 안되는 원인 파악을 위한 함수 추가
def find_cause(driver):
    els = driver.find_elements(By.XPATH, "//span[@data-ref='btnInnerEl' and normalize-space(.)='Remove']")
    print("found spans:", len(els))
    if els:
        # 버튼이 is displayed 인지 출력
        print("span[0] displayed:", els[0].is_displayed())
        # 버튼의 outerHTML 출력 (속성 정보)
        print("span[0] outerHTML:", els[0].get_attribute("outerHTML"))

    # 2) 부모 <a> 기준
    # 버튼이 위치한 부모 <a> 태그를 기준으로 다시 찾기
    as_ = driver.find_elements(By.XPATH, "//a[contains(@class,'x-btn')][.//span[@data-ref='btnInnerEl' and normalize-space(.)='Remove']]")
    print("found anchors:", len(as_))
    if as_:
        print("a[0] class:", as_[0].get_attribute("class"))
        print("a[0] displayed:", as_[0].is_displayed())
        print("a[0] outerHTML:", as_[0].get_attribute("outerHTML"))

    # 3) 비활성 여부(class에 x-btn-disabled 포함?)
    if as_:
        disabled = "x-btn-disabled" in (as_[0].get_attribute("class") or "")
        print("a[0] disabled? ->", disabled)
