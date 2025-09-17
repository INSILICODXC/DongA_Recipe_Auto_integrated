import time
import pyautogui
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException

# value를 찾아 클릭
def wait_and_click(driver, timeout, by, value):
    print(f"[DEBUG] wait_and_click CALLED xpath={value}, timeout={timeout}")
    element = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, value)))
    element.click()
    time.sleep(2)

# Remove 버튼을 클릭 못하는 경우, 마지막 요소를 클릭하는 함수를 추가로 정의함
def wait_and_click_last(driver, timeout, by, value):
    print(f"[DEBUG] wait_and_click_last CALLED xpath={value}, timeout={timeout}")
    element = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, f"({value})[last()]")))
    element.click()
    time.sleep(2)

# 보이는 버튼 후보만 골라 자바스크립트로 강제 클릭하는 함수
def wait_and_click_visible(driver, timeout, by, value):
    print(f"[DEBUG] wait_and_click_visible CALLED xpath={value}, timeout={timeout}")
    WebDriverWait(driver, timeout).until(EC.presence_of_all_elements_located((by, value)))
    cands = driver.find_elements(by, value)
    visible = [e for e in cands if e.is_displayed() and e.is_enabled()]
    if not visible:
        raise TimeoutException("No visible & enabled candidates")
    el = visible[-1]  # 보이는 후보 중 마지막(대개 툴바에 실제로 보이는 것)
    try:
        el.click()
    except Exception:
        driver.execute_script("arguments[0].click();", el)
    time.sleep(2)


# value를 찾아 내용 삭제 후 key 입력
def wait_and_send_keys(driver, timeout, by, value, keys):
    element = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, value)))
    element.clear()
    element.send_keys(keys)
    time.sleep(2)

# value를 찾아 num의 횟수만큼 down 버튼 입력
def wait_and_search(driver, timeout, by, value, num):
    search_box = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, value)))
    for _ in range(num):
        search_box.send_keys(Keys.ARROW_DOWN)
        time.sleep(2)
    
    search_box.send_keys(Keys.RETURN)
    time.sleep(2)

# 우측 파라미터 정보에 파라미터 탭, Expand All 버튼, Properties Veiw 버튼 차례로 클릭
def param_click(driver):
    wait_and_click(driver, 20, By.XPATH, "//span[normalize-space(text())='Parameters']")
    wait_and_click(driver, 20, By.XPATH, "//span[normalize-space(text())='Expand All']")
    wait_and_click(driver, 20, By.XPATH, "//span[contains(@class, 'x-btn-icon-el') and normalize-space(text())=' ']")

#num의 횟수만큼 현재 화면 줌 아웃
def zoom_out(num):
    for _ in range(num):
        pyautogui.keyDown('ctrl')
        pyautogui.press('-')
        pyautogui.keyUp('ctrl')
        time.sleep(0.1)
    time.sleep(5)

# num의 횟수만큼 현재 화면 down 버튼 입력
def actions_down(driver, num):
    actions = ActionChains(driver)
  
    for _ in range(num):
        actions.send_keys(Keys.ARROW_DOWN)
        time.sleep(0.5)  

    actions.perform()  
    time.sleep(5)

# Name에 input_value 입력
def input_param(driver, timeout, input_value):

    tabpanel = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, "//div[starts-with(@id, 'tabpanel-') and contains(@id, '-body')]")))      
    input_field = WebDriverWait(tabpanel, timeout).until(EC.presence_of_element_located((By.XPATH, ".//input[starts-with(@id, 'entitytextfield-') and contains(@id, '-inputEl')]")))     
    WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.ID, input_field.get_attribute('id'))))  
    input_field.clear()
    time.sleep(1) 
    input_field.send_keys(input_value)
    time.sleep(1)

# Display Name에 input_value + val 입력
def input_text(driver, input_value, val):
    tabpanel_divs = driver.find_elements(By.XPATH, "//div[starts-with(@id, 'tabpanel-') and contains(@id, '-body')]")
    for tabpanel_div in tabpanel_divs:
        textarea_fields = tabpanel_div.find_elements(By.XPATH, ".//textarea[starts-with(@id, 'textarea-') and contains(@id, '-inputEl')]")
        for i, textarea_field in enumerate(textarea_fields[:len(val)]):
            textarea_field.clear()
            text = input_value + val[i]
            textarea_field.send_keys(text)
            time.sleep(2)
        






