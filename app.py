from selenium import webdriver
import os
import pyautogui 
import time
from selenium.webdriver.common.by import By

DRIVER_PATH = os.path.join(os.path.dirname(__file__), 'chromedriver')
driver = webdriver.Chrome(DRIVER_PATH)

# 크롬드라이버가 켜졌을 때 일관성을 유지하기 위해 화면에 뜰 위치와 크기를 지정해준다.
driver.set_window_position(0, 0, windowHandle='current')
driver.set_window_size(1300, 800)


# 오피넷 홈페이지 접속
driver.get('https://www.opinet.co.kr')

time.sleep(5)


# 크롬드라이버에 포커싱이 안되있을 경우 마우스 올려도 안뜨기 때문에  포커스 자동으로 주기
# pyautogui.moveTo(14, 189, 1)
driver.switch_to.window(driver.window_handles[0])



# <싼 주유소 찾기> 좌표로 마우스 이동
pyautogui.moveTo(273, 189, 1)


# <지역별> 좌표로 마우스 이동
pyautogui.moveTo(245, 265, 1)
pyautogui.click()

time.sleep(3)


# 1st 리스트 박스를 찾아서 서울 선택 
driver.find_element(by=By.ID, value="SIDO_NM0").send_keys('서울특별시')


# 'xx구' 를 선택하는 element를 second_list_raw에 저장을 한다.
gu_select_box = driver.find_element(by=By.ID, value="SIGUNGU_NM0")


time.sleep(2)


# second_list_raw에 할당된 자식 요소들 중 tag 명이 option인 것들을 받아온다. (ppt 13p 참고)
gu_tags = gu_select_box.find_elements(by=By.TAG_NAME, value="option")


# 여기에 xx구 들을 저장할 것이다. 저장하려면 위의 gu_list에 저장된 태그들로부터 value 값을 빼오면 됨
gu_values = []


# gu_tags에 저장된 태그들을 하나하나 보면서 value 값을 뺴와서 gu_values 배열에 저장.
# 결과 : gu_values = ['', '강남구', '강동구', '강북구', '강서구', '관악구', '광진구', '구로구', '금천구', '노원구', '도봉구', '동대문구', '동작구', '마포구', '서대문구', '서초구', '성동구', '성북구', '송파구', '양천구', '영등포구', '용산구', '은평구', '종로구', '중구', '중랑구']
for gu_tag in gu_tags:
    gu_values.append(gu_tag.get_attribute('value'))


# 첫번째에 빈 문자열이 들어가 있어서, 저거 빼기
# 결과 : gu_value = ['강남구', '강동구', '강북구', '강서구', '관악구', '광진구', '구로구', '금천구', '노원구', '도봉구', '동대문구', '동작구', '마포구', '서대문구', '서초구', '성동구', '성북구', '송파구', '양천구', '영등포구', '용산구', '은평구', '종로구', '중구', '중랑구']
gu_values.remove('')


for gu in gu_values: # 모든 구에 대하여
    gu_select_box = driver.find_element(by=By.ID, value="SIGUNGU_NM0") # 위에서 이미 gu_select_box를 저장했는데 왜 이걸 해야하는지 나도 잘 모르겠음. 근데 이거 안쓰면 저 element가 없다고 오류가 남. 그냥 외우자.
    gu_select_box.send_keys(gu) # 구 하나하나 선택 태그에 값을 보내고, 
    
    time.sleep(3)
    file_down = driver.find_element(by=By.ID, value='glopopd_excel').click() # 엑셀 다운로드 버튼을 클릭한다.


time.sleep(10)