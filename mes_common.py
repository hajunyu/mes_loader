from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.service import Service
import time

def initialize_driver(show_browser=True):
    """웹드라이버 초기화 및 설정"""
    options = webdriver.ChromeOptions()
    
    if not show_browser:
        options.add_argument('--headless=new')
        options.add_argument('--window-size=1920,1080')
    else:
        options.add_argument('--start-maximized')
    
    # 성능 최적화 옵션
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-gpu')
    options.add_argument('--disable-extensions')
    options.add_argument('--disable-popup-blocking')
    options.add_argument('--blink-settings=imagesEnabled=true')  # 이미지 활성화 (MES 사이트 호환성)
    
    # 네트워크 최적화
    options.add_argument('--disable-features=OptimizationHints')
    options.add_argument('--enable-features=NetworkServiceInProcess')
    options.add_argument('--remote-allow-origins=*')
    
    # 로깅 및 기타 옵션
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    options.add_experimental_option('detach', True)
    
    try:
        driver = webdriver.Chrome(options=options)
        driver.set_page_load_timeout(30)
        driver.set_script_timeout(30)
        return driver
    except Exception as e:
        print(f"Chrome 드라이버 초기화 실패: {str(e)}")
        raise

def login_and_navigate(driver, user_id, password, log_callback=print):
    """로그인 및 일반모드까지 이동"""
    max_retry = 2  # 최대 재시도 횟수
    retry_count = 0
    
    while retry_count < max_retry:
        try:
            # MES 페이지 접속
            driver.get("https://mescj.unitus.co.kr/Login.aspx")
            time.sleep(1)
            
            # 혹시 모를 알림창 처리
            try:
                alert = driver.switch_to.alert
                alert.accept()
            except:
                pass
            
            # 로그인
            user_id_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "txtUserID"))
            )
            user_id_input.clear()
            time.sleep(0.1)
            for char in user_id.upper():  # 대문자로 변환
                user_id_input.send_keys(char)
                time.sleep(0.1)
            time.sleep(0.2)
            
            password_input = driver.find_element(By.ID, "txtPassword")
            password_input.clear()
            time.sleep(0.1)
            for char in password:
                password_input.send_keys(char)
                time.sleep(0.1)
            time.sleep(0.2)
            
            login_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "btnLogin"))
            )
            login_button.click()
            time.sleep(0.5)
            
            # 로그인 후 알림창 처리
            try:
                alert = driver.switch_to.alert
                alert_text = alert.text
                log_callback(f"로그인 알림: {alert_text}")
                alert.accept()
                if "비밀번호" in alert_text or "아이디" in alert_text:
                    return False
            except:
                pass
            
            # WCT 시스템으로 이동
            try:
                log_callback(f"WCT 시스템 이동 시도 중... (시도 {retry_count + 1}/{max_retry})")
                
                # JavaScript를 사용하여 WCT 시스템 링크 클릭
                wct_link = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//a[contains(text(), 'New WCT System')]"))
                )
                driver.execute_script("arguments[0].scrollIntoView(true); arguments[0].click();", wct_link)
                time.sleep(1)
                
                # 새로 열린 창으로 전환 (최대 10초 대기)
                try:
                    WebDriverWait(driver, 4).until(lambda d: len(d.window_handles) > 1)
                    driver.switch_to.window(driver.window_handles[-1])
                except TimeoutException:
                    log_callback(f"WCT 시스템 창 열기 실패 (시도 {retry_count + 1}/{max_retry})")
                    retry_count += 1
                    if retry_count >= max_retry:
                        log_callback(f"최대 재시도 횟수 초과 ({max_retry}회)")
                        return False
                    # 실패 시 현재 창 닫기 시도
                    try:
                        driver.close()
                    except:
                        pass
                    continue
                
                # WCT 시스템 알림창 처리
                try:
                    alert = driver.switch_to.alert
                    alert.accept()
                except:
                    pass
                
                # JavaScript를 사용하여 닫기 버튼 찾고 클릭
                time.sleep(2)  # 팝업이 완전히 로드될 때까지 대기
                driver.execute_script("""
                    function findCloseButton() {
                        // td 태그에서 찾기
                        var tds = document.getElementsByTagName('td');
                        for(var i = 0; i < tds.length; i++) {
                            if(tds[i].textContent.includes('닫기')) {
                                return tds[i];
                            }
                        }
                        
                        // span 태그에서 찾기
                        var spans = document.getElementsByTagName('span');
                        for(var i = 0; i < spans.length; i++) {
                            if(spans[i].textContent.includes('닫기')) {
                                return spans[i];
                            }
                        }
                        
                        // a 태그에서 찾기
                        var links = document.getElementsByTagName('a');
                        for(var i = 0; i < links.length; i++) {
                            if(links[i].textContent.includes('닫기')) {
                                return links[i];
                            }
                        }
                        
                        // button 태그에서 찾기
                        var buttons = document.getElementsByTagName('button');
                        for(var i = 0; i < buttons.length; i++) {
                            if(buttons[i].textContent.includes('닫기')) {
                                return buttons[i];
                            }
                        }
                        
                        return null;
                    }
                    
                    var closeButton = findCloseButton();
                    if(closeButton) {
                        closeButton.scrollIntoView(true);
                        closeButton.click();
                    }
                """)
                time.sleep(1)
                
                # JavaScript를 사용하여 비가동 등록/관리 링크 클릭
                try:
                    downtime_link = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, "//a[contains(@href, '/Downtime/WCT2040.aspx')]"))
                    )
                    driver.execute_script("arguments[0].click();", downtime_link)
                    time.sleep(1.5)
                except TimeoutException:
                    log_callback("비가동 등록/관리 링크를 찾을 수 없어 직접 URL로 이동합니다.")
                    driver.get("https://wct-cj1.unitus.co.kr/Downtime/WCT2040.aspx")
                    time.sleep(1)
                    
                    # URL 이동 후 알림창 처리
                    try:
                        alert = driver.switch_to.alert
                        alert.accept()
                    except:
                        pass
            
                time.sleep(0.8)
                
                # 일반모드 클릭 (여러 ID 시도)
                normal_mode_ids = [
                    "ContentPlaceHolder1_itbtnA_Normal",
                    "ContentPlaceHolder1_btnNormal",
                    "ctl00_ContentPlaceHolder1_btnNormal"
                ]
                
                normal_mode_clicked = False
                for mode_id in normal_mode_ids:
                    try:
                        normal_mode_button = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable((By.ID, mode_id))
                        )
                        driver.execute_script("arguments[0].click();", normal_mode_button)
                        normal_mode_clicked = True
                        time.sleep(0.8)
                        break
                    except:
                        continue
                
                # 마지막으로 JavaScript로 시도
                if not normal_mode_clicked:
                    try:
                        driver.execute_script("""
                            var elements = document.getElementsByTagName('td');
                            for(var i = 0; i < elements.length; i++) {
                                if(elements[i].textContent.includes('일반')) {
                                    elements[i].click();
                                    break;
                                }
                            }
                        """)
                        time.sleep(0.8)
                    except Exception as e:
                        log_callback(f"JavaScript를 통한 일반모드 선택 실패: {str(e)}")
                        return False
                
                # 모든 과정이 성공적으로 완료되면 루프 종료
                return True
                
            except Exception as e:
                log_callback(f"WCT 시스템 이동 중 오류: {str(e)}")
                retry_count += 1
                if retry_count >= max_retry:
                    return False
                continue
                
        except Exception as e:
            log_callback(f"로그인 및 페이지 이동 중 오류 발생: {str(e)}")
            retry_count += 1
            if retry_count >= max_retry:
                return False
            continue
    
    return False 