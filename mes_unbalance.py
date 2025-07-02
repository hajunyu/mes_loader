import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox, filedialog
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
from datetime import datetime
import time
import os
from mes_common import initialize_driver, login_and_navigate
import traceback

class UnbalanceLoader(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("Unbalance Loader")
        self.geometry("400x750")  # 높이를 늘림
        self.resizable(False, False)
        
        # 데이터 저장용 변수
        self.selected_lines = []
        self.start_date = ""
        self.end_date = ""
        self.is_first_search = True  # 첫 조회 여부를 체크하는 플래그 추가
        
        # 상태 메시지 변수
        self.status_text = ctk.StringVar(value="대기 중...")
        
        # 메인 프레임
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(expand=True, fill="both", padx=20, pady=20)

        # 헤더 프레임
        header_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        header_frame.pack(fill="x", pady=(0,20))
        ctk.CTkLabel(header_frame, text="언발런스 로더", font=("맑은 고딕", 22, "bold")).pack(anchor="center")

        # 선택된 라인 표시 영역
        line_frame = ctk.CTkFrame(main_frame)
        line_frame.pack(fill="x", pady=(0,20))
        
        ctk.CTkLabel(line_frame, text="선택된 라인", font=("맑은 고딕", 14, "bold")).pack(pady=(5,10))
        
        # 라인 목록을 표시할 리스트박스
        self.line_listbox = tk.Listbox(line_frame, height=5)
        self.line_listbox.pack(fill="x", padx=5, pady=5)

        # 기간 표시 영역
        date_frame = ctk.CTkFrame(main_frame)
        date_frame.pack(fill="x", pady=(0,20))
        
        ctk.CTkLabel(date_frame, text="선택된 기간", font=("맑은 고딕", 14, "bold")).pack(pady=(5,10))
        
        date_info_frame = ctk.CTkFrame(date_frame, fg_color="transparent")
        date_info_frame.pack(fill="x", padx=5)
        
        self.date_label = ctk.CTkLabel(date_info_frame, text="")
        self.date_label.pack()

        # 조회 버튼
        self.search_btn = ctk.CTkButton(main_frame, 
                                      text="데이터 조회",
                                      height=40,
                                      command=self.search_data)
        self.search_btn.pack(fill="x", pady=(0,20))

        # 상태 메시지
        self.status_label = ctk.CTkLabel(main_frame,
                                       textvariable=self.status_text,
                                       text_color="white",
                                       font=("맑은 고딕", 12))
        self.status_label.pack(pady=(0,10))

        # 로그 출력 영역
        log_label = ctk.CTkLabel(main_frame, text="로그 데이터", anchor="w")
        log_label.pack(fill="x", pady=(0,5))
        
        # 로그 텍스트 박스와 스크롤바를 담을 프레임
        log_frame = ctk.CTkFrame(main_frame)
        log_frame.pack(fill="both", expand=True)
        
        # 스크롤바
        scrollbar = tk.Scrollbar(log_frame)
        scrollbar.pack(side="right", fill="y")
        
        # 로그 텍스트 박스
        self.log_text = tk.Text(log_frame,
                              font=("Consolas", 9),
                              bg="#2b2b2b",
                              fg="#ffffff",
                              height=8,
                              yscrollcommand=scrollbar.set)
        self.log_text.pack(fill="both", expand=True)
        
        # 스크롤바 연결
        scrollbar.config(command=self.log_text.yview)

    def set_initial_data(self, selected_lines, start_date, end_date):
        """초기 데이터 설정"""
        self.selected_lines = selected_lines
        self.start_date = start_date
        self.end_date = end_date
        
        # 라인 목록 업데이트
        self.line_listbox.delete(0, tk.END)
        for line in selected_lines:
            self.line_listbox.insert(tk.END, line)
        
        # 기간 정보 업데이트
        self.date_label.configure(text=f"시작: {start_date}  ~  종료: {end_date}")
        
        self.log_step("초기 데이터 설정 완료")
    
    def search_data(self):
        """데이터 조회"""
        if not self.selected_lines:
            messagebox.showwarning("경고", "선택된 라인이 없습니다.")
            return
            
        if not self.start_date or not self.end_date:
            messagebox.showwarning("경고", "기간이 설정되지 않았습니다.")
            return
            
        self.log_step("데이터 조회 시작")
        driver = None
        
        try:
            # 브라우저 실행
            self.log_step("브라우저 실행 중...")
            driver = initialize_driver(self.parent.show_browser_var.get())
            
            if not self.parent.show_browser_var.get():
                self.log_step("헤드리스 모드로 실행 중...")
            
            # 로그인 및 페이지 이동
            self.log_step("로그인 시도 중...")
            if not login_and_navigate(driver, 
                                    self.parent.id_entry.get(),
                                    self.parent.pw_entry.get(), 
                                    self.log_step):
                self.log_step("로그인 실패")
                driver.quit()
                return
                
            self.log_step("로그인 성공")
            time.sleep(1.5)

            while True:  # 메인 처리 루프
                # 데이터 조회
                self.log_step("데이터 조회 시작...")
                
                if self.is_first_search:  # 첫 조회일 때만 라인타입과 라인 선택 수행
                    # 라인타입 선택
                    self.log_step("라인타입 [01] 제어기 선택 중...")
                    line_type_dropdown = driver.find_element(By.ID, "ContentPlaceHolder1_ddeLineType_B-1")
                    line_type_dropdown.click()
                    time.sleep(0.5)
                    
                    line_type_checkbox = driver.find_element(By.ID, "ContentPlaceHolder1_ddeLineType_DDD_listLineType_0_01_D")
                    line_type_checkbox.click()
                    time.sleep(0.5)

                    driver.find_element(By.TAG_NAME, "body").click()
                    time.sleep(1)

                    # 라인 선택
                    self.log_step("선택된 라인 체크 중...")
                    line_dropdown = driver.find_element(By.ID, "ContentPlaceHolder1_ddeLine_B-1")
                    line_dropdown.click()
                    time.sleep(0.5)

                    try:
                        filter_button = driver.find_element(By.ID, "ContentPlaceHolder1_ddeLine_DDD_listLine_0_LBSFEB")
                        filter_button.click()
                        time.sleep(0.5)

                        total_lines = len(self.selected_lines)
                        for idx, line_code in enumerate(self.selected_lines, 1):
                            try:
                                code = line_code.split("]")[0].strip("[")
                                self.log_step(f"라인 선택 중... ({idx}/{total_lines})")
                                
                                filter_input = driver.find_element(By.ID, "ContentPlaceHolder1_ddeLine_DDD_listLine_0_LBFE_I")
                                filter_input.clear()
                                time.sleep(0.3)
                                filter_input.send_keys(code)
                                time.sleep(0.3)
                                
                                checkbox_id = f"ContentPlaceHolder1_ddeLine_DDD_listLine_0_{code}_D"
                                click_script = f"""
                                    var checkbox = document.getElementById('{checkbox_id}');
                                    if (checkbox) {{
                                        checkbox.click();
                                        return true;
                                    }}
                                    return false;
                                """
                                success = driver.execute_script(click_script)
                                
                                if success:
                                    self.log_step(f"라인 선택 성공: {line_code}")
                                    time.sleep(0.3)
                                else:
                                    self.log_step(f"라인을 찾을 수 없음: {line_code}")
                                    
                                filter_input.clear()
                                time.sleep(0.3)
                                
                            except Exception as e:
                                self.log_step(f"라인 선택 중 오류 ({code}): {str(e)}")
                                continue

                        try:
                            filter_button.click()
                        except:
                            pass
                        
                        time.sleep(0.3)
                        webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                        time.sleep(0.3)
                        
                    except Exception as e:
                        self.log_step(f"라인 선택 처리 중 오류: {str(e)}")
                        try:
                            webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                        except:
                            pass

                    time.sleep(0.3)

                    # 기간 선택
                    self.log_step("기간 옵션 선택 중...")
                    period_dropdown = driver.find_element(By.ID, "ContentPlaceHolder1_ddl_Div_B-1")
                    period_dropdown.click()
                    time.sleep(0.1)

                    try:
                        period_option = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, "//td[contains(text(), '기간')]"))
                        )
                        period_option.click()
                        time.sleep(0.5)
                    except Exception as e:
                        self.log_step(f"기간 옵션 선택 중 오류: {str(e)}")

                    # 날짜 입력
                    self.log_step("시작 날짜 입력 중...")
                    time.sleep(0.3)
                    
                    try:
                        start_date = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, "input.dxeEditArea_DevEx[id='ContentPlaceHolder1_dxDateEditStart_I']"))
                        )
                        
                        driver.execute_script("""
                            arguments[0].value = '';
                            arguments[0].dispatchEvent(new Event('focus'));
                        """, start_date)
                        time.sleep(0.5)
                        
                        driver.execute_script("""
                            arguments[0].value = arguments[1];
                            arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                            arguments[0].dispatchEvent(new Event('blur', { bubbles: true }));
                            if(typeof ASPx !== 'undefined') {
                                ASPx.ETextChanged('ContentPlaceHolder1_dxDateEditStart');
                                ASPx.ELostFocus('ContentPlaceHolder1_dxDateEditStart');
                            }
                        """, start_date, self.start_date)
                        
                        time.sleep(0.5)
                    except Exception as e:
                        self.log_step(f"시작 날짜 입력 중 오류: {str(e)}")

                    self.log_step("종료 날짜 입력 중...")
                    try:
                        end_date = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, "input.dxeEditArea_DevEx[id='ContentPlaceHolder1_dxDateEditEnd_I']"))
                        )
                        
                        driver.execute_script("""
                            arguments[0].value = '';
                            arguments[0].dispatchEvent(new Event('focus'));
                        """, end_date)
                        time.sleep(0.1)
                        
                        driver.execute_script("""
                            arguments[0].value = arguments[1];
                            arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                            arguments[0].dispatchEvent(new Event('blur', { bubbles: true }));
                            if(typeof ASPx !== 'undefined') {
                                ASPx.ETextChanged('ContentPlaceHolder1_dxDateEditEnd');
                                ASPx.ELostFocus('ContentPlaceHolder1_dxDateEditEnd');
                            }
                        """, end_date, self.end_date)
                        
                        time.sleep(0.1)
                    except Exception as e:
                        self.log_step(f"종료 날짜 입력 중 오류: {str(e)}")

                # 조회 버튼 클릭
                self.log_step("데이터 조회 중...")
                search_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//td[contains(text(),'조회')]"))
                )
                search_button.click()
                
                # 로딩 완료 대기
                if not self.wait_for_loading_complete(driver):
                    self.log_step("로딩 대기 중 오류 발생")
                    break

                # 데이터 로드 대기를 위한 추가 시간
                time.sleep(2)

                # No data 체크
                if self.check_no_data(driver):
                    self.log_step("더 이상 처리할 데이터가 없습니다. 작업을 종료합니다.")
                    break

                # 언발런스 등록 프로세스 실행
                if not self.register_unbalance(driver):
                    self.log_step("언발런스 등록 실패")
                    break

                # 첫 조회 플래그 업데이트
                self.is_first_search = False
                
                time.sleep(1)  # 다음 조회 전 잠시 대기

            self.log_step("작업 완료")

        except Exception as e:
            self.log_step(f"오류 발생: {str(e)}")
        finally:
            if driver:
                try:
                    driver.quit()
                except:
                    pass
    
    def log_message(self, message):
        """상태 메시지를 업데이트하고 로그에 추가"""
        self.status_text.set(message)
        self.update()
    
    def log_step(self, message, detail_status=""):
        """로그 메시지를 시간과 함께 출력"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_message = f"[{timestamp}] {message}\n"
        
        # GUI 로그 업데이트
        self.log_text.insert(tk.END, log_message)
        self.log_text.see(tk.END)  # 스크롤을 최신 로그로 이동
        
        # 상태 메시지 업데이트
        self.status_text.set(message)
        self.update()
        
        # 콘솔에도 출력
        print(log_message, end='')

    def wait_for_loading(self, driver, timeout=60):
        """로딩 이미지가 나타났다가 사라질 때까지 대기"""
        try:
            start_time = time.time()
            loading_appeared = False
            last_log_time = 0
            
            while True:
                current_time = time.time()
                if current_time - start_time > timeout:
                    raise TimeoutException(f"로딩 시간 초과 ({timeout}초)")
                
                # UpdateProgress1 div의 style 속성 확인
                loading_div = driver.find_elements(By.ID, "UpdateProgress1")
                if loading_div:
                    style = loading_div[0].get_attribute("style")
                    if "display: block" in style or style == "":  # 로딩 중
                        loading_appeared = True
                        if current_time - last_log_time >= 1:
                            self.log_step("로딩 중...")
                            last_log_time = current_time
                    elif loading_appeared and "display: none" in style:  # 로딩 완료
                        self.log_step("로딩 완료")
                        time.sleep(1)  # 추가 대기
                        return True
                
                # 3초 동안 로딩이 나타나지 않으면 계속 진행
                if not loading_appeared and current_time - start_time > 3:
                    self.log_step("로딩 없이 계속 진행")
                    return True
                    
                time.sleep(0.1)
                
        except Exception as e:
            self.log_step(f"로딩 대기 중 오류: {str(e)}")
            return False

    def wait_for_loading_complete(self, driver, timeout=60):
        """로딩이 완전히 끝날 때까지 대기"""
        try:
            start_time = time.time()
            loading_appeared = False
            last_log_time = 0
            
            while True:
                current_time = time.time()
                if current_time - start_time > timeout:
                    raise TimeoutException(f"로딩 시간 초과 ({timeout}초)")
                
                # UpdateProgress1 div의 상태 확인
                loading_div = driver.find_elements(By.ID, "UpdateProgress1")
                if loading_div:
                    style = loading_div[0].get_attribute("style")
                    aria_hidden = loading_div[0].get_attribute("aria-hidden")
                    
                    if "display: block" in style or aria_hidden == "false":
                        loading_appeared = True
                        if current_time - last_log_time >= 1:
                            self.log_step("로딩 중...")
                            last_log_time = current_time
                    elif loading_appeared and ("display: none" in style or aria_hidden == "true"):
                        self.log_step("로딩 완료")
                        time.sleep(1)  # 추가 대기
                        return True
                
                # 3초 동안 로딩이 나타나지 않으면 계속 진행
                if not loading_appeared and current_time - start_time > 3:
                    self.log_step("로딩 없이 계속 진행")
                    return True
                    
                time.sleep(0.1)
                
        except Exception as e:
            self.log_step(f"로딩 대기 중 오류: {str(e)}")
            return False

    def wait_for_registration_complete(self, driver, timeout=120):
        """등록 완료될 때까지 대기 (더 긴 타임아웃)"""
        try:
            start_time = time.time()
            loading_appeared = False
            last_log_time = 0
            
            while True:
                current_time = time.time()
                if current_time - start_time > timeout:
                    raise TimeoutException(f"등록 시간 초과 ({timeout}초)")
                
                # UpdateProgress1 div의 상태 확인
                loading_div = driver.find_elements(By.ID, "UpdateProgress1")
                if loading_div:
                    style = loading_div[0].get_attribute("style")
                    aria_hidden = loading_div[0].get_attribute("aria-hidden")
                    
                    if "display: block" in style or aria_hidden == "false":
                        loading_appeared = True
                        # 10초마다 로그 출력
                        if current_time - last_log_time >= 10:
                            self.log_step("등록 처리 중...")
                            last_log_time = current_time
                    elif loading_appeared and ("display: none" in style or aria_hidden == "true"):
                        self.log_step("등록 완료")
                        time.sleep(2)  # 등록 완료 후 충분한 대기
                        return True
                
                time.sleep(0.1)
                
        except Exception as e:
            self.log_step(f"등록 대기 중 오류: {str(e)}")
            return False

    def check_no_data(self, driver):
        """No data 메시지 확인"""
        try:
            # 충분한 대기 시간 설정
            time.sleep(2)
            
            # 로딩이 완전히 끝날 때까지 대기
            if not self.wait_for_loading_complete(driver):
                self.log_step("로딩 완료 대기 실패")
                return False

            # 데이터 그리드가 완전히 로드될 때까지 추가 대기
            time.sleep(1)
            
            # 특정 테이블에서 No data 메시지 확인
            no_data_xpath = "//table[@id='ContentPlaceHolder1_ASPxSplitter1_dxGrid2_DXMainTable']//td[contains(@class, 'dxgv')]//div[contains(text(), 'No data')]"
            
            try:
                # 명시적 대기 추가
                WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, no_data_xpath))
                )
                self.log_step("No data 상태 확인됨")
                return True
            except:
                # 데이터가 있는지 한 번 더 확인 (특정 테이블 내에서)
                rows = driver.find_elements(
                    By.XPATH,
                    "//table[@id='ContentPlaceHolder1_ASPxSplitter1_dxGrid2_DXMainTable']//tr[contains(@id, 'DXDataRow')]"
                )
                if rows:
                    self.log_step("처리할 데이터가 있습니다")
                    return False
                else:
                    # 데이터 로드 대기를 위해 추가 시간 부여
                    time.sleep(2)
                    rows = driver.find_elements(
                        By.XPATH,
                        "//table[@id='ContentPlaceHolder1_ASPxSplitter1_dxGrid2_DXMainTable']//tr[contains(@id, 'DXDataRow')]"
                    )
                    if not rows:
                        self.log_step("데이터가 없습니다")
                        return True
                    self.log_step("처리할 데이터가 있습니다")
                    return False
            
        except Exception as e:
            self.log_step(f"No data 확인 중 오류: {str(e)}")
            return False

    def open_menu_safely(self, driver, menu_name, xpath, index, max_retries=3):
        """메뉴를 안전하게 열기"""
        for attempt in range(max_retries):
            try:
                self.log_step(f"{menu_name} 열  시도 {attempt + 1}/{max_retries})")
                
                # 먼저 요소가 보이도록 스크롤
                element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, xpath))
                )
                driver.execute_script("arguments[0].scrollIntoView(true);", element)
                time.sleep(1)  # 스크롤 후 대기

                # 요소가 클릭 가능할 때까지 대기
                element = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, xpath))
                )
                time.sleep(1)  # 추가 대기

                # 클릭 시도
                element.click()
                time.sleep(2)  # 클릭 후 충분히 대기

                # 로딩 완료 대기
                self.wait_for_loading_complete(driver)
                return True

            except Exception as e:
                self.log_step(f"{menu_name} 열기 {attempt + 1}번째 시도 실패: {str(e)}")
                time.sleep(2)  # 재시도 전 대기
                
                if attempt == max_retries - 1:
                    self.log_step(f"{menu_name} 열기 최종 실패")
                    return False

        return False

    def register_unbalance(self, driver):
        """언발런스 등록 프로세스"""
        try:
            # 입력 필드 값 설정 (전체 항목 선택)
            self.log_step("전체 항목 선택 설정 중...")
            try:
                driver.execute_script("""
                    var input = document.getElementById('ContentPlaceHolder1_ASPxSplitter1_ddlGubunDowntime_I');
                    if (input) {
                        input.value = '전체항목';
                        input.dispatchEvent(new Event('change', { bubbles: true }));
                        if(typeof ASPx !== 'undefined') {
                            ASPx.ETextChanged('ContentPlaceHolder1_ASPxSplitter1_ddlGubunDowntime');
                        }
                    }
                """)
                time.sleep(1)
                self.log_step("전체 항목 선택 설정 완료")
            except Exception as e:
                self.log_step(f"전체 항목 선택 설정 중 오류: {str(e)}")

            # 첫 번째 체크박스 클릭
            self.log_step("첫 번째 항목 선택 중...")
            first_checkbox = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_ASPxSplitter1_dxGrid2_DXSelBtn0_D"))
            )
            first_checkbox.click()
            time.sleep(1)
            
            # 로딩 완료 대기
            if not self.wait_for_loading_complete(driver):
                raise Exception("첫 번째 항목 선택 후 로딩 실패")

            # 일반비가동 오픈
            general_xpath = "//td[@class='dxgv']/img[@class='dxGridView_gvCollapsedButton' and contains(@onclick,\"ASPx.GVExpandRow('ContentPlaceHolder1_ASPxSplitter1_dxGrid3',0,\")]"
            if not self.open_menu_safely(driver, "일반비가동", general_xpath, 0):
                return False

            # 생산성 손실 오픈
            productivity_xpath = "//td[@class='dxgv']/img[@class='dxGridView_gvCollapsedButton' and contains(@onclick,\"ASPx.GVExpandRow('ContentPlaceHolder1_ASPxSplitter1_dxGrid3',5,\")]"
            if not self.open_menu_safely(driver, "생산성 손실", productivity_xpath, 2):
                return False

            # 언발런스 오픈
            unbalance_xpath = "//td[@class='dxgv']/img[@class='dxGridView_gvCollapsedButton' and contains(@onclick,\"ASPx.GVExpandRow('ContentPlaceHolder1_ASPxSplitter1_dxGrid3',6,\")]"
            if not self.open_menu_safely(driver, "언발런스", unbalance_xpath, 3):
                return False

            # 언발런스 선택
            self.log_step("언발런스 선택...")
            unbalance_cell = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//td[@class='dxgvIndentCell dxgv'][@style='width:0px;border-left-width:0px;']"))
            )
            unbalance_cell.click()
            time.sleep(1)

            # 원인 입력
            self.log_step("원인 입력 중...")
            cause_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_ASPxSplitter1_txtCause_I"))
            )
            cause_input.clear()
            cause_input.send_keys("생산성 손실")
            time.sleep(0.5)

            # 조치사항 입력
            self.log_step("조치사항 입력 중...")
            action_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_ASPxSplitter1_txtAction_I"))
            )
            action_input.clear()
            action_input.send_keys("생산성 손실")
            time.sleep(0.5)

            # 조회된 항목 전체 선택
            self.log_step("전체 항목 선택 중...")
            select_all = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_ASPxSplitter1_dxGrid2_header0_chkAll_0_S_D"))
            )
            select_all.click()
            time.sleep(1)

            # 적용 버튼 클릭
            self.log_step("적용 버튼 클릭...")
            apply_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//td[@class='btn01'][contains(text(),'적용')]"))
            )
            apply_button.click()
            self.wait_for_loading_complete(driver)

            # 전체 체크박스 체크
            self.log_step("최종 전체 선택...")
            final_select_all = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_ASPxSplitter1_dxGrid1_DXSelAllBtn0_D"))
            )
            final_select_all.click()
            time.sleep(1)

            # 등록 버튼 클릭
            self.log_step("등록 버튼 클릭...")
            register_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//td[@class='btn01'][contains(text(),'등록')]"))
            )
            register_button.click()

            # 확인 메시지 처리
            try:
                alert = WebDriverWait(driver, 10).until(EC.alert_is_present())
                alert.accept()
            except:
                self.log_step("확인 메시지가 나타나지 않음")

            # 등록 완료 대기 (더 긴 타임아웃으로 설정)
            if not self.wait_for_registration_complete(driver):
                raise Exception("등록 완료 대기 실패")

            self.log_step("등록 완료")
            return True

        except Exception as e:
            self.log_step(f"언발런스 등록 중 오류: {str(e)}")
            return False 