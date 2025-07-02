from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import traceback
import json
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
import tkinter.messagebox as messagebox

class Logger:
    def __init__(self):
        self.message_history = []
        self.max_history = 100  # 최대 히스토리 저장 개수
        
    def __call__(self, message, level="INFO"):
        """로그 메시지 출력 및 저장"""
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        formatted_message = f"[{timestamp}] [{level}] {message}"
        print(formatted_message)
        
        # 메시지 히스토리에 저장
        self.message_history.append(formatted_message)
        if len(self.message_history) > self.max_history:
            self.message_history.pop(0)  # 가장 오래된 메시지 제거
            
    def get_last_message(self):
        """마지막 로그 메시지 반환"""
        return self.message_history[-1] if self.message_history else None
        
    def get_messages(self, count=None):
        """최근 로그 메시지들 반환"""
        if count is None:
            return self.message_history[:]
        return self.message_history[-count:]
        
    def clear_history(self):
        """로그 히스토리 초기화"""
        self.message_history = []

class MESRegister:
    def __init__(self, driver, logger=None):
        self.driver = driver
        self.log = logger if logger else Logger()
        
        # 필수비가동 코드 매핑
        self.essential_downtime = {
            "model_change": "차종/기종교체",
            "tool_change": "치공구 교체",
            "material_change": "부자재 교체",
            "master_check": "마스터 점검",
            "daily_check": "일상 점검",
            "wip_fill": "재공 채움"
        }
        
        # 일반비가동 코드 매핑
        self.normal_downtime = {
            "work_delay": "작업지연",
            "etc": "기타",
            "experiment": "실험공정",
            "quality_process": "품질문제(공정)",
            "quality_material": "품질문제(자재)",
            "equipment_failure_mass": "설비고장(양산)",
            "equipment_failure_new": "설비고장(신규)",
            "material_shortage": "자재결품",
            "condition_adjust_normal": "조건조정(일반)"
        }

    def select_downtime_type(self, downtime_type, code):
        """비가동 유형 선택"""
        try:
            self.log("\n=== 비가동 유형 선택 시작 ===")
            
            # 비가동 유형에 따른 매핑 선택
            mapping = self.essential_downtime if downtime_type == "essential" else self.normal_downtime
            if code not in mapping:
                self.log(f"잘못된 비가동 코드: {code}", "ERROR")
                return False
                
            downtime_text = mapping[code]
            self.log(f"선택할 비가동: {downtime_text}")
            
            # 비가동 드롭다운 클릭
            try:
                # 드롭다운이 나타날 때까지 명시적으로 대기
                WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_ASPxSplitter1_ddlGubunDowntime_B-1"))
                )
                
                # JavaScript로 드롭다운 클릭
                self.driver.execute_script("""
                    var dropdown = document.getElementById('ContentPlaceHolder1_ASPxSplitter1_ddlGubunDowntime_B-1');
                    if(dropdown) {
                        dropdown.click();
                        if(typeof ASPx !== 'undefined') {
                            ASPx.GotFocus('ContentPlaceHolder1_ASPxSplitter1_ddlGubunDowntime');
                        }
                    }
                """)
                time.sleep(1)
                self.log("비가동 드롭다운 클릭 완료")
                
                # 드롭다운 목록이 표시될 때까지 대기
                WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_ASPxSplitter1_ddlGubunDowntime_DDD_L"))
                )
                
                # JavaScript로 항목 선택
                selection_script = f"""
                    var found = false;
                    var items = document.querySelectorAll('#ContentPlaceHolder1_ASPxSplitter1_ddlGubunDowntime_DDD_L td');
                    for(var i = 0; i < items.length; i++) {{
                        if(items[i].textContent.trim() === '{downtime_text}') {{
                            items[i].click();
                            found = true;
                            break;
                        }}
                    }}
                    return found;
                """
                
                success = self.driver.execute_script(selection_script)
                if success:
                    self.log(f"비가동 항목 '{downtime_text}' 선택 완료")
                    time.sleep(0.5)
                    return True
                else:
                    self.log(f"비가동 항목 '{downtime_text}'를 찾을 수 없음", "ERROR")
                    return False
                
            except Exception as e:
                self.log(f"비가동 선택 중 오류: {str(e)}", "ERROR")
                return False
                
        except Exception as e:
            self.log(f"비가동 유형 선택 중 오류: {str(e)}", "ERROR")
            traceback.print_exc()
            return False

    def select_process_by_code(self, process_code):
        """공정 코드로 공정 선택"""
        try:
            self.log(f"\n=== 공정 선택 시작 (코드: {process_code}) ===")
            
            # 공정명 리스트가 비어있으면 가져오기
            if not self.process_list:
                if not self.get_process_list():
                    return False
            
            # 공정 코드가 리스트에 있는지 확인
            if process_code not in self.process_list:
                self.log(f"잘못된 공정 코드: {process_code}", "ERROR")
                return False
                
            process_name = self.process_list[process_code]
            self.log(f"선택할 공정: [{process_code}] {process_name}")
            
            try:
                # 드롭다운 버튼 클릭
                dropdown_btn = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_ASPxSplitter1_ddlOperDowntime_B-1"))
                )
                self.driver.execute_script("arguments[0].click();", dropdown_btn)
                time.sleep(1)
                
                # 해당 코드를 가진 항목 찾기
                item_xpath = f"//td[contains(@class, 'dxeListBoxItem_DevEx') and contains(text(), '[{process_code}]')]"
                item = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, item_xpath))
                )
                
                # 항목 클릭
                self.driver.execute_script("arguments[0].click();", item)
                time.sleep(0.5)
                
                self.log(f"공정 '{process_code}' 선택 완료")
                return True
                
            except Exception as e:
                self.log(f"공정 선택 중 오류: {str(e)}", "ERROR")
                return False
                
        except Exception as e:
            self.log(f"공정 선택 중 오류: {str(e)}", "ERROR")
            traceback.print_exc()
            return False
            
    def select_process(self, process_name):
        """공정 선택"""
        try:
            self.log("\n=== 공정 선택 시작 ===")
            self.log(f"선택할 공정: {process_name}")
            
            # 공정 드롭다운 클릭
            try:
                dropdown = WebDriverWait(self.driver, 5).until(
                    EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_ASPxSplitter1_ddlProcess_B-1"))
                )
                dropdown.click()
                time.sleep(0.5)
                self.log("공정 드롭다운 클릭 완료")
                
                # 공정 항목 선택
                item = WebDriverWait(self.driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, f"//td[contains(text(), '{process_name}')]"))
                )
                item.click()
                time.sleep(0.5)
                self.log(f"공정 '{process_name}' 선택 완료")
                
                return True
                
            except Exception as e:
                self.log(f"공정 선택 중 오류: {str(e)}", "ERROR")
                return False
                
        except Exception as e:
            self.log(f"공정 선택 중 오류: {str(e)}", "ERROR")
            traceback.print_exc()
            return False
            
    def set_cause_and_action(self, cause, action):
        """원인과 조치사항 입력"""
        try:
            self.log("\n=== 원인/조치사항 입력 시작 ===")
            
            # 원인 입력
            try:
                cause_input = WebDriverWait(self.driver, 5).until(
                    EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_ASPxSplitter1_txtCause_I"))
                )
                cause_input.clear()
                cause_input.send_keys(cause)
                self.log("원인 입력 완료")
                
                # 조치사항 입력
                action_input = WebDriverWait(self.driver, 5).until(
                    EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_ASPxSplitter1_txtAction_I"))
                )
                action_input.clear()
                action_input.send_keys(action)
                self.log("조치사항 입력 완료")
                
                return True
                
            except Exception as e:
                self.log(f"입력 중 오류: {str(e)}", "ERROR")
                return False
                
        except Exception as e:
            self.log(f"원인/조치사항 입력 중 오류: {str(e)}", "ERROR")
            traceback.print_exc()
            return False
            
    def register_data(self):
        """등록 버튼 클릭"""
        try:
            self.log("\n=== 데이터 등록 시작 ===")
            
            # 등록 버튼 클릭
            try:
                register_button = WebDriverWait(self.driver, 5).until(
                    EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_ASPxSplitter1_btnInsert"))
                )
                register_button.click()
                self.log("등록 버튼 클릭 완료")
                
                # 알림창 처리
                try:
                    alert = WebDriverWait(self.driver, 3).until(EC.alert_is_present())
                    alert_text = alert.text
                    self.log(f"알림창 메시지: {alert_text}")
                    alert.accept()
                    
                    if "완료" in alert_text or "성공" in alert_text:
                        self.log("데이터 등록 성공", "SUCCESS")
                        return True
                    else:
                        self.log("데이터 등록 실패", "ERROR")
                        return False
                        
                except:
                    self.log("알림창이 없음", "WARNING")
                    return True
                
            except Exception as e:
                self.log(f"등록 버튼 클릭 중 오류: {str(e)}", "ERROR")
                return False
                
        except Exception as e:
            self.log(f"데이터 등록 중 오류: {str(e)}", "ERROR")
            traceback.print_exc()
            return False
            
    def handle_separate_mode(self, process_name, cause, action, normal_type=None):
        """분리모드 처리 로직"""
        try:
            self.log("\n=== 분리모드 처리 시작 ===")
            
            # 분리모드 체크
            separate_mode = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_ASPxSplitter1_dxGrid1_DXSelBtn1_D"))
            )
            separate_mode.click()
            time.sleep(0.5)
            
            # 일반비가동 처리
            if normal_type:
                self.handle_normal_downtime_gui(process_name, normal_type, cause, action)
            
            # 등록 버튼 클릭
            register_btn = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//td[@class='btn01'][contains(text(), '등록')]"))
            )
            register_btn.click()
            
            # 언밸런스 로딩 감지
            try:
                loading = WebDriverWait(self.driver, 3).until(
                    EC.presence_of_element_located((By.ID, "loadingUnbalance"))
                )
                # 로딩이 사라질 때까지 대기
                WebDriverWait(self.driver, 30).until(
                    EC.invisibility_of_element_located((By.ID, "loadingUnbalance"))
                )
                self.log("언밸런스 로딩 완료")
            except:
                self.log("언밸런스 로딩 없음")
                
            self.log("분리모드 처리 완료")
            return True
            
        except Exception as e:
            self.log(f"분리모드 처리 중 오류: {str(e)}", "ERROR")
            traceback.print_exc()
            return False

    def handle_separate_mode_gui(self, process_name, cause, action):
        """분리모드 GUI에서의 등록 처리"""
        try:
            self.log("\n=== 분리모드 GUI 처리 시작 ===")
            
            # 필수 비가동 열기
            essential_btn = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//img[@onclick=\"ASPx.GVExpandRow('ContentPlaceHolder1_ASPxSplitter1_dxGrid3',2,event)\"]"))
            )
            essential_btn.click()
            time.sleep(0.5)
            
            # 기종 변경 열기
            model_change_btn = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//img[@onclick=\"ASPx.GVExpandRow('ContentPlaceHolder1_ASPxSplitter1_dxGrid3',3,event)\"]"))
            )
            model_change_btn.click()
            time.sleep(0.5)
            
            # 차종/기종교체(계획) 열기
            planned_change_btn = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//img[@onclick=\"ASPx.GVExpandRow('ContentPlaceHolder1_ASPxSplitter1_dxGrid3',6,event)\"]"))
            )
            planned_change_btn.click()
            time.sleep(0.5)
            
            # 비가동 클릭
            downtime_cell = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//td[@class='dxgvIndentCell dxgv']"))
            )
            downtime_cell.click()
            time.sleep(0.5)
            
            # 공정명 입력
            process_input = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_ASPxSplitter1_ddlOperDowntime_I"))
            )
            process_input.clear()
            process_input.send_keys(process_name)
            process_input.send_keys(Keys.RETURN)
            time.sleep(0.5)
            
            # 원인 입력
            cause_input = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_ASPxSplitter1_txtCause_I"))
            )
            cause_input.clear()
            cause_input.send_keys(cause)
            
            # 조치사항 입력
            action_input = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_ASPxSplitter1_txtAction_I"))
            )
            action_input.clear()
            action_input.send_keys(action)
            
            # 적용 버튼 클릭
            apply_btn = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//td[@class='btn01'][contains(text(), '적용')]"))
            )
            apply_btn.click()
            time.sleep(0.5)
            
            # 언밸런스 로딩 감지
            try:
                loading = WebDriverWait(self.driver, 3).until(
                    EC.presence_of_element_located((By.ID, "loadingUnbalance"))
                )
                # 로딩이 사라질 때까지 대기
                WebDriverWait(self.driver, 30).until(
                    EC.invisibility_of_element_located((By.ID, "loadingUnbalance"))
                )
                self.log("언밸런스 로딩 완료")
            except:
                self.log("언밸런스 로딩 없음")
            
            self.log("분리모드 GUI 처리 완료")
            return True
            
        except Exception as e:
            self.log(f"분리모드 GUI 처리 중 오류: {str(e)}", "ERROR")
            traceback.print_exc()
            return False

    def register_model_change(self, process_name, cause, action):
        """차종/기종교체 필수비가동 등록 (분리모드에서 체크된 상태)"""
        try:
            self.log("\n=== 차종/기종교체 등록 시작 ===")
            
            # 필수 비가동 열기
            try:
                essential_btn = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//img[@onclick=\"ASPx.GVExpandRow('ContentPlaceHolder1_ASPxSplitter1_dxGrid3',2,event)\"]"))
                )
                self.driver.execute_script("arguments[0].click();", essential_btn)
                time.sleep(1)
                self.log("필수 비가동 열기 완료")
            except Exception as e:
                self.log(f"필수 비가동 열기 실패: {str(e)}", "ERROR")
                return False
            
            # 기종 변경 열기
            try:
                model_change_btn = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//img[@onclick=\"ASPx.GVExpandRow('ContentPlaceHolder1_ASPxSplitter1_dxGrid3',3,event)\"]"))
                )
                self.driver.execute_script("arguments[0].click();", model_change_btn)
                time.sleep(1)
                self.log("기종 변경 열기 완료")
            except Exception as e:
                self.log(f"기종 변경 열기 실패: {str(e)}", "ERROR")
                return False
            
            # 차종/기종교체(계획) 열기
            try:
                planned_change_btn = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//img[@onclick=\"ASPx.GVExpandRow('ContentPlaceHolder1_ASPxSplitter1_dxGrid3',6,event)\"]"))
                )
                self.driver.execute_script("arguments[0].click();", planned_change_btn)
                time.sleep(1)
                self.log("차종/기종교체(계획) 열기 완료")
            except Exception as e:
                self.log(f"차종/기종교체(계획) 열기 실패: {str(e)}", "ERROR")
                return False
            
            # 비가동 클릭
            try:
                downtime_cell = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//td[@class='dxgvIndentCell dxgv']"))
                )
                self.driver.execute_script("arguments[0].click();", downtime_cell)
                time.sleep(1)
                self.log("비가동 클릭 완료")
            except Exception as e:
                self.log(f"비가동 클릭 실패: {str(e)}", "ERROR")
                return False
            
            # 공정명 입력
            try:
                # 입력 필드가 나타날 때까지 대기
                process_input = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_ASPxSplitter1_ddlOperDowntime_I"))
                )
                # 기존 값 클리어
                self.driver.execute_script("arguments[0].value = '';", process_input)
                time.sleep(0.5)
                
                # 새 값 입력
                for char in process_name:
                    process_input.send_keys(char)
                    time.sleep(0.1)
                
                time.sleep(0.5)
                process_input.send_keys(Keys.RETURN)
                time.sleep(1)
                self.log("공정명 입력 완료")
            except Exception as e:
                self.log(f"공정명 입력 실패: {str(e)}", "ERROR")
                return False
            
            # 원인 입력
            try:
                cause_input = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_ASPxSplitter1_txtCause_I"))
                )
                self.driver.execute_script("arguments[0].value = '';", cause_input)
                time.sleep(0.5)
                cause_input.send_keys(cause)
                time.sleep(0.5)
                self.log("원인 입력 완료")
            except Exception as e:
                self.log(f"원인 입력 실패: {str(e)}", "ERROR")
                return False
            
            # 조치사항 입력
            try:
                action_input = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_ASPxSplitter1_txtAction_I"))
                )
                self.driver.execute_script("arguments[0].value = '';", action_input)
                time.sleep(0.5)
                action_input.send_keys(action)
                time.sleep(0.5)
                self.log("조치사항 입력 완료")
            except Exception as e:
                self.log(f"조치사항 입력 실패: {str(e)}", "ERROR")
                return False
            
            # 적용 버튼 클릭
            try:
                apply_btn = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//td[@class='btn01'][contains(text(), '적용')]"))
                )
                self.driver.execute_script("arguments[0].click();", apply_btn)
                time.sleep(1)
                self.log("적용 버튼 클릭 완료")
            except Exception as e:
                self.log(f"적용 버튼 클릭 실패: {str(e)}", "ERROR")
                return False
            
            # 언밸런스 로딩 감지
            try:
                loading = WebDriverWait(self.driver, 3).until(
                    EC.presence_of_element_located((By.ID, "loadingUnbalance"))
                )
                # 로딩이 사라질 때까지 대기
                WebDriverWait(self.driver, 30).until(
                    EC.invisibility_of_element_located((By.ID, "loadingUnbalance"))
                )
                self.log("언밸런스 로딩 완료")
            except:
                self.log("언밸런스 로딩 없음")
            
            self.log("차종/기종교체 등록 완료", "SUCCESS")
            return True
            
        except Exception as e:
            self.log(f"차종/기종교체 등록 중 오류: {str(e)}", "ERROR")
            traceback.print_exc()
            return False

    def register_downtime(self, process_name, cause, action, essential_type=None):
        """비가동 등록 처리"""
        try:
            # 공정명 입력
            self.enter_process_name(process_name)
            
            # 필수비가동 처리
            if essential_type:
                self.handle_essential_downtime(essential_type)
            
            # 원인과 조치사항 입력
            self.enter_details(cause, action)
            
            # 등록 처리
            return self.complete_registration()
            
        except Exception as e:
            self.log(f"비가동 등록 중 오류: {str(e)}", "ERROR")
            messagebox.showerror("오류", "비가동 등록 중 오류가 발생했습니다.")
            return False
            
    def enter_process_name(self, process_name):
        """공정명 입력"""
        self.log("공정명 입력 중...")
        process_input = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_ASPxSplitter1_ddlOperDowntime_I"))
        )
        
        process_input.send_keys(Keys.CONTROL + "a")
        process_input.send_keys(Keys.DELETE)
        time.sleep(0.5)
        
        for char in process_name:
            process_input.send_keys(char)
            time.sleep(0.1)
        
        process_input.send_keys(Keys.RETURN)
        time.sleep(2)
        
    def handle_essential_downtime(self, essential_type):
        """필수비가동 유형별 처리"""
        # 필수 비가동 메뉴 열기
        self.click_menu_item("필수 비가동", 2)
        
        # 유형별 처리
        if essential_type == "model_change":
            self.handle_model_change()
        elif essential_type == "tool_change":
            self.handle_tool_change()
        elif essential_type == "material_change":
            self.handle_material_change()
        elif essential_type == "master_check":
            self.handle_master_check()
        elif essential_type == "daily_check":
            self.handle_daily_check()
        elif essential_type == "wip_fill":
            self.handle_wip_fill()
            
    def handle_model_change(self):
        """차종/기종교체 처리"""
        self.click_menu_item("기종 변경", 3)
        self.click_menu_item("차종/기종교체", 6)
        self.select_downtime()
        
    def handle_tool_change(self):
        """치공구 교체 처리"""
        self.click_menu_item("소모성치공구/부자재교환", 5)
        self.click_menu_item("치공구 교체", 5)
        self.select_downtime()
        
    def handle_material_change(self):
        """부자재 교체 처리"""
        self.click_menu_item("소모성치공구/부자재교환", 5)
        self.click_menu_item("부자재 교체", 6)
        self.select_downtime()
        
    def handle_master_check(self):
        """마스터 점검 처리"""
        self.click_menu_item("생산조건셋업", 4)
        self.click_menu_item("마스터점검", 5)
        self.select_downtime()
        
    def handle_daily_check(self):
        """일상 점검 처리"""
        self.click_menu_item("생산조건셋업", 4)
        self.click_menu_item("일상점검", 7)
        self.select_downtime()
        
    def handle_wip_fill(self):
        """재공 채움 처리"""
        self.click_menu_item("재공채움", 6)
        self.click_menu_item("재공채움2", 7)
        self.select_downtime()
        
    def click_menu_item(self, menu_name, row_number):
        """메뉴 아이템 클릭"""
        self.log(f"{menu_name} 메뉴 열기...")
        xpath = f"//td[@class='dxgv']/img[@class='dxGridView_gvCollapsedButton' and contains(@onclick,\"ASPx.GVExpandRow('ContentPlaceHolder1_ASPxSplitter1_dxGrid3',{row_number},\")]"
        menu_button = WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, xpath))
        )
        menu_button.click()
        time.sleep(1)
        
    def select_downtime(self):
        """비가동 선택"""
        self.log("비가동 선택...")
        downtime_cell = WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//td[@class='dxgvIndentCell dxgv'][@style='width:0px;border-left-width:0px;']"))
        )
        downtime_cell.click()
        time.sleep(1)
        
    def enter_details(self, cause, action):
        """원인과 조치사항 입력"""
        # 원인 입력
        self.log("원인 입력 중...")
        cause_input = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_ASPxSplitter1_txtCause_I"))
        )
        cause_input.clear()
        cause_input.send_keys(cause)
        time.sleep(1)
        
        # 조치사항 입력
        self.log("조치사항 입력 중...")
        action_input = WebDriverWait(self.driver, 10).until(
            EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_ASPxSplitter1_txtAction_I"))
        )
        action_input.clear()
        action_input.send_keys(action)
        time.sleep(1)
        
    def complete_registration(self):
        """등록 완료 처리"""
        try:
            # 적용 버튼 클릭
            self.log("적용 버튼 클릭...")
            apply_button = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//td[@class='btn01'][contains(text(),'적용')]"))
            )
            apply_button.click()
            self.wait_for_loading()
            
            # 체크박스 처리
            if not self.handle_checkbox_selection():
                return False
            
            # 등록 버튼 클릭
            self.log("등록 버튼 클릭...")
            register_button = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//td[@class='btn01'][contains(text(),'등록')]"))
            )
            register_button.click()
            
            # 알림 처리
            if not self.handle_alert():
                return False
                
            return True
            
        except Exception as e:
            self.log(f"등록 완료 처리 중 오류: {str(e)}", "ERROR")
            return False
            
    def handle_checkbox_selection(self):
        """체크박스 선택 처리"""
        try:
            self.log("\n체크박스 선택 시도...", "INFO")
            checkbox_td = WebDriverWait(self.driver, 5).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "td[onclick*=\"ASPx.GVScheduleCommand('ContentPlaceHolder1_ASPxSplitter1_dxGrid1',['Select',1]\"]"))
            )
            
            # JavaScript로 체크박스 처리
            result = self.driver.execute_script("""
                function triggerCheckbox(tdElement) {
                    var span = document.getElementById('ContentPlaceHolder1_ASPxSplitter1_dxGrid1_DXSelBtn1_D');
                    var input = document.getElementById('ContentPlaceHolder1_ASPxSplitter1_dxGrid1_DXSelBtn1');
                    
                    if (span && input) {
                        span.className = 'dxWeb_edtCheckBoxChecked dxICheckBox dxichSys dx-not-acc';
                        input.value = 'C';
                        
                        var clickEvent = new MouseEvent('click', {
                            bubbles: true,
                            cancelable: true,
                            view: window
                        });
                        tdElement.dispatchEvent(clickEvent);
                        
                        if (typeof ASPx !== 'undefined') {
                            ASPx.GVScheduleCommand('ContentPlaceHolder1_ASPxSplitter1_dxGrid1', ['Select', 1], 1);
                        }
                        
                        var row = tdElement.closest('tr');
                        if (row) {
                            row.classList.add('dxgvFocusedRow_DevEx');
                            row.style.backgroundColor = '#FFE7A2';
                        }
                        
                        return true;
                    }
                    return false;
                }
                return triggerCheckbox(arguments[0]);
            """, checkbox_td)
            
            if not result:
                return self.try_alternative_checkbox_selection()
                
            time.sleep(0.1)
            return True
            
        except Exception as e:
            self.log(f"체크박스 선택 중 오류: {str(e)}", "ERROR")
            return self.try_alternative_checkbox_selection()
            
    def try_alternative_checkbox_selection(self):
        """대체 체크박스 선택 방법"""
        try:
            self.log("대체 방법으로 선택 시도...", "INFO")
            checkbox_td = self.driver.find_element(By.CSS_SELECTOR, 
                "td[onclick*=\"ASPx.GVScheduleCommand('ContentPlaceHolder1_ASPxSplitter1_dxGrid1',['Select',1]\"]")
            self.driver.execute_script("arguments[0].click();", checkbox_td)
            time.sleep(0.1)
            
            checkbox_state = self.driver.execute_script("""
                var input = document.getElementById('ContentPlaceHolder1_ASPxSplitter1_dxGrid1_DXSelBtn1');
                if (input) {
                    var span = document.getElementById('ContentPlaceHolder1_ASPxSplitter1_dxGrid1_DXSelBtn1_D');
                    span.className = 'dxWeb_edtCheckBoxChecked dxICheckBox dxichSys dx-not-acc';
                    input.value = 'C';
                    return input.value;
                }
                return null;
            """)
            
            return checkbox_state == 'C'
            
        except Exception as e:
            self.log(f"대체 방법도 실패: {str(e)}", "ERROR")
            return False
            
    def handle_alert(self):
        """알림 처리"""
        try:
            alert = WebDriverWait(self.driver, 10).until(EC.alert_is_present())
            alert_text = alert.text
            self.log(f"알림 메시지: {alert_text}")
            alert.accept()
            
            if "등록 대상이 없습니다" in alert_text:
                self.log("등록 대상이 없습니다.", "WARNING")
                return False
                
            return True
            
        except:
            self.log("확인 메시지가 나타나지 않음")
            return True
            
    def wait_for_loading(self, timeout=60):
        """로딩 대기"""
        try:
            loading_element = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "loadingPanel"))
            )
            WebDriverWait(self.driver, timeout).until(
                EC.invisibility_of_element(loading_element)
            )
        except TimeoutException:
            self.log("로딩창 감지 시간 초과", "WARNING")
        except Exception as e:
            self.log(f"로딩 대기 중 오류: {str(e)}", "ERROR")

    def handle_normal_downtime(self, normal_type):
        """일반비가동 유형별 처리"""
        try:
            if not normal_type:
                self.log("일반비가동 유형이 선택되지 않았습니다.", "ERROR")
                return False
                
            self.log(f"\n=== 일반비가동 처리 시작 ({normal_type}) ===")
            
            # 일반 비가동 메뉴 열기
            self.click_menu_item("일반 비가동", 0)
            time.sleep(1)
            
            # 유형별 처리
            if normal_type == "work_delay":
                self.handle_work_delay()
            elif normal_type == "etc":
                self.handle_etc()
            elif normal_type == "quality_process":
                self.handle_quality_process()
            elif normal_type == "quality_material":
                self.handle_quality_material()
            elif normal_type == "equipment_failure_mass":
                self.handle_equipment_failure_mass()
            elif normal_type == "equipment_failure_new":
                self.handle_equipment_failure_new()
            elif normal_type == "material_shortage":
                self.handle_material_shortage()
            elif normal_type == "condition_adjust_normal":
                self.handle_condition_adjust_normal()
            else:
                self.log(f"알 수 없는 일반비가동 유형: {normal_type}", "ERROR")
                return False
            
            return True
            
        except Exception as e:
            self.log(f"일반비가동 처리 중 오류: {str(e)}", "ERROR")
            traceback.print_exc()
            return False

    def handle_work_delay(self):
        """작업지연 처리"""
        self.click_menu_item("생산성손실", 5)
        time.sleep(1)
        self.click_menu_item("작업지연", 9)
        time.sleep(1)
        self.select_downtime()

    def handle_etc(self):
        """기타 처리"""
        self.click_menu_item("기타", 3)
        time.sleep(1)
        self.click_menu_item("기타", 4)
        time.sleep(1)
        self.select_downtime()

    def handle_quality_process(self):
        """품질문제(공정) 처리"""
        self.click_menu_item("품질문제", 9)
        time.sleep(1)
        self.click_menu_item("공정품질문제", 10)
        time.sleep(1)
        self.select_downtime()

    def handle_quality_material(self):
        """품질문제(자재) 처리"""
        self.click_menu_item("품질문제", 9)
        time.sleep(1)
        self.click_menu_item("공정품질문제", 11)
        time.sleep(1)
        self.select_downtime()

    def handle_equipment_failure_mass(self):
        """설비고장(양산) 처리"""
        self.click_menu_item("설비고장", 6)
        time.sleep(1)
        self.click_menu_item("양산설비고장", 8)
        time.sleep(1)
        self.select_downtime()

    def handle_equipment_failure_new(self):
        """설비고장(신규) 처리"""
        self.click_menu_item("설비고장", 6)
        time.sleep(1)
        self.click_menu_item("신규설비공장", 7)
        time.sleep(1)
        self.select_downtime()

    def handle_material_shortage(self):
        """자재결품 처리"""
        self.click_menu_item("자재결품", 7)
        time.sleep(1)
        self.click_menu_item("자재결품", 9)
        time.sleep(1)
        self.select_downtime()

    def handle_condition_adjust_normal(self):
        """조건조정(일반) 처리"""
        self.click_menu_item("생산 조건", 4)
        time.sleep(1)
        self.click_menu_item("장비/조건 조정", 6)
        time.sleep(1)
        self.select_downtime()

    def register_normal_downtime(self, process_name, normal_type, cause, action):
        """일반비가동 등록 처리"""
        try:
            if not normal_type:
                self.log("일반비가동 유형이 선택되지 않았습니다.", "ERROR")
                messagebox.showerror("오류", "일반비가동 유형을 선택해주세요.")
                return False
                
            self.log("\n=== 일반비가동 등록 시작 ===")
            self.log(f"선택된 유형: {normal_type}")
            
            # 공정명 입력
            self.enter_process_name(process_name)
            
            # 일반비가동 처리
            if not self.handle_normal_downtime(normal_type):
                return False
            
            # 원인과 조치사항 입력
            self.enter_details(cause, action)
            
            # 등록 완료 처리
            return self.complete_registration()
            
        except Exception as e:
            self.log(f"일반비가동 등록 중 오류: {str(e)}", "ERROR")
            messagebox.showerror("오류", "일반비가동 등록 중 오류가 발생했습니다.")
            return False

    def handle_normal_downtime_gui(self, process_name, normal_type, cause, action):
        """일반비가동 GUI에서의 등록 처리"""
        try:
            self.log("\n=== 일반비가동 GUI 처리 시작 ===")
            
            # 일반 비가동 열기
            normal_btn = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//img[@onclick=\"ASPx.GVExpandRow('ContentPlaceHolder1_ASPxSplitter1_dxGrid3',0,event)\"]"))
            )
            normal_btn.click()
            time.sleep(0.5)
            
            # 일반비가동 유형별 처리
            if normal_type == "work_delay":
                # 작업지연 처리
                self.click_menu_item("생산성손실", 5)
                time.sleep(0.5)
                self.click_menu_item("작업지연", 9)
            elif normal_type == "etc":
                # 기타 처리
                self.click_menu_item("기타", 3)
                time.sleep(0.5)
                self.click_menu_item("기타", 4)
            elif normal_type == "quality_process":
                # 품질문제(공정) 처리
                self.click_menu_item("품질문제", 9)
                time.sleep(0.5)
                self.click_menu_item("공정품질문제", 10)
            elif normal_type == "quality_material":
                # 품질문제(자재) 처리
                self.click_menu_item("품질문제", 9)
                time.sleep(0.5)
                self.click_menu_item("공정품질문제", 11)
            elif normal_type == "equipment_failure_mass":
                # 설비고장(양산) 처리
                self.click_menu_item("설비고장", 6)
                time.sleep(0.5)
                self.click_menu_item("양산설비고장", 8)
            elif normal_type == "equipment_failure_new":
                # 설비고장(신규) 처리
                self.click_menu_item("설비고장", 6)
                time.sleep(0.5)
                self.click_menu_item("신규설비공장", 7)
            elif normal_type == "material_shortage":
                # 자재결품 처리
                self.click_menu_item("자재결품", 7)
                time.sleep(0.5)
                self.click_menu_item("자재결품", 9)
            elif normal_type == "condition_adjust_normal":
                # 조건조정(일반) 처리
                self.click_menu_item("생산 조건", 4)
                time.sleep(0.5)
                self.click_menu_item("장비/조건 조정", 6)
            
            time.sleep(0.5)
            
            # 비가동 클릭
            downtime_cell = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//td[@class='dxgvIndentCell dxgv']"))
            )
            downtime_cell.click()
            time.sleep(0.5)
            
            # 공정명 입력
            process_input = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_ASPxSplitter1_ddlOperDowntime_I"))
            )
            process_input.clear()
            process_input.send_keys(process_name)
            process_input.send_keys(Keys.RETURN)
            time.sleep(0.5)
            
            # 원인 입력
            cause_input = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_ASPxSplitter1_txtCause_I"))
            )
            cause_input.clear()
            cause_input.send_keys(cause)
            
            # 조치사항 입력
            action_input = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_ASPxSplitter1_txtAction_I"))
            )
            action_input.clear()
            action_input.send_keys(action)
            
            # 적용 버튼 클릭
            apply_btn = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//td[@class='btn01'][contains(text(), '적용')]"))
            )
            apply_btn.click()
            time.sleep(0.5)
            
            # 언밸런스 로딩 감지
            try:
                loading = WebDriverWait(self.driver, 3).until(
                    EC.presence_of_element_located((By.ID, "loadingUnbalance"))
                )
                # 로딩이 사라질 때까지 대기
                WebDriverWait(self.driver, 30).until(
                    EC.invisibility_of_element_located((By.ID, "loadingUnbalance"))
                )
                self.log("언밸런스 로딩 완료")
            except:
                self.log("언밸런스 로딩 없음")
            
            self.log("일반비가동 GUI 처리 완료")
            return True
            
        except Exception as e:
            self.log(f"일반비가동 GUI 처리 중 오류: {str(e)}", "ERROR")
            traceback.print_exc()
            return False 