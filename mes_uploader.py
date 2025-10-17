import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import pyautogui
from mes_common import initialize_driver, login_and_navigate
from datetime import datetime
from selenium.common.exceptions import TimeoutException
import pandas as pd
import openpyxl
from pathlib import Path

class FileUploader(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("비가동 업로드")
        self.geometry("300x750")  # 전체 높이를 700px로 조정
        self.resizable(False, False)
        
        # 상태 메시지 변수
        self.status_text = ctk.StringVar(value="대기 중...")
        
        # 업로드 중지 플래그
        self.stop_upload = False

        # 메인 프레임
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(expand=True, fill="both", padx=20, pady=20)

        # 파일 선택 영역
        select_frame = ctk.CTkFrame(main_frame)
        select_frame.pack(fill="x", pady=(0,20))
        
        select_btn = ctk.CTkButton(select_frame, 
                                 text="파일 선택",
                                 command=self.select_files,
                                 height=32,
                                 width=80,  # 너비 조정
                                 fg_color="#27ae60",
                                 hover_color="#2ecc71")
        select_btn.pack(side="left", padx=(0,5))
        
        # 폴더 선택 버튼 추가
        folder_btn = ctk.CTkButton(select_frame,
                                text="폴더 선택",
                                command=self.select_directory,
                                height=32,
                                width=80,  # 너비 조정
                                fg_color="#27ae60",
                                hover_color="#2ecc71")
        folder_btn.pack(side="left", padx=(0,5))
        
        clear_btn = ctk.CTkButton(select_frame,
                                text="목록 지우기",
                                command=self.clear_files,
                                height=32,
                                width=80,  # 너비 조정
                                fg_color="#27ae60",
                                hover_color="#2ecc71")
        clear_btn.pack(side="left", padx=(5,0))

        # 파일 목록
        list_frame = ctk.CTkFrame(main_frame)
        list_frame.pack(fill="both", pady=(0,20))
        
        self.file_listbox = tk.Listbox(list_frame,
                                     font=("맑은 고딕", 10),
                                     selectmode=tk.SINGLE,
                                     height=8)
        self.file_listbox.pack(fill="both", padx=5, pady=5)

        # 상태 메시지
        self.status_label = ctk.CTkLabel(main_frame,
                                       textvariable=self.status_text,
                                       text_color="white",
                                       font=("맑은 고딕", 12))
        self.status_label.pack(pady=(0,10))

        # 업로드 버튼과 중지 버튼을 담을 프레임
        button_frame = ctk.CTkFrame(main_frame)
        button_frame.pack(fill="x", pady=(0,20))

        # 업로드 버튼
        self.upload_btn = ctk.CTkButton(button_frame,
                                      text="업로드 시작",
                                      command=self.start_upload,
                                      height=32,
                                      width=120,  # 너비 감소
                                      #font=("맑은 고딕", 12, "bold"),
                                      fg_color="#27ae60",
                                      hover_color="#2ecc71")
        self.upload_btn.pack(side="left", padx=(0,5))

        # 중지 버튼
        self.stop_btn = ctk.CTkButton(button_frame,
                                    text="업로드 중지",
                                    command=self.stop_upload_process,
                                    height=32,
                                    width=120,  # 너비 감소
                                    #font=("맑은 고딕", 12, "bold"),
                                    fg_color="#27ae60",
                                    hover_color="#2ecc71")                                    
        self.stop_btn.pack(side="left", padx=(5,0))

        # 로그 출력 영역
        log_label = ctk.CTkLabel(main_frame, text="로그 데이터", anchor="w") #, font=("맑은 고딕", 12, "bold"))
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

        self.files = []
        
        # 성공적으로 업로드된 파일들을 추적하기 위한 리스트 추가
        self.successful_uploads = []
    
    def check_excel_security(self, file_path):
        """개별 엑셀 파일의 보안 상태를 확인"""
        try:
            file_info = {
                'file_path': str(file_path),
                'file_name': os.path.basename(file_path),
                'file_size': os.path.getsize(file_path),
                'modified_time': datetime.fromtimestamp(os.path.getmtime(file_path)),
                'security_status': 'Unknown',
                'is_password_protected': False,
                'has_macros': False,
                'has_vba_code': False,
                'is_read_only': False,
                'protection_level': 'None',
                'error_message': None
            }
            
            # 파일 확장자 확인
            if not file_path.lower().endswith(('.xlsx', '.xls', '.xlsm', '.xlsb')):
                file_info['error_message'] = '엑셀 파일이 아닙니다.'
                return file_info
            
            # openpyxl을 사용한 보안 검사
            try:
                workbook = openpyxl.load_workbook(file_path, read_only=True)
                
                # 워크북 보호 상태 확인
                if hasattr(workbook, 'security') and workbook.security:
                    file_info['protection_level'] = 'Workbook Protected'
                
                # 워크시트 보호 상태 확인
                protected_sheets = []
                for sheet_name in workbook.sheetnames:
                    sheet = workbook[sheet_name]
                    if sheet.protection.sheet:
                        protected_sheets.append(sheet_name)
                
                if protected_sheets:
                    file_info['protection_level'] = f'Sheet Protected: {", ".join(protected_sheets)}'
                
                # 매크로 확인 (xlsm 파일)
                if file_path.lower().endswith('.xlsm'):
                    file_info['has_macros'] = True
                
                workbook.close()
                
            except Exception as e:
                file_info['error_message'] = f'openpyxl 오류: {str(e)}'
            
            # pandas를 사용한 추가 검사
            try:
                # 파일 읽기 시도 (여러 엔진으로 시도)
                engines = ['openpyxl', 'xlrd']
                df = None
                
                for engine in engines:
                    try:
                        if file_path.lower().endswith(('.xlsx', '.xlsm', '.xlsb')):
                            df = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
                        else:
                            df = pd.read_excel(file_path, sheet_name=None, engine='xlrd')
                        break
                    except Exception as engine_error:
                        if engine == engines[-1]:  # 마지막 엔진까지 실패
                            raise engine_error
                        continue
                
                if df is not None:
                    file_info['sheet_count'] = len(df)
                    file_info['security_status'] = 'Accessible'
                
            except Exception as e:
                error_msg = str(e).lower()
                if 'password' in error_msg or 'encrypted' in error_msg:
                    file_info['is_password_protected'] = True
                    file_info['security_status'] = 'Password Protected'
                elif 'permission' in error_msg or 'read-only' in error_msg:
                    file_info['is_read_only'] = True
                    file_info['security_status'] = 'Read Only'
                elif 'not a zip file' in error_msg:
                    # DRM 보호된 파일 또는 손상된 파일일 수 있음
                    file_info['security_status'] = 'DRM Protected'
                    file_info['error_message'] = 'DRM 보안 프로그램으로 보호된 파일이거나 파일이 손상되었습니다.'
                else:
                    file_info['security_status'] = 'Error'
                    file_info['error_message'] = str(e)
            
            # 보안 상태 분류
            if file_info['is_password_protected']:
                file_info['security_level'] = 'High'
            elif file_info['has_macros'] or file_info['protection_level'] != 'None':
                file_info['security_level'] = 'Medium'
            elif file_info['is_read_only']:
                file_info['security_level'] = 'Low'
            elif file_info['security_status'] == 'DRM Protected':
                file_info['security_level'] = 'DRM Protected'
            elif file_info['security_status'] == 'File Corrupted':
                file_info['security_level'] = 'Corrupted'
            elif file_info['security_status'] == 'Error':
                file_info['security_level'] = 'Error'
            else:
                file_info['security_level'] = 'None'
            
            return file_info
            
        except Exception as e:
            return {
                'file_path': str(file_path),
                'file_name': os.path.basename(file_path),
                'security_status': 'Error',
                'error_message': f'파일 처리 오류: {str(e)}',
                'security_level': 'Unknown'
            }
    
    def check_files_security(self):
        """선택된 모든 파일의 보안 상태를 검사"""
        if not self.files:
            return True, []
        
        self.log_step("파일 보안 상태 검사 시작...")
        self.status_text.set("파일 보안 상태 검사 중...")
        
        drm_protected_files = []
        total_files = len(self.files)
        
        for index, file_path in enumerate(self.files, 1):
            self.log_step(f"보안 검사 중... [{index}/{total_files}] {os.path.basename(file_path)}")
            
            try:
                security_info = self.check_excel_security(file_path)
                
                if security_info['security_level'] == 'DRM Protected':
                    drm_protected_files.append({
                        'file_name': security_info['file_name'],
                        'file_path': security_info['file_path'],
                        'error_message': security_info.get('error_message', 'DRM 보호됨')
                    })
                    self.log_step(f"⚠️ DRM 보호 파일 발견: {security_info['file_name']}")
                elif security_info['security_level'] == 'None':
                    self.log_step(f"✅ 보안 해제됨: {security_info['file_name']}")
                else:
                    self.log_step(f"⚠️ 보안 설정됨 ({security_info['security_level']}): {security_info['file_name']}")
                    
            except Exception as e:
                self.log_step(f"❌ 보안 검사 실패: {os.path.basename(file_path)} - {str(e)}")
                drm_protected_files.append({
                    'file_name': os.path.basename(file_path),
                    'file_path': file_path,
                    'error_message': f'보안 검사 실패: {str(e)}'
                })
        
        if drm_protected_files:
            self.log_step(f"❌ DRM 보호 파일 {len(drm_protected_files)}개 발견")
            return False, drm_protected_files
        else:
            self.log_step("✅ 모든 파일 보안 검사 완료 - 업로드 가능")
            return True, []
    
    def log_message(self, message):
        """상태 메시지를 업데이트하고 로그에 추가"""
        self.status_text.set(message)
        self.update()
    
    def log_step(self, message):
        """로그 메시지를 출력"""
        log_message = f"{message}\n"
        
        # GUI 로그 업데이트
        self.log_text.insert(tk.END, log_message)
        self.log_text.see(tk.END)  # 스크롤을 최신 로그로 이동
        
        # 상태 메시지 업데이트
        self.status_text.set(message)
        self.update()
        
        # 콘솔에도 출력
        print(log_message, end='')
    
    def select_files(self):
        """파일 선택 다이얼로그 열기"""
        files = filedialog.askopenfilenames(
            title="업로드할 파일 선택",
            filetypes=[("모든 파일", "*.*")]
        )
        
        for file in files:
            if file not in self.files:
                self.files.append(file)
                self.file_listbox.insert(tk.END, os.path.basename(file))
    
    def select_directory(self):
        """폴더를 선택하고 해당 폴더의 파일들을 리스트에 추가"""
        directory = filedialog.askdirectory(
            title="파일이 있는 폴더 선택"
        )
        if directory:
            try:
                # 폴더 내 모든 파일 가져오기
                files = [os.path.join(directory, f) for f in os.listdir(directory) 
                        if os.path.isfile(os.path.join(directory, f))]
                
                # 파일들을 리스트에 추가
                for file in files:
                    if file not in self.files:  # 중복 방지
                        self.files.append(file)
                        self.file_listbox.insert(tk.END, os.path.basename(file))
                
                self.status_text.set(f"폴더에서 {len(files)}개 파일을 불러왔습니다.")
                self.log_step(f"폴더 '{directory}'에서 {len(files)}개 파일 로드됨")
                
            except Exception as e:
                error_msg = f"폴더에서 파일 불러오기 실패: {str(e)}"
                self.log_step(error_msg)
                self.status_text.set(error_msg)

    def clear_files(self):
        """파일 목록 초기화"""
        self.files.clear()
        self.file_listbox.delete(0, tk.END)
        self.status_text.set("파일 목록이 초기화되었습니다.")
    
    def stop_upload_process(self):
        """업로드 중지"""
        self.stop_upload = True
        self.log_step("업로드 중지 요청됨")
        self.status_text.set("업로드 중지 중...")
        #self.stop_btn.configure(state="disabled")

    def start_upload(self):
        """파일 업로드 시작"""
        if not self.files:
            self.log_step("오류: 업로드할 파일이 없습니다.")
            self.status_text.set("업로드할 파일이 없습니다.")
            return
        
        # 파일 보안 상태 검사
        security_check_passed, drm_protected_files = self.check_files_security()
        
        if not security_check_passed:
            # DRM 보호 파일이 있는 경우 메시지 표시
            import tkinter.messagebox as messagebox
            
            file_list = "\n".join([f"- {file['file_name']}" for file in drm_protected_files])
            error_message = f"다음 파일들이 보안으로 보호되어 있습니다.\n\n{file_list}\n\n보안을 해제한 후 다시 시도해주세요."
            
            messagebox.showerror("보안 오류", error_message)
            self.log_step("업로드 중지: DRM 보호 파일 발견")
            self.status_text.set("DRM 보호 파일 발견 - 보안 해제 후 다시 시도하세요")
            return
        
        self.log_step("업로드를 시작합니다.")
        self.status_text.set("업로드 준비 중...")
            
        # 성공적으로 업로드된 파일 리스트 초기화
        self.successful_uploads = []
        
        # 로그인 정보 가져오기
        user_id = self.parent.id_entry.get()
        password = self.parent.pw_entry.get()
        
        if not user_id or not password:
            self.log_step("오류: 로그인 정보가 없습니다.")
            self.status_text.set("로그인 정보를 입력해주세요.")
            return

        # 업로드 중지 플래그 초기화
        self.stop_upload = False
        # 업로드 버튼 비활성화, 중지 버튼 활성화
        self.upload_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
            
        # 업로드 로직 시작
        try:
            self.log_step("브라우저 실행 시작")
            self.status_text.set("브라우저 실행 중...")
            driver = initialize_driver(self.parent.show_browser_var.get())
            
            # 로그인 및 페이지 이동
            self.log_step("로그인 시도 중...")
            self.status_text.set("로그인 및 페이지 이동 중...")
            if not login_and_navigate(driver, user_id, password, self.log_step):
                self.log_step("오류: 로그인 또는 페이지 이동 실패")
                self.status_text.set("로그인 또는 페이지 이동 실패")
                driver.quit()
                return
            
            self.log_step("로그인 성공")
            total_files = len(self.files)
            for index, file in enumerate(self.files, 1):
                # 중지 요청 확인
                if self.stop_upload:
                    self.log_step("업로드가 사용자에 의해 중지되었습니다.")
                    break

                try:
                    self.log_step(f"=== 파일 {index}/{total_files} 처리 시작: {file} ===")
                    self.status_text.set(f"[{index}/{total_files}] {file}\n파일 업로드 중...")
                    
                    # 업로드 버튼 클릭
                    self.log_step("업로드 버튼 찾는 중...")
                    upload_btn = driver.find_element(By.XPATH, '//td[@class="btn01"][contains(text(), "업로드")]')
                    upload_btn.click()
                    self.log_step("업로드 버튼 클릭 완료")
                    time.sleep(1)  # 파일 선택 창이 뜰 때까지 대기
                    
                    # 파일명만 입력 및 열기 버튼 클릭
                    filename = os.path.basename(file)
                    self.log_step(f"파일명 입력 중: {filename}")
                    pyautogui.typewrite(filename, interval=0.1)
                    time.sleep(0.5)
                    pyautogui.press('enter')
                    self.log_step("파일 선택 완료")
                    time.sleep(1)

                    try:
                        # 파일 로딩 상태 체크
                        self.log_step("파일 로딩 상태 체크 중...")
                        self.status_text.set(f"[{index}/{total_files}] {file}\n파일 로딩 중...")
                        
                        # 로딩 완료 확인을 위한 헤더 체크박스 확인
                        header_id = "ContentPlaceHolder1_ASPxSplitter1_dxGrid1_col0"
                        checkbox_id = "ContentPlaceHolder1_ASPxSplitter1_dxGrid1_DXSelAllBtn0_D"
                        max_attempts = 200  # 최대 시도 횟수를 200으로 증가
                        check_interval = 5  # 체크 간격 (초)
                        loading_success = False
                        
                        for attempt in range(max_attempts):
                            try:
                                self.log_step(f"데이터 로딩체크 중... {attempt + 1}/{max_attempts}")
                                
                                # 헤더 셀 찾기
                                header_cell = driver.find_element(By.ID, header_id)
                                if header_cell:
                                    # 체크박스 존재 여부 확인
                                    try:
                                        checkbox = header_cell.find_element(By.ID, checkbox_id)
                                        if checkbox.is_displayed():
                                            self.log_step("데이터 발견 - 로딩 완료")
                                            loading_success = True
                                            break
                                        else:
                                            self.log_step("체크박스가 존재하지만 표시되지 않음")
                                    except:
                                        self.log_step("데이터 로딩 중...")
                                else:
                                    self.log_step("헤더를 찾을 수 없음")
                                
                                if attempt < max_attempts - 1 and not loading_success:
                                    self.log_step(f"데이터 로딩상태, {check_interval}초 후 재확인")
                                    time.sleep(check_interval)
                                elif not loading_success:
                                    self.log_step("최대 시도 횟수 초과")
                            except Exception as e:
                                if attempt < max_attempts - 1:
                                    self.log_step(f"확인 실패 ({str(e)}), 재시도 중...")
                                    time.sleep(check_interval)
                                else:
                                    self.log_step(f"최대 시도 횟수 초과: {str(e)}")
                        
                        if not loading_success:
                            self.log_step("파일 로딩 실패 - 체크박스를 찾을 수 없음")
                            self.file_listbox.itemconfig(self.files.index(file), {'bg': 'red', 'fg': 'white'})
                            continue
                            
                        # 체크박스 클릭
                        self.log_step("체크박스 클릭 시도 중...")
                        try:
                            # 먼저 일반적인 방법으로 시도
                            checkbox = driver.find_element(By.ID, checkbox_id)
                            if checkbox.is_displayed() and checkbox.is_enabled():
                                driver.execute_script("arguments[0].click();", checkbox)
                                self.log_step("체크박스 클릭 성공")
                            else:
                                # JavaScript를 사용하여 강제로 클릭
                                driver.execute_script(f"""
                                    var checkbox = document.getElementById('{checkbox_id}');
                                    if (checkbox) {{
                                        checkbox.click();
                                        return true;
                                    }}
                                    return false;
                                """)
                                self.log_step("JavaScript를 통한 체크박스 클릭 성공")
                        except Exception as e:
                            self.log_step(f"체크박스 클릭 실패: {str(e)}")
                            self.file_listbox.itemconfig(self.files.index(file), {'bg': 'red', 'fg': 'white'})
                            continue

                        # 체크박스 상태 변경 확인
                        self.log_step("체크박스 상태 변경 확인 중...")
                        checked_box_xpath = "//span[contains(@class, 'dxWeb_edtCheckBoxChecked')]"
                        try:
                            wait = WebDriverWait(driver, 10)
                            wait.until(
                                EC.presence_of_element_located((By.XPATH, checked_box_xpath))
                            )
                            self.log_step("체크박스 상태 변경 확인됨")
                        except Exception as e:
                            self.log_step(f"체크박스 상태 변경 확인 실패: {str(e)}")
                            self.file_listbox.itemconfig(self.files.index(file), {'bg': 'red', 'fg': 'white'})
                            continue

                        # 등록 버튼 클릭
                        self.log_step("등록 버튼 찾는 중...")
                        try:
                            # 먼저 일반적인 방법으로 시도
                            register_btn = wait.until(
                                EC.element_to_be_clickable((By.XPATH, "//td[@class='btn01'][contains(text(), '등록')]"))
                            )
                            driver.execute_script("arguments[0].click();", register_btn)
                            self.log_step("등록 버튼 클릭 완료")
                        except Exception as e:
                            self.log_step(f"일반적인 클릭 실패, JavaScript로 시도: {str(e)}")
                            try:
                                # JavaScript를 사용하여 강제로 클릭
                                driver.execute_script("""
                                    var elements = document.getElementsByTagName('td');
                                    for(var i = 0; i < elements.length; i++) {
                                        if(elements[i].textContent.includes('등록')) {
                                            elements[i].click();
                                            return true;
                                        }
                                    }
                                    return false;
                                """)
                                self.log_step("JavaScript를 통한 등록 버튼 클릭 완료")
                            except Exception as e:
                                self.log_step(f"등록 버튼 클릭 실패: {str(e)}")
                                self.file_listbox.itemconfig(self.files.index(file), {'bg': 'red', 'fg': 'white'})
                                continue

                        # 확인 메시지 처리
                        try:
                            self.log_step("확인 메시지 대기 중...")
                            alert = wait.until(EC.alert_is_present())
                            alert.accept()
                            self.log_step("확인 메시지 처리 완료")
                            
                            # 로딩창 감지 및 대기
                            try:
                                start_time = time.time()
                                loading_appeared = False
                                max_wait_time = 30  # 최대 대기 시간
                                last_log_time = 0  # 마지막 로그 출력 시간
                                
                                while True:
                                    current_time = time.time()
                                    if current_time - start_time > max_wait_time:
                                        raise TimeoutException("로딩 대기 시간 초과")
                                    
                                    # UpdateProgress1 div의 상태 확인
                                    loading_div = driver.find_elements(By.ID, "UpdateProgress1")
                                    if loading_div:
                                        style = loading_div[0].get_attribute("style")
                                        aria_hidden = loading_div[0].get_attribute("aria-hidden")
                                        
                                        if "display: block" in style or aria_hidden == "false":
                                            loading_appeared = True
                                            # 5초마다 한 번씩만 로딩 메시지 출력
                                            if not last_log_time or current_time - last_log_time >= 5:
                                                self.log_step("등록 중...")
                                                last_log_time = current_time
                                        elif loading_appeared and ("display: none" in style or aria_hidden == "true"):
                                            self.log_step("등록 완료")
                                            time.sleep(2)  # 추가 안정화 대기
                                            break
                                    
                                    # 3초 동안 로딩이 나타나지 않으면 계속 진행
                                    if not loading_appeared and current_time - start_time > 3:
                                        self.log_step("로딩 없이 계속 진행")
                                        break
                                        
                                    time.sleep(0.1)
                                
                            except TimeoutException as te:
                                self.log_step(f"로딩 패널 처리 중 시간 초과: {str(te)}")
                            except Exception as e:
                                self.log_step(f"로딩 패널 처리 중 오류 발생: {str(e)}")

                        except Exception as e:
                            self.log_step(f"확인 메시지 처리 실패: {str(e)}")
                            self.file_listbox.itemconfig(self.files.index(file), {'bg': 'red', 'fg': 'white'})
                            continue

                        # 등록 완료 확인
                        try:
                            self.log_step("등록 완료 확인 중...")
                            max_wait_time = 60  # 최대 60초 대기
                            start_time = time.time()
                            completion_confirmed = False
                            
                            # No data 메시지 확인
                            while True:
                                if time.time() - start_time > max_wait_time:
                                    raise TimeoutException("등록 완료 확인 시간 초과")
                                
                                try:
                                    # No data 메시지가 있는지 확인
                                    no_data = driver.find_elements(By.XPATH, "//div[text()='No data']")
                                    if not no_data:  # No data 메시지가 없으면 데이터가 로드된 것
                                        self.log_step("데이터 등록 완료")
                                        completion_confirmed = True
                                        break
                                    else:
                                        self.log_step("데이터가 없습니다. 다시 확인합니다.")
                                        time.sleep(2)  # 잠시 대기 후 재시도
                                except Exception as e:
                                    self.log_step(f"데이터 확인 중 오류 발생: {str(e)}")
                                    time.sleep(2)  # 오류 발생시 잠시 대기
                            
                            if not completion_confirmed:
                                raise TimeoutException("등록 완료 확인 시간 초과")
                            
                            # 추가 안정화 대기
                            time.sleep(1)
                            
                            # 일반 모드로 전환
                            self.log_step("일반 모드로 전환 중...")
                            try:
                                # ID로 시도
                                normal_mode_ids = [
                                    "ContentPlaceHolder1_itbtnA_Normal",
                                    "ContentPlaceHolder1_btnNormal",
                                    "ctl00_ContentPlaceHolder1_btnNormal",
                                    "ContentPlaceHolder1_ASPxSplitter1_btnNormal"
                                ]
                                
                                normal_mode_clicked = False
                                for mode_id in normal_mode_ids:
                                    try:
                                        # 요소가 존재하는지 먼저 확인
                                        normal_mode_btn = WebDriverWait(driver, 5).until(
                                            EC.presence_of_element_located((By.ID, mode_id))
                                        )
                                        
                                        # 요소가 클릭 가능한지 확인
                                        if normal_mode_btn.is_displayed() and normal_mode_btn.is_enabled():
                                            # 스크롤하여 요소를 보이게 만들기
                                            driver.execute_script("arguments[0].scrollIntoView(true);", normal_mode_btn)
                                            time.sleep(0.5)  # 스크롤 후 잠시 대기
                                            
                                            # 클릭 시도
                                            try:
                                                normal_mode_btn.click()
                                            except:
                                                driver.execute_script("arguments[0].click();", normal_mode_btn)
                                            
                                            normal_mode_clicked = True
                                            self.log_step(f"일반 모드 버튼 클릭 성공 (ID: {mode_id})")
                                            break
                                    except:
                                        continue
                                
                                if not normal_mode_clicked:
                                    # 모든 ID 시도가 실패하면 JavaScript로 시도
                                    self.log_step("JavaScript로 일반 모드 버튼 찾기 시도...")
                                    clicked = driver.execute_script("""
                                        var ids = arguments[0];
                                        for(var i = 0; i < ids.length; i++) {
                                            var btn = document.getElementById(ids[i]);
                                            if(btn) {
                                                btn.scrollIntoView(true);
                                                setTimeout(function() {
                                                    btn.click();
                                                }, 100);
                                                return true;
                                            }
                                        }
                                        return false;
                                    """, normal_mode_ids)
                                    
                                    if clicked:
                                        self.log_step("JavaScript를 통한 일반 모드 버튼 클릭 성공")
                                        normal_mode_clicked = True
                                
                                if not normal_mode_clicked:
                                    raise Exception("일반 모드 버튼을 찾을 수 없거나 클릭할 수 없습니다.")
                                
                                # 일반 모드 전환 대기
                                time.sleep(2)
                                
                                # 업로드 버튼이 나타날 때까지 대기
                                try:
                                    upload_btn = WebDriverWait(driver, 10).until(
                                        EC.element_to_be_clickable((By.XPATH, '//td[@class="btn01"][contains(text(), "업로드")]'))
                                    )
                                    self.log_step("일반 모드 전환 완료")
                                    
                                    # 성공 표시 (녹색)
                                    self.file_listbox.itemconfig(self.files.index(file), {'bg': '#27ae60', 'fg': 'white'})
                                    self.log_step(f"=== 파일 {index}/{total_files} 처리 완료 ===\n")
                                    # 성공적으로 업로드된 파일 추가
                                    self.successful_uploads.append(file)
                                except Exception as e:
                                    self.log_step(f"업로드 버튼 확인 실패: {str(e)}")
                                    raise
                                    
                            except Exception as e:
                                self.log_step(f"일반 모드 버튼 클릭 실패: {str(e)}")
                                self.file_listbox.itemconfig(self.files.index(file), {'bg': 'red', 'fg': 'white'})
                                raise

                        except TimeoutException as e:
                            self.log_step(f"등록 완료 확인 시간 초과: {str(e)}")
                            self.file_listbox.itemconfig(self.files.index(file), {'bg': 'red', 'fg': 'white'})
                            continue
                        except Exception as e:
                            self.log_step(f"등록 완료 확인 또는 일반 모드 전환 실패: {str(e)}")
                            self.file_listbox.itemconfig(self.files.index(file), {'bg': 'red', 'fg': 'white'})
                            continue

                    except Exception as e:
                        self.log_step(f"등록 완료 확인 또는 일반 모드 전환 실패: {str(e)}")
                        self.file_listbox.itemconfig(self.files.index(file), {'bg': 'red', 'fg': 'white'})
                        continue

                    # 다음 파일 처리를 위한 대기
                    time.sleep(2)

                except Exception as e:
                    error_msg = f"[{index}/{total_files}] {file}\n예외 발생: {str(e)}"
                    self.log_step(f"치명적 오류 발생: {error_msg}")
                    self.status_text.set(error_msg)
                    print(error_msg)
                    self.file_listbox.itemconfig(self.files.index(file), {'bg': 'red', 'fg': 'white'})
                    if not self.stop_upload:  # 중지 요청이 아닌 경우에만 계속 진행
                        continue
                    else:
                        break

            self.log_step("=== 모든 파일 업로드 작업 완료 ===")
            self.status_text.set("모든 파일 업로드 완료")
            
            # 파일 삭제 여부 확인 및 처리
            self.delete_uploaded_files()
            
            driver.quit()

        except Exception as e:
            error_msg = f"예외 발생: {str(e)}"
            self.log_step(f"시스템 오류 발생: {error_msg}")
            self.status_text.set(error_msg)
            print(error_msg)
            if 'driver' in locals():
                driver.quit()
        finally:
            # 업로드 버튼 활성화, 중지 버튼 비활성화
            self.upload_btn.configure(state="normal")
            self.stop_btn.configure(state="disabled")
            # 중지 플래그 초기화
            self.stop_upload = False 

    def delete_uploaded_files(self):
        """성공적으로 업로드된 파일들의 삭제 여부를 확인하고 처리하는 메소드"""
        if not self.successful_uploads:  # 성공적으로 업로드된 파일이 없으면 처리하지 않음
            self.log_step("삭제할 파일이 없습니다. (성공적으로 업로드된 파일 없음)")
            return
            
        try:
            # 삭제 확인 대화상자 표시
            if tk.messagebox.askyesno("파일 삭제 확인", 
                "성공적으로 업로드된 파일들을 폴더에서 삭제하시겠습니까?"):
                
                deleted_count = 0
                failed_count = 0
                failed_files = []
                
                for file_path in self.successful_uploads:  # self.files 대신 self.successful_uploads 사용
                    try:
                        if os.path.exists(file_path):  # 파일이 실제로 존재하는지 확인
                            os.remove(file_path)
                            deleted_count += 1
                            # 파일 리스트에서도 제거
                            if file_path in self.files:
                                idx = self.files.index(file_path)
                                self.files.remove(file_path)
                                self.file_listbox.delete(idx)
                            self.log_step(f"파일 삭제 완료: {os.path.basename(file_path)}")
                    except Exception as e:
                        failed_count += 1
                        failed_files.append(os.path.basename(file_path))
                        self.log_step(f"파일 삭제 실패: {os.path.basename(file_path)} - {str(e)}")
                
                # 결과 메시지 생성
                result_msg = f"총 {deleted_count}개 파일 삭제 완료"
                if failed_count > 0:
                    result_msg += f"\n{failed_count}개 파일 삭제 실패"
                    if failed_files:
                        result_msg += f"\n실패한 파일: {', '.join(failed_files)}"
                
                self.log_step(result_msg)
                self.status_text.set(result_msg)
                
                # 성공적으로 업로드된 파일 리스트 초기화
                self.successful_uploads = []
        except Exception as e:
            error_msg = f"파일 삭제 중 오류 발생: {str(e)}"
            self.log_step(error_msg)
            self.status_text.set(error_msg) 