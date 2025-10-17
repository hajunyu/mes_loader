import customtkinter as ctk
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import threading
import time
import traceback  # 디버깅을 위해 추가
from selenium.common.exceptions import TimeoutException
import json
import os
from tkinter import messagebox
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import filedialog
from mes_uploader import FileUploader  # 업로더 클래스 import
from mes_common import initialize_driver, login_and_navigate
from mes_unbalance import UnbalanceLoader  # 언발런스 로더 import 추가
from mes_updater import UpdateManager, UpdateDialog  # 업데이터 import 추가
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")  # 기본 제공되는 테마 사용

class MESDownloader(ctk.CTk):
    def __init__(self):
        """GUI 초기화"""
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        super().__init__()
        self.title("MES Web Report")
        self.geometry("400x750")  # 창 크기 축소
        self.resizable(False, False)  # 창 크기 고정

        # 설정 파일 경로
        self.config_file = "mes_config.json"
        self.load_config()  # 설정 로드

        # 전체 UI 생성
        self.create_main_frame()


    def load_config(self):
        """설정 파일 로드"""
        default_config = {
            "version": UpdateManager.current_version,  # 버전 정보를 UpdateManager에서 가져옴
            "line_types": ["[01] 제어기"],
            "lines": {
                "[01] 제어기": [
                    "[AS01] HPCU A/S",
                    "[EC01] EPCU 1",
                    "[HC02] HPCU 2",
                    "[HC06] HPCU 6",
                    "[HC07] HPCU 7",
                    "[HC08] HPCU 8",
                    "[HC09] HPCU 9",
                    "[HC10] HPCU 10",
                    "[HC11] HPCU 11",
                    "[HC12] HPCU 12",
                    "[IC03] ICCU 3",
                    "[OC02] OBC 2",
                    "[OC03] OBC 3",
                    "[OC04] OBC 4",
                    "[PM01] 파워모듈 1",
                    "[PM05] 파워모듈 5",
                    "[PM06] 파워모듈 6",
                    "[PM07] 파워모듈 7",
                    "[PM08] 파워모듈 8",
                    "[PM09] 파워모듈 9",
                    "[PM10] 파워모듈 10",
                    "[PM11] 파워모듈 11",
                    "[PM12] 파워모듈 12",
                ]
            },
            "login_info": {
                "save_login": False,
                "username": "",
                "password": ""
            },
            "editor_path": "",  # 편집기 경로 추가
            "process_data": {
                'HPCU 2': ['HC02-010', 'HC02-020', 'HC02-030', 'HC02-040', 'HC02-050', 'HC02-060', 'HC02-070', 'HC02-080', 'HC02-090', 'HC02-100', 'HC02-110', 'HC02-120', 'HC02-130', 'HC02-140', 'HC02-150', 'HC02-160', 'HC02-180', 'HC02-200', 'HC02-210', 'HC02-215', 'HC02-220', 'HC02-230', 'HC02-900'],
                'HPCU 6': ['HC06-010', 'HC06-020', 'HC06-030', 'HC06-040', 'HC06-050', 'HC06-060', 'HC06-070', 'HC06-080', 'HC06-090', 'HC06-100', 'HC06-110', 'HC06-120', 'HC06-130', 'HC06-140', 'HC06-145', 'HC06-150', 'HC06-160', 'HC06-170', 'HC06-175', 'HC06-180', 'HC06-190', 'HC06-200', 'HC06-210', 'HC06-220', 'HC06-230', 'HC06-240', 'HC06-250', 'HC06-260', 'HC06-900'],
                'HPCU 7': ['HC07-010', 'HC07-020', 'HC07-030', 'HC07-040', 'HC07-050', 'HC07-060', 'HC07-065', 'HC07-067', 'HC07-070', 'HC07-080', 'HC07-085', 'HC07-090', 'HC07-100', 'HC07-110', 'HC07-120', 'HC07-130', 'HC07-140', 'HC07-150', 'HC07-160', 'HC07-170', 'HC07-900'],
                'HPCU 8': ['HC08-010', 'HC08-020', 'HC08-030', 'HC08-035', 'HC08-040', 'HC08-050', 'HC08-060', 'HC08-067', 'HC08-070', 'HC08-080', 'HC08-090', 'HC08-100', 'HC08-110', 'HC08-120', 'HC08-130', 'HC08-140', 'HC08-150', 'HC08-160', 'HC08-165', 'HC08-170', 'HC08-180', 'HC08-190', 'HC08-200', 'HC08-210', 'HC08-220', 'HC08-230', 'HC08-240', 'HC08-900'],
                'HPCU 9': ['HC09-010', 'HC09-020', 'HC09-030', 'HC09-035', 'HC09-040', 'HC09-050', 'HC09-060', 'HC09-070', 'HC09-080', 'HC09-090', 'HC09-100', 'HC09-110', 'HC09-120', 'HC09-130', 'HC09-140', 'HC09-150', 'HC09-160', 'HC09-165', 'HC09-170', 'HC09-180', 'HC09-190', 'HC09-200', 'HC09-210', 'HC09-220', 'HC09-230', 'HC09-240', 'HC09-900'],
                'HPCU 10': ['HC10-010', 'HC10-020', 'HC10-030', 'HC10-040', 'HC10-050', 'HC10-060', 'HC10-070', 'HC10-080', 'HC10-090', 'HC10-100', 'HC10-110', 'HC10-130', 'HC10-140', 'HC10-150', 'HC10-160', 'HC10-170', 'HC10-180', 'HC10-190', 'HC10-200', 'HC10-900'],
                'HPCU 11': ['HC11-010', 'HC11-020', 'HC11-030', 'HC11-040', 'HC11-050', 'HC11-060', 'HC11-070', 'HC11-080', 'HC11-090', 'HC11-100', 'HC11-110', 'HC11-120', 'HC11-130', 'HC11-140', 'HC11-150', 'HC11-160', 'HC11-170', 'HC11-180', 'HC11-190', 'HC11-900'],
                'ICCU 3': ['IC03-010', 'IC03-020', 'IC03-025', 'IC03-027', 'IC03-030', 'IC03-040', 'IC03-050', 'IC03-060', 'IC03-070', 'IC03-080', 'IC03-090', 'IC03-100', 'IC03-110', 'IC03-120', 'IC03-125', 'IC03-130', 'IC03-140', 'IC03-147', 'IC03-149', 'IC03-150', 'IC03-160', 'IC03-165', 'IC03-170', 'IC03-175', 'IC03-180', 'IC03-190', 'IC03-200', 'IC03-210', 'IC03-213', 'IC03-215', 'IC03-220', 'IC03-230', 'IC03-900'],
                '파워모듈 1': ['PM01-010', 'PM01-020', 'PM01-023', 'PM01-027', 'PM01-030', 'PM01-040', 'PM01-050', 'PM01-060', 'PM01-070', 'PM01-080', 'PM01-090', 'PM01-900'],
                '파워모듈 5': ['PM05-010', 'PM05-015', 'PM05-020', 'PM05-030', 'PM05-033', 'PM05-035', 'PM05-040', 'PM05-045', 'PM05-050', 'PM05-060', 'PM05-065', 'PM05-070', 'PM05-073', 'PM05-075', 'PM05-076', 'PM05-077', 'PM05-080', 'PM05-085', 'PM05-090', 'PM05-100', 'PM05-110', 'PM05-120', 'PM05-130', 'PM05-140', 'PM05-150', 'PM05-170', 'PM05-180', 'PM05-190', 'PM05-200', 'PM05-205', 'PM05-210', 'PM05-220', 'PM05-230', 'PM05-240', 'PM05-250', 'PM05-900'],
                '파워모듈 6': ['PM06-010', 'PM06-020', 'PM06-030', 'PM06-040', 'PM06-045', 'PM06-050', 'PM06-060', 'PM06-070', 'PM06-080', 'PM06-085', 'PM06-090', 'PM06-100', 'PM06-110', 'PM06-120', 'PM06-130', 'PM06-140', 'PM06-150', 'PM06-160', 'PM06-180', 'PM06-190', 'PM06-200', 'PM06-210', 'PM06-215', 'PM06-220', 'PM06-230', 'PM06-240', 'PM06-250', 'PM06-260', 'PM06-900'],
                '파워모듈 7': ['PM07-010', 'PM07-020', 'PM07-030', 'PM07-035', 'PM07-040', 'PM07-050', 'PM07-060', 'PM07-070', 'PM07-080', 'PM07-090', 'PM07-092', 'PM07-095', 'PM07-100', 'PM07-110', 'PM07-120', 'PM07-900'],
                '파워모듈 8': ['PM08-010', 'PM08-020', 'PM08-030', 'PM08-040', 'PM08-045', 'PM08-050', 'PM08-055', 'PM08-060', 'PM08-070', 'PM08-080', 'PM08-090', 'PM08-100', 'PM08-110', 'PM08-120', 'PM08-130', 'PM08-140', 'PM08-150', 'PM08-160', 'PM08-170', 'PM08-180', 'PM08-900'],
                '파워모듈 9': ['PM09-010', 'PM09-020', 'PM09-030', 'PM09-035', 'PM09-040', 'PM09-045', 'PM09-050', 'PM09-060', 'PM09-070', 'PM09-080', 'PM09-090', 'PM09-100', 'PM09-110', 'PM09-120', 'PM09-130', 'PM09-140', 'PM09-150', 'PM09-160', 'PM09-170', 'PM09-900'],
                '파워모듈 10': ['PM10-010', 'PM10-020', 'PM10-025', 'PM10-030', 'PM10-040', 'PM10-045', 'PM10-050', 'PM10-060', 'PM10-070', 'PM10-080', 'PM10-090', 'PM10-100', 'PM10-110', 'PM10-115', 'PM10-120', 'PM10-130', 'PM10-140', 'PM10-150', 'PM10-155', 'PM10-157', 'PM10-160', 'PM10-170', 'PM10-180', 'PM10-900'],
                '파워모듈 11': ['PM11-010', 'PM11-020', 'PM11-025', 'PM11-030', 'PM11-040', 'PM11-045', 'PM11-050', 'PM11-060', 'PM11-070', 'PM11-080', 'PM11-090', 'PM11-100', 'PM11-110', 'PM11-120', 'PM11-130', 'PM11-140', 'PM11-150', 'PM11-155', 'PM11-160', 'PM11-170', 'PM11-180', 'PM11-900'],
                'EPCU 1': ['EC01-001', 'EC01-005', 'EC01-010', 'EC01-015', 'EC01-020', 'EC01-030', 'EC01-040', 'EC01-050', 'EC01-060', 'EC01-070', 'EC01-075', 'EC01-080', 'EC01-900'],
                'OBC 2': ['OC02-010', 'OC02-020', 'OC02-030', 'OC02-040', 'OC02-050', 'OC02-060', 'OC02-065', 'OC02-070', 'OC02-080', 'OC02-090', 'OC02-100', 'OC02-110', 'OC02-120', 'OC02-125', 'OC02-130', 'OC02-900'],
                'OBC 3': ['OC03-010', 'OC03-020', 'OC03-030', 'OC03-040', 'OC03-050', 'OC03-060', 'OC03-065', 'OC03-070', 'OC03-080', 'OC03-090', 'OC03-100', 'OC03-110', 'OC03-120', 'OC03-130', 'OC03-140', 'OC03-900'],
                'OBC 4': ['OC04-010', 'OC04-020', 'OC04-030', 'OC04-040', 'OC04-050', 'OC04-070', 'OC04-080', 'OC04-090', 'OC04-100', 'OC04-105', 'OC04-110', 'OC04-120', 'OC04-130', 'OC04-140', 'OC04-900']
            }
        }
        
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    self.config = json.load(f)
                    # 기존 설정 파일에 login_info가 없는 경우 추가
                    if "login_info" not in self.config:
                        self.config["login_info"] = default_config["login_info"]
                    # 기존 설정 파일에 process_data가 없는 경우 추가
                    if "process_data" not in self.config:
                        self.config["process_data"] = default_config["process_data"]
                    # 기존 설정 파일에 version이 없는 경우 추가
                    if "version" not in self.config:
                        self.config["version"] = default_config["version"]
                    self.save_config()
            else:
                self.config = default_config
                self.save_config()
        except Exception as e:
            print(f"설정 로드 중 오류: {e}")
            self.config = default_config

    def save_config(self):
        """설정 파일 저장"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"설정 저장 중 오류: {e}")

    def create_main_frame(self):
        """메인 프레임 생성"""
        main_frame = ctk.CTkFrame(self)
        main_frame.pack(expand=True, fill="both", padx=20, pady=20)

        # MES 로그인 헤더
        header_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        header_frame.pack(fill="x", pady=(0,20))
        ctk.CTkLabel(header_frame, text="MES 웹리포트", font=("맑은 고딕", 22, "bold")).pack(anchor="center")

        # 로그인 프레임
        login_frame = ctk.CTkFrame(main_frame)
        login_frame.pack(fill="x", padx=20, pady=(0,20))

        # 입력 필드 프레임
        input_frame = ctk.CTkFrame(login_frame, fg_color="transparent")
        input_frame.pack(pady=10)

        # 아이디 입력 행
        id_frame = ctk.CTkFrame(input_frame, fg_color="transparent")
        id_frame.pack(pady=(0,5))
        ctk.CTkLabel(id_frame, text="아이디", width=80, anchor="e").pack(side="left", padx=(0,10))
        self.id_entry = ctk.CTkEntry(id_frame, height=32, width=200)
        self.id_entry.pack(side="left")

        # 비밀번호 입력 행
        pw_frame = ctk.CTkFrame(input_frame, fg_color="transparent")
        pw_frame.pack()
        ctk.CTkLabel(pw_frame, text="비밀번호", width=80, anchor="e").pack(side="left", padx=(0,10))
        self.pw_entry = ctk.CTkEntry(pw_frame, show="*", height=32, width=200)
        self.pw_entry.pack(side="left")

        # 체크박스 프레임
        checkbox_frame = ctk.CTkFrame(login_frame, fg_color="transparent")
        checkbox_frame.pack(pady=10, fill="x")
        
        # 체크박스들을 담을 프레임
        checks_frame = ctk.CTkFrame(checkbox_frame, fg_color="transparent")
        checks_frame.pack(pady=(0,10))
        
        # 브라우저 표시 체크박스
        self.show_browser_var = tk.BooleanVar(value=False)
        show_browser_check = ctk.CTkCheckBox(checks_frame, text="브라우저 표시", 
                                           variable=self.show_browser_var)
        show_browser_check.pack(side="left", padx=(0,10))

        # 로그인 정보 저장 체크박스
        self.save_login_var = tk.BooleanVar(value=self.config["login_info"]["save_login"])
        save_login_check = ctk.CTkCheckBox(checks_frame, text="로그인 정보 저장",
                                         variable=self.save_login_var,
                                         command=self.toggle_save_login)
        save_login_check.pack(side="left")

        # 버튼 프레임 (중앙 정렬)
        button_frame = ctk.CTkFrame(checkbox_frame, fg_color="transparent")
        button_frame.pack(fill="x")
        
        # 버튼을 담을 중앙 정렬 프레임
        center_button_frame = ctk.CTkFrame(button_frame, fg_color="transparent")
        center_button_frame.pack(expand=True)

        # 설정 버튼
        self.settings_btn = ctk.CTkButton(
            center_button_frame,
            text="설정",
            width=60,
            height=32,
            command=self.show_settings
        )
        self.settings_btn.pack(side="left", padx=5)

        # 업로드 버튼
        self.upload_btn = ctk.CTkButton(
            center_button_frame, 
            text="업로드",
            width=60,
            height=32,
            command=self.show_uploader,
           # fg_color="#27ae60",
           # hover_color="#2ecc71"
        )
        self.upload_btn.pack(side="left", padx=5)

        # 비가동편집 버튼
        self.edit_downtime_btn = ctk.CTkButton(
            center_button_frame,
            text="편집기",
            width=60,
            height=32,
            command=self.open_downtime_editor,
        )
        self.edit_downtime_btn.pack(side="left", padx=5)

        # 분리모드 버튼 추가
        self.split_mode_btn = ctk.CTkButton(
            center_button_frame,
            text="분리모드",
            width=60,
            command=self.run_split_mode
        )
        self.split_mode_btn.pack(side="left", padx=5)
        
        # 저장된 로그인 정보 불러오기
        if self.config["login_info"]["save_login"]:
            self.id_entry.insert(0, self.config["login_info"]["username"])
            self.pw_entry.insert(0, self.config["login_info"]["password"])
        
        # 라인타입 선택 프레임
        select_frame = ctk.CTkFrame(main_frame)
        select_frame.pack(fill="x", padx=20, pady=(0,20))

        # 라인타입 레이블과 드롭다운
        dropdown_frame = ctk.CTkFrame(select_frame, fg_color="transparent")
        dropdown_frame.pack(side="left", fill="x", expand=True)
        ctk.CTkLabel(dropdown_frame, text="라인타입:", width=80, anchor="e").pack(side="left", padx=(0,10))
        self.line_type = ctk.CTkComboBox(
            dropdown_frame,
            values=self.config["line_types"],
            command=self.on_line_type_change,
            width=180,
            height=32
        )
        self.line_type.pack(side="left")
        self.line_type.set(self.config["line_types"][0])  # 첫 번째 라인타입으로 초기화

        # 라인 선택 영역 전체 프레임
        lines_container = ctk.CTkFrame(main_frame, fg_color="transparent")
        lines_container.pack(fill="x", pady=(0,20))

        # 가운데 정렬을 위한 컨테이너
        center_container = ctk.CTkFrame(lines_container, fg_color="transparent")
        center_container.pack(expand=True)

        # 컨테이너 크기 계산 (전체 너비의 90%)
        container_width = int(380 * 0.9)  # 전체 창 너비 400의 90%
        frame_width = int(container_width / 2 - 10)  # 각 프레임의 너비 (여백 20px 고려)

        # 왼쪽: 라인 선택 영역 (50%)
        left_frame = ctk.CTkFrame(center_container, fg_color="transparent")
        left_frame.pack(side="left", padx=10)
        left_frame.pack_propagate(False)
        left_frame.configure(width=frame_width)
        
        # 라인 선택 헤더
        header_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
        header_frame.pack(fill="x", pady=(0,5))
        ctk.CTkLabel(header_frame, text="라인 선택", font=("맑은 고딕", 12, "bold")).pack(anchor="center")
        
        # 라인 체크박스 스크롤 영역 - 배경색만 살짝 어둡게
        self.lines_frame = ctk.CTkScrollableFrame(left_frame, height=150, width=frame_width,
                                                fg_color=("gray92", "gray13"))
        self.lines_frame.pack(fill="both", expand=True)
        
        # 라인 체크박스 변수 저장용 딕셔너리
        self.line_vars = {}
        
        # 초기 라인 체크박스 생성
        self.create_line_checkboxes(self.line_type.get())  # 초기 라인 목록 생성

        # 오른쪽: 선택된 라인 목록 (50%)
        right_frame = ctk.CTkFrame(center_container, fg_color="transparent")
        right_frame.pack(side="left", padx=10)
        right_frame.pack_propagate(False)
        right_frame.configure(width=frame_width)
        
        # 선택된 라인 목록 헤더
        header_frame = ctk.CTkFrame(right_frame, fg_color="transparent")
        header_frame.pack(fill="x", pady=(0,5))
        ctk.CTkLabel(header_frame, text="선택된 라인 목록", font=("맑은 고딕", 12, "bold")).pack(anchor="center")
        
        # 선택된 라인 목록 텍스트 영역 - 배경색만 살짝 어둡게
        self.selected_lines_text = ctk.CTkTextbox(right_frame, height=150, width=frame_width,
                                                fg_color=("gray92", "gray13"))
        self.selected_lines_text.pack(fill="both", expand=True)

        # 날짜 선택 프레임
        date_frame = ctk.CTkFrame(main_frame)
        date_frame.pack(fill="x", pady=(0,20))

        # 시작 날짜
        start_date_frame = ctk.CTkFrame(date_frame)
        start_date_frame.pack(fill="x", pady=(0,5))
        
        ctk.CTkLabel(start_date_frame, text="시작 날짜 (YYYY-MM-DD)").pack(side="left")
        self.start_date_entry = ctk.CTkEntry(start_date_frame, height=32)
        self.start_date_entry.pack(side="right", expand=True, fill="x", padx=(10,0))
        
        # 종료 날짜
        end_date_frame = ctk.CTkFrame(date_frame)
        end_date_frame.pack(fill="x", pady=(5,0))
        
        ctk.CTkLabel(end_date_frame, text="종료 날짜 (YYYY-MM-DD)").pack(side="left")
        self.end_date_entry = ctk.CTkEntry(end_date_frame, height=32)
        self.end_date_entry.pack(side="right", expand=True, fill="x", padx=(10,0))

        # 기본 날짜 설정 (어제 날짜)
        today = datetime.now()
        yesterday = today - timedelta(days=1)
        self.start_date_entry.insert(0, yesterday.strftime("%Y-%m-%d"))
        self.end_date_entry.insert(0, yesterday.strftime("%Y-%m-%d"))

        # 버튼을 담을 프레임 생성
        button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        button_frame.pack(fill="x")

        # 다운로드 버튼 (왼쪽)
        self.download_btn = ctk.CTkButton(button_frame, text="조회 및 다운로드",
                                        command=self.start_download,
                                        height=40,
                                        width=180)  # 너비 지정
        self.download_btn.pack(side="left", expand=True, padx=5)

        # 언발런스 등록 버튼 (오른쪽)
        self.unbalance_register_btn = ctk.CTkButton(button_frame, text="언발런스 등록",
                                                   command=self.start_unbalance_register,
                                                   height=40,
                                                   width=180)  # 너비 지정
        self.unbalance_register_btn.pack(side="right", expand=True, padx=5)

        # 상태 메시지
        self.status_label = ctk.CTkLabel(main_frame, text="대기 중...",
                                       text_color="white",
                                       font=("맑은 고딕", 12))
        self.status_label.pack(pady=(10,0))

        # 상세 상태 메시지
        self.detail_status_label = ctk.CTkLabel(main_frame, text="",
                                              text_color="gray",
                                              font=("맑은 고딕", 10))
        self.detail_status_label.pack(pady=(5,0))

        # 선택된 라인 목록 초기화
        self.selected_lines = []

    def create_line_checkboxes(self, line_type):
        """라인 체크박스 생성"""
        # 기존 체크박스 제거
        for widget in self.lines_frame.winfo_children():
            widget.destroy()
        
        # 새 체크박스 생성
        self.line_vars = {}
        
        # 설정에서 해당 라인타입의 라인 목록 가져오기
        if line_type in self.config["lines"]:
            lines = sorted(self.config["lines"][line_type])  # 정렬된 라인 목록 사용
            
            for line in lines:
                var = ctk.BooleanVar()
                self.line_vars[line] = var
                checkbox = ctk.CTkCheckBox(self.lines_frame, text=line, 
                                         variable=var,
                                         command=self.update_selected_lines,
                                         font=("맑은 고딕", 11,"bold"))  # 폰트 크기를 10으로 설정
                checkbox.pack(anchor="w", pady=2)

    def on_line_type_change(self, value):
        """라인타입 변경 시 체크박스 업데이트"""
        self.create_line_checkboxes(value)
        self.update_selected_lines()

    def update_selected_lines(self):
        """선택된 라인 목록 업데이트"""
        self.selected_lines = [line for line, var in self.line_vars.items() if var.get()]
        self.selected_lines_text.delete("1.0", "end")
        self.selected_lines_text.insert("1.0", "\n".join(sorted(self.selected_lines)))

    def show_settings(self):
        """설정 관리 창 열기"""
        settings_window = ctk.CTkToplevel(self)
        settings_window.title("설정 관리")
        settings_window.geometry("500x900")  # 높이를 900으로 증가
        settings_window.resizable(False, False)

        # 라인타입 관리
        line_type_frame = ctk.CTkFrame(settings_window)
        line_type_frame.pack(pady=20, padx=20, fill="x")

        ctk.CTkLabel(line_type_frame, text="라인타입 관리", font=("맑은 고딕", 14, "bold")).pack(pady=5)

        # 라인타입 목록 프레임
        list_frame = ctk.CTkFrame(line_type_frame)
        list_frame.pack(pady=5, fill="x")

        # 라인타입 목록
        list_content_frame = ctk.CTkFrame(list_frame)
        list_content_frame.pack(side="left", fill="x", expand=True)

        self.line_type_listbox = ctk.CTkTextbox(list_content_frame, height=100)
        self.line_type_listbox.pack(fill="x", expand=True)

        # 라인타입 버튼 프레임
        type_btn_frame = ctk.CTkFrame(list_frame)
        type_btn_frame.pack(side="right", padx=5)

        # 위/아래 이동 버튼
        move_up_type_btn = ctk.CTkButton(type_btn_frame, text="↑", width=40, height=25,
                                        command=lambda: self.move_item("type", "up"))
        move_up_type_btn.pack(pady=2)
        
        move_down_type_btn = ctk.CTkButton(type_btn_frame, text="↓", width=40, height=25,
                                          command=lambda: self.move_item("type", "down"))
        move_down_type_btn.pack(pady=2)
        
        # 삭제 버튼
        delete_type_btn = ctk.CTkButton(type_btn_frame, text="삭제", width=40, height=25,
                                       command=self.delete_selected_line_type,
                                       fg_color="red", hover_color="darkred")
        delete_type_btn.pack(pady=2)

        # 라인타입 추가
        add_type_frame = ctk.CTkFrame(line_type_frame)
        add_type_frame.pack(pady=5, fill="x")
        
        self.new_type_entry = ctk.CTkEntry(add_type_frame, placeholder_text="새 라인타입 입력")
        self.new_type_entry.pack(side="left", padx=5, expand=True, fill="x")
        
        add_type_btn = ctk.CTkButton(add_type_frame, text="추가", width=60,
                                    command=lambda: self.add_line_type(self.new_type_entry.get()))
        add_type_btn.pack(side="right", padx=5)

        # 라인 관리
        line_frame = ctk.CTkFrame(settings_window)
        line_frame.pack(pady=10, padx=20, fill="x")

        ctk.CTkLabel(line_frame, text="라인 관리", font=("맑은 고딕", 14, "bold")).pack(pady=5)

        # 라인타입 선택
        self.type_select = ctk.CTkComboBox(line_frame, values=list(self.config["line_types"]),
                                          command=self.update_line_list)
        self.type_select.pack(pady=5, fill="x")

        # 라인 목록 프레임
        line_list_frame = ctk.CTkFrame(line_frame)
        line_list_frame.pack(pady=5, fill="x")

        # 라인 목록
        list_content_frame = ctk.CTkFrame(line_list_frame)
        list_content_frame.pack(side="left", fill="x", expand=True)

        self.line_listbox = ctk.CTkTextbox(list_content_frame, height=200)
        self.line_listbox.pack(fill="x", expand=True)

        # 라인 버튼 프레임
        line_btn_frame = ctk.CTkFrame(line_list_frame)
        line_btn_frame.pack(side="right", padx=5)

        # 위/아래 이동 버튼
        move_up_line_btn = ctk.CTkButton(line_btn_frame, text="↑", width=40, height=25,
                                        command=lambda: self.move_item("line", "up"))
        move_up_line_btn.pack(pady=2)
        
        move_down_line_btn = ctk.CTkButton(line_btn_frame, text="↓", width=40, height=25,
                                          command=lambda: self.move_item("line", "down"))
        move_down_line_btn.pack(pady=2)
        
        # 삭제 버튼
        delete_line_btn = ctk.CTkButton(line_btn_frame, text="삭제", width=40, height=25,
                                       command=self.delete_selected_line,
                                       fg_color="red", hover_color="darkred")
        delete_line_btn.pack(pady=2)

        # 라인 추가
        add_line_frame = ctk.CTkFrame(line_frame)
        add_line_frame.pack(pady=5, fill="x")
        
        self.new_line_entry = ctk.CTkEntry(add_line_frame, placeholder_text="새 라인 입력")
        self.new_line_entry.pack(side="left", padx=5, expand=True, fill="x")
        
        add_line_btn = ctk.CTkButton(add_line_frame, text="추가", width=60,
                                    command=lambda: self.add_line(self.type_select.get(), 
                                                                self.new_line_entry.get()))
        add_line_btn.pack(side="right", padx=5)

        # 편집기 경로 프레임 추가
        editor_frame = ctk.CTkFrame(settings_window)
        editor_frame.pack(pady=10, padx=20, fill="x")
        
        ctk.CTkLabel(editor_frame, text="비가동 편집기 경로", font=("맑은 고딕", 14, "bold")).pack(pady=5)
        
        # 경로 표시 및 선택 프레임
        path_frame = ctk.CTkFrame(editor_frame)
        path_frame.pack(fill="x", pady=5)
        
        self.editor_path_var = tk.StringVar(value=self.config.get("editor_path", ""))
        self.editor_path_entry = ctk.CTkEntry(path_frame, textvariable=self.editor_path_var)
        self.editor_path_entry.pack(side="left", fill="x", expand=True, padx=(5,5))
        
        ctk.CTkButton(path_frame, text="찾아보기", width=60,
                     command=self.select_editor_path).pack(side="right", padx=5)

        # 저장 버튼과 업데이트 확인 버튼을 위한 프레임
        button_frame = ctk.CTkFrame(settings_window, fg_color="transparent")
        button_frame.pack(pady=5)

        # 저장 버튼
        save_btn = ctk.CTkButton(button_frame, text="저장", 
                                command=lambda: self.save_settings(settings_window),
                                height=40, width=150, 
                                font=("맑은 고딕", 14, "bold"))
        save_btn.pack(side="left", padx=5)

        # 업데이트 확인 버튼
        update_check_btn = ctk.CTkButton(button_frame, text="업데이트 확인",
                                       command=self.check_update_manually,
                                       height=40, width=150,
                                       font=("맑은 고딕", 14, "bold"))
        update_check_btn.pack(side="left", padx=5)
        
        # 비가동관리 버튼
        dt_history_btn = ctk.CTkButton(button_frame, text="비가동관리",
                                      command=self.open_dt_history,
                                      height=40, width=150,
                                      font=("맑은 고딕", 14, "bold"),
                                      fg_color="green", hover_color="darkgreen")
        dt_history_btn.pack(side="left", padx=5)
        
        # 만든이 정보 프레임
        info_frame = ctk.CTkFrame(settings_window)
        info_frame.pack(pady=10, padx=20, fill="x")
        
        # 구분선 추가
        separator = ctk.CTkFrame(info_frame, height=2, fg_color="gray50")
        separator.pack(fill="x", pady=10)
        
        # 만든이 정보
        creator_label = ctk.CTkLabel(info_frame, 
                                   text="MES WEB REPORT LOADER",
                                   font=("맑은 고딕", 16, "bold"))
        creator_label.pack(pady=(5,0))
        
        version_label = ctk.CTkLabel(info_frame,
                                   text=f"충주 제어기 생산팀 / Version {UpdateManager.current_version}",
                                   font=("맑은 고딕", 12))
        version_label.pack()
        
        copyright_label = ctk.CTkLabel(info_frame,
                                     text="© 2025 유하준. All rights reserved.",
                                     font=("맑은 고딕", 12))
        copyright_label.pack(pady=(0,5))
        
        contact_label = ctk.CTkLabel(info_frame,
                                   text="Contact: y@unitus.co.kr",
                                   font=("맑은 고딕", 12))
        contact_label.pack(pady=(0,10))

        # 초기 목록 업데이트
        self.update_line_type_list()
        if self.config["line_types"]:
            self.type_select.set(self.config["line_types"][0])
            self.update_line_list(self.config["line_types"][0])

    def select_editor_path(self):
        """편집기 파일 선택"""
        file_path = filedialog.askopenfilename(
            title="비가동 편집기 파일 선택",
            filetypes=[("Excel 파일", "*.xls;*.xlsx;*.xlsm"), ("모든 파일", "*.*")]
        )
        if file_path:
            self.editor_path_var.set(file_path)

    def save_settings(self, settings_window=None):
        """설정 저장 및 GUI 업데이트"""
        # 기존 설정 저장
        self.update_line_type_list()
        
        # 편집기 경로 저장
        self.config["editor_path"] = self.editor_path_var.get()
        
        self.save_config()
        
        # 메인 창의 라인타입과 라인 선택 콤보박스 업데이트
        self.line_type.configure(values=self.config["line_types"])
        if self.config["line_types"]:
            self.line_type.set(self.config["line_types"][0])
            self.create_line_checkboxes(self.config["line_types"][0])
        
        # 설정 창이 전달된 경우에만 닫기
        if settings_window:
            settings_window.destroy()
        
        messagebox.showinfo("알림", "설정이 저장되었습니다.")

    def open_downtime_editor(self):
        """비가동편집 Excel 파일 실행"""
        try:
            editor_path = self.config.get("editor_path", "")
            if not editor_path:
                messagebox.showerror("오류", "비가동 편집기 경로가 설정되지 않았습니다.\n설정에서 경로를 지정해주세요.")
                return
                
            if os.path.exists(editor_path):
                os.startfile(editor_path)
            else:
                messagebox.showerror("오류", f"설정된 경로에서 파일을 찾을 수 없습니다.\n경로: {editor_path}")
        except Exception as e:
            messagebox.showerror("오류", f"파일 실행 중 오류가 발생했습니다: {str(e)}")

    def move_item(self, list_type, direction):
        """항목 위/아래로 이동"""
        try:
            if list_type == "type":
                # 라인타입 이동
                current_line = self.line_type_listbox.get("insert linestart", "insert lineend").strip()
                if not current_line:
                    return
                
                idx = self.config["line_types"].index(current_line)
                if direction == "up" and idx > 0:
                    # 위로 이동
                    self.config["line_types"][idx], self.config["line_types"][idx-1] = \
                        self.config["line_types"][idx-1], self.config["line_types"][idx]
                    self.update_line_type_list()
                    # 커서 위치 이동
                    self.line_type_listbox.mark_set("insert", f"{idx}.0")
                    self.line_type_listbox.see(f"{idx}.0")
                elif direction == "down" and idx < len(self.config["line_types"]) - 1:
                    # 아래로 이동
                    self.config["line_types"][idx], self.config["line_types"][idx+1] = \
                        self.config["line_types"][idx+1], self.config["line_types"][idx]
                    self.update_line_type_list()
                    # 커서 위치 이동
                    self.line_type_listbox.mark_set("insert", f"{idx+2}.0")
                    self.line_type_listbox.see(f"{idx+2}.0")
            
            else:
                # 라인 이동
                current_type = self.type_select.get()
                if not current_type:
                    return
                
                current_line = self.line_listbox.get("insert linestart", "insert lineend").strip()
                if not current_line:
                    return
                
                lines = self.config["lines"][current_type]
                idx = lines.index(current_line)
                
                if direction == "up" and idx > 0:
                    # 위로 이동
                    lines[idx], lines[idx-1] = lines[idx-1], lines[idx]
                    self.update_line_list(current_type)
                    # 커서 위치 이동
                    self.line_listbox.mark_set("insert", f"{idx}.0")
                    self.line_listbox.see(f"{idx}.0")
                elif direction == "down" and idx < len(lines) - 1:
                    # 아래로 이동
                    lines[idx], lines[idx+1] = lines[idx+1], lines[idx]
                    self.update_line_list(current_type)
                    # 커서 위치 이동
                    self.line_listbox.mark_set("insert", f"{idx+2}.0")
                    self.line_listbox.see(f"{idx+2}.0")
                    
        except Exception as e:
            print(f"항목 이동 중 오류: {str(e)}")

    def update_line_type_list(self):
        """라인타입 목록 업데이트"""
        self.line_type_listbox.delete("1.0", "end")
        for line_type in self.config["line_types"]:  # sorted 제거
            self.line_type_listbox.insert("end", f"{line_type}\n")

    def update_line_list(self, selected_type=None):
        """선택된 라인타입의 라인 목록 업데이트"""
        if selected_type is None:
            selected_type = self.type_select.get()
        
        self.line_listbox.delete("1.0", "end")
        if selected_type in self.config["lines"]:
            for line in self.config["lines"][selected_type]:  # sorted 제거
                self.line_listbox.insert("end", f"{line}\n")

    def delete_selected_line_type(self):
        """선택된 라인타입 삭제"""
        try:
            # 현재 커서 위치의 라인 가져오기
            current_line = self.line_type_listbox.get("insert linestart", "insert lineend").strip()
            if current_line and current_line in self.config["line_types"]:
                # 라인타입과 관련된 라인들 삭제
                del self.config["lines"][current_line]
                self.config["line_types"].remove(current_line)
                
                # UI 업데이트
                self.update_line_type_list()
                self.type_select.configure(values=list(self.config["line_types"]))
                if self.config["line_types"]:
                    self.type_select.set(self.config["line_types"][0])
                    self.update_line_list(self.config["line_types"][0])
                else:
                    self.line_listbox.delete("1.0", "end")
        except Exception as e:
            print(f"라인타입 삭제 중 오류: {str(e)}")

    def delete_selected_line(self):
        """선택된 라인 삭제"""
        try:
            current_type = self.type_select.get()
            if not current_type:
                return
                
            # 현재 커서 위치의 라인 가져오기
            current_line = self.line_listbox.get("insert linestart", "insert lineend").strip()
            if current_line and current_line in self.config["lines"][current_type]:
                # 라인 삭제
                self.config["lines"][current_type].remove(current_line)
                # UI 업데이트
                self.update_line_list(current_type)
        except Exception as e:
            print(f"라인 삭제 중 오류: {str(e)}")

    def add_line_type(self, new_type):
        """라인타입 추가"""
        if new_type and new_type not in self.config["line_types"]:
            self.config["line_types"].append(new_type)
            self.config["lines"][new_type] = []
            self.update_line_type_list()
            self.type_select.configure(values=list(self.config["line_types"]))
            self.new_type_entry.delete(0, "end")

    def add_line(self, line_type, new_line):
        """라인 추가"""
        if line_type and new_line and new_line not in self.config["lines"][line_type]:
            self.config["lines"][line_type].append(new_line)
            self.update_line_list(line_type)
            self.new_line_entry.delete(0, "end")

    def update_status(self, main_status, detail_status="", color="yellow"):
        """상태 메시지 업데이트"""
        self.status_label.configure(text=main_status, text_color=color)
        self.detail_status_label.configure(text=detail_status, text_color="gray")
        self.update()  # GUI 즉시 업데이트

    def start_download(self):
        """다운로드 시작"""
        if not self.validate_inputs():
            return
            
        try:
            self.update_status("브라우저 시작", "크롬 브라우저를 초기화하고 있습니다...")
            
            # mes_common의 initialize_driver 사용
            driver = initialize_driver(show_browser=self.show_browser_var.get())
            
            if not self.show_browser_var.get():
                self.update_status("브라우저 시작", "헤드리스 모드로 실행 중...")
            
            # mes_common의 login_and_navigate 사용
            self.update_status("MES 로그인 및 WCT 시스템 진입 중...")
            login_success = login_and_navigate(
                driver, 
                self.id_entry.get(), 
                self.pw_entry.get(),
                log_callback=lambda msg: self.update_status("진행 중", msg)
            )
            
            if not login_success:
                raise Exception("로그인 또는 WCT 시스템 진입 실패")
            
            time.sleep(1.5)

            # 라인타입 선택
            self.update_status("라인 설정", "라인타입 [01] 제어기 선택 중...")
            line_type_dropdown = driver.find_element(By.ID, "ContentPlaceHolder1_ddeLineType_B-1")
            line_type_dropdown.click()
            time.sleep(0.5)
            
            line_type_checkbox = driver.find_element(By.ID, "ContentPlaceHolder1_ddeLineType_DDD_listLineType_0_01_D")
            line_type_checkbox.click()
            time.sleep(0.5)

            driver.find_element(By.TAG_NAME, "body").click()
            time.sleep(1)

            # 라인 선택
            self.update_status("선택된 라인 체크 중...")
            line_dropdown = driver.find_element(By.ID, "ContentPlaceHolder1_ddeLine_B-1")
            line_dropdown.click()
            time.sleep(0.5)

            try:
                filter_button = driver.find_element(By.ID, "ContentPlaceHolder1_ddeLine_DDD_listLine_0_LBSFEB")
                filter_button.click()
                time.sleep(0.5)

                selected_lines = [line for line, var in self.line_vars.items() if var.get()]
                total_lines = len(selected_lines)
                for idx, line_code in enumerate(selected_lines, 1):
                    try:
                        code = line_code.split("]")[0].strip("[")
                        self.update_status(f"라인 선택 중... ({idx}/{total_lines})")
                        
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
                            print(f"라인 선택 성공: {line_code}")
                            time.sleep(0.3)
                        else:
                            print(f"라인을 찾을 수 없음: {line_code}")
                            
                        filter_input.clear()
                        time.sleep(0.3)
                        
                    except Exception as e:
                        print(f"라인 선택 중 오류 ({code}): {str(e)}")
                        continue

                try:
                    filter_button.click()
                except:
                    pass
                
                time.sleep(0.3)
                webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                time.sleep(0.3)
                
            except Exception as e:
                print(f"라인 선택 처리 중 오류: {str(e)}")
                try:
                    webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                except:
                    pass

            time.sleep(0.3)

            # 기간 선택
            self.update_status("기간 옵션 선택 중...")
            period_dropdown = driver.find_element(By.ID, "ContentPlaceHolder1_ddl_Div_B-1")
            period_dropdown.click()
            time.sleep(0.3)

            try:
                period_option = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//td[contains(text(), '기간')]"))
                )
                period_option.click()
                time.sleep(0.5)
            except Exception as e:
                print(f"기간 옵션 선택 중 오류: {str(e)}")

            # 날짜 입력
            self.update_status("시작 날짜 입력 중...")
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
                
                start_date_value = self.start_date_entry.get()
                driver.execute_script("""
                    arguments[0].value = arguments[1];
                    arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                    arguments[0].dispatchEvent(new Event('blur', { bubbles: true }));
                    if(typeof ASPx !== 'undefined') {
                        ASPx.ETextChanged('ContentPlaceHolder1_dxDateEditStart');
                        ASPx.ELostFocus('ContentPlaceHolder1_dxDateEditStart');
                    }
                """, start_date, start_date_value)
                
                time.sleep(0.5)
            except Exception as e:
                print(f"시작 날짜 입력 중 오류: {str(e)}")

            self.update_status("종료 날짜 입력 중...")
            try:
                end_date = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "input.dxeEditArea_DevEx[id='ContentPlaceHolder1_dxDateEditEnd_I']"))
                )
                
                driver.execute_script("""
                    arguments[0].value = '';
                    arguments[0].dispatchEvent(new Event('focus'));
                """, end_date)
                time.sleep(0.1)
                
                end_date_value = self.end_date_entry.get()
                driver.execute_script("""
                    arguments[0].value = arguments[1];
                    arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                    arguments[0].dispatchEvent(new Event('blur', { bubbles: true }));
                    if(typeof ASPx !== 'undefined') {
                        ASPx.ETextChanged('ContentPlaceHolder1_dxDateEditEnd');
                        ASPx.ELostFocus('ContentPlaceHolder1_dxDateEditEnd');
                    }
                """, end_date, end_date_value)
                
                time.sleep(0.1)
            except Exception as e:
                print(f"종료 날짜 입력 중 오류: {str(e)}")

            # 조회 버튼 클릭
            self.update_status("데이터 조회", "조회 버튼 클릭...")
            search_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//td[contains(text(),'조회')]"))
            )
            search_button.click()

            # 데이터 로딩 대기
            try:
                self.update_status("데이터를 불러오는 중...")
                
                max_wait = 60
                start_time = time.time()
                
                while True:
                    elapsed_time = time.time() - start_time
                    if elapsed_time > max_wait:
                        raise TimeoutException(f"데이터 로딩 시간 초과 ({max_wait}초)")
                    
                    try:
                        no_data = driver.find_elements(By.XPATH, "//div[text()='No data']")
                        if not no_data:
                            print("데이터 로딩 완료")
                            break
                    except:
                        pass
                    
                    remaining_time = max_wait - int(elapsed_time)
                    self.update_status(f"데이터를 불러오는 중... (남은 시간: {remaining_time}초)")
                    time.sleep(1)
                
                time.sleep(1)
                
                # 선택된 라인 수에 따라 대기 시간 설정
                selected_rows = len(driver.find_elements(By.XPATH, "//tr[contains(@id, 'ContentPlaceHolder1_grdWork_DXDataRow')]"))
                wait_time_after_loading = 10 if selected_rows >= 5 else 5
                wait_time_after_format = 10 if selected_rows >= 5 else 3
                
                time.sleep(wait_time_after_loading)

                # 포멧 버튼 클릭 (파일 다운로드)
                self.update_status("포맷 파일 다운로드 중...")
                format_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_itbtnR_Format"))
                )
                format_button.click()
                time.sleep(wait_time_after_format)
                
                self.update_status("완료", "새로운 다운로드를 원하시면 다시 시작해주세요.", "lightgreen")
                return

            except TimeoutException as te:
                error_msg = str(te)
                print(f"시간 초과 오류: {error_msg}")
                self.update_status("오류", f"시간 초과: {error_msg}", "red")
            except Exception as e:
                print(f"데이터 로딩 중 오류: {str(e)}")
                self.update_status("오류", f"데이터 로딩 오류: {str(e)}", "red")

        except Exception as e:
            error_msg = str(e)
            print(f"상세 오류: {traceback.format_exc()}")
            self.update_status("오류", f"오류 발생: {error_msg}", "red")
        finally:
            if driver:
                try:
                    driver.quit()
                except:
                    pass

    def validate_inputs(self):
        """입력값 검증"""
        if not self.id_entry.get() or not self.pw_entry.get():
            messagebox.showerror("오류", "아이디와 비밀번호를 모두 입력해주세요.")
            return False
            
        # 로그인 정보 저장
        if self.save_login_var.get():
            self.config["login_info"]["save_login"] = True
            self.config["login_info"]["username"] = self.id_entry.get()
            self.config["login_info"]["password"] = self.pw_entry.get()
        else:
            self.config["login_info"]["save_login"] = False
            self.config["login_info"]["username"] = ""
            self.config["login_info"]["password"] = ""
        
        self.save_config()
        return True

    def show_uploader(self):
        """업로드 창 열기"""
        # 브라우저 표시 체크박스를 체크
        self.show_browser_var.set(True)
        # 업로더 창 열기
        uploader = FileUploader(self)
        uploader.grab_set()  # 모달 창으로 설정

    def toggle_save_login(self):
        """로그인 정보 저장 토글"""
        if self.save_login_var.get():
            # 현재 입력된 정보 저장
            self.config["login_info"]["save_login"] = True
            self.config["login_info"]["username"] = self.id_entry.get()
            self.config["login_info"]["password"] = self.pw_entry.get()
        else:
            # 저장된 정보 삭제
            self.config["login_info"]["save_login"] = False
            self.config["login_info"]["username"] = ""
            self.config["login_info"]["password"] = ""
        
        self.save_config()

    def show_unbalance_loader(self):
        """언발런스 로더 창 표시"""
        unbalance_window = UnbalanceLoader(self)
        unbalance_window.focus()

    def start_unbalance_register(self):
        """언발런스 등록 시작"""
        # 기존 다운로더의 라인 선택과 기간 정보를 가져옴
        selected_lines = []
        for line, var in self.line_vars.items():
            if var.get():
                selected_lines.append(line)

        if not selected_lines:
            messagebox.showwarning("경고", "라인을 선택해주세요.")
            return

        # 언발런스 로더 창 생성 및 정보 전달
        unbalance_window = UnbalanceLoader(self)
        unbalance_window.focus()
        
        # 선택된 라인 정보와 기간 정보 전달
        unbalance_window.set_initial_data(
            selected_lines=selected_lines,
            start_date=self.start_date_entry.get(),
            end_date=self.end_date_entry.get()
        )

    def run_split_mode(self):
        """분리모드 실행"""
        from mes_split import SplitMode
        app = SplitMode(self)
        app.grab_set()

    def open_dt_history(self):
        """비가동관리 창 열기"""
        try:
            from mes_DTHistory import DTHistory
            app = DTHistory(self)
            app.grab_set()
        except Exception as e:
            messagebox.showerror("오류", f"비가동관리 모듈 실행 중 오류가 발생했습니다.\n\n{str(e)}")

    def check_update_manually(self):
        """업데이트 확인 수동 실행"""
        try:
            from mes_updater import run_updater
            run_updater()
        except Exception as e:
            messagebox.showerror("오류", f"업데이트 모듈 실행 중 오류가 발생했습니다.\n\n{str(e)}")

if __name__ == "__main__":
    app = MESDownloader()
    app.mainloop()