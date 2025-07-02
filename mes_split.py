import customtkinter as ctk
from tkinter import messagebox, filedialog, ttk
import pandas as pd
import os
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import threading
import time
from mes_common import initialize_driver, login_and_navigate
from selenium.webdriver.common.keys import Keys
import traceback
import json
from selenium.common.exceptions import NoSuchElementException
from mes_register import MESRegister  # 새로운 모듈 import

class LoadingWindow(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("데이터 로딩 중")
        
        # 창 크기와 위치 설정
        window_width = 300
        window_height = 100
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")
        
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        
        # 프로그레스 바와 레이블 추가
        self.progress_label = ctk.CTkLabel(self, text="데이터를 불러오는 중...")
        self.progress_label.pack(pady=10)
        
        self.progress_bar = ctk.CTkProgressBar(self)
        self.progress_bar.pack(pady=10, padx=20, fill="x")
        self.progress_bar.set(0)
        
        # 진행 상태 업데이트를 위한 변수
        self.current_progress = 0
        
    def update_progress(self, value, text=None):
        """진행 상태 업데이트"""
        self.progress_bar.set(value)
        if text:
            self.progress_label.configure(text=text)
            
    def update_status(self, text):
        """상태 텍스트 업데이트"""
        self.progress_label.configure(text=text)

class SplitMode(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        
        # 현재 체크된 체크박스의 인덱스를 저장하는 변수 추가
        self.current_checked_idx = None
        
        # 비가동 관련 UI 요소들의 활성화 상태를 관리하는 변수 추가
        self.registration_enabled = False
        
        # 기본 창 설정
        self.title("MES 비가동 분리 도구")
        self.geometry("1500x800")
        self.resizable(True, True)
        
        # 모달 창으로 설정
        self.transient(self.parent)  # 부모 창과 연결
        self.grab_set()  # 모달 상태 설정
        
        # 다크 테마 설정
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        
        # Selenium 드라이버 초기화
        self.driver = None
        self.crawling_thread = None
        self.is_crawling = False
        self.is_connected = False
        
        # 체크박스 모니터링 관련 변수
        self.monitoring_thread = None
        self.stop_monitoring = False
        
        # 데이터 저장용 변수
        self.current_data = None
        self.line_vars = {}
        
        # 열 너비 설정
        self.column_widths = {
            0: 60,    # 체크박스
            1: 70,    # 상태
            2: 0,    # 등록구분
            3: 0,    # 진행현황
            4: 120,   # 생성일
            5: 120,   # 작업일
            6: 120,   # 라인명
            7: 150,   # 제품코드
            8: 100,   # 중단시간(분)
            9: 150,   # 시작시간
            10: 150,  # 종료시간
            11: 100,  # 공정
            12: 120   # 공정명
        }

        # 열 정렬 방식 설정 ('w': 왼쪽, 'center': 가운데, 'e': 오른쪽)
        self.column_alignments = {
            0: 'w',    # 체크박스
            1: 'center',        # 상태
            2: 'center',   # 등록구분
            3: 'center',   # 진행현황
            4: 'center',        # 생성일
            5: 'center',        # 작업일
            6: 'center',        # 라인명
            7: 'center',   # 제품코드
            8: 'center',        # 중단시간(분)
            9: 'center',   # 시작시간
            10: 'center',  # 종료시간
            11: 'center',  # 공정
            12: 'center'   # 공정명
        }

        # 데이터 행의 정렬 방식 (헤더와 별도로 설정 가능)
        self.data_alignments = {
            0: 'center',    # 체크박스
            1: 'center',    # 상태
            2: 'center',    # 등록구분
            3: 'center',    # 진행현황
            4: 'center',    # 생성일
            5: 'center',    # 작업일
            6: 'center',    # 라인명
            7: 'center',    # 제품코드
            8: 'center',    # 중단시간(분)
            9: 'center',    # 시작시간
            10: 'center',   # 종료시간
            11: 'center',   # 공정
            12: 'center'    # 공정명
        }
        
        # 컬럼 매핑 설정
        self.column_mapping = {
            "ContentPlaceHolder1_ASPxSplitter1_dxGrid2_col0": 0,    # 체크박스
            "ContentPlaceHolder1_ASPxSplitter1_dxGrid2_col31": 1,   # 상태
            "ContentPlaceHolder1_ASPxSplitter1_dxGrid2_col2": 2,    # 등록구분
            "ContentPlaceHolder1_ASPxSplitter1_dxGrid2_col3": 3,    # 진행현황
            "ContentPlaceHolder1_ASPxSplitter1_dxGrid2_col4": 4,    # 생성일
            "ContentPlaceHolder1_ASPxSplitter1_dxGrid2_col5": 5,    # 작업일
            "ContentPlaceHolder1_ASPxSplitter1_dxGrid2_col6": 6,    # 라인명
            "ContentPlaceHolder1_ASPxSplitter1_dxGrid2_col7": 7,    # 제품코드
            "ContentPlaceHolder1_ASPxSplitter1_dxGrid2_col8": 8,    # 중단시간(분)
            "ContentPlaceHolder1_ASPxSplitter1_dxGrid2_col9": 9,    # 시작시간
            "ContentPlaceHolder1_ASPxSplitter1_dxGrid2_col10": 10,  # 종료시간
            "ContentPlaceHolder1_ASPxSplitter1_dxGrid2_col11": 11,  # 공정
            "ContentPlaceHolder1_ASPxSplitter1_dxGrid2_col12": 12,  # 공정명
        }
        
        # 설정 파일 로드
        self.config_file = "mes_config.json"
        self.load_config()
        
        # UI 생성
        self.create_main_frame()
        
        # 창이 닫힐 때 정리
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        
    def load_config(self):
        """설정 파일 로드"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    self.config = json.load(f)
            else:
                self.config = self.parent.config
        except Exception as e:
            print(f"설정 로드 중 오류: {e}")
            self.config = self.parent.config
        
    def create_main_frame(self):
        """메인 프레임 생성"""
        # 전체 레이아웃을 위한 메인 프레임
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.pack(expand=True, fill="both")
        
        # 좌우 분할을 위한 프레임
        left_panel = ctk.CTkFrame(self.main_frame, fg_color="#2B2B2B", width=250)  # 너비 조정
        left_panel.pack(side="left", fill="y", padx=0, pady=0)
        left_panel.pack_propagate(False)  # 크기 고정
        
        # 오른쪽 패널을 상하로 분할
        right_panel = ctk.CTkFrame(self.main_frame, fg_color="#2B2B2B")
        right_panel.pack(side="right", fill="both", expand=True, padx=0, pady=0)
        
        # 상단 패널 (여백용)
        self.top_panel = ctk.CTkFrame(right_panel, fg_color="#1E1E1E", height=400)
        self.top_panel.pack(side="top", fill="both", expand=True, padx=1, pady=1)
        
        # 상단 패널을 좌우로 분할
        self.top_left_panel = ctk.CTkFrame(self.top_panel, fg_color="#2B2B2B")
        self.top_left_panel.pack(side="left", fill="both", expand=True, padx=1, pady=1)
        
        self.top_right_panel = ctk.CTkFrame(self.top_panel, fg_color="#2B2B2B")
        self.top_right_panel.pack(side="right", fill="both", expand=True, padx=1, pady=1)
        
        # 기존 GUI 요소들을 top_left_panel로 이동
        # 바로가기 버튼들을 감싸는 메인 프레임
        shortcuts_frame = ctk.CTkFrame(self.top_left_panel, fg_color="#2B2B2B")
        shortcuts_frame.pack(fill="x", padx=10, pady=5)

        # 필수비가동 바로가기 프레임 (왼쪽에 배치)
        essential_shortcut_frame = ctk.CTkFrame(shortcuts_frame, fg_color="#1E1E1E")
        essential_shortcut_frame.pack(side="left", fill="both", expand=True, padx=(0,5), pady=5)
        
        # 필수비가동 바로가기 내부 프레임
        radio_group1 = ctk.CTkFrame(essential_shortcut_frame, fg_color="#1E1E1E")
        radio_group1.pack(fill="x", padx=10, pady=5)  # expand=True 제거하고 fill="x"로 변경
        
        radio_label1 = ctk.CTkLabel(radio_group1, text="필수비가동 바로가기", font=("Arial", 11))
        radio_label1.pack(anchor="w", pady=(0,5))
        
        # 필수비가동 라디오 버튼 변수 추가
        self.essential_var = ctk.StringVar(value="")
        
        radio_buttons1 = [
            ("차종/기종교체", "model_change"),
            ("치공구 교체", "tool_change"),
            ("부자재 교체", "material_change"),
            ("마스터 점검", "master_check"),
            ("일상 점검", "daily_check"),
            ("재공 채움", "wip_fill")
        ]
        
        # 2열 레이아웃을 위한 컨테이너 프레임
        buttons_container1 = ctk.CTkFrame(radio_group1, fg_color="#1E1E1E")
        buttons_container1.pack(fill="x")  # expand=True 제거
        
        # 첫 번째 열 (3개)
        col1_frame1 = ctk.CTkFrame(buttons_container1, fg_color="#1E1E1E")
        col1_frame1.pack(side="left", fill="x", expand=True)
        
        # 필수비가동 라디오 버튼들을 저장할 리스트 추가
        self.essential_radios = []
        
        for text, value in radio_buttons1[:3]:
            radio = ctk.CTkRadioButton(col1_frame1, 
                                     text=text, 
                                     value=value, 
                                     variable=self.essential_var,
                                     command=lambda v=value: self.on_essential_select(v),
                                     font=("Arial", 11),
                                     state="disabled")  # 초기 상태 비활성화
            radio.pack(anchor="w", pady=2)
            self.essential_radios.append(radio)
            
        # 두 번째 열 (3개)
        col2_frame1 = ctk.CTkFrame(buttons_container1, fg_color="#1E1E1E")
        col2_frame1.pack(side="left", fill="x", expand=True)
        
        for text, value in radio_buttons1[3:]:
            radio = ctk.CTkRadioButton(col2_frame1, 
                                     text=text, 
                                     value=value, 
                                     variable=self.essential_var,
                                     command=lambda v=value: self.on_essential_select(v),
                                     font=("Arial", 11),
                                     state="disabled")  # 초기 상태 비활성화
            radio.pack(anchor="w", pady=2)
            self.essential_radios.append(radio)

        # 일반비가동 바로가기 프레임 (가운데 배치)
        normal_shortcut_frame = ctk.CTkFrame(shortcuts_frame, fg_color="#1E1E1E")
        normal_shortcut_frame.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        
        # 일반비가동 바로가기 내부 프레임
        radio_group2 = ctk.CTkFrame(normal_shortcut_frame, fg_color="#1E1E1E")
        radio_group2.pack(fill="x", padx=10, pady=5)  # expand=True 제거하고 fill="x"로 변경
        
        radio_label2 = ctk.CTkLabel(radio_group2, text="일반비가동 바로가기", font=("Arial", 11))
        radio_label2.pack(anchor="w", pady=(0,5))
        
        # 일반비가동 라디오 버튼 변수 추가
        self.normal_var = ctk.StringVar(value="")
        
        radio_buttons2 = [
            ("작업지연", "work_delay"),
            ("기타", "etc"),
            ("실험공정", "experiment"),
            ("품질문제(공정)", "quality_process"),
            ("품질문제(자재)", "quality_material"),
            ("설비고장(양산)", "equipment_failure_mass"),
            ("설비고장(신규)", "equipment_failure_new"),
            ("자재결품", "material_shortage"),
            ("조건조정(일반)", "condition_adjust_normal")
        ]
        
        # 2열 레이아웃을 위한 컨테이너 프레임
        buttons_container2 = ctk.CTkFrame(radio_group2, fg_color="#1E1E1E")
        buttons_container2.pack(fill="x")
        
        # 첫 번째 열 (5개)
        col1_frame2 = ctk.CTkFrame(buttons_container2, fg_color="#1E1E1E")
        col1_frame2.pack(side="left", fill="x", expand=True, anchor="n")  # anchor="n" 추가
        
        # 일반비가동 라디오 버튼들을 저장할 리스트 추가
        self.normal_radios = []
        
        for text, value in radio_buttons2[:5]:
            radio = ctk.CTkRadioButton(col1_frame2, 
                                     text=text, 
                                     value=value, 
                                     variable=self.normal_var,
                                     command=lambda v=value: self.on_normal_select(v),
                                     font=("Arial", 11),
                                     state="disabled")  # 초기 상태 비활성화
            radio.pack(anchor="w", pady=2)
            self.normal_radios.append(radio)
            
        # 두 번째 열 (4개)
        col2_frame2 = ctk.CTkFrame(buttons_container2, fg_color="#1E1E1E")
        col2_frame2.pack(side="left", fill="x", expand=True, anchor="n")  # anchor="n" 추가
        
        # 더미 프레임으로 4개 버튼을 위로 정렬
        dummy_frame2 = ctk.CTkFrame(col2_frame2, fg_color="#1E1E1E", height=30)
        dummy_frame2.pack(side="bottom", fill="x", expand=True)
        
        for text, value in radio_buttons2[5:]:
            radio = ctk.CTkRadioButton(col2_frame2, 
                                     text=text, 
                                     value=value, 
                                     variable=self.normal_var,
                                     command=lambda v=value: self.on_normal_select(v),
                                     font=("Arial", 11),
                                     state="disabled")  # 초기 상태 비활성화
            radio.pack(anchor="w", pady=2)
            self.normal_radios.append(radio)

        # 원인/조치사항 입력을 위한 컨테이너 프레임 (오른쪽 배치)
        input_container = ctk.CTkFrame(shortcuts_frame, fg_color="#1E1E1E")
        input_container.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        
        # 공정명 선택 프레임
        process_frame = ctk.CTkFrame(input_container, fg_color="#1E1E1E")
        process_frame.pack(fill="x", pady=(0,2))
        
        process_label = ctk.CTkLabel(process_frame, text="공정명", font=("Arial", 11))
        process_label.pack(anchor="w", padx=5, pady=(0,2))
        
        # 공정명 드롭다운 초기 상태 비활성화
        self.process_dropdown = ctk.CTkComboBox(process_frame, 
                                              values=["공정1", "공정2", "공정3"],  # 임시 값들
                                              font=("Arial", 11),
                                              height=28,
                                              width=250,
                                              state="disabled")  # 초기 상태 비활성화
        self.process_dropdown.pack(side="left", padx=5)
        
        # 공정불러오기 버튼 추가
        self.load_process_btn = ctk.CTkButton(
            process_frame,
            text="공정불러오기",
            font=("Arial", 11),
            width=100,
            height=28,
            command=self.load_process_list,
            state="disabled"  # 초기 상태 비활성화
        )
        self.load_process_btn.pack(side="left", padx=5)
        
        # 원인 입력 프레임
        cause_frame = ctk.CTkFrame(input_container, fg_color="#1E1E1E")
        cause_frame.pack(fill="x", pady=(2,0))
        
        cause_label = ctk.CTkLabel(cause_frame, text="원        인", font=("Arial", 11))
        cause_label.pack(anchor="w", padx=5, pady=(0,2))
        
        self.cause_textbox = ctk.CTkEntry(cause_frame, font=("Arial", 11), height=28, width=350)
        self.cause_textbox.pack(anchor="w", padx=5)
        
        # 조치사항 입력 프레임
        action_frame = ctk.CTkFrame(input_container, fg_color="#1E1E1E")
        action_frame.pack(fill="x", pady=(2,0))
        
        action_label = ctk.CTkLabel(action_frame, text="조치사항", font=("Arial", 11))
        action_label.pack(anchor="w", padx=5, pady=(0,2))
        
        self.action_textbox = ctk.CTkEntry(action_frame, font=("Arial", 11), height=28, width=350)
        self.action_textbox.pack(anchor="w", padx=5)

        # 등록 버튼 프레임
        register_frame = ctk.CTkFrame(input_container, fg_color="#1E1E1E")
        register_frame.pack(fill="x", pady=(5,0))

        # 등록 버튼 초기 상태 비활성화
        self.register_btn = ctk.CTkButton(register_frame, 
                                        text="등록",
                                        command=self.register_data,
                                        fg_color="#3B8ED0",
                                        font=("Arial", 11),
                                        width=80,
                                        height=28,
                                        state="disabled")  # 초기 상태 비활성화
        self.register_btn.pack(side="right", padx=5)

        
        # 하단 패널 (기존 데이터 그리드용)
        self.bottom_panel = ctk.CTkFrame(right_panel, fg_color="#1E1E1E", height=400)
        self.bottom_panel.pack(side="bottom", fill="both", expand=True, padx=1, pady=1)
        
        # 왼쪽 패널 구성
        # 상단 연결 버튼
        header_frame = ctk.CTkFrame(left_panel, fg_color="#2B2B2B")
        header_frame.pack(fill="x", padx=10, pady=(10,5))
        
        self.connection_label = ctk.CTkLabel(header_frame, text="⚫", font=("Arial", 16))
        self.connection_label.configure(text_color="red")
        self.connection_label.pack(side="left", padx=5)
        
        self.connect_btn = ctk.CTkButton(header_frame, text="MES 연결",
                                       command=self.connect_to_mes,
                                       fg_color="#3B8ED0",
                                       font=("Arial", 11),
                                       width=80,
                                       height=28)
        self.connect_btn.pack(side="left", padx=5)

        # 브라우저 활성화 체크박스 추가
        self.chrome_visible_var = ctk.BooleanVar(value=False)  # 기본값 True로 설정
        self.chrome_visible_checkbox = ctk.CTkCheckBox(header_frame,
                                                     text="브라우저 표시",
                                                     variable=self.chrome_visible_var,
                                                     font=("Arial", 11),
                                                     height=28,
                                                     fg_color="#3B8ED0",
                                                     border_color="#666666")
        self.chrome_visible_checkbox.pack(side="left", padx=5)
        
        # 구분선 추가
        separator1 = ctk.CTkFrame(left_panel, height=1, fg_color="#404040")
        separator1.pack(fill="x", padx=10, pady=5)
        
        # 라인타입 선택
        line_type_frame = ctk.CTkFrame(left_panel, fg_color="#2B2B2B")
        line_type_frame.pack(fill="x", padx=10, pady=5)
        
        line_type_label = ctk.CTkLabel(line_type_frame, text="라인타입:", font=("Arial", 11))
        line_type_label.pack(side="left")
        
        self.line_type_var = ctk.StringVar(value=self.parent.line_type_var.get() if hasattr(self.parent, 'line_type_var') else self.config["line_types"][0])
        self.line_type_combo = ctk.CTkComboBox(line_type_frame, 
                                              variable=self.line_type_var,
                                              values=self.config["line_types"],
                                              command=self.on_line_type_change,
                                              font=("Arial", 11),
                                              height=28,
                                              dropdown_font=("Arial", 11))
        self.line_type_combo.pack(side="left", padx=5, fill="x", expand=True)
        
        # 구분선 추가
        separator2 = ctk.CTkFrame(left_panel, height=1, fg_color="#404040")
        separator2.pack(fill="x", padx=10, pady=5)
        
        # 라인 선택 영역
        line_label_frame = ctk.CTkFrame(left_panel, fg_color="#2B2B2B")
        line_label_frame.pack(fill="x", padx=10, pady=(5,0))
        
        ctk.CTkLabel(line_label_frame, text="라인 선택", anchor="w", font=("Arial", 11)).pack(side="left")
        
        # 라인 목록 스크롤 영역
        self.line_scrolled_frame = ctk.CTkScrollableFrame(left_panel, fg_color="#1E1E1E", height=50)
        self.line_scrolled_frame.pack(fill="x", padx=10, pady=5)
        
        # 구분선 추가
        separator3 = ctk.CTkFrame(left_panel, height=1, fg_color="#404040")
        separator3.pack(fill="x", padx=10, pady=5)
        
        # 날짜 선택 영역
        date_frame = ctk.CTkFrame(left_panel, fg_color="#2B2B2B")
        date_frame.pack(fill="x", padx=10, pady=5)
        
        # 시작 날짜
        start_date_frame = ctk.CTkFrame(date_frame, fg_color="#2B2B2B")
        start_date_frame.pack(fill="x", pady=2)
        
        ctk.CTkLabel(start_date_frame, text="시작 날짜:", anchor="w", font=("Arial", 11)).pack(side="left", padx=(0,5))
        self.start_date_var = ctk.StringVar(value=self.parent.start_date_entry.get() if hasattr(self.parent, 'start_date_entry') else datetime.now().strftime("%Y-%m-%d"))
        self.start_date_entry = ctk.CTkEntry(start_date_frame, 
                                           textvariable=self.start_date_var, 
                                           font=("Arial", 11),
                                           height=28)
        self.start_date_entry.pack(side="left", fill="x", expand=True)
        
        # 종료 날짜
        end_date_frame = ctk.CTkFrame(date_frame, fg_color="#2B2B2B")
        end_date_frame.pack(fill="x", pady=2)
        
        ctk.CTkLabel(end_date_frame, text="종료 날짜:", anchor="w", font=("Arial", 11)).pack(side="left", padx=(0,5))
        self.end_date_var = ctk.StringVar(value=self.parent.end_date_entry.get() if hasattr(self.parent, 'end_date_entry') else datetime.now().strftime("%Y-%m-%d"))
        self.end_date_entry = ctk.CTkEntry(end_date_frame, 
                                         textvariable=self.end_date_var, 
                                         font=("Arial", 11),
                                         height=28)
        self.end_date_entry.pack(side="left", fill="x", expand=True)
        
        # 구분선 추가
        separator4 = ctk.CTkFrame(left_panel, height=1, fg_color="#404040")
        separator4.pack(fill="x", padx=10, pady=5)
        
        # 데이터 조회 버튼
        self.search_btn = ctk.CTkButton(left_panel, 
                                      text="데이터 조회",
                                      command=self.start_data_search,
                                      fg_color="#3B8ED0",
                                      font=("Arial", 11),
                                      height=32)
        self.search_btn.pack(pady=10, padx=10, fill="x")
        self.search_btn.configure(state="disabled")
        
        # 구분선 추가
        separator5 = ctk.CTkFrame(left_panel, height=1, fg_color="#404040")
        separator5.pack(fill="x", padx=10, pady=5)
        
        # 시간 분리 프레임
        split_time_frame = ctk.CTkFrame(left_panel, fg_color="#2B2B2B")
        split_time_frame.pack(fill="x", padx=10, pady=5)
        
        # 분리시간 입력 프레임
        split_input_frame = ctk.CTkFrame(split_time_frame, fg_color="#2B2B2B")
        split_input_frame.pack(fill="x", pady=2)
        
        ctk.CTkLabel(split_input_frame, text="분리시간(분):", anchor="w", font=("Arial", 11)).pack(side="left", padx=(0,5))
        self.split_time_var = ctk.StringVar(value="")
        self.split_time_entry = ctk.CTkEntry(split_input_frame, 
                                           textvariable=self.split_time_var,
                                           font=("Arial", 11),
                                           height=28)
        self.split_time_entry.pack(side="left", fill="x", expand=True)
        
        # 시간분리 버튼
        self.split_time_btn = ctk.CTkButton(split_time_frame,
                                          text="시간분리",
                                          command=self.split_time,
                                          fg_color="#3B8ED0",
                                          font=("Arial", 11),
                                          height=32)
        self.split_time_btn.pack(pady=(5,0), fill="x")
        self.split_time_btn.configure(state="disabled")
        
        # 구분선 추가
        separator6 = ctk.CTkFrame(left_panel, height=1, fg_color="#404040")
        separator6.pack(fill="x", padx=10, pady=5)
        
        # 로그창 프레임
        log_frame = ctk.CTkFrame(left_panel, fg_color="#1E1E1E")
        log_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        # 로그창 레이블
        log_label = ctk.CTkLabel(log_frame, text="데이터 로그", anchor="w", font=("Arial", 11))
        log_label.pack(fill="x", padx=5, pady=2)
        
        # 로그창 텍스트 영역
        self.log_text = ctk.CTkTextbox(log_frame, 
                                     font=("Consolas", 10),
                                     fg_color="#2B2B2B",
                                     text_color="#FFFFFF",
                                     height=150)
        self.log_text.pack(fill="both", expand=True, padx=2, pady=2)
        
        # 라인 목록 생성
        self.create_line_checkboxes(self.line_type_var.get())
        
        # 데이터 표시를 위한 프레임
        self.grid_frame = ctk.CTkFrame(self.bottom_panel, fg_color="#1E1E1E")
        self.grid_frame.pack(expand=True, fill="both", padx=1, pady=1)

        # 헤더 프레임 (고정)
        self.header_frame = ctk.CTkFrame(self.grid_frame, fg_color="#2B2B2B")
        self.header_frame.pack(fill="x", padx=1, pady=0)

        # 데이터 영역 프레임 (스크롤)
        self.data_container = ctk.CTkFrame(self.grid_frame, fg_color="#1E1E1E")
        self.data_container.pack(expand=True, fill="both", padx=1, pady=1)
        
        # 스크롤바가 있는 캔버스 생성
        self.canvas = ctk.CTkCanvas(self.data_container, bg="#1E1E1E", highlightthickness=0)
        self.scrollbar_y = ttk.Scrollbar(self.data_container, orient="vertical", command=self.canvas.yview)
        self.scrollbar_x = ttk.Scrollbar(self.data_container, orient="horizontal", command=self.canvas.xview)
        
        # 실제 데이터를 표시할 프레임
        self.data_frame = ctk.CTkFrame(self.canvas, fg_color="#1E1E1E")
        
        # 스크롤바 설정
        self.canvas.configure(yscrollcommand=self.scrollbar_y.set, xscrollcommand=self.scrollbar_x.set)
        
        # 스크롤바 배치
        self.scrollbar_y.pack(side="right", fill="y")
        self.scrollbar_x.pack(side="bottom", fill="x")
        self.canvas.pack(side="left", fill="both", expand=True)
        
        # 캔버스에 데이터 프레임 추가
        self.canvas_frame = self.canvas.create_window((0, 0), window=self.data_frame, anchor="nw")
        
        # 이벤트 바인딩
        self.canvas.bind("<Configure>", self.on_canvas_configure)
        self.data_frame.bind("<Configure>", self.on_frame_configure)
        self.canvas.bind_all("<MouseWheel>", self.on_mousewheel)
        
        # 헤더 생성
        self.create_headers()
        
        # 공정명 라벨
        process_label = ctk.CTkLabel(self.main_frame, text="공정명:")
        process_label.grid(row=current_row, column=0, padx=5, pady=5, sticky="e")
        
        # 공정명 드롭다운
        self.process_dropdown = ctk.CTkComboBox(self.main_frame, values=[], width=200)
        self.process_dropdown.grid(row=current_row, column=1, padx=5, pady=5, sticky="w")
        
        current_row += 1
        
    def create_headers(self):
        """헤더 생성"""
        # 기존 헤더 제거
        for widget in self.header_frame.winfo_children():
            widget.destroy()

        # 헤더 스타일 (기본)
        style = {
            "fg_color": "#2B2B2B",
            "text_color": "white",
            "height": 25,
            "font": ("Arial", 11, "bold")
        }

        # 체크박스 열 헤더
        if self.column_widths[0] > 0:
            checkbox_header = ctk.CTkLabel(
                self.header_frame,
                text="선택",
                width=self.column_widths[0],
                anchor=self.column_alignments[0],
                padx=10 if self.column_alignments[0] == "w" else 0,
                **style
            )
            checkbox_header.grid(row=0, column=0, sticky="nsew", padx=1, pady=1)

        # 나머지 헤더
        headers = ["상태", "등록구분", "진행현황", "생성일", "작업일", "라인명", 
                  "제품코드", "중단시간(분)", "시작시간", "종료시간", "공정", "공정명"]
        
        for idx, header in enumerate(headers, 1):
            if self.column_widths[idx] > 0:  # 너비가 0보다 큰 경우만 표시
                label = ctk.CTkLabel(
                    self.header_frame,
                    text=header,
                    width=self.column_widths[idx],
                    anchor=self.column_alignments[idx],
                    padx=10 if self.column_alignments[idx] == "w" else 0,
                    **style
                )
                label.grid(row=0, column=idx, sticky="nsew", padx=1, pady=1)
                self.header_frame.grid_columnconfigure(idx, weight=1)

    def on_frame_configure(self, event=None):
        """데이터 프레임 크기가 변경될 때 호출"""
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        
    def on_canvas_configure(self, event):
        """캔버스 크기가 변경될 때 호출"""
        # 데이터 프레임의 너비를 캔버스에 맞춤
        self.canvas.itemconfig(self.canvas.find_withtag("all")[0], width=event.width)
        
    def on_mousewheel(self, event):
        """마우스 휠 스크롤 처리"""
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
    def clear_data_grid(self):
        """데이터 그리드 초기화"""
        # 헤더를 제외한 모든 위젯 제거
        for widget in self.data_frame.grid_slaves():
            if int(widget.grid_info()["row"]) > 0:
                widget.destroy()
        
        # 체크박스 변수들 초기화
        for attr in list(self.__dict__.keys()):
            if attr.startswith("checkbox_var_"):
                delattr(self, attr)
        
        # 현재 체크된 인덱스 초기화
        self.current_checked_idx = None
        # 시간분리 버튼 비활성화
        self.split_time_btn.configure(state="disabled")

    def add_data_row(self, row_idx, data, bg_color=None):
        """데이터 행 추가"""
        # 기본 스타일
        style = {
            "fg_color": "transparent",
            "text_color": "white",
            "height": 22,
            "font": ("Arial", 10)
        }
            
        # 각 셀 생성
        visible_col = 0  # 실제로 표시되는 열의 인덱스
        for col, value in enumerate(data):
            if self.column_widths[col] > 0:  # 너비가 0보다 큰 경우만 표시
                if col == 0:  # 체크박스 열
                    checkbox_var = ctk.BooleanVar()
                    setattr(self, f"checkbox_var_{row_idx}", checkbox_var)
                    
                    cell = ctk.CTkCheckBox(
                        self.data_frame,
                        text="",
                        variable=checkbox_var,
                        width=self.column_widths[col],
                        height=16,
                        fg_color="#3B8ED0",
                        border_color="#666666"
                    )
                    
                    # 체크박스 클릭 이벤트 연결
                    cell.configure(command=lambda idx=row_idx: self.on_checkbox_click(idx))
                    
                else:  # 일반 데이터 열
                    cell = ctk.CTkLabel(
                        self.data_frame,
                        text=str(value),
                        width=self.column_widths[col],
                        anchor=self.data_alignments[col],
                        padx=10 if self.data_alignments[col] == "w" else 0,
                        **style
                    )
                cell.grid(row=row_idx, column=visible_col, sticky="nsew", padx=1, pady=0)
                visible_col += 1  # 표시된 열만 카운트
            
        # 열 가중치 설정 (표시되는 열만)
        visible_columns = sum(1 for width in self.column_widths.values() if width > 0)
        for i in range(visible_columns):
            self.data_frame.grid_columnconfigure(i, weight=1)

    def on_checkbox_click(self, row_idx):
        """체크박스 클릭 이벤트 핸들러"""
        try:
            # 체크박스 상태 가져오기
            checkbox_var = getattr(self, f"checkbox_var_{row_idx}")
            is_checked = checkbox_var.get()
            
            # 현재 체크된 체크박스가 있고, 다른 체크박스를 체크하려는 경우
            if is_checked and self.current_checked_idx is not None and self.current_checked_idx != row_idx:
                # 이전에 체크된 체크박스 해제
                prev_checkbox_var = getattr(self, f"checkbox_var_{self.current_checked_idx}")
                prev_checkbox_var.set(False)
                
                # UI 비활성화
                self.disable_registration_controls()
                
                # 웹페이지의 이전 체크박스 해제
                try:
                    prev_row_id = f"ContentPlaceHolder1_ASPxSplitter1_dxGrid2_DXDataRow{self.current_checked_idx}"
                    prev_row = self.driver.find_element(By.ID, prev_row_id)
                    prev_checkbox_cell = prev_row.find_elements(By.TAG_NAME, "td")[0]
                    if prev_checkbox_cell:
                        self.driver.execute_script("""
                            var cell = arguments[0];
                            var input = cell.querySelector('input.dxAIFFE');
                            if (input && input.value === 'C') {
                                cell.click();
                            }
                        """, prev_checkbox_cell)
                except Exception as e:
                    print(f"이전 체크박스 해제 중 오류: {str(e)}")
            
            # 현재 체크박스 상태 업데이트
            if is_checked:
                self.current_checked_idx = row_idx
                self.split_time_btn.configure(state="normal")  # 시간분리 버튼 활성화
                
                # 체크된 행의 라인명 값 출력
                try:
                    cells = self.data_frame.grid_slaves(row=row_idx, column=6)
                    if cells:
                        line_name = cells[0].cget("text").strip()
                        print(f"\n=== 체크된 행 정보 ===")
                        print(f"행 번호: {row_idx}")
                        print(f"라인명: {line_name}")
                except Exception as e:
                    print(f"라인명 값 출력 중 오류: {str(e)}")
                
            elif self.current_checked_idx == row_idx:
                self.current_checked_idx = None
                self.split_time_btn.configure(state="disabled")  # 시간분리 버튼 비활성화
                # UI 비활성화
                self.disable_registration_controls()
            
            # 웹페이지의 체크박스 찾기
            row_id = f"ContentPlaceHolder1_ASPxSplitter1_dxGrid2_DXDataRow{row_idx}"
            row = self.driver.find_element(By.ID, row_id)
            checkbox_cell = row.find_elements(By.TAG_NAME, "td")[0]
            
            # 실제 클릭 이벤트 발생
            self.driver.execute_script("arguments[0].click();", checkbox_cell)
            
            # UI 상태 동기화
            self.driver.execute_script("""
                var cell = arguments[0];
                var isChecked = arguments[1];
                
                // input 값 변경
                var input = cell.querySelector('input.dxAIFFE');
                if (input) {
                    input.value = isChecked ? 'C' : 'U';
                }
                
                // span 클래스 변경
                var span = cell.querySelector('span.dxICheckBox');
                if (span) {
                    span.className = isChecked ? 
                        'dxWeb_edtCheckBoxChecked dxICheckBox dxichSys dx-not-acc' : 
                        'dxWeb_edtCheckBoxUnchecked dxICheckBox dxichSys dx-not-acc';
                }
            """, checkbox_cell, is_checked)
            
        except Exception as e:
            print(f"체크박스 상태 변경 중 오류: {str(e)}")

    def update_data_grid(self, data):
        """데이터 그리드 업데이트"""
        self.clear_data_grid()
        for idx, row in enumerate(data):
            self.add_data_row(idx, row)

    def update_status(self, message, color="white"):
        """상태 표시 업데이트"""
        if color == "red":
            self.connection_label.configure(text="⚫", text_color="red")
        elif color == "green":
            self.connection_label.configure(text="⚫", text_color="green")
        elif color == "green":
            self.connection_label.configure(text="⚫", text_color="green")

    def connect_to_mes(self):
        """MES 연결 버튼 클릭 핸들러"""
        if not self.is_connected:
            self.update_status("연결 중...", "yellow")
            threading.Thread(target=self.establish_connection, daemon=True).start()
        else:
            if self.driver:
                self.driver.quit()
                self.driver = None
            self.is_connected = False
            self.update_status("연결 해제됨", "red")
            self.search_btn.configure(state="disabled")

    def login_and_navigate_split_mode(self, driver, username, password):
        """분리모드 전용 로그인 및 페이지 이동"""
        try:
            self.log("\n=== MES 로그인 시작 ===", "INFO")
            # MES 시스템 접속
            self.log("1. MES 시스템 접속 중...", "INFO")
            driver.get("https://mescj.unitus.co.kr/Login.aspx")
            time.sleep(2)
            self.log("   ✓ MES 시스템 접속 완료", "SUCCESS")

            # 혹시 모를 알림창 처리
            try:
                alert = driver.switch_to.alert
                alert.accept()
                self.log("   → 알림창 처리됨", "INFO")
            except:
                pass

            # 아이디/비밀번호 입력
            self.log("\n2. 로그인 정보 입력 중...", "INFO")
            id_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "txtUserID"))
            )
            id_input.clear()
            time.sleep(0.1)
            for char in username.upper():  # 대문자로 변환
                id_input.send_keys(char)
                time.sleep(0.1)
            time.sleep(0.2)
            self.log("   ✓ 아이디 입력 완료", "SUCCESS")
            
            pw_input = driver.find_element(By.ID, "txtPassword")
            pw_input.clear()
            time.sleep(0.1)
            for char in password:
                pw_input.send_keys(char)
                time.sleep(0.1)
            time.sleep(0.2)
            self.log("   ✓ 비밀번호 입력 완료", "SUCCESS")
            
            # 로그인 버튼 클릭
            self.log("\n3. 로그인 시도 중...", "INFO")
            login_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "btnLogin"))
            )
            login_button.click()
            time.sleep(0.5)
            self.log("   ✓ 로그인 버튼 클릭 완료", "SUCCESS")

            # 로그인 후 알림창 처리
            try:
                alert = driver.switch_to.alert
                alert_text = alert.text
                self.log(f"   → 로그인 알림: {alert_text}", "WARNING")
                alert.accept()
                if "비밀번호" in alert_text or "아이디" in alert_text:
                    self.log("   !!! 로그인 실패", "ERROR")
                    return False
            except:
                pass

            # WCT 시스템으로 이동
            self.log("\n4. WCT 시스템 이동 중...", "INFO")
            wct_link = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//a[contains(text(), 'New WCT System')]"))
            )
            driver.execute_script("arguments[0].scrollIntoView(true); arguments[0].click();", wct_link)
            time.sleep(1)
            self.log("   ✓ WCT 시스템 링크 클릭 완료", "SUCCESS")

            # 새로 열린 창으로 전환
            self.log("\n5. 새 창으로 전환 중...", "INFO")
            WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) > 1)
            driver.switch_to.window(driver.window_handles[-1])
            self.log("   ✓ 새 창 전환 완료", "SUCCESS")

            # WCT 시스템 알림창 처리
            try:
                alert = driver.switch_to.alert
                alert.accept()
                self.log("   → WCT 시스템 알림창 처리됨", "INFO")
            except:
                pass

            # 닫기 버튼 클릭
            self.log("\n6. 시작 페이지 닫기 중...", "INFO")
            time.sleep(1)
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
            self.log("   ✓ 시작 페이지 닫기 완료", "SUCCESS")

            # 분리모드 페이지로 이동
            self.log("\n7. 분리모드 페이지 이동 중...", "INFO")
            try:
                split_mode_link = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, "//a[contains(@href, '/Downtime/WCT2040.aspx')]"))
                )
                driver.execute_script("arguments[0].click();", split_mode_link)
                time.sleep(1.5)
                self.log("   ✓ 분리모드 링크 클릭 완료", "SUCCESS")
            except TimeoutException:
                self.log("   → 분리모드 링크를 찾을 수 없어 직접 URL로 이동합니다", "WARNING")
                driver.get("https://wct-cj1.unitus.co.kr/Downtime/WCT2040.aspx")
                time.sleep(1)

            # URL 이동 후 알림창 처리
            try:
                alert = driver.switch_to.alert
                alert.accept()
                self.log("   → 페이지 이동 알림창 처리됨", "INFO")
            except:
                pass

            # 분리모드 버튼 클릭
            self.log("\n8. 분리모드 활성화 중...", "INFO")
            split_mode_btn = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_itbtnA_Divide_CD"))
            )
            split_mode_btn.click()
            time.sleep(0.5)
            self.log("   ✓ 분리모드 활성화 완료", "SUCCESS")
            self.log("\n=== MES 로그인 완료 ===", "SUCCESS")

            return True

        except Exception as e:
            self.log(f"\n!!! 로그인 실패: {str(e)}", "ERROR")
            traceback.print_exc()
            return False

    def establish_connection(self):
        """MES 연결 수행"""
        try:
            self.log("\n=== MES 연결 시작 ===")
            # 드라이버 초기화
            self.log("1. 웹 드라이버 초기화 중...")
            
            # 브라우저 표시 여부 확인
            show_browser = self.chrome_visible_var.get()
            if not show_browser:
                self.log("   → 크롬 창 비활성화 모드로 실행")
            self.driver = initialize_driver(show_browser=show_browser)
            self.log("   → 웹 드라이버 초기화 완료")
            
            # 분리모드 전용 로그인 및 페이지 이동
            self.log("\n2. MES 로그인 시도 중...", "INFO")
            login_success = self.login_and_navigate_split_mode(
                self.driver,
                self.parent.id_entry.get(),
                self.parent.pw_entry.get()
            )
            
            if not login_success:
                self.log("!!! MES 로그인 실패", "ERROR")
                raise Exception("로그인 실패")
            
            # 연결 상태 업데이트
            self.is_connected = True
            self.update_status("MES 연결 성공", "green")
            self.search_btn.configure(state="normal")
            self.log("\n=== MES 연결 완료 ===", "SUCCESS")
            
        except Exception as e:
            self.log(f"\n!!! MES 연결 실패: {str(e)}", "ERROR")
            traceback.print_exc()
            self.update_status("MES 연결 실패", "red")
            if self.driver:
                self.driver.quit()
                self.driver = None
            self.is_connected = False
            self.search_btn.configure(state="disabled")
            messagebox.showerror("오류", "MES 연결에 실패했습니다.")



    def start_data_search(self):
        """데이터 조회 시작"""
        if not self.validate_inputs():
            return
            
        print("\n=== 데이터 조회 시작 ===")
        print("1. 입력값 검증 완료")
        self.update_status("데이터 조회 중...", "yellow")
        self.search_btn.configure(state="disabled")
        print("2. 검색 버튼 비활성화")
        self.crawling_thread = threading.Thread(target=self.search_data, daemon=True)
        self.crawling_thread.start()
        print("3. 크롤링 스레드 시작")

    def validate_inputs(self):
        """입력값 검증"""
        # 라인 선택 확인
        selected_lines = [line for line, var in self.line_vars.items() if var.get()]
        if not selected_lines:
            messagebox.showwarning("경고", "하나 이상의 라인을 선택해주세요.")
            return False
            
        # 날짜 형식 검증
        try:
            datetime.strptime(self.start_date_var.get(), "%Y-%m-%d")
            datetime.strptime(self.end_date_var.get(), "%Y-%m-%d")
        except ValueError:
            messagebox.showwarning("경고", "올바른 날짜 형식을 입력해주세요 (YYYY-MM-DD).")
            return False
            
        return True
        
    def clear_tree(self):
        """트리뷰 데이터 초기화"""
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.tree["columns"] = ()
        
    def search_data(self):
        """데이터 조회 수행"""
        try:
            # 두 번째 이후 데이터 조회시 페이지 새로고침 및 분리모드 활성화
            self.driver.refresh()
            time.sleep(2)
            
            try:
                split_mode_btn = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_itbtnA_Divide_CD"))
                )
                split_mode_btn.click()
                time.sleep(2)
            except:
                self.log("분리모드 버튼 클릭 실패", "WARNING")
                pass

            self.log("=== 조회 프로세스 시작 ===")
            # 라인타입 선택
            self.update_status("라인 설정", "라인타입 선택 중...")
            self.log("1. 라인타입 선택 중...")
            line_type_dropdown = self.driver.find_element(By.ID, "ContentPlaceHolder1_ddeLineType_B-1")
            line_type_dropdown.click()
            time.sleep(0.1)
            
            line_type = self.line_type_var.get().split("]")[0].strip("[")
            line_type_checkbox = self.driver.find_element(By.ID, f"ContentPlaceHolder1_ddeLineType_DDD_listLineType_0_{line_type}_D")
            line_type_checkbox.click()
            time.sleep(0.1)
            self.log(f"   → 라인타입 [{line_type}] 선택됨")

            self.driver.find_element(By.TAG_NAME, "body").click()
            time.sleep(0.5)

            # 라인 선택
            self.log("\n2. 라인 선택 중...")
            self.update_status("선택된 라인 체크 중...")
            line_dropdown = self.driver.find_element(By.ID, "ContentPlaceHolder1_ddeLine_B-1")
            line_dropdown.click()
            time.sleep(0.1)

            try:
                filter_button = self.driver.find_element(By.ID, "ContentPlaceHolder1_ddeLine_DDD_listLine_0_LBSFEB")
                filter_button.click()
                time.sleep(0.1)

                selected_lines = [line for line, var in self.line_vars.items() if var.get()]
                total_lines = len(selected_lines)
                self.log(f"   → 선택된 라인 수: {total_lines}")
                
                for idx, line_code in enumerate(selected_lines, 1):
                    try:
                        code = line_code.split("]")[0].strip("[")
                        self.update_status(f"라인 선택 중... ({idx}/{total_lines})")
                        self.log(f"   → 라인 선택: {line_code}")
                        
                        filter_input = self.driver.find_element(By.ID, "ContentPlaceHolder1_ddeLine_DDD_listLine_0_LBFE_I")
                        filter_input.clear()
                        time.sleep(0.1)
                        filter_input.send_keys(code)
                        time.sleep(0.5)
                        
                        checkbox_id = f"ContentPlaceHolder1_ddeLine_DDD_listLine_0_{code}_D"
                        click_script = f"""
                            var checkbox = document.getElementById('{checkbox_id}');
                            if (checkbox) {{
                                checkbox.click();
                                return true;
                            }}
                            return false;
                        """
                        success = self.driver.execute_script(click_script)
                        
                        if not success:
                            self.log(f"      → 라인을 찾을 수 없음: {line_code}")
                            
                        filter_input.clear()
                        time.sleep(0.1)
                        
                    except Exception as e:
                        self.log(f"      → 라인 선택 중 오류: {line_code}")
                        continue

                try:
                    filter_button.click()
                except:
                    pass
                
                time.sleep(0.1)
                webdriver.ActionChains(self.driver).send_keys(Keys.ESCAPE).perform()
                time.sleep(0.1)
                self.log("   → 라인 선택 완료")
                
            except Exception as e:
                self.log(f"   → 라인 선택 처리 중 오류 발생")
                try:
                    webdriver.ActionChains(self.driver).send_keys(Keys.ESCAPE).perform()
                except:
                    pass

            time.sleep(0.5)

            # 기간 선택
            self.log("\n3. 기간 옵션 선택 중...")
            self.update_status("기간 옵션 선택 중...")
            period_dropdown = self.driver.find_element(By.ID, "ContentPlaceHolder1_ddl_Div_B-1")
            period_dropdown.click()
            time.sleep(0.1)

            try:
                period_option = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//td[contains(text(), '기간')]"))
                )
                period_option.click()
                time.sleep(0.1)
                self.log("   → 기간 옵션 선택됨")
            except Exception as e:
                self.log(f"   → 기간 옵션 선택 중 오류: {str(e)}")

            # 날짜 입력
            self.log("\n4. 날짜 입력 중...")
            self.log(f"   → 검색 기간: {self.start_date_entry.get()} ~ {self.end_date_entry.get()}")
            
            # 시작 날짜 입력
            self.log("   → 시작 날짜 입력 중...")
            self.update_status("시작 날짜 입력 중...")
            time.sleep(0.5)
            
            try:
                start_date = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "input.dxeEditArea_DevEx[id='ContentPlaceHolder1_dxDateEditStart_I']"))
                )
                
                self.driver.execute_script("""
                    arguments[0].value = '';
                    arguments[0].dispatchEvent(new Event('focus'));
                """, start_date)
                time.sleep(0.5)
                
                start_date_value = self.start_date_entry.get()
                self.driver.execute_script("""
                    arguments[0].value = arguments[1];
                    arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                    arguments[0].dispatchEvent(new Event('blur', { bubbles: true }));
                    if(typeof ASPx !== 'undefined') {
                        ASPx.ETextChanged('ContentPlaceHolder1_dxDateEditStart');
                        ASPx.ELostFocus('ContentPlaceHolder1_dxDateEditStart');
                    }
                """, start_date, start_date_value)
                
                time.sleep(0.5)
                self.log("      ✓ 시작 날짜 입력 완료")
            except Exception as e:
                self.log(f"      → 시작 날짜 입력 중 오류: {str(e)}")

            # 종료 날짜 입력
            self.log("   → 종료 날짜 입력 중...")
            self.update_status("종료 날짜 입력 중...")
            try:
                end_date = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "input.dxeEditArea_DevEx[id='ContentPlaceHolder1_dxDateEditEnd_I']"))
                )
                
                self.driver.execute_script("""
                    arguments[0].value = '';
                    arguments[0].dispatchEvent(new Event('focus'));
                """, end_date)
                time.sleep(0.1)
                
                end_date_value = self.end_date_entry.get()
                self.driver.execute_script("""
                    arguments[0].value = arguments[1];
                    arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                    arguments[0].dispatchEvent(new Event('blur', { bubbles: true }));
                    if(typeof ASPx !== 'undefined') {
                        ASPx.ETextChanged('ContentPlaceHolder1_dxDateEditEnd');
                        ASPx.ELostFocus('ContentPlaceHolder1_dxDateEditEnd');
                    }
                """, end_date, end_date_value)
                
                time.sleep(0.5)
                self.log("      ✓ 종료 날짜 입력 완료")
            except Exception as e:
                self.log(f"      → 종료 날짜 입력 중 오류: {str(e)}")

            # 조회 버튼 클릭
            self.log("\n5. 데이터 조회 중...")
            self.update_status("데이터 조회", "조회 버튼 클릭...")
            search_button = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//td[contains(text(),'조회')]"))
            )
            search_button.click()
            self.log("   → 조회 버튼 클릭됨")

            # 로딩 상태 확인
            self.log("\n6. 데이터 로딩 중...")
            try:
                # 로딩 이미지 감지
                loading_element = WebDriverWait(self.driver, 3).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "dxgvLoadingPanel_DevEx"))
                )
                self.log("   → 데이터 로딩 시작됨")
                
                # 로딩 이미지가 사라질 때까지 대기
                WebDriverWait(self.driver, 30).until_not(
                    EC.presence_of_element_located((By.CLASS_NAME, "dxgvLoadingPanel_DevEx"))
                )
                self.log("   → 데이터 로딩 완료")
                
            except TimeoutException:
                self.log("   → 데이터 로딩 완료")
            except Exception as e:
                self.log("   → 데이터 로딩 중 오류 발생")

            # 데이터 파싱
            self.log("\n7. 데이터 파싱 중...")
            self.wait_for_data_and_parse()

        except Exception as e:
            self.log(f"\n!!! 데이터 조회 중 오류 발생")
            traceback.print_exc()
            self.update_status("데이터 조회 실패", "red")
            messagebox.showerror("오류", "데이터 조회에 실패했습니다.")

    def wait_for_loading(self, timeout=30):
        """로딩 완료를 기다립니다."""
        try:
            print("\n=== 로딩 감지 프로세스 시작 ===")
            # 로딩 이미지 대기
            loading_element = None
            print("1. 로딩 이미지 감지 시도 중...")
            try:
                loading_element = WebDriverWait(self.driver, 3).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "dxgvLoadingPanel_DevEx"))
                )
                print("   ✓ 로딩 이미지 감지됨")
            except:
                print("   → 로딩 이미지를 찾을 수 없음 (정상적인 상황일 수 있음)")
                pass

            if loading_element:
                print("2. 로딩 이미지 사라짐 대기 중...")
                try:
                    # 로딩 이미지가 사라질 때까지 대기
                    WebDriverWait(self.driver, timeout).until_not(
                        EC.presence_of_element_located((By.CLASS_NAME, "dxgvLoadingPanel_DevEx"))
                    )
                    print("   ✓ 로딩 완료 감지")
                except:
                    print("   → 로딩 완료 감지 실패 (타임아웃)")
                    pass

            print("3. 데이터 그리드 로드 확인 중...")
            # 데이터 그리드가 로드될 때까지 대기
            WebDriverWait(self.driver, timeout).until(
                EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_ASPxSplitter1_dxGrid2_DXMainTable"))
            )
            print("   ✓ 데이터 그리드 요소 발견")
            
            start_time = time.time()
            print("4. 데이터 검증 시작...")
            while True:
                if time.time() - start_time > timeout:
                    print("   → 데이터 검증 시간 초과")
                    raise TimeoutException(f"데이터 로딩 시간 초과 ({timeout}초)")
                
                try:
                    # No data 체크
                    no_data = self.driver.find_elements(By.XPATH, "//div[text()='No data']")
                    if no_data:
                        print("   → 데이터가 없음")
                        return True
                        
                    # 데이터 그리드 체크
                    grid = self.driver.find_element(By.ID, "ContentPlaceHolder1_ASPxSplitter1_dxGrid2_DXMainTable")
                    if grid:
                        # 추가: 데이터가 실제로 로드되었는지 확인
                        rows = grid.find_elements(By.CLASS_NAME, "dxgvDataRow")
                        if len(rows) > 0:
                            print(f"   ✓ 데이터 확인됨: {len(rows)}행 발견")
                            time.sleep(0.5)  # 데이터가 완전히 로드될 때까지 추가 대기
                            print("5. 로딩 완료")
                            return True
                        else:
                            print("   → 데이터 행을 찾을 수 없음, 재시도...")
                        
                except Exception as e:
                    print(f"   → 데이터 확인 중 오류: {str(e)}")
                    time.sleep(0.5)
                    continue
                
        except TimeoutException as e:
            print(f"!!! 로딩 시간 초과: {str(e)}")
            raise
        except Exception as e:
            print(f"!!! 로딩 대기 중 오류: {str(e)}")
            raise

    def wait_for_loading(self, timeout=30):
        """언밸런스 모듈의 적용 버튼 클릭 후 로딩 완료를 기다립니다."""
        try:
            print("\n=== 언밸런스 로딩 감지 프로세스 시작 ===")
            
            # 적용 버튼 찾기 (정확한 HTML 구조 반영)
            apply_button = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//span[@id='ContentPlaceHolder1_ASPxSplitter1_itbtnR_Apply1']//td[@class='btn01' and contains(text(), '적용')]"))
            )
            
            # JavaScript로 클릭 이벤트와 함수 실행
            self.driver.execute_script("""
                var result = ApplyClick1();
                if (result) {
                    __doPostBack('ctl00$ContentPlaceHolder1$ASPxSplitter1$itbtnR_Apply1','');
                }
                return result;
            """)
            print("   ✓ 적용 버튼 클릭됨")
            
            # 로딩 이미지 대기
            print("1. 로딩 이미지 감지 시도 중...")
            try:
                loading_element = WebDriverWait(self.driver, 3).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "dxgvLoadingPanel_DevEx"))
                )
                print("   ✓ 로딩 이미지 감지됨")
                
                # 로딩 이미지가 사라질 때까지 대기
                WebDriverWait(self.driver, timeout).until_not(
                    EC.presence_of_element_located((By.CLASS_NAME, "dxgvLoadingPanel_DevEx"))
                )
                print("   ✓ 로딩 완료 감지")
            except:
                print("   → 로딩 이미지를 찾을 수 없음 (정상적인 상황일 수 있음)")
            
            # 데이터 그리드 업데이트 확인
            print("2. 데이터 그리드 업데이트 확인 중...")
            time.sleep(1)  # 데이터 그리드 업데이트를 위한 추가 대기
            
            return True
            
        except Exception as e:
            print(f"!!! 언밸런스 로딩 대기 중 오류: {str(e)}")
            raise

    def wait_for_data_and_parse(self):
        """데이터를 파싱하고 표시합니다."""
        try:
            print("\n=== 데이터 파싱 프로세스 시작 ===")
            
            # 로딩 창 생성
            loading_window = LoadingWindow(self)
            loading_window.update_status("데이터 로딩 중...")
            print("1. 로딩 창 생성됨")
            
            # 데이터 그리드가 로드될 때까지 대기
            print("\n2. 데이터 그리드 로딩 대기 중...")
            self.wait_for_loading()
            loading_window.update_progress(0.2, "데이터 크롤링 중...")
            print("   ✓ 데이터 그리드 로딩 완료")
            
            print("\n3. 데이터 크롤링 준비...")
            # 메인 테이블 찾기
            print("   → 메인 테이블 검색 중...")
            main_table = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_ASPxSplitter1_dxGrid2_DXMainTable"))
            )
            print("   ✓ 메인 테이블 발견")
            
            # 헤더 업데이트
            self.create_headers()
            print("   ✓ 헤더 업데이트 완료")
            
            # 실제 데이터 행 수 확인
            data_rows = main_table.find_elements(By.CLASS_NAME, "dxgvDataRow")
            total_rows = len(data_rows)
            print(f"\n   → 총 {total_rows}개의 데이터 행 발견")
            
            all_data = []  # 모든 행 데이터를 저장할 리스트
            print("\n4. 데이터 크롤링 시작...")
            
            # 실제 데이터 행 수만큼만 처리
            for row_idx in range(total_rows):
                try:
                    row_id = f"ContentPlaceHolder1_ASPxSplitter1_dxGrid2_DXDataRow{row_idx}"
                    print(f"\n   → 행 {row_idx + 1}/{total_rows} 처리 중... (ID: {row_id})")
                    
                    row = WebDriverWait(self.driver, 2).until(
                        EC.presence_of_element_located((By.ID, row_id))
                    )
                    
                    row_data = [""] * 13  # 13개의 빈 값으로 초기화
                    bg_color = None
                    
                    # 모든 td 요소 찾기
                    cells = row.find_elements(By.TAG_NAME, "td")
                    
                    # 각 셀 처리 (13개까지만)
                    for col_idx, cell in enumerate(cells):
                        if col_idx >= 13:  # 13개 컬럼까지만 처리
                            break
                            
                        try:
                            if col_idx == 0:  # 체크박스 열
                                input_elem = cell.find_element(By.TAG_NAME, "input")
                                row_data[col_idx] = input_elem.get_attribute("value") or "U"
                                print(f"     ✓ 체크박스 값: {row_data[col_idx]}")
                                
                            elif col_idx == 1:  # 상태 열
                                row_data[col_idx] = cell.text.strip()
                                bg_color = cell.value_of_css_property("background-color")
                                print(f"     ✓ 상태: {row_data[col_idx]} (배경색: {bg_color})")
                                
                            else:  # 나머지 열
                                row_data[col_idx] = cell.text.strip()
                                print(f"     ✓ 열 {col_idx}: {row_data[col_idx]}")
                                
                        except Exception as e:
                            print(f"     !!! 열 {col_idx} 처리 중 오류: {str(e)}")
                            continue
                    
                    all_data.append((row_data, bg_color))
                    print(f"     ✓ 행 {row_idx + 1} 데이터 저장 완료")
                    loading_window.update_progress((row_idx + 1) / total_rows)
                    
                except Exception as e:
                    print(f"     !!! 행 {row_idx} 처리 중 오류: {str(e)}")
                    continue
            
            # 데이터 표시
            print("\n5. 데이터 표시 중...")
            print("   → 기존 데이터 초기화")
            self.clear_data_grid()
            
            print("   → 새 데이터 추가 중...")
            for idx, (row_data, bg_color) in enumerate(all_data):
                self.add_data_row(idx, row_data, bg_color)
                print(f"     ✓ 행 {idx + 1}/{len(all_data)} 표시됨")
            
            print("\n6. 마무리 작업...")
            loading_window.destroy()
            print("   ✓ 로딩 창 제거됨")
            
            self.data_frame.update_idletasks()
            print("   ✓ 데이터 프레임 새로고침됨")
            
            # 입력 필드 값 설정
            try:
                print("\n7. 입력 필드 값 설정 중...")
                self.driver.execute_script("""
                    var input = document.getElementById('ContentPlaceHolder1_ASPxSplitter1_ddlGubunDowntime_I');
                    if (input) {
                        input.value = '전체항목';
                        input.dispatchEvent(new Event('change', { bubbles: true }));
                        if(typeof ASPx !== 'undefined') {
                            ASPx.ETextChanged('ContentPlaceHolder1_ASPxSplitter1_ddlGubunDowntime');
                        }
                    }
                """)
                print("   ✓ 입력 필드 값 설정 완료")
            except Exception as e:
                print(f"   !!! 입력 필드 값 설정 중 오류: {str(e)}")
            
            print("\n=== 데이터 파싱 완료 ===")
            self.search_btn.configure(state="normal")
            print("   ✓ 조회 버튼 다시 활성화됨")
            
            return all_data, self.column_mapping.keys()
            
        except Exception as e:
            print("\n!!! 데이터 파싱 중 오류 발생 !!!")
            print(str(e))
            traceback.print_exc()
            
            if 'loading_window' in locals():
                loading_window.destroy()
            
            self.search_btn.configure(state="normal")
            print("   ✓ 오류 발생 후 조회 버튼 다시 활성화됨")
            return None, None

    def monitor_web_checkboxes(self):
        """웹페이지 체크박스 상태 모니터링"""
        if not self.driver:
            return
            
        try:
            # 모든 행에 대해 체크박스 상태 확인
            for row_idx in range(100):  # 충분히 큰 범위 설정
                try:
                    row_id = f"ContentPlaceHolder1_ASPxSplitter1_dxGrid2_DXDataRow{row_idx}"
                    row = self.driver.find_element(By.ID, row_id)
                    checkbox_cell = row.find_elements(By.TAG_NAME, "td")[0]
                    
                    # 체크박스 상태 확인
                    input_elem = checkbox_cell.find_element(By.CSS_SELECTOR, "input.dxAIFFE")
                    is_checked = input_elem.get_attribute("value") == "C"
                    
                    # UI 체크박스 상태 업데이트
                    checkbox_var = getattr(self, f"checkbox_var_{row_idx}", None)
                    if checkbox_var and checkbox_var.get() != is_checked:
                        if is_checked:
                            checkbox_var.set(True)
                        else:
                            checkbox_var.set(False)
                            
                except NoSuchElementException:
                    break  # 더 이상 행이 없으면 중단
                    
        except Exception as e:
            print(f"체크박스 상태 모니터링 중 오류: {e}")

    def get_cell_data(self, cell, include_html=False):
        """셀 데이터를 추출하는 헬퍼 함수"""
        try:
            if include_html:
                return cell.get_attribute('innerHTML')
            
            # DevExpress 그리드의 숨겨진 값 가져오기
            try:
                hidden_value = self.driver.execute_script("""
                    var cell = arguments[0];
                    var cellContent = cell.querySelector('td[class*="dxgv"]');
                    if (cellContent) {
                        return cellContent.innerText;
                    }
                    return null;
                """, cell)
                if hidden_value:
                    return hidden_value.strip()
            except:
                pass
            
            # 기본 텍스트 값 가져오기
            text = cell.text.strip()
            if not text:
                # title 속성 확인
                title = cell.get_attribute('title')
                if title:
                    return title.strip()
                    
                # aria-label 속성 확인
                aria_label = cell.get_attribute('aria-label')
                if aria_label:
                    return aria_label.strip()
                    
                # 내부 input 요소 확인
                inputs = cell.find_elements(By.TAG_NAME, 'input')
                if inputs:
                    for input_elem in inputs:
                        value = input_elem.get_attribute('value')
                        if value:
                            return value.strip()
            
            return text
        except Exception as e:
            print(f"셀 데이터 추출 중 오류: {str(e)}")
            return ""

    def display_parsed_data(self, parsed_data, column_headers):
        """파싱된 데이터를 표시합니다."""
        if parsed_data:
            self.update_data_grid(parsed_data)
            # 체크박스 모니터링 시작
            self.start_checkbox_monitoring()
            
    def on_closing(self):
        """창이 닫힐 때의 처리"""
        # 체크박스 모니터링 중지
        self.stop_checkbox_monitoring()
        
        # 기존 정리 작업 수행
        if self.driver:
            self.driver.quit()
        self.destroy()

    def create_line_checkboxes(self, line_type):
        """라인 체크박스 생성"""
        # 기존 체크박스 제거
        for widget in self.line_scrolled_frame.winfo_children():
            widget.destroy()
        
        # 새 체크박스 생성
        if line_type in self.config["lines"]:
            for line in self.config["lines"][line_type]:
                var = ctk.BooleanVar(value=False)
                if line in self.parent.line_vars:
                    var.set(self.parent.line_vars[line].get())
                self.line_vars[line] = var
                cb = ctk.CTkCheckBox(self.line_scrolled_frame, 
                                   text=line, 
                                   variable=var,
                                   fg_color="#3B8ED0",
                                   border_color="#666666",
                                   text_color="#FFFFFF",
                                   font=("Arial", 11),
                                   height=24)
                cb.pack(anchor="w", padx=5, pady=2)

    def on_line_type_change(self, value):
        """라인타입 변경 시 호출"""
        self.create_line_checkboxes(value)

    def start_checkbox_monitoring(self):
        """체크박스 모니터링 시작"""
        if self.monitoring_thread and self.monitoring_thread.is_alive():
            return
            
        self.stop_monitoring = False
        self.monitoring_thread = threading.Thread(target=self.checkbox_monitoring_loop, daemon=True)
        self.monitoring_thread.start()
        
    def stop_checkbox_monitoring(self):
        """체크박스 모니터링 중지"""
        self.stop_monitoring = True
        if self.monitoring_thread:
            self.monitoring_thread.join(timeout=1.0)
            
    def checkbox_monitoring_loop(self):
        """체크박스 상태 모니터링 루프"""
        while not self.stop_monitoring:
            if self.driver and self.is_connected:
                self.monitor_web_checkboxes()
            time.sleep(0.5)  # 0.5초마다 체크

    def set_column_width(self, column_index, width):
        """특정 열의 너비를 설정"""
        if 0 <= column_index <= 12:  # 유효한 열 인덱스 확인
            old_width = self.column_widths[column_index]
            self.column_widths[column_index] = width
            
            # 너비가 0이 되거나 0에서 변경되는 경우 전체 그리드 새로 그리기
            if (old_width == 0 and width > 0) or (old_width > 0 and width == 0):
                self.refresh_grid()
            else:
                # 단순 너비 변경인 경우 해당 열만 업데이트
                self.update_column_width(column_index, width)

    def update_column_width(self, column_index, width):
        """열 너비만 업데이트"""
        if width > 0:
            # 헤더 업데이트
            header = self.header_frame.grid_slaves(row=0, column=column_index)
            if header:
                header[0].configure(width=width)
            # 데이터 셀 업데이트
            for row in range(len(self.data_frame.winfo_children()) // len(self.column_widths)):
                cell = self.data_frame.grid_slaves(row=row, column=column_index)
                if cell:
                    cell[0].configure(width=width)

    def refresh_grid(self):
        """그리드 전체 새로 그리기"""
        # 현재 데이터 백업
        current_data = []
        rows = len(self.data_frame.winfo_children()) // len(self.column_widths)
        for row in range(rows):
            row_data = []
            for col in range(len(self.column_widths)):
                cells = self.data_frame.grid_slaves(row=row, column=col)
                if cells:
                    cell = cells[0]
                    if isinstance(cell, ctk.CTkCheckBox):
                        row_data.append(cell.get())
                    else:
                        row_data.append(cell.cget("text"))
            current_data.append(row_data)
        
        # 그리드 초기화
        self.clear_data_grid()
        self.create_headers()
        
        # 데이터 다시 표시
        for idx, row_data in enumerate(current_data):
            self.add_data_row(idx, row_data)
            
    def set_column_alignment(self, column_index, header_alignment, data_alignment=None):
        """열의 정렬 방식을 설정"""
        if 0 <= column_index <= 12:
            self.column_alignments[column_index] = header_alignment
            if data_alignment is not None:
                self.data_alignments[column_index] = data_alignment
            self.refresh_grid()
            
    def register_data(self):
        """등록 버튼 클릭 핸들러"""
        try:
            # 입력값 검증
            if not self.validate_inputs():
                return
                
            # 공정명, 원인, 조치사항 가져오기
            process_name = self.process_dropdown.get()
            cause = self.cause_textbox.get()
            action = self.action_textbox.get()
            
            # 필수비가동 유형 가져오기
            essential_type = self.essential_var.get() if hasattr(self, 'essential_var') else None
            
            # 일반비가동 유형 가져오기
            normal_type = self.normal_type if hasattr(self, 'normal_type') else None
            
            # MESRegister 모듈 한 번만 import
            from mes_register import MESRegister
            register = MESRegister(self.driver, self.log)
            
            # 필수비가동이 선택된 경우 새로운 모듈 사용
            if essential_type:
                if register.register_downtime(process_name, cause, action, essential_type=essential_type):
                    # 입력 필드 초기화
                    self.process_dropdown.set("선택")
                    self.cause_textbox.delete(0, 'end')
                    self.action_textbox.delete(0, 'end')
                    self.split_time_var.set("")  # 분리 비가동 시간 초기화
                    self.essential_var.set("")  # 필수비가동 라디오 버튼 초기화
                    self.clear_data_grid()
                    self.disable_registration_controls()  # UI 요소 비활성화 추가
                    self.log("모든 초기화 완료", "SUCCESS")
                    self.grab_release()  # 메시지 박스를 위해 잠시 모달 해제
                    messagebox.showinfo("알림", "분리 필수비가동 등록되었습니다.")
                    self.grab_set()  # 메시지 박스 닫힌 후 다시 모달 설정
                return
            # 일반비가동이 선택된 경우 새로운 모듈 사용
            elif normal_type:
                if register.register_normal_downtime(process_name, normal_type, cause, action):
                    # 입력 필드 초기화
                    self.process_dropdown.set("선택")
                    self.cause_textbox.delete(0, 'end')
                    self.action_textbox.delete(0, 'end')
                    self.split_time_var.set("")  # 분리 비가동 시간 초기화
                    if hasattr(self, 'normal_type'):
                        self.normal_type = None  # normal_type 초기화
                    self.normal_var.set("")  # 일반비가동 라디오 버튼 초기화
                    self.clear_data_grid()
                    self.disable_registration_controls()  # UI 요소 비활성화 추가
                    self.log("모든 초기화 완료", "SUCCESS")
                    self.grab_release()  # 메시지 박스를 위해 잠시 모달 해제
                    messagebox.showinfo("알림", "분리 일반비가동 등록되었습니다.")
                    self.grab_set()  # 메시지 박스 닫힌 후 다시 모달 설정
                return
                
            # 기존 등록 로직 실행
            try:
                # 여기에 기존 등록 로직이 있음
                # ... existing code ...

                # 공정명 입력
                self.log("공정명 입력 중...")
                process_input = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_ASPxSplitter1_ddlOperDowntime_I"))
                )
                
                # 기존 value 삭제를 위해 Ctrl+A 후 Delete
                process_input.send_keys(Keys.CONTROL + "a")
                process_input.send_keys(Keys.DELETE)
                time.sleep(0.5)
                
                # 공정명을 한 글자씩 입력
                for char in process_name:
                    process_input.send_keys(char)
                    time.sleep(0.1)  # 각 글자 사이에 약간의 딜레이
                
                process_input.send_keys(Keys.RETURN)
                time.sleep(2)

                # 필수 비가동 열기
                self.log("필수 비가동 메뉴 열기...")
                essential_xpath = "//td[@class='dxgv']/img[@class='dxGridView_gvCollapsedButton' and contains(@onclick,\"ASPx.GVExpandRow('ContentPlaceHolder1_ASPxSplitter1_dxGrid3',2,\")]"
                essential_button = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, essential_xpath))
                )
                essential_button.click()
                time.sleep(1)
                
                # 기종 변경 열기
                self.log("기종 변경 메뉴 열기...")
                model_xpath = "//td[@class='dxgv']/img[@class='dxGridView_gvCollapsedButton' and contains(@onclick,\"ASPx.GVExpandRow('ContentPlaceHolder1_ASPxSplitter1_dxGrid3',3,\")]"
                model_button = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, model_xpath))
                )
                model_button.click()
                time.sleep(1)
                
                # 차종/기종교체(계획) 열기
                self.log("차종/기종교체(계획) 메뉴 열기...")
                planned_xpath = "//td[@class='dxgv']/img[@class='dxGridView_gvCollapsedButton' and contains(@onclick,\"ASPx.GVExpandRow('ContentPlaceHolder1_ASPxSplitter1_dxGrid3',6,\")]"
                planned_button = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, planned_xpath))
                )
                planned_button.click()
                time.sleep(1)
                
                # 비가동 클릭
                self.log("비가동 선택...")
                downtime_cell = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//td[@class='dxgvIndentCell dxgv'][@style='width:0px;border-left-width:0px;']"))
                )
                downtime_cell.click()
                time.sleep(1)
                                                
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
                
                # 적용 버튼 클릭
                self.log("적용 버튼 클릭...")
                apply_button = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//td[@class='btn01'][contains(text(),'적용')]"))
                )
                apply_button.click()
                self.wait_for_loading()
                self.log("적용 버튼 클릭 성공")

                # 클릭 후 로딩 대기
                self.wait_for_loading()

                # 체크박스 클릭
                self.log("\n체크박스 선택 시도...", "INFO")
                
                try:
                    # 체크박스 요소 찾기
                    checkbox_td = WebDriverWait(self.driver, 5).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, "td[onclick*=\"ASPx.GVScheduleCommand('ContentPlaceHolder1_ASPxSplitter1_dxGrid1',['Select',1]\"]"))
                    )
                    
                    # JavaScript로 체크박스 상태 변경 및 이벤트 트리거
                    self.driver.execute_script("""
                        function triggerCheckbox(tdElement) {
                            // 체크박스 span과 input 찾기
                            var span = document.getElementById('ContentPlaceHolder1_ASPxSplitter1_dxGrid1_DXSelBtn1_D');
                            var input = document.getElementById('ContentPlaceHolder1_ASPxSplitter1_dxGrid1_DXSelBtn1');
                            
                            if (span && input) {
                                // 체크박스 상태 변경
                                span.className = 'dxWeb_edtCheckBoxChecked dxICheckBox dxichSys dx-not-acc';
                                input.value = 'C';
                                
                                // 클릭 이벤트 시뮬레이션
                                var clickEvent = new MouseEvent('click', {
                                    bubbles: true,
                                    cancelable: true,
                                    view: window
                                });
                                tdElement.dispatchEvent(clickEvent);
                                
                                // DevExpress 명령 직접 실행
                                if (typeof ASPx !== 'undefined') {
                                    ASPx.GVScheduleCommand('ContentPlaceHolder1_ASPxSplitter1_dxGrid1', ['Select', 1], 1);
                                }
                                
                                // 행 스타일 적용
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
                    
                    time.sleep(0.1)  # 상태 변경 대기
                    
                    # 체크박스 상태 확인
                    checkbox_state = self.driver.execute_script("""
                        var input = document.getElementById('ContentPlaceHolder1_ASPxSplitter1_dxGrid1_DXSelBtn1');
                        return input ? input.value : null;
                    """)
                    
                    if checkbox_state == 'C':
                        self.log("체크박스 선택 완료", "SUCCESS")
                        
                        time.sleep(0.1)  # 상태 변경 대기

                except Exception as e:
                    self.log(f"체크박스 선택 중 오류: {str(e)}", "ERROR")
                    # 대체 방법 시도
                    try:
                        self.log("대체 방법으로 선택 시도...", "INFO")
                        # 직접 td 요소 클릭
                        checkbox_td = self.driver.find_element(By.CSS_SELECTOR, 
                            "td[onclick*=\"ASPx.GVScheduleCommand('ContentPlaceHolder1_ASPxSplitter1_dxGrid1',['Select',1]\"]")
                        self.driver.execute_script("arguments[0].click();", checkbox_td)
                        time.sleep(0.1)
                        
                        # 상태 변경 확인
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
                        
                        if checkbox_state == 'C':
                            self.log("대체 방법으로 선택 완료", "SUCCESS")
                            # 체크박스 선택 완료 후 UI 활성화
                            self.enable_registration_controls()
                        else:
                            self.log("대체 방법으로도 선택 실패", "ERROR")
                            # UI 비활성화
                            self.disable_registration_controls()
                        
                    except Exception as e2:
                        self.log(f"대체 방법도 실패: {str(e2)}", "ERROR")
                        self.disable_registration_controls()
                
                # 등록 버튼 클릭
                self.log("등록 버튼 클릭...")
                register_button = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//td[@class='btn01'][contains(text(),'등록')]"))
                )
                register_button.click()
                
                # 확인 메시지 처리
                try:
                    alert = WebDriverWait(self.driver, 10).until(EC.alert_is_present())
                    alert_text = alert.text
                    self.log(f"알림 메시지: {alert_text}")
                    alert.accept()
                    
                    if "등록 대상이 없습니다" in alert_text:
                        self.log("등록 대상이 없습니다.", "WARNING")
                        return
                except:
                    self.log("확인 메시지가 나타나지 않음")
                
                # 등록 완료 대기
                try:
                    self.log("등록 진행 중...", "INFO")
                    loading_element = WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located((By.ID, "loadingPanel"))
                    )
                    self.log("로딩 패널 감지됨", "INFO")
                    
                    # 로딩창이 사라질 때까지 대기
                    WebDriverWait(self.driver, 30).until(
                        EC.invisibility_of_element(loading_element)
                    )
                    self.log("로딩 완료", "INFO")
                    
                except TimeoutException:
                    self.log("등록이 완료되었습니다.", "SUCCESS")
                
                self.log("입력 필드 초기화 중...", "INFO")
                # 입력 필드 초기화
                self.process_dropdown.set("선택")
                self.cause_textbox.delete(0, 'end')
                self.action_textbox.delete(0, 'end')
                # 데이터 그리드 초기화
                self.log("데이터 그리드 초기화 중...", "INFO")
                self.clear_data_grid()
                self.log("모든 초기화 완료", "SUCCESS")
                # 메시지 박스를 마지막에 표시
                self.log("작업 완료 메시지 표시", "INFO")
                messagebox.showinfo("알림", "분리 비가동 등록되었습니다.")
                self.log("모든 프로세스 완료", "SUCCESS")                
                return
                
            except Exception as e:
                self.log(f"차종/기종교체 등록 중 오류: {str(e)}", "ERROR")
                messagebox.showerror("오류", "차종/기종교체 등록 중 오류가 발생했습니다.")
                return
            
        except Exception as e:
            self.log(f"데이터 등록 중 오류: {str(e)}", "ERROR")
            messagebox.showerror("오류", "데이터 등록 중 오류가 발생했습니다.")

    def log(self, message, level="INFO"):
        """로그 메시지를 로그창에 출력"""
        try:
            # 첫 로그 메시지인 경우 초기화 메시지 출력
            if not hasattr(self, 'log_initialized') or not self.log_initialized:
                self.log_text.delete("1.0", "end")
                self.log_initialized = True
            
            # 모든 메시지를 흰색으로 통일
            level_color = "#FFFFFF"
            
            # 로그 메시지 포맷팅 (레벨 표시 제거)
            log_message = f"{message}\n"
            
            # 텍스트 위젯에 메시지 추가
            self.log_text.insert("end", log_message)
            
            # 색상 태그 설정
            last_line_start = f"end-{len(log_message)}c"
            self.log_text.tag_add("log", last_line_start, "end")
            self.log_text.tag_config("log", foreground=level_color)
            
            # 스크롤을 항상 최하단으로
            self.log_text.see("end")
            
            # 로그창 업데이트
            self.log_text.update()
            
            # 디버깅을 위해 print도 함께 사용
            print(message)
        except Exception as e:
            print(f"로그 출력 중 오류: {str(e)}")

    def validate_inputs(self):
        """입력값 검증"""
        # 라인 선택 확인
        selected_lines = [line for line, var in self.line_vars.items() if var.get()]
        if not selected_lines:
            self.log("경고: 라인이 선택되지 않음", "WARNING")
            messagebox.showwarning("경고", "하나 이상의 라인을 선택해주세요.")
            return False
            
        # 날짜 형식 검증
        try:
            start_date = datetime.strptime(self.start_date_var.get(), "%Y-%m-%d")
            end_date = datetime.strptime(self.end_date_var.get(), "%Y-%m-%d")
            
            # 선택된 라인 정보 로깅
            self.log("\n=== 검색 조건 ===", "INFO")
            self.log(f"선택된 라인 타입: {self.line_type_var.get()}", "INFO")
            self.log(f"선택된 라인 수: {len(selected_lines)}", "INFO")
            for line in selected_lines:
                self.log(f"  - {line}", "INFO")
            self.log(f"검색 기간: {start_date.strftime('%Y-%m-%d')} ~ {end_date.strftime('%Y-%m-%d')}", "INFO")
            
        except ValueError:
            self.log("경고: 잘못된 날짜 형식", "WARNING")
            messagebox.showwarning("경고", "올바른 날짜 형식을 입력해주세요 (YYYY-MM-DD).")
            return False
            
        self.log("입력값 검증 완료", "SUCCESS")
        return True

    def split_time(self):
        """시간 분리 실행"""
        try:
            self.log("=== 시간 분리 시작 ===", "INFO")
            # 체크된 행이 있는지 확인
            if self.current_checked_idx is None:
                self.log("오류: 선택된 행이 없음", "ERROR")
                messagebox.showwarning("경고", "분리할 행을 선택해주세요.")
                return
            self.log(f"선택된 행 번호: {self.current_checked_idx}", "INFO")
                
            # 분리시간 입력값 확인
            try:
                split_minutes = float(self.split_time_var.get())
                self.log(f"입력된 분리시간: {split_minutes}분", "INFO")
                if split_minutes <= 0:
                    self.log("오류: 분리시간이 0보다 작거나 같음", "ERROR")
                    messagebox.showwarning("경고", "분리시간은 0보다 커야 합니다.")
                    return
            except ValueError:
                self.log("오류: 잘못된 분리시간 형식", "ERROR")
                messagebox.showwarning("경고", "올바른 분리시간을 입력해주세요.")
                return
                
            # 체크된 행의 중단시간(분) 가져오기
            cells = self.data_frame.grid_slaves(row=self.current_checked_idx, column=6)
            if not cells:
                self.log("오류: 중단시간을 찾을 수 없음", "ERROR")
                messagebox.showerror("오류", "중단시간을 찾을 수 없습니다.")
                return
                
            # 체크된 행의 라인명 가져오기
            line_cells = self.data_frame.grid_slaves(row=self.current_checked_idx, column=6)
            if line_cells:
                line_name = line_cells[0].cget("text").strip()
                self.log(f"가져온 시간: {line_name}", "INFO")
            
            stop_time_str = cells[0].cget("text").strip()
            try:
                # 숫자가 아닌 문자 제거 후 float 변환 시도
                stop_time_str = ''.join(c for c in stop_time_str if c.isdigit() or c == '.')
                if not stop_time_str:
                    self.log("오류: 중단시간이 비어있음", "ERROR")
                    messagebox.showerror("오류", "중단시간이 비어있습니다.")
                    return
                stop_time = float(stop_time_str)
            except ValueError:
                self.log("오류: 중단시간이 올바른 형식이 아님", "ERROR")
                messagebox.showerror("오류", "중단시간이 올바른 형식이 아닙니다.")
                return
                
            # 초로 변환 및 계산
            self.log("\n=== 시간 계산 과정 ===", "INFO")
            self.log(f"1. 전체 중단시간: {stop_time}분", "INFO")
            total_seconds = stop_time * 60
            self.log(f"2. 전체 중단시간 초 변환: {stop_time} × 60 = {total_seconds}초", "INFO")
            
            self.log(f"\n3. 분리할 시간: {split_minutes}분", "INFO")
            split_seconds = split_minutes * 60
            self.log(f"4. 분리할 시간 초 변환: {split_minutes} × 60 = {split_seconds}초", "INFO")
            
            # 제외 시간 계산 (전체 시간 - 분리 시간)
            remaining_seconds = total_seconds - split_seconds
            self.log(f"\n5. 제외 시간 계산: {total_seconds} - {split_seconds} = {remaining_seconds}초", "INFO")
            
            self.log("\n=== 최종 결과 ===", "INFO")
            self.log(f"입력예정값: {remaining_seconds}초", "SUCCESS")

            # 웹페이지의 입력 필드에 값 설정
            try:
                # 입력 필드 찾기
                self.log("\n입력 필드 설정 시도...", "INFO")
                input_field = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_ASPxSplitter1_dxDivideSecond_I"))
                )
                self.log("입력 필드 찾음", "SUCCESS")
                
                # JavaScript로 값 설정
                self.driver.execute_script("""
                    var input = arguments[0];
                    var value = arguments[1];
                    input.value = value;
                    
                    // ASPx 이벤트 트리거
                    if(typeof ASPx !== 'undefined') {
                        ASPx.EValueChanged('ContentPlaceHolder1_ASPxSplitter1_dxDivideSecond');
                    }
                """, input_field, str(remaining_seconds))
                
                self.log(f"입력 필드에 값 설정 완료: {remaining_seconds}", "SUCCESS")
                
                # 확인 메시지 표시
                result_message = f"=== 시간 분리 계산 결과 ===\n\n" \
                               f"전체 중단시간: {stop_time}분 ({total_seconds}초)\n" \
                               f"분리할 시간: {split_minutes}분 ({split_seconds}초)\n" \
                               f"제외 시간: {remaining_seconds/60:.2f}분 ({remaining_seconds}초)\n\n" \
                               f"입력예정값: {remaining_seconds}초\n\n" \
                               f"입력하시겠습니까?"
                
                if messagebox.askyesno("확인", result_message):
                    # 추가 버튼 클릭
                    add_button = self.driver.find_element(By.ID, "ContentPlaceHolder1_ASPxSplitter1_btnDivideInsert_I")
                    if add_button:
                        self.driver.execute_script("arguments[0].click();", add_button)
                        self.log("추가 버튼 클릭 완료", "SUCCESS")
                        
                        # 로딩 대기 및 체크박스 선택
                        try:
                            self.log("\n로딩 완료 대기 중...", "INFO")
                            time.sleep(1)  # 초기 대기
                            
                            try:
                                # 로딩 이미지 감지 시도
                                loading_element = WebDriverWait(self.driver, 2).until(
                                    EC.presence_of_element_located((By.CLASS_NAME, "dxgvLoadingPanel_DevEx"))
                                )
                                self.log("로딩 시작 감지됨", "INFO")
                                
                                # 로딩 이미지가 사라질 때까지 대기
                                WebDriverWait(self.driver, 30).until_not(
                                    EC.presence_of_element_located((By.CLASS_NAME, "dxgvLoadingPanel_DevEx"))
                                )
                                self.log("로딩 완료 감지됨", "SUCCESS")
                            except:
                                self.log("로딩 이미지를 찾을 수 없음 (정상적인 상황일 수 있음)", "INFO")
                                time.sleep(2)  # 추가 대기
                            
                            # 체크박스 클릭
                            self.log("\n체크박스 선택 시도...", "INFO")
                            
                            try:
                                # 체크박스 요소 찾기
                                checkbox_td = WebDriverWait(self.driver, 5).until(
                                    EC.presence_of_element_located((By.CSS_SELECTOR, "td[onclick*=\"ASPx.GVScheduleCommand('ContentPlaceHolder1_ASPxSplitter1_dxGrid1',['Select',1]\"]"))
                                )
                                
                                # JavaScript로 체크박스 상태 변경 및 이벤트 트리거
                                self.driver.execute_script("""
                                    function triggerCheckbox(tdElement) {
                                        // 체크박스 span과 input 찾기
                                        var span = document.getElementById('ContentPlaceHolder1_ASPxSplitter1_dxGrid1_DXSelBtn1_D');
                                        var input = document.getElementById('ContentPlaceHolder1_ASPxSplitter1_dxGrid1_DXSelBtn1');
                                        
                                        if (span && input) {
                                            // 체크박스 상태 변경
                                            span.className = 'dxWeb_edtCheckBoxChecked dxICheckBox dxichSys dx-not-acc';
                                            input.value = 'C';
                                            
                                            // 클릭 이벤트 시뮬레이션
                                            var clickEvent = new MouseEvent('click', {
                                                bubbles: true,
                                                cancelable: true,
                                                view: window
                                            });
                                            tdElement.dispatchEvent(clickEvent);
                                            
                                            // DevExpress 명령 직접 실행
                                            if (typeof ASPx !== 'undefined') {
                                                ASPx.GVScheduleCommand('ContentPlaceHolder1_ASPxSplitter1_dxGrid1', ['Select', 1], 1);
                                            }
                                            
                                            // 행 스타일 적용
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
                                
                                time.sleep(0.1)  # 상태 변경 대기
                                
                                # 체크박스 상태 확인
                                checkbox_state = self.driver.execute_script("""
                                    var input = document.getElementById('ContentPlaceHolder1_ASPxSplitter1_dxGrid1_DXSelBtn1');
                                    return input ? input.value : null;
                                """)
                                
                                if checkbox_state == 'C':
                                    self.log("체크박스 선택 완료", "SUCCESS")
                                    # 체크박스 선택 완료 후 UI 활성화
                                    self.enable_registration_controls()
                                else:
                                    self.log("체크박스 상태 확인 필요", "WARNING")
                                    # UI 비활성화
                                    self.disable_registration_controls()
                                    
                                # 한번 더 클릭 이벤트 발생
                                self.driver.execute_script("""
                                    var td = arguments[0];
                                    var event = new MouseEvent('click', {
                                        bubbles: true,
                                        cancelable: true,
                                        view: window
                                    });
                                    td.dispatchEvent(event);
                                """, checkbox_td)
                                
                            except Exception as e:
                                self.log(f"체크박스 선택 중 오류: {str(e)}", "ERROR")
                                # 대체 방법 시도
                                try:
                                    self.log("대체 방법으로 선택 시도...", "INFO")
                                    # 직접 td 요소 클릭
                                    checkbox_td = self.driver.find_element(By.CSS_SELECTOR, 
                                        "td[onclick*=\"ASPx.GVScheduleCommand('ContentPlaceHolder1_ASPxSplitter1_dxGrid1',['Select',1]\"]")
                                    self.driver.execute_script("arguments[0].click();", checkbox_td)
                                    time.sleep(0.1)
                                    
                                    # 상태 변경 확인
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
                                    
                                    if checkbox_state == 'C':
                                        self.log("대체 방법으로 선택 완료", "SUCCESS")
                                        # 체크박스 선택 완료 후 UI 활성화
                                        self.enable_registration_controls()
                                    else:
                                        self.log("대체 방법으로도 선택 실패", "ERROR")
                                        # UI 비활성화
                                        self.disable_registration_controls()
                                    
                                except Exception as e2:
                                    self.log(f"대체 방법도 실패: {str(e2)}", "ERROR")
                                    self.disable_registration_controls()
                            
                        except Exception as e:
                            self.log(f"로딩 대기 또는 체크박스 선택 중 오류: {str(e)}", "ERROR")
                            traceback.print_exc()
                            self.disable_registration_controls()
                else:
                    self.log("사용자가 취소를 선택함", "WARNING")
                    self.disable_registration_controls()
                    
            except Exception as e:
                self.log(f"입력 필드 설정 중 오류: {str(e)}", "ERROR")
                self.disable_registration_controls()
                raise
            
        except Exception as e:
            self.log(f"시간 분리 중 오류: {str(e)}", "ERROR")
            self.disable_registration_controls()
            self.grab_release()  # 메시지 박스를 위해 잠시 모달 해제
            messagebox.showerror("오류", "시간 분리 중 오류가 발생했습니다.")
            self.grab_set()  # 메시지 박스 닫힌 후 다시 모달 설정

    def on_essential_select(self, value):
        """필수비가동 라디오 버튼 선택 핸들러"""
        # 일반비가동 선택 해제
        self.normal_var.set("")

    def on_normal_select(self, value):
        """일반비가동 라디오 버튼 선택 처리"""
        self.normal_type = value  # normal_type을 인스턴스 변수로 저장
        if value:
            self.essential_var.set("")  # 필수비가동 선택 해제
            self.enable_registration_controls()
        else:
            self.disable_registration_controls()

    def enable_registration_controls(self):
        """비가동 등록 관련 UI 요소들을 활성화"""
        if not self.registration_enabled:
            self.log("\n=== 비가동 등록 UI 활성화 ===", "INFO")
            
            # 필수비가동 라디오 버튼 활성화
            for radio in self.essential_radios:
                radio.configure(state="normal")
            self.log("필수비가동 라디오 버튼 활성화됨", "SUCCESS")
            
            # 일반비가동 라디오 버튼 활성화
            for radio in self.normal_radios:
                radio.configure(state="normal")
            self.log("일반비가동 라디오 버튼 활성화됨", "SUCCESS")
            
            # 공정명 드롭다운 활성화
            self.process_dropdown.configure(state="normal")
            self.log("공정명 드롭다운 활성화됨", "SUCCESS")
            
            # 공정불러오기 버튼 활성화
            self.load_process_btn.configure(state="normal")
            self.log("공정불러오기 버튼 활성화됨", "SUCCESS")
            
            # 등록 버튼 활성화
            self.register_btn.configure(state="normal")
            self.log("등록 버튼 활성화됨", "SUCCESS")
            
            self.registration_enabled = True
            self.log("=== UI 활성화 완료 ===\n", "SUCCESS")

    def disable_registration_controls(self):
        """비가동 등록 관련 UI 요소들을 비활성화"""
        if self.registration_enabled:
            self.log("\n=== 비가동 등록 UI 비활성화 ===", "INFO")
            
            # 필수비가동 라디오 버튼 비활성화
            for radio in self.essential_radios:
                radio.configure(state="disabled")
            self.log("필수비가동 라디오 버튼 비활성화됨", "SUCCESS")
            
            # 일반비가동 라디오 버튼 비활성화
            for radio in self.normal_radios:
                radio.configure(state="disabled")
            self.log("일반비가동 라디오 버튼 비활성화됨", "SUCCESS")
            
            # 공정명 드롭다운 비활성화
            self.process_dropdown.configure(state="disabled")
            self.log("공정명 드롭다운 비활성화됨", "SUCCESS")
            
            # 공정불러오기 버튼 비활성화
            self.load_process_btn.configure(state="disabled")
            self.log("공정불러오기 버튼 비활성화됨", "SUCCESS")
            
            # 등록 버튼 비활성화
            self.register_btn.configure(state="disabled")
            self.log("등록 버튼 비활성화됨", "SUCCESS")
            
            # 선택 초기화
            self.essential_var.set("")
            self.normal_var.set("")
            self.process_dropdown.set("선택")
            
            self.registration_enabled = False
            self.log("=== UI 비활성화 완료 ===\n", "SUCCESS")

    def load_process_list(self):
        """공정불러오기 버튼 클릭 핸들러"""
        self.log("\n=== 공정 불러오기 시작 ===", "INFO")
        
        # 체크된 행의 4번째 열 값 수집
        selected_lines = set()
        
        # data_frame의 모든 행을 순회
        for row_idx in range(len(self.data_frame.grid_slaves())):
            # 해당 행의 체크박스 변수 가져오기
            checkbox_var = getattr(self, f"checkbox_var_{row_idx}", None)
            if checkbox_var and checkbox_var.get():  # 체크된 행만 처리
                try:
                    # 해당 행의 모든 셀 데이터 출력
                    cells = self.data_frame.grid_slaves(row=row_idx)
                    cells.reverse()  # grid_slaves는 역순으로 반환하므로 순서 뒤집기
                    
                    for i, cell in enumerate(cells):
                        if isinstance(cell, ctk.CTkLabel):
                            cell_value = cell.cget("text").strip()
                            self.log(f"열 {i}: {cell_value}", "INFO")
                            
                            # 4번째 열(라인 정보)인 경우
                            if i == 4:
                                self.log(f"선택된 라인: {cell_value}", "INFO")
                                if cell_value:
                                    selected_lines.add(cell_value)
                                
                except Exception as e:
                    self.log(f"행 {row_idx} 처리 중 오류: {str(e)}", "ERROR")
        
        if not selected_lines:
            self.log("선택된 행이 없습니다.", "WARNING")
            return
        
        self.log(f"찾은 라인들: {selected_lines}", "INFO")
        
        # 공정명 데이터 정의
        process_data = self.config.get("process_data", {})
        if not process_data:
            # 설정 파일에 공정명 데이터가 없는 경우 기본값 사용
            process_data = {
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
            # 설정 파일에 공정명 데이터 저장
            self.config["process_data"] = process_data
            self.save_config()
        
        # 선택된 라인에 해당하는 공정명만 필터링
        filtered_processes = []
        for line in selected_lines:
            self.log(f"라인 '{line}' 처리 중...", "INFO")
            line_name = line.split("] ")[1] if "] " in line else line  # "[HC02] HPCU 2" -> "HPCU 2"
            if line_name in process_data:
                filtered_processes.extend(process_data[line_name])
                self.log(f"라인 '{line_name}'의 공정 {len(process_data[line_name])}개 추가됨", "SUCCESS")
            else:
                self.log(f"라인 '{line_name}'에 대한 공정 데이터가 없습니다.", "WARNING")
        
        # 중복 제거 및 정렬
        filtered_processes = sorted(list(set(filtered_processes)))
        
        # 드롭다운 업데이트
        if filtered_processes:
            self.process_dropdown.configure(values=filtered_processes)
            self.process_dropdown.set(filtered_processes[0])  # 첫 번째 값 선택
            self.log(f"총 {len(filtered_processes)}개의 공정이 로드됨", "SUCCESS")
        else:
            self.process_dropdown.configure(values=["공정 없음"])
            self.process_dropdown.set("공정 없음")
            self.log("로드된 공정이 없습니다.", "WARNING")
        
        self.log("=== 공정 불러오기 완료 ===\n", "INFO")

    def save_config(self):
        """설정 파일 저장"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"설정 저장 중 오류: {e}")










