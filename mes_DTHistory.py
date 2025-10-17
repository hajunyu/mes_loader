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

class DTHistory(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        
        # 현재 체크된 체크박스들의 인덱스를 저장하는 변수 추가 (다중 선택)
        self.current_checked_indices = []
        
        # 비가동 관련 UI 요소들의 활성화 상태를 관리하는 변수 추가
        self.registration_enabled = False
        
        # 기본 창 설정
        self.title("MES 비가동관리 도구")
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
        
        # 필요한 컬럼만 매핑 (사용자 요청에 따라)
        self.column_mapping = {
            0: "체크박스",           # col0
            1: "실적상태",          # col31  
            2: "등록구분",          # col32
            3: "진행현황",          # col33
            5: "작업일",            # col3
            6: "라인명",            # col5
            8: "중단시간(분)",      # col55
            11: "공정",             # col34
            12: "공정명",           # col35
            34: "대분류명",         # col8
            36: "중분류명",         # col10
            38: "소분류명",         # col12
            45: "원인"              # col42
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
        
        # 오른쪽 패널 (데이터 그리드용)
        right_panel = ctk.CTkFrame(self.main_frame, fg_color="#1E1E1E")
        right_panel.pack(side="right", fill="both", expand=True, padx=1, pady=1)
        
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
        self.line_scrolled_frame = ctk.CTkScrollableFrame(left_panel, fg_color="#1E1E1E", height=150)
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
                                      font=("Arial", 12, "bold"),
                                      height=38)
        self.search_btn.pack(pady=10, padx=10, fill="x")
        self.search_btn.configure(state="disabled")
        
        # 구분선 추가
        separator5 = ctk.CTkFrame(left_panel, height=1, fg_color="#404040")
        separator5.pack(fill="x", padx=10, pady=5)
        
        # 작업 버튼 프레임
        action_buttons_frame = ctk.CTkFrame(left_panel, fg_color="#2B2B2B")
        action_buttons_frame.pack(fill="x", padx=10, pady=5)
        
        # 첫 번째 줄: 등록취소, 확정
        action_row1 = ctk.CTkFrame(action_buttons_frame, fg_color="#2B2B2B")
        action_row1.pack(fill="x", pady=(0,3))
        
        # 등록취소 버튼
        self.cancel_register_btn = ctk.CTkButton(action_row1,
                                                 text="등록취소",
                                                 command=self.cancel_register,
                                                 fg_color="#DC3545",
                                                 hover_color="#C82333",
                                                 font=("Arial", 11),
                                                 height=32,
                                                 width=100)
        self.cancel_register_btn.pack(side="left", padx=2, fill="both", expand=True)
        self.cancel_register_btn.configure(state="disabled")
        
        # 확정 버튼
        self.confirm_btn = ctk.CTkButton(action_row1,
                                        text="확정",
                                        command=self.confirm_data,
                                        fg_color="#28A745",
                                        hover_color="#218838",
                                        font=("Arial", 11),
                                        height=32,
                                        width=100)
        self.confirm_btn.pack(side="left", padx=2, fill="both", expand=True)
        self.confirm_btn.configure(state="disabled")
        
        # 두 번째 줄: 확정취소
        action_row2 = ctk.CTkFrame(action_buttons_frame, fg_color="#2B2B2B")
        action_row2.pack(fill="x", pady=0)
        
        # 확정취소 버튼
        self.cancel_confirm_btn = ctk.CTkButton(action_row2,
                                               text="확정취소",
                                               command=self.cancel_confirm,
                                               fg_color="#FFC107",
                                               hover_color="#E0A800",
                                               font=("Arial", 11),
                                               height=32)
        self.cancel_confirm_btn.pack(fill="x", padx=2)
        self.cancel_confirm_btn.configure(state="disabled")
        
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
        self.grid_frame = ctk.CTkFrame(right_panel, fg_color="#1E1E1E")
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
            "font": ("Arial", 10, "bold")  # 폰트 크기를 줄여서 더 많은 컬럼 표시
        }

        # 필요한 컬럼만 헤더 생성
        visible_col = 0
        for idx in sorted(self.column_mapping.keys()):
            header_text = self.column_mapping[idx]
            # 헤더 텍스트가 너무 길면 줄임
            if len(header_text) > 10:
                header_text = header_text[:10] + "..."
            
            label = ctk.CTkLabel(
                self.header_frame,
                text=header_text,
                width=100,  # 조금 더 넓게 설정
                anchor="center",
                **style
            )
            label.grid(row=0, column=visible_col, sticky="nsew", padx=1, pady=1)
            self.header_frame.grid_columnconfigure(visible_col, weight=1)
            visible_col += 1

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
        
        # 현재 체크된 인덱스들 초기화
        self.current_checked_indices = []

    def add_data_row(self, row_idx, data, bg_color=None):
        """데이터 행 추가"""
        # 기본 스타일
        style = {
            "fg_color": "transparent",
            "text_color": "white",
            "height": 22,
            "font": ("Arial", 9)  # 폰트 크기를 줄여서 더 많은 컬럼 표시
        }
            
        # 필요한 컬럼만 표시
        visible_col = 0  # 실제로 표시되는 열의 인덱스
        
        for col_idx in sorted(self.column_mapping.keys()):
            if col_idx < len(data):
                value = data[col_idx]
                
                if col_idx == 0:  # 체크박스 열
                    checkbox_var = ctk.BooleanVar()
                    setattr(self, f"checkbox_var_{row_idx}", checkbox_var)
                    
                    cell = ctk.CTkCheckBox(
                        self.data_frame,
                        text="",
                        variable=checkbox_var,
                        width=80,  # 조금 더 넓게
                        height=16,
                        fg_color="#3B8ED0",
                        border_color="#666666"
                    )
                    
                    # 체크박스 클릭 이벤트 연결
                    cell.configure(command=lambda idx=row_idx: self.on_checkbox_click(idx))
                    
                else:  # 일반 데이터 열
                    # 텍스트가 너무 길면 줄임
                    display_text = str(value)
                    if len(display_text) > 15:
                        display_text = display_text[:15] + "..."
                    
                    cell = ctk.CTkLabel(
                        self.data_frame,
                        text=display_text,
                        width=100,  # 조금 더 넓게
                        anchor="center",
                        **style
                    )
                
                cell.grid(row=row_idx, column=visible_col, sticky="nsew", padx=1, pady=0)
                visible_col += 1
            
        # 열 가중치 설정
        for i in range(visible_col):
            self.data_frame.grid_columnconfigure(i, weight=1)

    def on_checkbox_click(self, row_idx):
        """체크박스 클릭 이벤트 핸들러 (다중 선택 지원)"""
        try:
            # 체크박스 상태 가져오기
            checkbox_var = getattr(self, f"checkbox_var_{row_idx}")
            is_checked = checkbox_var.get()
            
            # 현재 체크박스 상태 업데이트 (다중 선택)
            if is_checked:
                # 체크된 경우 리스트에 추가
                if row_idx not in self.current_checked_indices:
                    self.current_checked_indices.append(row_idx)
                
                # 체크된 행의 정보 출력 (새로운 컬럼 구조에 맞게 수정)
                try:
                    # 라인명 찾기 (컬럼 매핑에서 라인명의 실제 인덱스 찾기)
                    line_name_col = None
                    for col_idx, col_name in self.column_mapping.items():
                        if col_name == "라인명":
                            line_name_col = col_idx
                            break
                    
                    if line_name_col is not None:
                        # 라인명이 표시되는 컬럼 위치 찾기
                        visible_cols = sorted(self.column_mapping.keys())
                        line_visible_col = visible_cols.index(line_name_col)
                        
                        cells = self.data_frame.grid_slaves(row=row_idx, column=line_visible_col)
                        if cells:
                            line_name = cells[0].cget("text").strip()
                            print(f"\n=== 체크된 행 정보 (총 {len(self.current_checked_indices)}개 선택) ===")
                            print(f"행 번호: {row_idx}")
                            print(f"라인명: {line_name}")
                            
                            # 추가 정보 출력 (처음 몇 개 컬럼)
                            for i, col_idx in enumerate(visible_cols[:5]):
                                try:
                                    col_cells = self.data_frame.grid_slaves(row=row_idx, column=i)
                                    if col_cells:
                                        col_value = col_cells[0].cget("text").strip()
                                        col_name = self.column_mapping.get(col_idx, f"컬럼{col_idx}")
                                        print(f"{col_name}: {col_value}")
                                except:
                                    continue
                except Exception as e:
                    print(f"라인명 값 출력 중 오류: {str(e)}")
                
            else:
                # 체크 해제된 경우 리스트에서 제거
                if row_idx in self.current_checked_indices:
                    self.current_checked_indices.remove(row_idx)
                print(f"\n체크 해제 (행: {row_idx}), 남은 선택: {len(self.current_checked_indices)}개")
            
            # 웹페이지의 체크박스 찾기
            row_id = f"ContentPlaceHolder1_ASPxSplitter1_dxGrid2_DXDataRow{row_idx}"
            row = self.driver.find_element(By.ID, row_id)
            checkbox_cell = row.find_elements(By.TAG_NAME, "td")[0]
            
            # 개선된 체크박스 처리 로직
            self.driver.execute_script("""
                var cell = arguments[0];
                var isChecked = arguments[1];
                
                // 현재 상태 확인
                var input = cell.querySelector('input.dxAIFFE');
                var span = cell.querySelector('span.dxICheckBox');
                var currentState = input ? input.value : 'U';
                
                // 원하는 상태와 다른 경우에만 처리
                if ((isChecked && currentState !== 'C') || (!isChecked && currentState === 'C')) {
                    // 먼저 UI 상태 변경
                    if (input) {
                        input.value = isChecked ? 'C' : 'U';
                    }
                    
                    if (span) {
                        span.className = isChecked ? 
                            'dxWeb_edtCheckBoxChecked dxICheckBox dxichSys dx-not-acc' : 
                            'dxWeb_edtCheckBoxUnchecked dxICheckBox dxichSys dx-not-acc';
                    }
                    
                    // 약간의 지연 후 클릭 이벤트 발생
                    setTimeout(function() {
                        cell.click();
                    }, 50);
                }
            """, checkbox_cell, is_checked)
            
            # 상태 변경이 완료될 때까지 대기
            time.sleep(0.15)
            
            # 최종 상태 확인 및 재시도
            final_state = self.driver.execute_script("""
                var cell = arguments[0];
                var isChecked = arguments[1];
                var input = cell.querySelector('input.dxAIFFE');
                var currentState = input ? input.value : 'U';
                
                // 상태가 일치하지 않으면 다시 시도
                if ((isChecked && currentState !== 'C') || (!isChecked && currentState === 'C')) {
                    if (input) {
                        input.value = isChecked ? 'C' : 'U';
                    }
                    var span = cell.querySelector('span.dxICheckBox');
                    if (span) {
                        span.className = isChecked ? 
                            'dxWeb_edtCheckBoxChecked dxICheckBox dxichSys dx-not-acc' : 
                            'dxWeb_edtCheckBoxUnchecked dxICheckBox dxichSys dx-not-acc';
                    }
                    return false;
                }
                return true;
            """, checkbox_cell, is_checked)
            
            if not final_state:
                print(f"체크박스 상태 불일치 발견 - 강제 동기화 완료 (행: {row_idx})")
            
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
            self.cancel_register_btn.configure(state="disabled")
            self.confirm_btn.configure(state="disabled")
            self.cancel_confirm_btn.configure(state="disabled")

    def login_and_navigate_dt_history(self, driver, username, password):
        """비가동관리 전용 로그인 및 페이지 이동"""
        max_retry = 2  # 최대 재시도 횟수
        retry_count = 0
        
        while retry_count < max_retry:
            try:
                self.log("\n=== MES 로그인 시작 ===", "INFO")
                # MES 시스템 접속
                self.log("1. MES 시스템 접속 중...", "INFO")
                driver.get("https://mescj.unitus.co.kr/Login.aspx")
                time.sleep(1)
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
                try:
                    self.log(f"WCT 시스템 이동 시도 중... (시도 {retry_count + 1}/{max_retry})", "INFO")
                    
                    # JavaScript를 사용하여 WCT 시스템 링크 클릭
                    wct_link = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, "//a[contains(text(), 'New WCT System')]"))
                    )
                    driver.execute_script("arguments[0].scrollIntoView(true); arguments[0].click();", wct_link)
                    time.sleep(1)
                    
                    # 새로 열린 창으로 전환 (최대 4초 대기)
                    try:
                        WebDriverWait(driver, 4).until(lambda d: len(d.window_handles) > 1)
                        driver.switch_to.window(driver.window_handles[-1])
                    except TimeoutException:
                        self.log(f"WCT 시스템 창 열기 실패 (시도 {retry_count + 1}/{max_retry})", "WARNING")
                        retry_count += 1
                        if retry_count >= max_retry:
                            self.log(f"최대 재시도 횟수 초과 ({max_retry}회)", "ERROR")
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
                        self.log("   → WCT 시스템 알림창 처리됨", "INFO")
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
                        self.log("비가동 등록/관리 링크를 찾을 수 없어 직접 URL로 이동합니다.", "WARNING")
                        driver.get("https://wct-cj1.unitus.co.kr/Downtime/WCT2040.aspx")
                        time.sleep(1)
                        
                        # URL 이동 후 알림창 처리
                        try:
                            alert = driver.switch_to.alert
                            alert.accept()
                        except:
                            pass
                
                    time.sleep(0.8)
                    
                    # 비가동관리 버튼 클릭
                    self.log("\n7. 비가동관리 활성화 중...", "INFO")
                    split_mode_btn = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnHistory_CD"))
                    )
                    split_mode_btn.click()
                    time.sleep(0.5)
                    self.log("   ✓ 비가동관리 활성화 완료", "SUCCESS")
                    
                    # 확정 버튼 클릭
                    self.log("\n8. 확정모드 활성화 중...", "INFO")
                    try:
                        confirm_btn = WebDriverWait(driver, 10).until(
                            EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_itbtnA_Confirm_CD"))
                        )
                        confirm_btn.click()
                        time.sleep(0.5)
                        self.log("   ✓ 확정모드 활성화 완료", "SUCCESS")
                    except Exception as e:
                        self.log(f"   → 확정 버튼 클릭 중 오류: {str(e)}", "WARNING")
                        # 대체 방법으로 시도
                        try:
                            confirm_input = driver.find_element(By.ID, "ContentPlaceHolder1_itbtnA_Confirm_I")
                            driver.execute_script("arguments[0].click();", confirm_input)
                            time.sleep(0.5)
                            self.log("   ✓ 확정모드 활성화 완료 (대체 방법)", "SUCCESS")
                        except Exception as e2:
                            self.log(f"   → 확정 버튼 대체 방법도 실패: {str(e2)}", "ERROR")
                    
                    # 페이지 로드 후 줌 레벨을 25%로 설정 (주석 처리)
                    # try:
                    #     driver.execute_script("document.body.style.zoom='1'")
                    #     self.log("   ✓ 브라우저 줌 레벨을 25%로 설정 완료", "SUCCESS")
                    # except Exception as e:
                    #     self.log(f"   → 줌 레벨 설정 중 오류: {str(e)}", "WARNING")
                    
                    # 모든 과정이 성공적으로 완료되면 루프 종료
                    self.log("\n=== MES 로그인 완료 ===", "SUCCESS")
                    return True
                    
                except Exception as e:
                    self.log(f"WCT 시스템 이동 중 오류: {str(e)}", "ERROR")
                    retry_count += 1
                    if retry_count >= max_retry:
                        return False
                    continue
                    
            except Exception as e:
                self.log(f"로그인 및 페이지 이동 중 오류 발생: {str(e)}", "ERROR")
                retry_count += 1
                if retry_count >= max_retry:
                    return False
                continue
        
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
            
            # 비가동관리 전용 로그인 및 페이지 이동
            self.log("\n2. MES 로그인 시도 중...", "INFO")
            login_success = self.login_and_navigate_dt_history(
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
            self.cancel_register_btn.configure(state="normal")
            self.confirm_btn.configure(state="normal")
            self.cancel_confirm_btn.configure(state="normal")
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
            self.cancel_register_btn.configure(state="disabled")
            self.confirm_btn.configure(state="disabled")
            self.cancel_confirm_btn.configure(state="disabled")
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
    
    def cancel_register(self):
        """등록취소 버튼 클릭"""
        # 버튼 비활성화
        self.cancel_register_btn.configure(state="disabled")
        self.confirm_btn.configure(state="disabled")
        self.cancel_confirm_btn.configure(state="disabled")
        
        try:
            # 체크된 항목이 있는지 확인
            if not self.current_checked_indices:
                messagebox.showwarning("경고", "등록취소할 항목을 선택해주세요.")
                return
            
            # 확인 메시지
            if not messagebox.askyesno("확인", f"{len(self.current_checked_indices)}개 항목을 등록취소하시겠습니까?"):
                return
            
            self.log(f"\n=== 등록취소 시작 ({len(self.current_checked_indices)}개 항목) ===", "INFO")
            
            # 웹페이지에서 등록취소 버튼 찾아서 클릭
            try:
                # XPath로 "등록취소" 텍스트를 포함한 tr 요소 찾기
                cancel_register_button = self.driver.find_element(
                    By.XPATH, 
                    "//tr[@onclick=\"javascript:__doPostBack('ctl00$ContentPlaceHolder1$itbtnR_Register_Cancel','')\"]"
                )
                self.driver.execute_script("arguments[0].click();", cancel_register_button)
                self.log("   → 등록취소 버튼 클릭", "INFO")
            except:
                # 대체 방법: td 요소에서 "등록취소" 텍스트 찾기
                cancel_register_button = self.driver.find_element(
                    By.XPATH,
                    "//td[contains(text(), '등록취소')]"
                )
                self.driver.execute_script("arguments[0].click();", cancel_register_button)
                self.log("   → 등록취소 버튼 클릭 (대체 방법)", "INFO")
            
            time.sleep(1)
            
            # 확인 알림창 처리
            try:
                alert = self.driver.switch_to.alert
                alert_text = alert.text
                self.log(f"   알림: {alert_text}", "INFO")
                alert.accept()
            except:
                pass
            
            # 로딩창 감지 및 대기
            self.log("   → 등록취소 처리 대기 중...", "INFO")
            loading_completed = False
            try:
                start_time = time.time()
                loading_appeared = False
                max_attempts = 20  # 최대 20번 확인
                check_interval = 10  # 10초마다 확인
                attempt = 0
                last_log_time = 0
                
                while attempt < max_attempts:
                    current_time = time.time()
                    
                    # UpdateProgress1 div의 상태 확인
                    loading_div = self.driver.find_elements(By.ID, "UpdateProgress1")
                    if loading_div:
                        style = loading_div[0].get_attribute("style")
                        aria_hidden = loading_div[0].get_attribute("aria-hidden")
                        
                        if "display: block" in style or aria_hidden == "false":
                            loading_appeared = True
                            # 10초마다 한 번씩 로그 출력
                            if current_time - last_log_time >= check_interval:
                                attempt += 1
                                self.log(f"   → 등록취소 처리 중... ({attempt}/{max_attempts})", "INFO")
                                last_log_time = current_time
                        elif loading_appeared and ("display: none" in style or aria_hidden == "true"):
                            self.log("   ✓ 등록취소 처리 완료", "SUCCESS")
                            loading_completed = True
                            time.sleep(2)  # 추가 안정화 대기
                            break
                    
                    # 3초 동안 로딩이 나타나지 않으면 계속 진행
                    if not loading_appeared and current_time - start_time > 3:
                        self.log("   → 로딩창 없이 완료", "INFO")
                        loading_completed = True
                        break
                        
                    time.sleep(0.5)
                
                if attempt >= max_attempts:
                    self.log("   → 최대 대기 시간 초과", "WARNING")
                    
            except Exception as e:
                self.log(f"   → 로딩 감지 중 오류: {str(e)}", "WARNING")
            
            # 등록취소가 완료된 경우에만 처리
            if loading_completed:
                self.log("   ✓ 등록취소 완료", "SUCCESS")
                
                # 체크된 항목 삭제
                self.current_checked_indices = []
                self.log("   → 체크된 항목 초기화", "INFO")
                
                # UI 데이터 그리드 전체 삭제 (웹페이지 데이터 변경으로 인한 오류 방지)
                self.clear_data_grid()
                self.log("   → UI 데이터 그리드 전체 삭제", "INFO")
                
                # 완료 메시지 표시
                messagebox.showinfo("완료", "등록취소가 완료되었습니다.\n데이터가 변경되어 목록을 초기화했습니다.")
            else:
                self.log("   ✗ 등록취소 처리 실패", "ERROR")
                messagebox.showerror("오류", "등록취소 처리가 완료되지 않았습니다.")
            
        except Exception as e:
            self.log(f"   ✗ 등록취소 중 오류: {str(e)}", "ERROR")
            traceback.print_exc()
            try:
                messagebox.showerror("오류", f"등록취소 중 오류가 발생했습니다:\n{str(e)}")
            except:
                pass
        finally:
            # 버튼 다시 활성화
            try:
                self.cancel_register_btn.configure(state="normal")
                self.confirm_btn.configure(state="normal")
                self.cancel_confirm_btn.configure(state="normal")
            except:
                pass
    
    def confirm_data(self):
        """확정 버튼 클릭"""
        # 버튼 비활성화
        self.cancel_register_btn.configure(state="disabled")
        self.confirm_btn.configure(state="disabled")
        self.cancel_confirm_btn.configure(state="disabled")
        
        try:
            # 체크된 항목이 있는지 확인
            if not self.current_checked_indices:
                messagebox.showwarning("경고", "확정할 항목을 선택해주세요.")
                return
            
            # 확인 메시지
            if not messagebox.askyesno("확인", f"{len(self.current_checked_indices)}개 항목을 확정하시겠습니까?"):
                return
            
            self.log(f"\n=== 확정 시작 ({len(self.current_checked_indices)}개 항목) ===", "INFO")
            
            # 웹페이지에서 확정 버튼 찾아서 클릭
            try:
                # XPath로 "확정" 텍스트를 포함한 tr 요소 찾기
                confirm_button = self.driver.find_element(
                    By.XPATH, 
                    "//tr[@onclick=\"javascript:__doPostBack('ctl00$ContentPlaceHolder1$itbtnR_Confirm','')\"]"
                )
                self.driver.execute_script("arguments[0].click();", confirm_button)
                self.log("   → 확정 버튼 클릭", "INFO")
            except:
                # 대체 방법: td 요소에서 "확정" 텍스트 찾기
                confirm_button = self.driver.find_element(
                    By.XPATH,
                    "//td[contains(text(), '확정') and not(contains(text(), '취소'))]"
                )
                self.driver.execute_script("arguments[0].click();", confirm_button)
                self.log("   → 확정 버튼 클릭 (대체 방법)", "INFO")
            
            time.sleep(1)
            
            # 확인 알림창 처리
            try:
                alert = self.driver.switch_to.alert
                alert_text = alert.text
                self.log(f"   알림: {alert_text}", "INFO")
                alert.accept()
            except:
                pass
            
            # 로딩창 감지 및 대기
            self.log("   → 확정 처리 대기 중...", "INFO")
            loading_completed = False
            try:
                start_time = time.time()
                loading_appeared = False
                max_attempts = 20  # 최대 20번 확인
                check_interval = 10  # 10초마다 확인
                attempt = 0
                last_log_time = 0
                
                while attempt < max_attempts:
                    current_time = time.time()
                    
                    # UpdateProgress1 div의 상태 확인
                    loading_div = self.driver.find_elements(By.ID, "UpdateProgress1")
                    if loading_div:
                        style = loading_div[0].get_attribute("style")
                        aria_hidden = loading_div[0].get_attribute("aria-hidden")
                        
                        if "display: block" in style or aria_hidden == "false":
                            loading_appeared = True
                            # 10초마다 한 번씩 로그 출력
                            if current_time - last_log_time >= check_interval:
                                attempt += 1
                                self.log(f"   → 확정 처리 중... ({attempt}/{max_attempts})", "INFO")
                                last_log_time = current_time
                        elif loading_appeared and ("display: none" in style or aria_hidden == "true"):
                            self.log("   ✓ 확정 처리 완료", "SUCCESS")
                            loading_completed = True
                            time.sleep(2)  # 추가 안정화 대기
                            break
                    
                    # 3초 동안 로딩이 나타나지 않으면 계속 진행
                    if not loading_appeared and current_time - start_time > 3:
                        self.log("   → 로딩창 없이 완료", "INFO")
                        loading_completed = True
                        break
                        
                    time.sleep(0.5)
                
                if attempt >= max_attempts:
                    self.log("   → 최대 대기 시간 초과", "WARNING")
                    
            except Exception as e:
                self.log(f"   → 로딩 감지 중 오류: {str(e)}", "WARNING")
            
            # 확정이 완료된 경우에만 처리
            if loading_completed:
                self.log("   ✓ 확정 완료", "SUCCESS")
                
                # 체크된 항목 삭제
                self.current_checked_indices = []
                self.log("   → 체크된 항목 초기화", "INFO")
                
                # UI 데이터 그리드 전체 삭제 (웹페이지 데이터 변경으로 인한 오류 방지)
                self.clear_data_grid()
                self.log("   → UI 데이터 그리드 전체 삭제", "INFO")
                
                # 완료 메시지 표시
                messagebox.showinfo("완료", "확정이 완료되었습니다.\n데이터가 변경되어 목록을 초기화했습니다.")
            else:
                self.log("   ✗ 확정 처리 실패", "ERROR")
                messagebox.showerror("오류", "확정 처리가 완료되지 않았습니다.")
            
        except Exception as e:
            self.log(f"   ✗ 확정 중 오류: {str(e)}", "ERROR")
            traceback.print_exc()
            try:
                messagebox.showerror("오류", f"확정 중 오류가 발생했습니다:\n{str(e)}")
            except:
                pass
        finally:
            # 버튼 다시 활성화
            try:
                self.cancel_register_btn.configure(state="normal")
                self.confirm_btn.configure(state="normal")
                self.cancel_confirm_btn.configure(state="normal")
            except:
                pass
    
    def cancel_confirm(self):
        """확정취소 버튼 클릭"""
        # 버튼 비활성화
        self.cancel_register_btn.configure(state="disabled")
        self.confirm_btn.configure(state="disabled")
        self.cancel_confirm_btn.configure(state="disabled")
        
        try:
            # 체크된 항목이 있는지 확인
            if not self.current_checked_indices:
                messagebox.showwarning("경고", "확정취소할 항목을 선택해주세요.")
                return
            
            # 확인 메시지
            if not messagebox.askyesno("확인", f"{len(self.current_checked_indices)}개 항목을 확정취소하시겠습니까?"):
                return
            
            self.log(f"\n=== 확정취소 시작 ({len(self.current_checked_indices)}개 항목) ===", "INFO")
            
            # 웹페이지에서 확정취소 버튼 찾아서 클릭
            try:
                # XPath로 "확정취소" 텍스트를 포함한 tr 요소 찾기
                cancel_confirm_button = self.driver.find_element(
                    By.XPATH, 
                    "//tr[@onclick=\"javascript:__doPostBack('ctl00$ContentPlaceHolder1$itbtnR_Confirm_Cancel','')\"]"
                )
                self.driver.execute_script("arguments[0].click();", cancel_confirm_button)
                self.log("   → 확정취소 버튼 클릭", "INFO")
            except:
                # 대체 방법: td 요소에서 "확정취소" 텍스트 찾기
                cancel_confirm_button = self.driver.find_element(
                    By.XPATH,
                    "//td[contains(text(), '확정취소')]"
                )
                self.driver.execute_script("arguments[0].click();", cancel_confirm_button)
                self.log("   → 확정취소 버튼 클릭 (대체 방법)", "INFO")
            
            time.sleep(1)
            
            # 확인 알림창 처리
            try:
                alert = self.driver.switch_to.alert
                alert_text = alert.text
                self.log(f"   알림: {alert_text}", "INFO")
                alert.accept()
            except:
                pass
            
            # 로딩창 감지 및 대기
            self.log("   → 확정취소 처리 대기 중...", "INFO")
            loading_completed = False
            try:
                start_time = time.time()
                loading_appeared = False
                max_attempts = 20  # 최대 20번 확인
                check_interval = 10  # 10초마다 확인
                attempt = 0
                last_log_time = 0
                
                while attempt < max_attempts:
                    current_time = time.time()
                    
                    # UpdateProgress1 div의 상태 확인
                    loading_div = self.driver.find_elements(By.ID, "UpdateProgress1")
                    if loading_div:
                        style = loading_div[0].get_attribute("style")
                        aria_hidden = loading_div[0].get_attribute("aria-hidden")
                        
                        if "display: block" in style or aria_hidden == "false":
                            loading_appeared = True
                            # 10초마다 한 번씩 로그 출력
                            if current_time - last_log_time >= check_interval:
                                attempt += 1
                                self.log(f"   → 확정취소 처리 중... ({attempt}/{max_attempts})", "INFO")
                                last_log_time = current_time
                        elif loading_appeared and ("display: none" in style or aria_hidden == "true"):
                            self.log("   ✓ 확정취소 처리 완료", "SUCCESS")
                            loading_completed = True
                            time.sleep(2)  # 추가 안정화 대기
                            break
                    
                    # 3초 동안 로딩이 나타나지 않으면 계속 진행
                    if not loading_appeared and current_time - start_time > 3:
                        self.log("   → 로딩창 없이 완료", "INFO")
                        loading_completed = True
                        break
                        
                    time.sleep(0.5)
                
                if attempt >= max_attempts:
                    self.log("   → 최대 대기 시간 초과", "WARNING")
                    
            except Exception as e:
                self.log(f"   → 로딩 감지 중 오류: {str(e)}", "WARNING")
            
            # 확정취소가 완료된 경우에만 처리
            if loading_completed:
                self.log("   ✓ 확정취소 완료", "SUCCESS")
                
                # 체크된 항목 삭제
                self.current_checked_indices = []
                self.log("   → 체크된 항목 초기화", "INFO")
                
                # UI 데이터 그리드 전체 삭제 (웹페이지 데이터 변경으로 인한 오류 방지)
                self.clear_data_grid()
                self.log("   → UI 데이터 그리드 전체 삭제", "INFO")
                
                # 완료 메시지 표시
                messagebox.showinfo("완료", "확정취소가 완료되었습니다.\n데이터가 변경되어 목록을 초기화했습니다.")
            else:
                self.log("   ✗ 확정취소 처리 실패", "ERROR")
                messagebox.showerror("오류", "확정취소 처리가 완료되지 않았습니다.")
            
        except Exception as e:
            self.log(f"   ✗ 확정취소 중 오류: {str(e)}", "ERROR")
            traceback.print_exc()
            try:
                messagebox.showerror("오류", f"확정취소 중 오류가 발생했습니다:\n{str(e)}")
            except:
                pass
        finally:
            # 버튼 다시 활성화
            try:
                self.cancel_register_btn.configure(state="normal")
                self.confirm_btn.configure(state="normal")
                self.cancel_confirm_btn.configure(state="normal")
            except:
                pass
        
    def clear_tree(self):
        """트리뷰 데이터 초기화"""
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.tree["columns"] = ()
        
    def search_data(self):
        """데이터 조회 수행"""
        try:
            # 두 번째 이후 데이터 조회시 페이지 새로고침 및 비가동관리 활성화
            self.driver.refresh()
            time.sleep(2)
            
            # 조회 시작 시 브라우저 줌 레벨을 100%로 복구
            try:
                self.driver.execute_script("document.body.style.zoom='1'")
                self.log("   ✓ 조회를 위해 브라우저 줌 레벨을 100%로 설정 완료", "SUCCESS")
            except Exception as e:
                self.log(f"   → 줌 레벨 설정 중 오류: {str(e)}", "WARNING")
            
            try:
                split_mode_btn = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.ID, "ContentPlaceHolder1_btnHistory_CD"))
                )
                split_mode_btn.click()
                time.sleep(1.5)
                self.log("   ✓ 비가동관리 버튼 클릭 완료", "SUCCESS")
            except:
                self.log("비가동관리 버튼 클릭 실패", "WARNING")
                pass

            # 확정모드 클릭
            try:
                self.log("\n확정모드 전환 중...")
                confirm_mode_clicked = False
                
                # 여러 가능한 ID 시도
                confirm_mode_ids = [
                    "ContentPlaceHolder1_itbtnA_Confirm",
                    "ContentPlaceHolder1_btnConfirm",
                    "ctl00_ContentPlaceHolder1_btnConfirm"
                ]
                
                for mode_id in confirm_mode_ids:
                    try:
                        confirm_mode_button = WebDriverWait(self.driver, 5).until(
                            EC.element_to_be_clickable((By.ID, mode_id))
                        )
                        self.driver.execute_script("arguments[0].click();", confirm_mode_button)
                        confirm_mode_clicked = True
                        self.log(f"   ✓ 확정모드 전환 완료 (ID: {mode_id})", "SUCCESS")
                        time.sleep(0.8)
                        break
                    except:
                        continue
                
                # ID로 찾기 실패시 JavaScript로 텍스트 기반 검색
                if not confirm_mode_clicked:
                    try:
                        self.driver.execute_script("""
                            var elements = document.getElementsByTagName('td');
                            for(var i = 0; i < elements.length; i++) {
                                if(elements[i].textContent.includes('확정')) {
                                    elements[i].click();
                                    break;
                                }
                            }
                        """)
                        self.log("   ✓ 확정모드 전환 완료 (텍스트 검색)", "SUCCESS")
                        time.sleep(0.8)
                        confirm_mode_clicked = True
                    except Exception as e:
                        self.log(f"   → JavaScript를 통한 확정모드 선택 실패: {str(e)}", "WARNING")
                
                if not confirm_mode_clicked:
                    self.log("   ⚠ 확정모드 전환 실패 - 계속 진행", "WARNING")
                    
            except Exception as e:
                self.log(f"   → 확정모드 전환 중 오류: {str(e)}", "WARNING")

            self.log("\n=== 조회 프로세스 시작 ===")
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

            # 날짜 입력 완료 후 브라우저 줌 레벨을 25%로 변경
            self.log("\n   → 브라우저 줌 레벨을 25%로 변경 중...")
            self.update_status("화면 줌 조정", "줌 레벨 25%로 변경 중...")
            try:
                self.driver.execute_script("document.body.style.zoom='0.25'")
                self.log("   ✓ 브라우저 줌 레벨 25% 설정 완료")
                time.sleep(1.0)  # 줌 적용 대기
                self.log("   ✓ 줌 레벨 변경 적용 완료")
            except Exception as e:
                self.log(f"   → 줌 레벨 설정 중 오류: {str(e)}")

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
            data, columns = self.wait_for_data_and_parse()
            
            # 데이터 유무 확인 및 메시지 표시
            if data and len(data) > 0:
                self.log(f"\n✓ 조회 완료: {len(data)}개의 데이터를 불러왔습니다", "SUCCESS")
            elif data is not None and len(data) == 0:
                self.log("\nℹ 조회 완료: 데이터가 없습니다", "INFO")
                messagebox.showinfo("알림", "조회가 완료되었습니다.\n조회된 데이터가 없습니다.")

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
        """로딩 완료를 기다립니다 (적용 버튼 없이도 작동)."""
        try:
            print("\n=== 로딩 감지 프로세스 시작 ===")
            
            # 적용 버튼 찾기 시도 (없어도 계속 진행)
            try:
                apply_button = WebDriverWait(self.driver, 3).until(
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
                
            except:
                print("   → 적용 버튼을 찾을 수 없음 (정상적인 상황일 수 있음)")
            
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
            print(f"!!! 로딩 대기 중 오류: {str(e)}")
            # 오류가 발생해도 계속 진행 (로딩이 이미 완료되었을 수 있음)
            return True

    def simple_wait_for_loading(self, timeout=10):
        """간단한 로딩 대기 (로딩 이미지만 확인)"""
        try:
            print("   → 로딩 이미지 확인 중...")
            
            # 로딩 이미지가 있는지 확인
            try:
                loading_element = WebDriverWait(self.driver, 2).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "dxgvLoadingPanel_DevEx"))
                )
                print("   ✓ 로딩 이미지 감지됨")
                
                # 로딩 이미지가 사라질 때까지 대기
                WebDriverWait(self.driver, timeout).until_not(
                    EC.presence_of_element_located((By.CLASS_NAME, "dxgvLoadingPanel_DevEx"))
                )
                print("   ✓ 로딩 완료 감지")
            except:
                print("   → 로딩 이미지 없음 (이미 로딩 완료)")
            
            # 짧은 대기 시간
            time.sleep(1)
            return True
            
        except Exception as e:
            print(f"   → 로딩 대기 중 오류 (무시): {str(e)}")
            return True

    def wait_for_data_and_parse(self):
        """데이터를 파싱하고 표시합니다."""
        loading_window = None
        try:
            print("\n=== 데이터 파싱 프로세스 시작 ===")
            
            # 로딩 창 생성
            try:
                loading_window = LoadingWindow(self)
                loading_window.update_status("데이터 로딩 중...")
                print("1. 로딩 창 생성됨")
            except Exception as e:
                print(f"   → 로딩 창 생성 중 오류 (계속 진행): {str(e)}")
            
            # 데이터 그리드가 로드될 때까지 대기
            print("\n2. 데이터 그리드 로딩 대기 중...")
            self.simple_wait_for_loading()
            if loading_window:
                try:
                    loading_window.update_progress(0.2, "데이터 확인 중...")
                except:
                    pass
            print("   ✓ 데이터 그리드 로딩 완료")
            
            # No data 확인
            print("\n3. 데이터 존재 여부 확인 중...")
            try:
                no_data_div = self.driver.find_elements(By.CLASS_NAME, "dxgvFCW")
                for div in no_data_div:
                    if "No data" in div.text:
                        print("   → No data 발견 - 데이터가 없습니다")
                        if loading_window:
                            try:
                                loading_window.destroy()
                            except:
                                pass
                        self.log("   ℹ 조회된 데이터가 없습니다", "INFO")
                        return [], []
            except Exception as e:
                print(f"   → No data 확인 중 오류 (계속 진행): {str(e)}")
            
            print("\n4. 데이터 크롤링 준비...")
            
            # 크롤링 테스트 먼저 실행
            if not self.test_crawling():
                print("   !!! 크롤링 테스트 실패!")
                if loading_window:
                    try:
                        loading_window.destroy()
                    except:
                        pass
                self.search_btn.configure(state="normal")
                messagebox.showerror("오류", "데이터 크롤링 테스트에 실패했습니다.")
                return [], []
            
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
            
            # 크롤링 시작 (줌은 이미 25%로 설정됨)
            if loading_window:
                try:
                    loading_window.update_progress(0.35, f"데이터 크롤링 중... (총 {total_rows}행)")
                except:
                    pass
            print(f"   → 총 {total_rows}개 행 크롤링 시작")
            
            # 실제 데이터 행 수만큼만 처리
            for row_idx in range(total_rows):
                try:
                    row_id = f"ContentPlaceHolder1_ASPxSplitter1_dxGrid2_DXDataRow{row_idx}"
                    print(f"\n   → 행 {row_idx + 1}/{total_rows} 처리 중... (ID: {row_id})")
                    
                    row = WebDriverWait(self.driver, 2).until(
                        EC.presence_of_element_located((By.ID, row_id))
                    )
                    
                    # 모든 td 요소 찾기
                    cells = row.find_elements(By.TAG_NAME, "td")
                    print(f"     → 발견된 셀 수: {len(cells)}")
                    
                    # 실제 컬럼 수에 맞게 데이터 배열 초기화 (55개 컬럼)
                    row_data = [""] * len(cells)
                    bg_color = None
                    
                    # 디버깅: 각 셀의 정보 출력 (처음 5개만)
                    for i, cell in enumerate(cells[:5]):
                        try:
                            cell_text = cell.text.strip()
                            print(f"       셀 {i}: '{cell_text}'")
                        except:
                            print(f"       셀 {i}: [텍스트 추출 실패]")
                    
                    if len(cells) > 5:
                        print(f"       ... 총 {len(cells)}개 셀 중 처음 5개만 표시")
                    
                    # 각 셀 처리 - 더 안정적인 방식으로 변경
                    for col_idx, cell in enumerate(cells):
                        try:
                            # 셀 텍스트 추출 (더 안정적인 방법)
                            cell_text = ""
                            try:
                                cell_text = cell.text.strip()
                            except:
                                try:
                                    cell_text = cell.get_attribute("innerText").strip()
                                except:
                                    cell_text = ""
                            
                            row_data[col_idx] = cell_text
                            
                            # 첫 번째 셀(실적상태)의 배경색 확인
                            if col_idx == 1:
                                try:
                                    bg_color = cell.value_of_css_property("background-color")
                                    print(f"     ✓ 실적상태: {cell_text} (배경색: {bg_color})")
                                except:
                                    print(f"     ✓ 실적상태: {cell_text}")
                            elif col_idx < 5:  # 처음 5개만 디버깅 출력
                                print(f"     ✓ 열 {col_idx}: '{cell_text}'")
                                
                        except Exception as e:
                            print(f"     !!! 열 {col_idx} 처리 중 오류: {str(e)}")
                            row_data[col_idx] = ""  # 오류 시 빈 문자열로 설정
                            continue
                    
                    all_data.append((row_data, bg_color))
                    print(f"     ✓ 행 {row_idx + 1} 데이터 저장 완료")
                    if loading_window:
                        try:
                            loading_window.update_progress((row_idx + 1) / total_rows)
                        except:
                            pass
                    
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
            if loading_window:
                try:
                    loading_window.destroy()
                    print("   ✓ 로딩 창 제거됨")
                except Exception as e:
                    print(f"   → 로딩 창 제거 중 오류 (무시): {str(e)}")
            
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
            
            # 크롤링 완료 후에도 줌 25% 유지 (다음 조회 시 복구됨)
            print("   ✓ 브라우저 줌 레벨 25% 유지 (다음 조회 시 복구)")
            
            self.search_btn.configure(state="normal")
            print("   ✓ 조회 버튼 다시 활성화됨")
            
            return all_data, self.column_mapping.keys()
            
        except Exception as e:
            print("\n!!! 데이터 파싱 중 오류 발생 !!!")
            print(str(e))
            traceback.print_exc()
            
            # 오류 발생 시에도 줌 25% 유지 (다음 조회 시 복구됨)
            print("   ✓ 오류 발생 후에도 브라우저 줌 레벨 25% 유지")
            
            # 로딩창 안전하게 닫기
            try:
                if 'loading_window' in locals() and loading_window:
                    loading_window.destroy()
            except Exception as close_error:
                print(f"   → 로딩창 닫기 중 오류 (무시): {str(close_error)}")
            
            # 버튼 상태 복구
            try:
                self.search_btn.configure(state="normal")
                print("   ✓ 오류 발생 후 조회 버튼 다시 활성화됨")
            except Exception as btn_error:
                print(f"   → 버튼 상태 복구 중 오류 (무시): {str(btn_error)}")
            
            return None, None

    def test_crawling(self):
        """크롤링 테스트 함수"""
        try:
            print("\n=== 크롤링 테스트 시작 ===")
            
            # 메인 테이블 찾기
            main_table = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, "ContentPlaceHolder1_ASPxSplitter1_dxGrid2_DXMainTable"))
            )
            print("✓ 메인 테이블 발견")
            
            # 데이터 행 찾기
            data_rows = main_table.find_elements(By.CLASS_NAME, "dxgvDataRow")
            print(f"✓ {len(data_rows)}개의 데이터 행 발견")
            
            if len(data_rows) > 0:
                # 첫 번째 행 테스트
                first_row = data_rows[0]
                cells = first_row.find_elements(By.TAG_NAME, "td")
                print(f"✓ 첫 번째 행에서 {len(cells)}개의 셀 발견")
                
                # 처음 10개 셀의 텍스트 출력
                for i, cell in enumerate(cells[:10]):
                    try:
                        text = cell.text.strip()
                        print(f"  셀 {i}: '{text}'")
                    except Exception as e:
                        print(f"  셀 {i}: 오류 - {str(e)}")
                
                print("✓ 크롤링 테스트 완료")
                return True
            else:
                print("! 데이터 행이 없습니다")
                return False
                
        except Exception as e:
            print(f"! 크롤링 테스트 실패: {str(e)}")
            return False

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
        """특정 열의 너비를 설정 (새로운 구조에서는 사용하지 않음)"""
        print(f"set_column_width 호출됨: 컬럼 {column_index}, 너비 {width}")
        # 새로운 구조에서는 고정 너비를 사용하므로 이 함수는 더 이상 사용하지 않음
        pass

    def update_column_width(self, column_index, width):
        """열 너비만 업데이트 (새로운 구조에서는 사용하지 않음)"""
        print(f"update_column_width 호출됨: 컬럼 {column_index}, 너비 {width}")
        # 새로운 구조에서는 고정 너비를 사용하므로 이 함수는 더 이상 사용하지 않음
        pass

    def refresh_grid(self):
        """그리드 전체 새로 그리기 (새로운 구조에서는 사용하지 않음)"""
        print("refresh_grid 호출됨 - 새로운 구조에서는 사용하지 않음")
        # 새로운 구조에서는 고정 레이아웃을 사용하므로 이 함수는 더 이상 사용하지 않음
        pass
            
    def set_column_alignment(self, column_index, header_alignment, data_alignment=None):
        """열의 정렬 방식을 설정 (새로운 구조에서는 사용하지 않음)"""
        print(f"set_column_alignment 호출됨: 컬럼 {column_index}, 헤더 정렬 {header_alignment}, 데이터 정렬 {data_alignment}")
        # 새로운 구조에서는 고정 정렬을 사용하므로 이 함수는 더 이상 사용하지 않음
        pass
            



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





    def save_config(self):
        """설정 파일 저장"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"설정 저장 중 오류: {e}")

if __name__ == "__main__":
    # 테스트용으로 실행할 때 사용
    root = ctk.CTk()
    root.title("MES 비가동관리 테스트")
    root.geometry("400x300")
    
    # 테스트용 버튼
    test_btn = ctk.CTkButton(root, text="비가동관리 창 열기", 
                           command=lambda: DTHistory(root))
    test_btn.pack(pady=50)
    
    root.mainloop()






