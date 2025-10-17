import os
import json
import requests
import sys
import shutil
import zipfile
import io
from tkinter import messagebox
import customtkinter as ctk
from datetime import datetime
import urllib3

class UpdateManager:
    # 클래스 변수로 변경
    current_version = "2.1.3"

    def __init__(self):
        """업데이트 매니저 초기화"""
        # 인스턴스 변수에서 클래스 변수로 이동했으므로 제거
        # self.current_version = "2.1.3"  # 현재 버전
        self.github_user = "hajunyu"    # GitHub 사용자 이름
        self.github_repo = "mes_loader" # GitHub 저장소 이름
        self.config_file = "mes_config.json"
        # SSL 경고 메시지 비활성화
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        self.mes_modules = [
            'mes_downloader.py',
            'mes_common.py',
            'mes_unbalance.py',
            'mes_updater.py',
            'mes_register.py',
            'mes_split.py',
            'mes_uploader.py' ,
            'mes_DTHistory.py'
        ]
        self.current_dir = os.path.dirname(os.path.abspath(__file__))
        
    def check_for_updates(self):
        """GitHub에서 최신 버전 확인"""
        try:
            # GitHub API를 통해 version.txt 파일 내용 가져오기
            version_url = f"https://raw.githubusercontent.com/{self.github_user}/{self.github_repo}/main/version.txt"
            print("\n=== GitHub 업데이트 확인 ===")
            print(f"버전 확인 URL: {version_url}")
            print(f"현재 버전: {self.current_version}")
            
            # API 요청 헤더 추가
            headers = {
                'Accept': 'application/vnd.github.v3.raw',
                'User-Agent': 'mes_updater'
            }
            
            # SSL 검증을 비활성화하고 요청
            response = requests.get(version_url, headers=headers, verify=False)
            
            if response.status_code == 200:
                latest_version = response.text.strip()
                print(f"최신 버전: {latest_version}")
                
                # 버전 비교
                if self._compare_versions(latest_version, self.current_version) > 0:
                    return True, latest_version, "새로운 버전이 있습니다."
                
                return False, latest_version, ""
            else:
                print(f"\n오류: 버전 정보를 가져올 수 없습니다. (상태 코드: {response.status_code})")
                return False, self.current_version, ""
            
        except Exception as e:
            print(f"\n오류: 업데이트 확인 중 예외 발생\n{str(e)}")
            return False, self.current_version, ""
    
    def _compare_versions(self, version1, version2):
        """버전 문자열 비교"""
        v1_parts = list(map(int, version1.split(".")))
        v2_parts = list(map(int, version2.split(".")))
        
        for i in range(max(len(v1_parts), len(v2_parts))):
            v1 = v1_parts[i] if i < len(v1_parts) else 0
            v2 = v2_parts[i] if i < len(v2_parts) else 0
            
            if v1 > v2:
                return 1
            elif v1 < v2:
                return -1
        
        return 0
    
    def download_update(self, version):
        """업데이트 다운로드 및 적용"""
        try:
            # GitHub에서 직접 exe 파일 다운로드
            exe_name = "MES_WEB_REPORT.exe"
            download_url = f"https://github.com/{self.github_user}/{self.github_repo}/blob/main/{exe_name}?raw=true"
            print(f"다운로드 URL: {download_url}")
            
            # 현재 실행 중인 exe 파일명 확인
            current_exe = sys.executable if getattr(sys, 'frozen', False) else None
            if not current_exe or not os.path.exists(current_exe):
                raise Exception("현재 실행 중인 exe 파일을 찾을 수 없습니다.")
            
            print(f"현재 실행 중인 exe 파일: {current_exe}")
            
            # 새 버전 다운로드
            print("새 버전 다운로드 중...")
            response = requests.get(download_url, verify=False, stream=True)
            
            if response.status_code != 200:
                raise Exception(f"다운로드 실패 (상태 코드: {response.status_code})")
            
            # 파일 크기 확인
            total_size = int(response.headers.get('content-length', 0))
            print(f"다운로드 크기: {total_size:,} bytes")
            
            if total_size < 1000000:  # 1MB 미만이면 의심스러운 파일
                raise Exception("다운로드된 파일이 너무 작습니다. 올바른 exe 파일이 아닐 수 있습니다.")
            
            # 임시 파일로 다운로드 (버전 번호 포함)
            file_name, file_ext = os.path.splitext(current_exe)
            temp_exe = f"{file_name}_v{version}{file_ext}"
            try:
                with open(temp_exe, 'wb') as f:
                    for chunk in response.iter_content(chunk_size=8192):
                        if chunk:
                            f.write(chunk)
                print("다운로드 완료")
                
                # 파일이 실제로 존재하는지 확인
                if not os.path.exists(temp_exe):
                    raise Exception("다운로드된 파일이 없습니다.")
                
                # 파일 크기 확인
                if os.path.getsize(temp_exe) != total_size:
                    raise Exception("다운로드된 파일이 손상되었습니다.")
                
                # 현재 파일 백업
                backup_exe = current_exe + '.bak'
                shutil.copy2(current_exe, backup_exe)
                print(f"현재 버전 백업 완료: {backup_exe}")
                
                # 배치 파일 생성
                update_bat = os.path.join(os.path.dirname(current_exe), 'update.bat')
                with open(update_bat, 'w', encoding='utf-8') as bat:
                    bat.write('@echo off\n')
                    bat.write('timeout /t 1 /nobreak >nul\n')
                    bat.write(f'move /y "{temp_exe}" "{current_exe}"\n')
                    bat.write(f'del "{backup_exe}"\n')
                    bat.write('del "%~f0"\n')
                    bat.write(f'start "" "{current_exe}"\n')
                
                print("업데이트 준비 완료")
                os.startfile(update_bat)
                print("프로그램을 종료합니다...")
                sys.exit(0)  # 프로그램 즉시 종료
                return True
                
            except Exception as e:
                print(f"파일 처리 중 오류: {e}")
                if os.path.exists(temp_exe):
                    os.remove(temp_exe)
                raise
            
        except Exception as e:
            error_msg = str(e)
            print(f"업데이트 실패: {error_msg}")
            messagebox.showerror("업데이트 오류", error_msg)
            return False

class UpdateDialog(ctk.CTkToplevel):
    def __init__(self, parent, current_version, latest_version, release_notes):
        super().__init__(parent)
        
        self.title("업데이트 확인")
        self.geometry("400x300")
        
        # 업데이트 정보 표시
        info_text = f"현재 버전: {current_version}\n"
        info_text += f"최신 버전: {latest_version}\n\n"
        info_text += "업데이트 내용:\n"
        info_text += release_notes
        
        info_label = ctk.CTkLabel(self, text=info_text, wraplength=350)
        info_label.pack(pady=20)
        
        # 업데이트 선택
        self.update_choice = False
        
        def update_now():
            self.update_choice = True
            self.destroy()
        
        def update_later():
            self.update_choice = False
            self.destroy()
        
        # 버튼 프레임
        button_frame = ctk.CTkFrame(self)
        button_frame.pack(pady=20)
        
        update_btn = ctk.CTkButton(button_frame, text="지금 업데이트",
                                 command=update_now)
        update_btn.pack(side="left", padx=10)
        
        later_btn = ctk.CTkButton(button_frame, text="나중에",
                               command=update_later)
        later_btn.pack(side="left", padx=10)
        
        # 모달 창으로 설정
        self.transient(parent)
        self.grab_set()
        parent.wait_window(self)

class UpdaterApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title("MES 업데이트")
        self.geometry("400x250")
        
        # 업데이트 매니저 초기화
        self.update_manager = UpdateManager()
        
        # UI 구성
        self.create_widgets()
    
    def create_widgets(self):
        # 현재 버전 표시
        version_frame = ctk.CTkFrame(self)
        version_frame.pack(pady=20, padx=20, fill="x")
        
        version_label = ctk.CTkLabel(version_frame, 
                                   text=f"현재 버전: {self.update_manager.current_version}",
                                   font=("Arial", 14))
        version_label.pack(pady=10)
        
        # 업데이트 확인 버튼
        check_button = ctk.CTkButton(self, 
                                   text="업데이트 확인",
                                   font=("Arial", 12),
                                   width=200,
                                   height=40,
                                   command=self.check_for_updates)
        check_button.pack(pady=20)
        
        # 상태 메시지 표시 (2줄로 표시)
        self.status_label = ctk.CTkLabel(self, 
                                       text="업데이트 확인을 눌러주세요.",
                                       font=("Arial", 12),
                                       wraplength=350)
        self.status_label.pack(pady=10)
        
        # 상세 상태 메시지
        self.detail_label = ctk.CTkLabel(self,
                                       text="",
                                       font=("Arial", 10),
                                       text_color="gray",
                                       wraplength=350)
        self.detail_label.pack(pady=5)
    
    def update_status(self, status, detail=""):
        self.status_label.configure(text=status)
        self.detail_label.configure(text=detail)
        self.update()  # UI 즉시 업데이트
    
    def check_for_updates(self):
        try:
            self.update_status("GitHub에서 업데이트 확인 중...")
            has_update, latest_version, release_notes = self.update_manager.check_for_updates()
            
            if has_update:
                self.update_status(
                    f"새 버전 {latest_version}이 있습니다.",
                    f"현재 버전: {self.update_manager.current_version}"
                )
                response = messagebox.askyesno("업데이트 확인", 
                                             f"새로운 버전 {latest_version}이 있습니다.\n"
                                             "지금 업데이트하시겠습니까?")
                if response:
                    self.update_status("업데이트 파일 다운로드 중...", "잠시만 기다려주세요.")
                    if self.update_manager.download_update(latest_version):
                        self.update_status("업데이트 완료!", "프로그램이 곧 재시작됩니다.")
                        self.after(2000, self.quit)  # 2초 후 종료
                    else:
                        self.update_status("업데이트 실패", "GitHub에서 파일을 다운로드할 수 없습니다.")
            else:
                self.update_status(
                    "현재 최신 버전입니다.",
                    f"현재 버전: {self.update_manager.current_version}\n"
                    f"최신 버전: {latest_version}"
                )
        except Exception as e:
            self.update_status(
                "업데이트 확인 실패",
                f"오류: {str(e)}"
            )

def run_updater():
    """업데이트 앱을 실행하는 함수"""
    ctk.set_appearance_mode("dark")
    app = UpdaterApp()
    app.mainloop()

if __name__ == "__main__":
    run_updater() 