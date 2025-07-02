import os
import json
import requests
import sys
import shutil
from tkinter import messagebox
import customtkinter as ctk
from datetime import datetime

class UpdateManager:
    def __init__(self, current_version="2.1.0"):
        self.current_version = current_version
        self.config_file = "mes_config.json"
        
        # GitHub 저장소 사용자명과 저장소 이름을 본인 것으로 수정
        self.github_username = "hajunyu"  # 예: yhj1129
        self.github_repo = "mes_loader"
        
        # GitHub API URL
        self.github_url = f"https://api.github.com/repos/{self.github_username}/{self.github_repo}/releases/latest"
        
        # 다운로드 URL (GitHub 릴리즈 페이지)
        self.download_base_url = f"https://github.com/{self.github_username}/{self.github_repo}/releases/download"

    def load_config(self):
        """설정 파일에서 버전 정보 로드"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    return config.get("version", self.current_version)
            return self.current_version
        except Exception as e:
            print(f"버전 정보 로드 중 오류: {str(e)}")
            return self.current_version

    def save_version_to_config(self, version):
        """설정 파일에 버전 정보 저장"""
        try:
            config = {}
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
            
            config["version"] = version
            config["last_update_check"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"버전 정보 저장 중 오류: {str(e)}")

    def check_for_updates(self):
        """업데이트 확인"""
        try:
            # GitHub API를 통해 최신 릴리즈 정보 가져오기
            response = requests.get(self.github_url)
            if response.status_code == 200:
                latest_release = response.json()
                latest_version = latest_release["tag_name"].replace("v", "")
                
                current_version_tuple = tuple(map(int, self.current_version.split(".")))
                latest_version_tuple = tuple(map(int, latest_version.split(".")))
                
                if latest_version_tuple > current_version_tuple:
                    return True, latest_version, latest_release["body"]
            return False, self.current_version, ""
        except Exception as e:
            print(f"업데이트 확인 중 오류: {str(e)}")
            return False, self.current_version, ""

    def download_update(self, version):
        """업데이트 다운로드 및 설치"""
        try:
            # 임시 디렉토리에 다운로드
            temp_dir = "temp_update"
            os.makedirs(temp_dir, exist_ok=True)
            
            # GitHub 릴리즈에서 update.zip 다운로드
            download_url = f"{self.download_base_url}/v{version}/update.zip"
            response = requests.get(download_url, stream=True)
            
            if response.status_code != 200:
                raise Exception("업데이트 파일을 찾을 수 없습니다.")
                
            update_file = os.path.join(temp_dir, "update.zip")
            
            with open(update_file, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
            
            # 압축 해제 및 파일 교체
            shutil.unpack_archive(update_file, temp_dir)
            
            # 실행 중인 프로그램 종료 후 업데이트를 위한 배치 파일 생성
            with open("update.bat", "w", encoding='utf-8') as f:
                f.write(f"""@echo off
timeout /t 2 /nobreak
xcopy /y /s "{temp_dir}\\*.*" "."
rmdir /s /q "{temp_dir}"
del update.bat
start "" "{sys.executable}"
exit
""")
            
            # 배치 파일 실행
            os.system("start update.bat")
            sys.exit()
            
        except Exception as e:
            print(f"업데이트 다운로드 중 오류: {str(e)}")
            messagebox.showerror("오류", f"업데이트 다운로드 중 오류가 발생했습니다: {str(e)}")
            return False
        
        return True

class UpdateDialog(ctk.CTkToplevel):
    def __init__(self, parent, current_version, latest_version, release_notes):
        super().__init__(parent)
        
        self.title("업데이트 확인")
        self.geometry("400x300")
        
        # 메시지
        message = f"새로운 버전이 있습니다!\n\n현재 버전: {current_version}\n최신 버전: {latest_version}\n\n업데이트 내용:\n{release_notes}"
        
        ctk.CTkLabel(self, text=message, wraplength=350).pack(pady=20)
        
        # 버튼 프레임
        button_frame = ctk.CTkFrame(self)
        button_frame.pack(pady=20)
        
        ctk.CTkButton(button_frame, text="업데이트", 
                     command=lambda: self.update_response(True)).pack(side="left", padx=10)
        ctk.CTkButton(button_frame, text="나중에", 
                     command=lambda: self.update_response(False)).pack(side="left", padx=10)
        
        self.update_choice = None
        self.wait_window()
    
    def update_response(self, response):
        self.update_choice = response
        self.destroy() 