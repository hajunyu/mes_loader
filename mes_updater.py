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
    def __init__(self):
        self.current_version = "2.1.1"  # 현재 버전
        self.github_user = "hajunyu"    # GitHub 사용자 이름
        self.github_repo = "mes_loader" # GitHub 저장소 이름
        self.config_file = "mes_config.json"
        
    def check_for_updates(self):
        """GitHub에서 최신 버전 확인"""
        try:
            # GitHub API를 통해 최신 릴리즈 정보 가져오기
            api_url = f"https://api.github.com/repos/{self.github_user}/{self.github_repo}/releases/latest"
            print("\n=== GitHub 업데이트 확인 ===")
            print(f"API URL: {api_url}")
            print(f"현재 버전: {self.current_version}")
            
            # API 요청 헤더 추가
            headers = {
                'Accept': 'application/vnd.github.v3+json',
                'User-Agent': 'mes_updater'
            }
            print("API 요청 시작...")
            
            # SSL 검증을 비활성화하고 요청
            response = requests.get(api_url, headers=headers, timeout=10, verify=False)
            
            # SSL 경고 메시지 무시
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
            
            print(f"\nAPI 응답 상태 코드: {response.status_code}")
            print(f"API 응답 헤더:\n{json.dumps(dict(response.headers), indent=2)}")
            
            if response.status_code == 200:
                release_info = response.json()
                print(f"\nAPI 응답 데이터:\n{json.dumps(release_info, indent=2, ensure_ascii=False)}")
                
                if "tag_name" not in release_info:
                    print("\n오류: 릴리스 정보에 tag_name이 없습니다")
                    return False, self.current_version, ""
                
                latest_version = release_info["tag_name"].replace("v", "").strip()
                print(f"\n최신 버전: {latest_version}")
                
                # 버전 비교
                comparison = self._compare_versions(latest_version, self.current_version)
                print(f"버전 비교 결과: {comparison} (양수: 업데이트 필요, 0: 최신, 음수: 현재가 더 높은 버전)")
                
                if comparison > 0:
                    return True, latest_version, release_info.get("body", "릴리스 노트가 없습니다.")
                
                return False, latest_version, ""
            elif response.status_code == 404:
                print("\n오류: 저장소를 찾을 수 없거나 릴리스가 없습니다")
                print(f"응답 내용:\n{response.text}")
                return False, self.current_version, ""
            elif response.status_code == 403:
                print("\n오류: API 사용량 제한 초과 또는 접근이 거부되었습니다")
                print(f"응답 내용:\n{response.text}")
                return False, self.current_version, ""
            else:
                print(f"\n오류: 예상치 못한 응답 코드 {response.status_code}")
                print(f"응답 내용:\n{response.text}")
                raise Exception(f"GitHub API 응답 오류: {response.status_code}")
            
        except requests.exceptions.Timeout:
            print("\n오류: GitHub API 요청 시간 초과")
            return False, self.current_version, ""
        except requests.exceptions.RequestException as e:
            print(f"\n오류: 네트워크 문제 발생\n{str(e)}")
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
            # GitHub API를 통해 릴리즈 정보 가져오기
            api_url = f"https://api.github.com/repos/{self.github_user}/{self.github_repo}/releases/latest"
            
            # API 요청 헤더 추가
            headers = {
                'Accept': 'application/vnd.github.v3+json',
                'User-Agent': 'mes_updater'
            }
            
            # SSL 검증 비활성화하여 요청
            response = requests.get(api_url, headers=headers, verify=False)
            
            if response.status_code == 200:
                release_info = response.json()
                
                # assets에서 다운로드 URL 찾기
                if 'assets' in release_info and release_info['assets']:
                    # 첫 번째 asset의 다운로드 URL 사용
                    zip_url = release_info['assets'][0]['browser_download_url']
                else:
                    # assets가 없으면 소스코드 ZIP 사용
                    zip_url = release_info["zipball_url"]
                    print("경고: Release assets를 찾을 수 없어 소스코드 ZIP을 사용합니다.")
                
                print(f"다운로드 URL: {zip_url}")
                
                # 임시 디렉토리 생성
                temp_dir = "temp_update"
                if os.path.exists(temp_dir):
                    shutil.rmtree(temp_dir)
                os.makedirs(temp_dir)
                
                try:
                    # zip 파일 다운로드
                    print("업데이트 파일 다운로드 중...")
                    # SSL 검증 비활성화하여 ZIP 파일 다운로드
                    response = requests.get(zip_url, verify=False)
                    
                    if response.status_code == 200:
                        # ZIP 파일 압축 해제
                        print("압축 파일 해제 중...")
                        zip_content = zipfile.ZipFile(io.BytesIO(response.content))
                        zip_content.extractall(temp_dir)
                        
                        # 압축 해제된 첫 번째 디렉토리 찾기 (GitHub ZIP의 경우 하나의 루트 디렉토리 포함)
                        extracted_dir = os.path.join(temp_dir, os.listdir(temp_dir)[0])
                        
                        # 현재 설정 파일 백업
                        if os.path.exists(self.config_file):
                            with open(self.config_file, "r", encoding="utf-8") as f:
                                current_config = json.load(f)
                        
                        # 파일 업데이트
                        print("파일 업데이트 중...")
                        current_dir = os.path.dirname(os.path.abspath(__file__))
                        for root, dirs, files in os.walk(extracted_dir):
                            for file in files:
                                if file.endswith('.py'):  # Python 파일만 업데이트
                                    src_file = os.path.join(root, file)
                                    dst_file = os.path.join(current_dir, file)
                                    shutil.copy2(src_file, dst_file)
                        
                        # 설정 파일 복원
                        if os.path.exists(self.config_file):
                            with open(self.config_file, "w", encoding="utf-8") as f:
                                json.dump(current_config, f, ensure_ascii=False, indent=4)
                        
                        # 임시 파일 정리
                        shutil.rmtree(temp_dir)
                        
                        messagebox.showinfo("업데이트 완료", 
                                          f"버전 {version}으로 업데이트가 완료되었습니다.\n프로그램을 다시 시작해주세요.")
                        sys.exit(0)  # 프로그램 종료
                
                except Exception as e:
                    print(f"업데이트 적용 중 오류: {str(e)}")
                    if os.path.exists(temp_dir):
                        shutil.rmtree(temp_dir)
                    raise e
            
            return False
            
        except Exception as e:
            print(f"업데이트 다운로드 중 오류: {str(e)}")
            messagebox.showerror("업데이트 오류", f"업데이트 중 오류가 발생했습니다:\n{str(e)}")
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
        self.geometry("300x150")
        
        # 업데이트 매니저 초기화
        self.update_manager = UpdateManager()
        
        # 업데이트 확인 시작
        self.check_for_updates()
    
    def check_for_updates(self):
        try:
            has_update, latest_version, release_notes = self.update_manager.check_for_updates()
            
            if has_update:
                dialog = UpdateDialog(self, self.update_manager.current_version, latest_version, release_notes)
                if dialog.update_choice:
                    self.update_manager.download_update(latest_version)
                self.quit()
            else:
                messagebox.showinfo("업데이트 확인", 
                                  f"현재 최신 버전을 사용 중입니다.\n\n"
                                  f"현재 버전: {self.update_manager.current_version}\n"
                                  f"GitHub 버전: {latest_version}")
                self.quit()
        except Exception as e:
            messagebox.showerror("오류", f"업데이트 확인 중 오류가 발생했습니다.\n\n{str(e)}")
            self.quit()

if __name__ == "__main__":
    ctk.set_appearance_mode("dark")
    app = UpdaterApp()
    app.mainloop() 