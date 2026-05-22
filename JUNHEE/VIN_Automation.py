import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
import win32com.client
import logging
from datetime import datetime

# ==========================================
# Hankook Company Style & Vibe Coding Rules
# ==========================================

# GUI 설정: 한국컴퍼니 브랜드 컬러 (예시: 오렌지 & 블랙/다크 그레이)
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue") # 기본 테마

class VINAutomationApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # 로그 설정
        self.log_path = self.setup_logging()
        self.app_com = None
        self.project = None

        # 윈도우 설정
        self.title("HANKOOK COMPANY - VIN Automation Tool")
        self.geometry("800x500")

        # 그리드 구성
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # 상단 로고/제목 바
        self.header_frame = ctk.CTkFrame(self, height=60, corner_radius=0, fg_color="#FF6600")
        self.header_frame.grid(row=0, column=0, sticky="nsew")
        self.header_label = ctk.CTkLabel(self.header_frame, text="VIN AUTOMATION SYSTEM", 
                                        font=ctk.CTkFont(size=20, weight="bold"), text_color="white")
        self.header_label.pack(pady=15)

        # 메인 컨텐츠 영역
        self.main_frame = ctk.CTkFrame(self, corner_radius=10)
        self.main_frame.grid(row=1, column=0, padx=20, pady=20, sticky="nsew")
        self.main_frame.grid_columnconfigure(0, weight=1)

        # 1. Testlab 연결 섹션
        self.conn_frame = ctk.CTkFrame(self.main_frame)
        self.conn_frame.grid(row=0, column=0, padx=20, pady=10, sticky="ew")
        
        self.btn_connect = ctk.CTkButton(self.conn_frame, text="Connect to Active Testlab", command=self.connect_testlab, fg_color="green", hover_color="#006400")
        self.btn_connect.pack(side="left", padx=10, pady=10)
        
        self.lbl_status = ctk.CTkLabel(self.conn_frame, text="Status: Disconnected", text_color="red")
        self.lbl_status.pack(side="left", padx=10)

        # 2. 데이터 선택 (Section & Run)
        self.data_frame = ctk.CTkFrame(self.main_frame)
        self.data_frame.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")
        self.data_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(self.data_frame, text="Select Section:").grid(row=0, column=0, padx=10, pady=10)
        self.combo_section = ctk.CTkComboBox(self.data_frame, values=["Connect first"], command=self.on_section_selected)
        self.combo_section.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

        ctk.CTkLabel(self.data_frame, text="Select Run:").grid(row=1, column=0, padx=10, pady=10)
        self.combo_run = ctk.CTkComboBox(self.data_frame, values=["Select Section first"])
        self.combo_run.grid(row=1, column=1, padx=10, pady=10, sticky="ew")

        # 3. 추출 버튼
        self.btn_extract = ctk.CTkButton(self.main_frame, text="EXTRACT DATA", state="disabled", font=ctk.CTkFont(size=15, weight="bold"), command=self.extract_data)
        self.btn_extract.grid(row=2, column=0, padx=20, pady=20)

        # 하단 로그 요약
        self.log_textbox = ctk.CTkTextbox(self, height=100)
        self.log_textbox.grid(row=2, column=0, padx=20, pady=10, sticky="nsew")
        self.log_textbox.insert("0.0", "Ready to connect to Simcenter Testlab\n")

    def setup_logging(self):
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        log_file = os.path.join(desktop_path, f"VIN_Automation_Log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
        logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(message)s',
                            handlers=[logging.FileHandler(log_file, encoding='utf-8')])
        return log_file

    def log(self, message):
        logging.info(message)
        self.log_textbox.insert("end", f"{datetime.now().strftime('%H:%M:%S')} - {message}\n")
        self.log_textbox.see("end")

    def connect_testlab(self):
        """Testlab 연결 및 ActiveBook 섹션 접근 최적화"""
        prog_id = "LMSTestLabAutomation.Application"
        try:
            self.log(f"Attempting to dispatch '{prog_id}'...")
            # Dispatch로 고정하여 이미 실행 중인 인스턴스에 연결
            self.app_com = win32com.client.Dispatch(prog_id)
            self.log("Successfully dispatched LMSTestLabAutomation.Application.")

            # ActiveBook 확인
            try:
                self.project = self.app_com.ActiveBook
            except Exception:
                self.project = None

            if self.project is None:
                self.log("Error: No Active Project found in Testlab.")
                self.lbl_status.configure(text="Status: Connected (No Project)", text_color="orange")
                messagebox.showerror("Project Error", "Testlab에 활성화된 프로젝트(.lms)가 없습니다.\n프로젝트를 먼저 열어주세요.")
                return

            # 프로젝트 이름 가져오기 시도 (여러 속성 시도)
            project_name = "Unknown"
            try:
                # 1. Name 속성 시도
                project_name = self.project.Name
            except Exception:
                try:
                    # 2. 파일 경로에서 이름 추출 시도
                    full_path = self.project.FullName
                    project_name = os.path.basename(full_path)
                except Exception:
                    project_name = "Connected Project"

            self.lbl_status.configure(text=f"Status: Connected ({project_name})", text_color="lightgreen")
            self.log(f"Active Project identified: {project_name}")
            
            self.update_sections()
            self.btn_extract.configure(state="normal")
                
        except Exception as e:
            error_details = str(e)
            self.log(f"Connection failed: {error_details}")
            self.lbl_status.configure(text="Status: Error", text_color="red")
            messagebox.showerror("Connection Error", f"Failed to connect:\n{error_details}")

    def update_sections(self):
        try:
            # Sections 컬렉션 접근
            sections_obj = self.project.Sections
            count = sections_obj.Count
            self.log(f"Found {count} sections in project.")
            
            if count == 0:
                self.log("Warning: Project contains 0 sections.")
                self.combo_section.configure(values=["No sections found"])
                return

            section_names = []
            for i in range(1, count + 1):
                section_names.append(sections_obj.Item(i).Name)
            
            self.combo_section.configure(values=section_names)
            self.combo_section.set(section_names[0])
            self.log(f"Sections loaded: {section_names}")
            self.on_section_selected(section_names[0])
            
        except Exception as e:
            self.log(f"Failed to load sections: {str(e)}")
            # 상세 디버깅 정보 로그
            import traceback
            self.log(traceback.format_exc())

    def on_section_selected(self, section_name):
        try:
            section = self.project.Sections.Item(section_name)
            runs_obj = section.Runs
            count = runs_obj.Count
            
            run_names = []
            for i in range(1, count + 1):
                run_names.append(runs_obj.Item(i).Name)
            
            self.combo_run.configure(values=run_names)
            if run_names:
                self.combo_run.set(run_names[0])
            self.log(f"Runs loaded for '{section_name}': {run_names}")
        except Exception as e:
            self.log(f"Failed to load runs: {str(e)}")

    def extract_data(self):
        section_name = self.combo_section.get()
        run_name = self.combo_run.get()
        
        # 추출 대상 채널 리스트
        channels = [
            "AutoPower FLW (A)",
            "AutoPower FLI (A)",
            "AutoPower RCC (A)",
            "AutoPower RRW (A)"
        ]
        
        # 데이터 저장용 딕셔너리 (Key: 채널명, Value: 데이터 리스트)
        extracted_results = {}
        frequencies = None

        # 기본 경로 구성
        target_sub_path = "Fixed sampling/Stationary Free run/Sections/Map statistics/Spectrum averaged"
        
        self.log(f"Starting extraction for Section: {section_name}, Run: {run_name}")
        
        try:
            for channel in channels:
                # 전체 데이터 경로 생성
                # {Section}/{Run}/{SubPath}/{ChannelName}
                full_data_path = f"{section_name}/{run_name}/{target_sub_path}/{channel}"
                
                try:
                    # Testlab COM을 통한 데이터 접근
                    data_item = self.app_com.GetDataItem(full_data_path)
                    
                    if data_item:
                        # 주파수 축(X축) 데이터 가져오기 (첫 채널에서만 1회 수행)
                        if frequencies is None:
                            x_values = data_item.XValues
                            frequencies = [x_values.Value(i) for i in range(1, x_values.Count + 1)]
                        
                        # 진폭(Y축) 데이터 가져오기
                        y_values = data_item.YValues
                        amplitudes = [y_values.Value(i) for i in range(1, y_values.Count + 1)]
                        
                        extracted_results[channel] = amplitudes
                        self.log(f"Successfully extracted: {channel}")
                    else:
                        self.log(f"Warning: Data item not found for {channel}")
                except Exception as e:
                    self.log(f"Failed to extract {channel}: {str(e)}")

            if not extracted_results:
                messagebox.showwarning("Warning", "No data was extracted. Please check the paths.")
                return

            # CSV 파일 저장 (바탕화면)
            self.save_to_csv(frequencies, extracted_results, section_name, run_name)

        except Exception as e:
            self.log(f"Critical error during extraction: {e}")
            messagebox.showerror("Error", f"Extraction failed:\n{e}")

    def save_to_csv(self, frequencies, results, section, run):
        import csv
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_name = f"VIN_Data_{section}_{run}_{timestamp}.csv"
        file_path = os.path.join(desktop_path, file_name)

        try:
            with open(file_path, mode='w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                
                # 헤더 작성 (첫 열은 Frequency, 이후 채널명)
                header = ["Frequency (Hz)"] + list(results.keys())
                writer.writerow(header)
                
                # 데이터 행 작성
                for i in range(len(frequencies)):
                    row = [frequencies[i]]
                    for channel in results.keys():
                        row.append(results[channel][i])
                    writer.writerow(row)
            
            self.log(f"CSV saved successfully: {file_path}")
            messagebox.showinfo("Success", f"Data exported to CSV:\n{file_name}")
        except Exception as e:
            self.log(f"CSV save failed: {e}")
            messagebox.showerror("Error", f"Failed to save CSV:\n{e}")

if __name__ == "__main__":
    app = VINAutomationApp()
    app.mainloop()
