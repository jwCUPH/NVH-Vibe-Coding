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
        """Testlab 연결 및 Database 핸들 확보 (Automation_260223.py 방식)"""
        prog_id = "LMSTestLabAutomation.Application"
        try:
            self.log(f"Attempting to connect to '{prog_id}'...")
            self.app_com = win32com.client.Dispatch(prog_id)
            
            # Database 핸들 확보 로직
            db_candidate = None
            try:
                book = self.app_com.ActiveBook
                if book: db_candidate = book.Database
            except: pass

            if db_candidate is None:
                try:
                    proj = self.app_com.ActiveProject
                    if proj: db_candidate = proj.Database
                except: pass

            if db_candidate is None:
                raise RuntimeError("Database 핸들을 찾지 못했습니다. 프로젝트가 열려 있는지 확인하세요.")

            self.db = db_candidate
            self.log("Successfully connected to Testlab Database.")

            # 프로젝트명 표시
            try:
                name = self.app_com.ActiveBook.Name
            except:
                name = "Active Project"
            self.lbl_status.configure(text=f"Status: Connected ({name})", text_color="lightgreen")
            
            self.update_sections()
            self.btn_extract.configure(state="normal")
                
        except Exception as e:
            self.log(f"Connection failed: {str(e)}")
            self.lbl_status.configure(text="Status: Error", text_color="red")
            messagebox.showerror("Connection Error", f"Failed to connect:\n{str(e)}")

    def to_string_list(self, maybe_list):
        """COM 리스트 객체를 Python 문자열 리스트로 변환"""
        if maybe_list is None: return []
        try:
            return [str(x) for x in list(maybe_list)]
        except:
            try:
                out = []
                n = int(maybe_list.Count)
                for i in range(1, n + 1):
                    out.append(str(maybe_list.Item(i)))
                return out
            except:
                return []

    def update_sections(self):
        """db.SectionNames를 사용하여 섹션 목록을 가져옵니다."""
        try:
            self.log("Loading section names from Database...")
            # Automation_260223.py 방식: db.SectionNames 사용
            sections = self.to_string_list(self.db.SectionNames)
            
            if not sections:
                self.log("Warning: No sections found via db.SectionNames.")
                # 폴백
                try:
                    active = self.app_com.ActiveBook.ActiveSectionName
                    if active: sections = [active]
                except: pass

            if not sections:
                raise AttributeError("섹션 정보를 가져올 수 없습니다.")

            self.combo_section.configure(values=sections)
            self.combo_section.set(sections[0])
            self.on_section_selected(sections[0])
            self.log(f"Successfully loaded {len(sections)} sections.")
            
        except Exception as e:
            self.log(f"Section Loading Failed: {str(e)}")

    def on_section_selected(self, section_name):
        """db.ElementNames 및 db.ElementType을 사용하여 Run 목록을 가져옵니다."""
        try:
            run_names = []
            self.log(f"Loading runs for section: {section_name}")

            # Automation_260223.py 방식: ElementNames 가져온 후 ElementType으로 필터링
            elements = self.to_string_list(self.db.ElementNames(section_name))
            for item in elements:
                try:
                    path = f"{section_name}/{item}"
                    if str(self.db.ElementType(path)) == "Run":
                        run_names.append(item)
                except:
                    continue

            if not run_names:
                # 폴백: ActiveRunName
                try:
                    active_run = self.app_com.ActiveBook.ActiveRunName
                    if active_run: run_names = [active_run]
                except: pass

            if run_names:
                run_names = sorted(list(set(run_names)))
                self.combo_run.configure(values=run_names)
                self.combo_run.set(run_names[0])
                self.log(f"Successfully loaded {len(run_names)} runs.")
            else:
                self.combo_run.configure(values=["No runs found"])

        except Exception as e:
            self.log(f"Run Loading Error: {str(e)}")

    def extract_data(self):
        """db.GetItem 및 IBlock2를 사용하여 데이터를 추출합니다."""
        section_name = self.combo_section.get()
        run_name = self.combo_run.get()
        
        channels = ["AutoPower FLW (A)", "AutoPower FLI (A)", "AutoPower RCC (A)", "AutoPower RRW (A)"]
        extracted_results = {}
        frequencies = None

        target_sub_path = "Fixed sampling/Stationary Free run/Sections/Map statistics/Spectrum averaged"
        self.log(f"Extracting data from {section_name}/{run_name}...")
        
        try:
            for channel in channels:
                full_path = f"{section_name}/{run_name}/{target_sub_path}/{channel}"
                
                try:
                    # Automation_260223.py 방식: db.GetItem 사용
                    obj = self.db.GetItem(full_path)
                    if obj:
                        # IBlock2 인터페이스로 캐스팅 (win32com.client가 자동으로 처리하지 못할 경우 대비)
                        try:
                            # 1. X축 데이터 (주파수)
                            if frequencies is None:
                                x_vals = obj.XValues
                                frequencies = [x_vals.Value(i) for i in range(1, x_vals.Count + 1)]
                            
                            # 2. Y축 데이터 (진폭)
                            y_vals = obj.YValues
                            amplitudes = [y_vals.Value(i) for i in range(1, y_vals.Count + 1)]
                            
                            extracted_results[channel] = amplitudes
                            self.log(f"Extracted: {channel}")
                        except Exception as de:
                            self.log(f"Block access failed for {channel}: {de}")
                    else:
                        self.log(f"Item not found: {channel}")
                except Exception as e:
                    self.log(f"Error accessing {channel}: {e}")

            if not extracted_results:
                messagebox.showwarning("Warning", "추출된 데이터가 없습니다. 경로를 확인하세요.")
                return

            self.save_to_csv(frequencies, extracted_results, section_name, run_name)

        except Exception as e:
            self.log(f"Extraction critical error: {e}")
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
