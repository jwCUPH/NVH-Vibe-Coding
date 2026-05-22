import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import win32com.client
import logging
from datetime import datetime

# Testlab Automation 가이드라인 준수
# 1차 목표: .lms 파일 선택 및 Testlab 연결 확인

def setup_logging():
    """바탕화면에 로그 파일을 설정합니다."""
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    log_file = os.path.join(desktop_path, f"VIN_Automation_Log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler()
        ]
    )
    logging.info(f"로그 파일이 생성되었습니다: {log_file}")
    return log_file

def select_file():
    """GUI를 통해 .lms 프로젝트 파일을 선택합니다."""
    root = tk.Tk()
    root.withdraw()  # 메인 윈도우 숨김
    
    file_path = filedialog.askopenfilename(
        title="Testlab 프로젝트 파일(.lms)을 선택하세요",
        filetypes=[("Testlab Project Files", "*.lms"), ("All Files", "*.*")]
    )
    
    root.destroy()
    if file_path:
        logging.info(f"선택된 파일: {file_path}")
    return file_path

def connect_testlab():
    """Simcenter Testlab의 Window Automation Support에 연결합니다. 여러 알려진 ProgID를 시도합니다."""
    # 알려진 Testlab Automation ProgID 리스트
    prog_ids = [
        "LMSTestLabAutomation.Application",
        "LMS.TestLab.Application",
        "SimcenterTestLab.Application",
        "TestLab.Application"
    ]
    
    app = None
    for prog_id in prog_ids:
        try:
            logging.info(f"'{prog_id}'으로 연결을 시도 중입니다...")
            # 이미 실행 중인 객체 확인
            app = win32com.client.GetActiveObject(prog_id)
            logging.info(f"'{prog_id}'으로 성공적으로 연결되었습니다 (실행 중인 객체 사용).")
            return app
        except Exception:
            try:
                # 새 객체 생성 시도
                app = win32com.client.Dispatch(prog_id)
                logging.info(f"'{prog_id}'으로 성공적으로 연결되었습니다 (새 객체 생성).")
                return app
            except Exception:
                continue

    error_msg = "모든 알려진 ProgID로 Testlab 연결에 실패했습니다.\n\n" \
                "1. Simcenter Testlab이 실행 중이며 'Window Automation Support'가 켜져 있는지 확인하세요.\n" \
                "2. 관리자 권한으로 실행 중인지 확인이 필요할 수 있습니다."
    logging.error(error_msg)
    messagebox.showerror("연결 오류", error_msg)
    return None

def main():
    # 로그 설정
    log_path = setup_logging()
    
    # 1. 파일 선택
    lms_file_path = select_file()
    
    if not lms_file_path:
        logging.warning("파일이 선택되지 않았습니다. 프로그램을 종료합니다.")
        return

    # 2. Testlab 연결
    app = connect_testlab()
    
    if app:
        try:
            logging.info("데이터 추출을 위한 준비가 완료되었습니다.")
            
            # TODO: Autopower (A) 데이터 경로 접근 및 추출 로직 추가
            
        except Exception as e:
            logging.error(f"작업 중 오류 발생: {e}")
    else:
        logging.error("연결 실패로 인해 작업을 중단합니다.")

if __name__ == "__main__":
    main()
