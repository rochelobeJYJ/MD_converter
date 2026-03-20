import os
import sys

# 스크립트가 위치한 폴더를 작업 디렉토리로 설정 (icon.ico 경로 추적 용이)
current_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(current_dir)
sys.path.append(current_dir)

from app import PDFtoMDApp

if __name__ == "__main__":
    app = PDFtoMDApp()
    app.run()
