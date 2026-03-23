import os
import sys
import shutil
import subprocess

APP_NAME = "Markdown_Converter"
MAIN_SCRIPT = "app.py"
ICON_FILE = "icon.ico"

def clean_old_builds():
    """기존 빌드 잔재들을 삭제하여 깨끗한 상태에서 빌드 시작"""
    print(">>> Cleaning old build and dist folders...")
    paths_to_remove = ["build", "dist", f"{APP_NAME}.spec", f"{APP_NAME}.zip"]
    for p in paths_to_remove:
        if os.path.exists(p):
            if os.path.isdir(p):
                shutil.rmtree(p)
            else:
                os.remove(p)

def get_conda_components():
    """sys.prefix를 이용해 현재 가상환경의 DLL 및 PyInstaller hooks를 활용한 패키지 데이터 자동 수집"""
    from PyInstaller.utils.hooks import collect_data_files
    env_base = sys.prefix
    
    binaries = []
    datas = []

    # 1. OpenSSL DLL 동적 검색 (libcrypto, libssl)
    bin_dir = os.path.join(env_base, "Library", "bin")
    if os.path.exists(bin_dir):
        for f in os.listdir(bin_dir):
            if (f.startswith("libcrypto") or f.startswith("libssl")) and f.endswith(".dll"):
                binaries.append(f"{os.path.join(bin_dir, f)};.")

    # 2. PyInstaller 내장 훅 대신 CLI 전달 플래그 활용 (명령어 길이 한계 방지)
    # 3. 아이콘 파일 추가
    if os.path.exists(ICON_FILE):
        datas.append(f"{ICON_FILE};.")

    return binaries, datas

def build_exe():
    clean_old_builds()
    
    binaries, datas = get_conda_components()
    
    # 명령어 파라미터를 리스트 형태로 작성 (띄어쓰기/따옴표 문제 원천 차단)
    pyinstaller_args = [
        sys.executable, "-m", "PyInstaller",
        "--name", APP_NAME,
        "--onedir",
        "--noconfirm",
        "--noconsole",
    ]
    
    if os.path.exists(ICON_FILE):
        pyinstaller_args.extend(["--icon", ICON_FILE])
        
    for b in binaries:
        pyinstaller_args.extend(["--add-binary", f"{b}"])
        
    for d in datas:
        pyinstaller_args.extend(["--add-data", f"{d}"])
        
    # 명시적 의존성 및 데이터 수집 추가 (WinError 206 방지용 네이티브 플래그)
    hidden_imports = [
        "tkinter", 
        "tkinterdnd2", 
        "opendataloader_pdf", 
        "PIL",
        "subprocess",
        "markitdown",
        "magika",
        "onnxruntime",
        "mammoth",
        "pdfminer"
    ]
    for hi in hidden_imports:
        pyinstaller_args.extend(["--hidden-import", hi])
        
    collect_all_pkgs = [
        'opendataloader_pdf',
        'tkinterdnd2',
        'markitdown',
        'magika',
        'onnxruntime',
        'mammoth',
        'pdfminer',
        'speech_recognition',
    ]
    for pkg in collect_all_pkgs:
        pyinstaller_args.extend(["--collect-all", pkg])

    pyinstaller_args.append(MAIN_SCRIPT)
    
    print(">>> Running PyInstaller with args sequence...")
    # subprocess list 실행 (shell=False) 이스케이프 완벽 대비
    subprocess.run(pyinstaller_args, check=True)

def zip_release():
    print(">>> Zipping the dist directory without nested top-level folder...")
    # 배포용 폴더 경로 (예: dist/PDF_to_Markdown)
    target_dir = os.path.join("dist", APP_NAME)
    zip_name = APP_NAME 
    
    if os.path.exists(target_dir):
        # root_dir을 target_dir로 설정해 안에 들어있는 '내용물'들만 바로 압축
        shutil.make_archive(zip_name, 'zip', root_dir=target_dir)
        print(f">>> Successfully created {zip_name}.zip")
    else:
        print(">>> ERROR: Target build folder not found. Skipping ZIP.")

if __name__ == "__main__":
    build_exe()
    zip_release()
