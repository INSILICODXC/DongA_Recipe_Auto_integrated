import os, sys, io, datetime, atexit, glob
import pandas as pd
import pyautogui
import importlib
import time  # ← 소요시간 디버그 로직을 위해 추가
import utility.config as config
from operation.generator import extract_from_excel, precheck_all_sheets
from operation import  recipe_rebuild   # recipe_rebuild 고정 실행용 추가

# ----------------------------------------------------------------------
# --- Recipe 복제 테스트 시 Timestamp 기록과 소요시간을 나타내기 위한 디버그 로직 -----------

ENABLE_FILE_LOG = True  # ← True면 파일 기록, False면 콘솔만 기록

LOG_DIR = r"D:\DongA_Recipe_Auto_integrated\Debug_Logs"
if ENABLE_FILE_LOG:
    os.makedirs(LOG_DIR, exist_ok=True)

    # 실행 시각 (초 단위까지)
    timestamp = datetime.datetime.now().strftime("%Y_%m_%d_%H-%M-%S")

    # 최종 로그 경로
    LOG_PATH = os.path.join(LOG_DIR, f"run_{timestamp}.log")

    # 파일 핸들 열기
    _log_fp = open(LOG_PATH, "w", encoding="utf-8", buffering=1)
else:
    LOG_PATH = None
    _log_fp = None

# 원래 스트림 백업
_orig_stdout, _orig_stderr = sys.stdout, sys.stderr

class _Tee(io.TextIOBase):
    def __init__(self, stream, logfp=None):
        self.stream = stream
        self.logfp = logfp
        self._at_line_start = True  # 줄 시작 여부 추적

    def _stamp(self):
        # HH:MM:SS 형태의 타임스탬프
        return datetime.datetime.now().strftime("[%H:%M:%S] ")

    def write(self, s):
        if not s:
            return 0
        if isinstance(s, bytes):
            s = s.decode("utf-8", errors="replace")

        out = []
        i = 0
        while i < len(s):
            if self._at_line_start:
                out.append(self._stamp())
                self._at_line_start = False
            ch = s[i]
            out.append(ch)
            if ch == "\n":
                self._at_line_start = True
            i += 1

        payload = "".join(out)
        self.stream.write(payload)
        if self.logfp:  # 파일 기록은 옵션
            self.logfp.write(payload)
            self.logfp.flush()
        self.stream.flush()
        return len(s)

    def flush(self):
        self.stream.flush()
        if self.logfp:
            self.logfp.flush()

# 콘솔+파일 동시 기록 (파일 로그는 옵션)
sys.stdout = _Tee(_orig_stdout, _log_fp if ENABLE_FILE_LOG else None)
sys.stderr = _Tee(_orig_stderr, _log_fp if ENABLE_FILE_LOG else None)

@atexit.register
def _close_tee():
    try:
        sys.stdout = _orig_stdout
        sys.stderr = _orig_stderr
    finally:
        if _log_fp:
            try:
                _log_fp.flush()
                _log_fp.close()
            except Exception:
                pass

if ENABLE_FILE_LOG:
    print(f"[TEE] Logging to: {LOG_PATH}")
else:
    print("[TEE] File logging disabled (console only)")
# ----------------------------------------------------------------------

# 계정 및 엑셀 경로 (ID/PW는 여기에 고정)
ID = "scitegicadmin"
PW = "Qwer1234!"
excel_path = r"D:\DongA_Recipe_Auto_integrated\Excel_Read_TEST_20250917.xlsx"

def run_selected_module():
    xls = pd.ExcelFile(excel_path)

    for repeat, sheet_name in enumerate(xls.sheet_names, start=1):
    #     df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
    #     choice = str(df.iloc[0, 0]).strip()  # A1 셀 값
    #     recipe_name = df.iloc[6, 1]  # B7 셀 값

        # --- [프리체크] Key-Value에서 '시험항목' 오른쪽 값 (choice) 만 빠르게 읽어 X 여부 판단 ---
        # ※ 간소화: 'peek_choice'로 엑셀을 한 번 더 읽어서 미리 확인하던 로직은 제거.
        #    이제는 extract_from_excel() 한 번만 호출해서 반환된 data['choice']로만 판단합니다.

        # drop-down 시트는 스킵 (ssgwak : Excel 시트 변경으로 인해 drop-down 목록 시트 추가, 해당 시트는 검사하지 않음)
        if sheet_name.lower() == config.IGNORE_SHEET_NAME.lower():
            print(f"[스킵] '{sheet_name}' 시트는 참조용이므로 실행하지 않음")
            continue

        # --- [프리체크] --------------------------------------------------------------------------------
        try:
            # extract_from_excel 함수 이용하여 엑셀 데이터 추출
            data = extract_from_excel(excel_path, sheet_name)
        except Exception as e:
            print(f"[스킵] {sheet_name}: 추출 실패 → {e}")
            continue

        # extract_from_excel 함수에서 Return 받은 딕셔너리 중 "choice" key의 value 값을 choice로
        choice = str(data["choice"]).strip()
        # extract_from_excel 함수에서 Return 받은 딕셔너리 중 "Recipe Name" key의 value 값을 recipe_name으로
        recipe_name = str(data["Recipe Name"]).strip()

        print("진행률 :", repeat, "/", len(xls.sheet_names))
        print(f"[{sheet_name}] 시험 항목: {choice} → Recipe: {recipe_name}")
        if choice.strip().upper().startswith("X"):
            print(f"[건너뜀] '{sheet_name}' 시트는 'X'로 시작하므로 실행하지 않음 (시험 항목: {choice})")
            continue

        try:
            _s = time.perf_counter()
            _w = datetime.datetime.now()
            print(f"[시작] {sheet_name} at {_w.strftime('%H:%M:%S')}")

            # recipe_copy.py 의 run_recipe_copy 로 단일화
            recipe_rebuild.run_recipe_rebuild(sheet_name, excel_path, ID, PW)

            _e = time.perf_counter()
            _w2 = datetime.datetime.now()
            print(f"[완료] {sheet_name} → recipe_copy 처리 완료 (elapsed: {_e - _s:.2f}s)")
        except Exception as e:
            print(f"[오류] {sheet_name} 시트 실행 중 오류 발생: {e}")

if __name__ == "__main__":
    pyautogui.FAILSAFE = False

    # ---- 전체 실행 타이밍 기록 시작 ----
    _start_wall = datetime.datetime.now()
    _start_perf = time.perf_counter()
    print(f"[RUN] Started at {_start_wall.strftime('%Y-%m-%d %H:%M:%S')}")

    try:
        # 1) Pre-check 먼저 실행
        precheck_all_sheets(excel_path)
        print("[PRECHECK] All sheets pre-check completed.")

        # 2) 이상 없으면 그때부터 Recipe 복제 실행
        run_selected_module()
        
    except KeyboardInterrupt:
        print("[RUN] Interrupted by user (KeyboardInterrupt).")
        raise
    finally:
        # ---- 전체 실행 타이밍 기록 종료/요약 ----
        _end_wall = datetime.datetime.now()
        _elapsed_td = _end_wall - _start_wall           # timedelta
        _elapsed_sec = time.perf_counter() - _start_perf # seconds (float)
        print(f"[RUN] Ended   at {_end_wall.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"[RUN] Elapsed = {_elapsed_td} ({_elapsed_sec:.2f}s)")