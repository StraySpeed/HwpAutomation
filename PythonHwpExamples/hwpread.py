import os, sys
sys.path.insert(1, os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
import HwpAutomation

# 읽어들일 파일 경로를 지정
path = r""

# PythonHwp 열기
hwp = HwpAutomation.openhwp(path)

# 출력 테스트 후 종료
hwp._readtest()
