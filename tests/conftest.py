import os
import sys

# 프로젝트 루트를 import 경로에 추가 (python -m pytest tests 로 실행)
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
