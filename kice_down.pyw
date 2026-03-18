import sys
import os

# 현재 폴더를 시스템 경로에 추가
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from main import KICEDownApp

if __name__ == "__main__":
    app = KICEDownApp()
    app.run()
