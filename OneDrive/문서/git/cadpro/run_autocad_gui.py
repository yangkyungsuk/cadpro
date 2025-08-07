"""
AutoCAD 물량 산출 도구 - GUI 자동 실행
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from autocad_plugin import AutoCADQuantityTool, create_gui

if __name__ == "__main__":
    print("="*60)
    print("AutoCAD 물량 산출 도구 - GUI 모드")
    print("="*60)
    print("\nGUI 창을 실행합니다...")
    print("AutoCAD에서 객체를 선택하고 계산을 수행하세요.\n")
    
    # GUI 실행
    create_gui()