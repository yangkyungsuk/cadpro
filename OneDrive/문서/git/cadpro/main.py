"""
CAD 물량 산출 프로그램 메인 모듈
건축, 토목, 조경 통합 물량 산출 시스템
"""

import os
import sys
from pathlib import Path

# src 디렉토리를 Python 경로에 추가
sys.path.append(str(Path(__file__).parent / "src"))

from dwg_parser import DWGParser
from quantity_calculator import QuantityCalculator
from report_generator import ReportGenerator


def main():
    """메인 실행 함수"""
    print("="*60)
    print("CAD 물량 산출 시스템 v1.0")
    print("건축, 토목, 조경 통합 물량 산출")
    print("="*60)
    
    # 테스트용 샘플 파일 경로 (실제 사용시 변경 필요)
    cad_file = input("\nCAD 파일 경로를 입력하세요 (DWG/DXF): ").strip()
    
    if not os.path.exists(cad_file):
        print(f"오류: 파일을 찾을 수 없습니다 - {cad_file}")
        return
    
    try:
        # 1. CAD 파일 파싱
        print("\n[1/4] CAD 파일 로드 중...")
        parser = DWGParser(cad_file)
        
        if not parser.load_file():
            print("CAD 파일 로드 실패")
            return
        
        print("[2/4] 엔티티 파싱 중...")
        parser.parse_entities()
        
        # 파싱 결과 요약
        summary = parser.get_summary()
        print(f"\n파싱 완료:")
        print(f"  - 레이어: {summary['레이어수']}개")
        print(f"  - 엔티티: {summary['총엔티티']}개")
        
        # 2. 분야별 분류
        print("\n[3/4] 분야별 분류 중...")
        classified = parser.classify_by_field()
        
        # 분류 결과 출력
        for field in ["건축", "토목", "조경"]:
            count = sum(len(entities) for cat in classified[field].values() 
                       for entities in cat.values())
            if count > 0:
                print(f"  - {field}: {count}개 엔티티")
        
        # 3. 물량 계산
        print("\n[4/4] 물량 계산 중...")
        calculator = QuantityCalculator(classified)
        quantities = calculator.calculate_all()
        
        # 4. 리포트 생성
        print("\n리포트 생성 중...")
        
        # output 디렉토리 생성
        os.makedirs("output", exist_ok=True)
        
        generator = ReportGenerator(quantities, summary)
        
        # 콘솔 출력
        generator.print_summary()
        
        # Excel 리포트 생성
        excel_path = generator.generate_excel()
        
        # JSON 리포트 생성
        json_path = generator.generate_json()
        
        print(f"\n완료! 리포트가 생성되었습니다:")
        print(f"  - Excel: {excel_path}")
        print(f"  - JSON: {json_path}")
        
    except Exception as e:
        print(f"\n오류 발생: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()