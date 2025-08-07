# -*- coding: utf-8 -*-
"""
간단한 DWG/DXF 테스트 프로그램
"""

import os
import sys
from pathlib import Path

# src 디렉토리를 Python 경로에 추가
sys.path.append(str(Path(__file__).parent / "src"))

def create_sample_dxf():
    """테스트용 샘플 DXF 파일 생성"""
    print("\n샘플 DXF 파일 생성 중...")
    try:
        import ezdxf
        from ezdxf import units
        
        # 새 DXF 문서 생성
        doc = ezdxf.new('R2018')
        doc.units = units.M  # 미터 단위
        
        msp = doc.modelspace()
        
        # === 건축 요소 ===
        # 벽체 레이어
        doc.layers.add("A-WALL", color=5)
        wall_points = [(0, 0), (10, 0), (10, 8), (0, 8), (0, 0)]
        msp.add_lwpolyline(wall_points, dxfattribs={'layer': 'A-WALL', 'closed': True})
        
        # 내부 벽
        msp.add_line((5, 0), (5, 8), dxfattribs={'layer': 'A-WALL'})
        msp.add_line((0, 4), (10, 4), dxfattribs={'layer': 'A-WALL'})
        
        # 기둥
        doc.layers.add("A-COLS", color=3)
        column_positions = [(2, 2), (8, 2), (2, 6), (8, 6), (5, 4)]
        for pos in column_positions:
            msp.add_circle(pos, radius=0.25, dxfattribs={'layer': 'A-COLS'})
        
        # 슬래브 (해치)
        doc.layers.add("A-SLAB", color=251)
        hatch = msp.add_hatch(color=251, dxfattribs={'layer': 'A-SLAB'})
        hatch.paths.add_polyline_path([(0.2, 0.2), (4.8, 0.2), (4.8, 3.8), (0.2, 3.8)])
        hatch.set_pattern_fill('CONCRETE', scale=0.05)
        
        # === 토목 요소 ===
        # 도로
        doc.layers.add("C-ROAD", color=8)
        road_points = [(12, 0), (30, 0), (30, 4), (12, 4), (12, 0)]
        msp.add_lwpolyline(road_points, dxfattribs={'layer': 'C-ROAD', 'closed': True})
        
        # 배관 (여러 직경)
        doc.layers.add("C-PIPE", color=4)
        # 주배관
        msp.add_line((12, 2), (30, 2), dxfattribs={'layer': 'C-PIPE'})
        # 배관 단면 표시 (원)
        for x in [15, 20, 25]:
            msp.add_circle((x, 2), radius=0.15, dxfattribs={'layer': 'C-PIPE'})
        
        # 철근
        doc.layers.add("C-RBAR", color=1)
        # 수평 철근
        for y in [5.5, 6.0, 6.5]:
            msp.add_line((12, y), (30, y), dxfattribs={'layer': 'C-RBAR'})
        # 수직 철근
        for x in range(13, 30, 2):
            msp.add_line((x, 5), (x, 7), dxfattribs={'layer': 'C-RBAR'})
        
        # === 조경 요소 ===
        # 수목 (교목)
        doc.layers.add("L-PLNT", color=2)
        tree_positions = [(15, 10), (20, 10), (25, 10), (15, 14), (20, 14), (25, 14)]
        for pos in tree_positions:
            msp.add_circle(pos, radius=0.8, dxfattribs={'layer': 'L-PLNT'})
            # 수목 중심 표시
            msp.add_point(pos, dxfattribs={'layer': 'L-PLNT'})
        
        # 관목
        shrub_positions = [(13, 12), (14, 12), (26, 12), (27, 12)]
        for pos in shrub_positions:
            msp.add_circle(pos, radius=0.3, dxfattribs={'layer': 'L-PLNT'})
        
        # 잔디
        doc.layers.add("L-PLNT-TURF", color=82)
        grass_points = [(12, 8), (30, 8), (30, 16), (12, 16), (12, 8)]
        grass = msp.add_lwpolyline(grass_points, dxfattribs={'layer': 'L-PLNT-TURF', 'closed': True})
        
        # 포장 (보도)
        doc.layers.add("L-PAVE", color=253)
        paving_points = [(0, 10), (10, 10), (10, 12), (0, 12), (0, 10)]
        paving = msp.add_lwpolyline(paving_points, dxfattribs={'layer': 'L-PAVE', 'closed': True})
        
        # 텍스트 라벨
        msp.add_text('Building Area', height=0.5, dxfattribs={'layer': '0'}).set_placement((5, 9))
        msp.add_text('Civil Area', height=0.5, dxfattribs={'layer': '0'}).set_placement((21, 7.5))
        msp.add_text('Landscape Area', height=0.5, dxfattribs={'layer': '0'}).set_placement((21, 17))
        
        # 파일 저장
        doc.saveas('sample_test.dxf')
        print("  [OK] sample_test.dxf 파일 생성 완료")
        return True
        
    except Exception as e:
        print(f"  [ERROR] 샘플 생성 실패: {e}")
        return False

def test_dxf_file():
    """DXF 파일로 물량 산출 테스트"""
    print("\nDXF 파일 물량 산출 테스트...")
    
    try:
        from dwg_parser import DWGParser
        from quantity_calculator import QuantityCalculator
        from report_generator import ReportGenerator
        
        # DXF 파일 파싱
        parser = DWGParser('sample_test.dxf')
        
        if not parser.load_file():
            print("  [ERROR] 파일 로드 실패")
            return False
        
        print("  [OK] 파일 로드 성공")
        
        # 엔티티 파싱
        parser.parse_entities()
        summary = parser.get_summary()
        
        print(f"\n파일 정보:")
        print(f"  - CAD 버전: {summary['CAD버전']}")
        print(f"  - 레이어 수: {summary['레이어수']}")
        print(f"  - 엔티티 수:")
        for entity_type, count in summary['엔티티수'].items():
            if count > 0:
                print(f"    * {entity_type}: {count}")
        
        # 분야별 분류
        classified = parser.classify_by_field()
        
        # 물량 계산
        calculator = QuantityCalculator(classified)
        quantities = calculator.calculate_all()
        
        # 결과 출력
        print("\n물량 산출 결과:")
        print("-" * 40)
        
        for field in ["건축", "토목", "조경"]:
            if field in quantities and quantities[field]:
                print(f"\n[{field} 분야]")
                for category, values in quantities[field].items():
                    print(f"  {category}:")
                    if isinstance(values, dict):
                        for key, value in values.items():
                            if key == "단위":
                                continue
                            elif isinstance(value, dict):
                                print(f"    - {key}:")
                                for sub_key, sub_value in value.items():
                                    print(f"        {sub_key}: {sub_value}")
                            else:
                                unit = values.get("단위", {}).get(key, "")
                                print(f"    - {key}: {value} {unit}")
        
        # 리포트 생성
        print("\n리포트 생성 중...")
        os.makedirs("output", exist_ok=True)
        
        generator = ReportGenerator(quantities, summary)
        excel_path = generator.generate_excel()
        json_path = generator.generate_json()
        
        print(f"\n리포트 생성 완료:")
        print(f"  - Excel: {excel_path}")
        print(f"  - JSON: {json_path}")
        
        return True
        
    except Exception as e:
        print(f"  [ERROR] 테스트 실패: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_dwg_file():
    """DWG 파일 처리 안내"""
    print("\nDWG 파일 처리 방법:")
    print("-" * 40)
    print("test.dwg 파일을 처리하려면 다음 방법 중 하나를 선택하세요:")
    print("\n1. ODA File Converter 사용 (무료)")
    print("   - https://www.opendesign.com/guestfiles/oda_file_converter")
    print("   - test.dwg -> test.dxf 변환")
    print("   - 변환 후 main.py 실행")
    print("\n2. 온라인 변환 서비스")
    print("   - https://cloudconvert.com/dwg-to-dxf")
    print("   - https://convertio.co/kr/dwg-dxf/")
    print("\n3. AutoCAD 또는 DWG 지원 CAD 프로그램")
    print("   - 파일 열기 -> 다른 이름으로 저장 -> DXF 형식")

def main():
    print("="*60)
    print("CAD 물량 산출 시스템 테스트")
    print("="*60)
    
    # 1. 샘플 DXF 생성
    if not os.path.exists('sample_test.dxf'):
        if not create_sample_dxf():
            return
    
    # 2. DXF 파일 테스트
    test_dxf_file()
    
    # 3. DWG 파일 안내
    test_dwg_file()
    
    print("\n" + "="*60)
    print("테스트 완료!")
    print("="*60)

if __name__ == "__main__":
    main()