"""
간단한 DWG 파일 테스트
DWG 파일을 처리하는 여러 방법 시도
"""

import os
import sys
from pathlib import Path

# src 디렉토리를 Python 경로에 추가
sys.path.append(str(Path(__file__).parent / "src"))

def test_with_ezdxf():
    """ezdxf로 DWG 직접 읽기 시도"""
    print("\n1. ezdxf로 DWG 직접 읽기 시도...")
    try:
        import ezdxf
        from ezdxf import recover
        from ezdxf.addons import odafc
        
        # DWG 파일 경로
        dwg_file = "test.dwg"
        
        # 방법 1: recover 모드로 시도
        try:
            doc, auditor = recover.readfile(dwg_file)
            print(f"  ✓ recover 모드로 읽기 성공!")
            print(f"    - DXF 버전: {doc.dxfversion}")
            print(f"    - 레이어 수: {len(doc.layers)}")
            return doc
        except:
            pass
        
        # 방법 2: 일반 읽기 시도
        try:
            doc = ezdxf.readfile(dwg_file)
            print(f"  ✓ 일반 모드로 읽기 성공!")
            return doc
        except:
            pass
            
        print("  ✗ ezdxf로 DWG 직접 읽기 실패 (DXF 변환 필요)")
        
    except ImportError:
        print("  ✗ ezdxf가 설치되지 않음")
    except Exception as e:
        print(f"  ✗ 오류: {e}")
    
    return None

def test_with_matplotlib():
    """matplotlib-dwg 시도"""
    print("\n2. 다른 라이브러리 확인...")
    try:
        # DWG는 독점 형식이라 직접 읽기 어려움
        print("  - DWG는 Autodesk 독점 형식")
        print("  - 오픈소스로는 제한적 지원")
        print("  - DXF 변환 권장")
    except Exception as e:
        print(f"  ✗ 오류: {e}")

def create_sample_dxf():
    """테스트용 샘플 DXF 파일 생성"""
    print("\n3. 테스트용 샘플 DXF 생성...")
    try:
        import ezdxf
        from ezdxf import units
        
        # 새 DXF 문서 생성
        doc = ezdxf.new('R2018')
        doc.units = units.M  # 미터 단위
        
        msp = doc.modelspace()
        
        # 건축 요소 추가 (벽체)
        doc.layers.add("A-WALL", color=5)
        # 사각형 벽체
        wall_points = [(0, 0), (10, 0), (10, 8), (0, 8), (0, 0)]
        msp.add_lwpolyline(wall_points, dxfattribs={'layer': 'A-WALL', 'closed': True})
        
        # 기둥 추가
        doc.layers.add("A-COLS", color=3)
        msp.add_circle((2, 2), radius=0.3, dxfattribs={'layer': 'A-COLS'})
        msp.add_circle((8, 2), radius=0.3, dxfattribs={'layer': 'A-COLS'})
        msp.add_circle((2, 6), radius=0.3, dxfattribs={'layer': 'A-COLS'})
        msp.add_circle((8, 6), radius=0.3, dxfattribs={'layer': 'A-COLS'})
        
        # 문 추가 (블록)
        doc.layers.add("A-DOOR", color=6)
        door = doc.blocks.new(name='DOOR')
        door.add_line((0, 0), (1, 0))
        door.add_arc((0, 0), radius=1, start_angle=0, end_angle=90)
        msp.add_blockref('DOOR', (5, 0), dxfattribs={'layer': 'A-DOOR'})
        
        # 토목 요소 추가 (도로)
        doc.layers.add("C-ROAD", color=8)
        road_points = [(12, 0), (25, 0), (25, 3), (12, 3), (12, 0)]
        msp.add_lwpolyline(road_points, dxfattribs={'layer': 'C-ROAD', 'closed': True})
        
        # 배관 추가
        doc.layers.add("C-PIPE", color=4)
        msp.add_line((12, 1.5), (25, 1.5), dxfattribs={'layer': 'C-PIPE'})
        
        # 조경 요소 추가 (수목)
        doc.layers.add("L-PLNT", color=2)
        tree_positions = [(15, 5), (18, 5), (21, 5), (15, 7), (18, 7), (21, 7)]
        for pos in tree_positions:
            msp.add_circle(pos, radius=0.5, dxfattribs={'layer': 'L-PLNT'})
        
        # 잔디 영역
        doc.layers.add("L-PLNT-TURF", color=82)
        grass_points = [(12, 4), (25, 4), (25, 8), (12, 8), (12, 4)]
        grass = msp.add_lwpolyline(grass_points, dxfattribs={'layer': 'L-PLNT-TURF', 'closed': True})
        
        # 해치 추가 (콘크리트 패턴)
        doc.layers.add("A-SLAB", color=251)
        hatch = msp.add_hatch(color=251, dxfattribs={'layer': 'A-SLAB'})
        hatch.paths.add_polyline_path([(0.5, 0.5), (9.5, 0.5), (9.5, 7.5), (0.5, 7.5), (0.5, 0.5)])
        hatch.set_pattern_fill('ANSI31', scale=0.1)
        
        # 텍스트 추가
        msp.add_text('건축 구역', height=0.5, dxfattribs={'layer': '0'}).set_placement((5, 9))
        msp.add_text('토목 구역', height=0.5, dxfattribs={'layer': '0'}).set_placement((18, 9))
        
        # 치수 추가
        dim = msp.add_linear_dim(base=(5, -1), p1=(0, 0), p2=(10, 0))
        dim.render()
        
        # 파일 저장
        doc.saveas('sample_test.dxf')
        print("  ✓ sample_test.dxf 파일 생성 완료")
        print("    - 건축: 벽체, 기둥, 문, 슬래브")
        print("    - 토목: 도로, 배관")
        print("    - 조경: 수목, 잔디")
        
        return True
        
    except Exception as e:
        print(f"  ✗ 샘플 생성 실패: {e}")
        return False

def test_sample_with_main():
    """생성된 샘플 DXF로 메인 프로그램 테스트"""
    print("\n4. 샘플 DXF로 물량 산출 테스트...")
    
    if not os.path.exists('sample_test.dxf'):
        print("  샘플 파일이 없습니다. 먼저 생성합니다.")
        if not create_sample_dxf():
            return
    
    try:
        from dwg_parser import DWGParser
        from quantity_calculator import QuantityCalculator
        from report_generator import ReportGenerator
        
        # DXF 파일 파싱
        parser = DWGParser('sample_test.dxf')
        
        if parser.load_file():
            print("  ✓ 파일 로드 성공")
            
            # 엔티티 파싱
            parser.parse_entities()
            summary = parser.get_summary()
            
            print(f"  ✓ 파싱 완료:")
            print(f"    - 레이어: {summary['레이어수']}개")
            print(f"    - 총 엔티티: {summary['총엔티티']}개")
            
            # 분야별 분류
            classified = parser.classify_by_field()
            
            # 물량 계산
            calculator = QuantityCalculator(classified)
            quantities = calculator.calculate_all()
            
            # 결과 출력
            print("\n  ✓ 물량 계산 결과:")
            for field in ["건축", "토목", "조경"]:
                if field in quantities and quantities[field]:
                    print(f"\n    [{field}]")
                    for category, values in quantities[field].items():
                        if isinstance(values, dict):
                            for key, value in values.items():
                                if key not in ["단위", "타입별"]:
                                    print(f"      {category}-{key}: {value}")
            
            # 리포트 생성
            os.makedirs("output", exist_ok=True)
            generator = ReportGenerator(quantities, summary)
            excel_path = generator.generate_excel()
            json_path = generator.generate_json()
            
            print(f"\n  ✓ 리포트 생성 완료:")
            print(f"    - Excel: {excel_path}")
            print(f"    - JSON: {json_path}")
            
        else:
            print("  ✗ 파일 로드 실패")
            
    except Exception as e:
        print(f"  ✗ 테스트 실패: {e}")
        import traceback
        traceback.print_exc()

def main():
    print("="*60)
    print("DWG 파일 테스트 프로그램")
    print("="*60)
    
    # 1. DWG 직접 읽기 시도
    doc = test_with_ezdxf()
    
    # 2. 다른 방법 안내
    test_with_matplotlib()
    
    # 3. 샘플 DXF 생성
    create_sample_dxf()
    
    # 4. 샘플로 테스트
    test_sample_with_main()
    
    print("\n" + "="*60)
    print("테스트 완료!")
    print("\n권장사항:")
    print("1. DWG 파일은 DXF로 변환 후 처리")
    print("2. 변환 도구: ODA File Converter (무료)")
    print("3. 또는 Aspose.CAD 등 상용 라이브러리 사용")
    print("="*60)

if __name__ == "__main__":
    main()