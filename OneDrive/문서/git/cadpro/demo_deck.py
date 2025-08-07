"""
데크 공사 도면 물량 산출 데모
AutoCAD에서 실행
"""

import win32com.client
import pythoncom
import math

def demo_deck_calculation():
    """데크 공사 물량 계산 데모"""
    print("="*60)
    print("데크 공사 물량 산출 데모")
    print("="*60)
    
    try:
        # AutoCAD 연결
        acad = win32com.client.Dispatch("AutoCAD.Application")
        doc = acad.ActiveDocument
        model = doc.ModelSpace
        
        print(f"\n현재 도면: {doc.Name}")
        print(f"객체 수: {model.Count}개")
        
        # 레이어별 분석
        print("\n[레이어 분석]")
        layer_objects = {}
        
        for obj in model:
            layer = obj.Layer
            if layer not in layer_objects:
                layer_objects[layer] = {
                    'count': 0,
                    'types': [],
                    'lines': [],
                    'polylines': [],
                    'circles': []
                }
            
            layer_objects[layer]['count'] += 1
            
            # 타입별 분류
            if "Line" in obj.ObjectName:
                layer_objects[layer]['lines'].append(obj)
            elif "Polyline" in obj.ObjectName:
                layer_objects[layer]['polylines'].append(obj)
            elif "Circle" in obj.ObjectName:
                layer_objects[layer]['circles'].append(obj)
        
        # 상위 10개 레이어 표시
        print("\n주요 레이어 (객체 수 기준):")
        sorted_layers = sorted(layer_objects.items(), key=lambda x: x[1]['count'], reverse=True)
        
        for i, (layer, data) in enumerate(sorted_layers[:10], 1):
            print(f"  {i}. {layer}: {data['count']}개")
        
        # 데크 관련 레이어 찾기
        print("\n[데크 관련 요소 검색]")
        
        deck_keywords = ['DECK', '데크', 'BOARD', '장선', 'JOIST', '난간', 'RAIL', '기둥', 'POST']
        deck_layers = []
        
        for layer in layer_objects.keys():
            layer_upper = layer.upper()
            for keyword in deck_keywords:
                if keyword in layer_upper or keyword in layer:
                    if layer not in deck_layers:
                        deck_layers.append(layer)
                        print(f"  - 발견: {layer}")
        
        # 물량 계산 예시
        print("\n[물량 계산 예시]")
        
        # 1. 전체 선 길이 (데크보드나 장선으로 가정)
        total_line_length = 0
        line_count = 0
        
        for layer_data in layer_objects.values():
            for line in layer_data['lines']:
                try:
                    start = line.StartPoint
                    end = line.EndPoint
                    length = math.sqrt((end[0]-start[0])**2 + (end[1]-start[1])**2)
                    total_line_length += length
                    line_count += 1
                except:
                    pass
        
        print(f"\n1. 선형 요소:")
        print(f"   - 총 개수: {line_count}개")
        print(f"   - 총 길이: {total_line_length:.2f} 도면단위")
        
        # mm를 m로 변환 (도면이 mm 단위라고 가정)
        if total_line_length > 0:
            length_m = total_line_length / 1000
            print(f"   - 실제 길이: {length_m:.2f} m")
        
        # 2. 폐곡선 면적 (데크 바닥 면적)
        total_area = 0
        area_count = 0
        
        for layer_data in layer_objects.values():
            for poly in layer_data['polylines']:
                try:
                    if poly.Closed:
                        area = poly.Area
                        total_area += area
                        area_count += 1
                except:
                    pass
        
        print(f"\n2. 면적 요소:")
        print(f"   - 폐곡선 개수: {area_count}개")
        print(f"   - 총 면적: {total_area:.2f} 도면단위²")
        
        if total_area > 0:
            area_m2 = total_area / 1000000  # mm² → m²
            print(f"   - 실제 면적: {area_m2:.2f} m²")
            
            # 데크보드 물량 계산
            print(f"\n3. 데크보드 물량 (7% 손실 포함):")
            deck_board_area = area_m2 * 1.07
            print(f"   - 필요 면적: {deck_board_area:.2f} m²")
            
            # 데크 스크류 계산 (m²당 8개)
            screws = int(area_m2 * 8)
            print(f"   - 데크 스크류: {screws}개")
        
        # 3. 원형 객체 (기둥)
        total_circles = 0
        for layer_data in layer_objects.values():
            total_circles += len(layer_data['circles'])
        
        if total_circles > 0:
            print(f"\n4. 기둥 (원형 객체):")
            print(f"   - 개수: {total_circles}개")
            print(f"   - 기초 콘크리트: {total_circles * 0.096:.3f} m³ (400x400x600 기준)")
        
        # 간단한 대화형 옵션
        print("\n" + "="*60)
        print("추가 계산 옵션:")
        print("  실제 AutoCAD 플러그인(autocad_plugin.py)을 사용하면")
        print("  - 특정 영역 선택")
        print("  - 레이어별 필터링")
        print("  - 사용자 정의 계산")
        print("  - Excel/AutoCAD 테이블 출력")
        print("  등이 가능합니다.")
        
        return True
        
    except Exception as e:
        print(f"\n[ERROR] 오류 발생: {e}")
        return False

if __name__ == "__main__":
    demo_deck_calculation()