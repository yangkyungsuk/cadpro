"""
창호도 객체 타입 분석
창호가 LINE인지 POLYLINE인지 확인
"""

import win32com.client
import pythoncom

def analyze_window_objects():
    """창호도 객체 타입 분석"""
    
    print("=" * 60)
    print("창호도 객체 타입 분석")
    print("=" * 60)
    
    try:
        pythoncom.CoInitialize()
        
        # AutoCAD 연결
        acad = win32com.client.dynamic.Dispatch("AutoCAD.Application")
        doc = acad.ActiveDocument
        print(f"도면: {doc.Name}\n")
        
        # 객체 타입별 카운트
        object_types = {}
        
        # 창호 하나 선택해보기
        print("창호 하나를 선택하세요 (사각형 전체를 드래그)...")
        
        selection = doc.SelectionSets.Add("TestWindow")
        selection.SelectOnScreen()
        
        print(f"\n선택된 객체: {selection.Count}개\n")
        
        # 각 객체 분석
        for i in range(selection.Count):
            obj = selection.Item(i)
            
            # 동적 디스패치로 재래핑
            try:
                obj = win32com.client.dynamic.Dispatch(obj)
            except:
                pass
            
            obj_type = str(obj.ObjectName)
            
            # 타입별 카운트
            if obj_type not in object_types:
                object_types[obj_type] = 0
            object_types[obj_type] += 1
            
            # 상세 정보 출력 (처음 10개만)
            if i < 10:
                print(f"[객체 {i+1}] {obj_type}")
                
                # LINE인 경우
                if "Line" in obj_type and "Polyline" not in obj_type:
                    try:
                        start = obj.StartPoint
                        end = obj.EndPoint
                        length = ((end[0]-start[0])**2 + (end[1]-start[1])**2) ** 0.5
                        print(f"  LINE 길이: {length:.1f}mm")
                        
                        # 수평/수직 판단
                        if abs(start[1] - end[1]) < 1:
                            print(f"  → 수평선")
                        elif abs(start[0] - end[0]) < 1:
                            print(f"  → 수직선")
                        else:
                            print(f"  → 대각선")
                    except:
                        pass
                
                # POLYLINE인 경우
                elif "Polyline" in obj_type:
                    try:
                        coords = obj.Coordinates
                        vertex_count = len(coords) // 2
                        print(f"  POLYLINE 점 개수: {vertex_count}개")
                        
                        # 폐합 여부
                        try:
                            closed = obj.Closed
                            print(f"  폐합: {'예' if closed else '아니오'}")
                        except:
                            pass
                        
                        # 사각형인지 확인
                        if vertex_count == 4 or vertex_count == 5:
                            print(f"  → 사각형 가능성")
                            
                            # 크기 계산
                            x_coords = [coords[j] for j in range(0, len(coords), 2)]
                            y_coords = [coords[j] for j in range(1, len(coords), 2)]
                            width = max(x_coords) - min(x_coords)
                            height = max(y_coords) - min(y_coords)
                            print(f"  크기: {width:.1f} x {height:.1f}mm")
                    except Exception as e:
                        print(f"  오류: {e}")
                
                print()
        
        # 요약
        print("=" * 60)
        print("객체 타입 요약:")
        print("-" * 60)
        for obj_type, count in sorted(object_types.items()):
            print(f"  {obj_type}: {count}개")
        
        print("\n" + "=" * 60)
        
        # 분석 결과
        if "AcDbLine" in object_types and object_types.get("AcDbLine", 0) >= 4:
            print("✅ LINE으로 구성된 창호입니다.")
            print("   → 4개 이상의 LINE이 하나의 창호를 구성")
            print("   → LINE 그룹화 알고리즘 필요")
            
        elif "AcDbPolyline" in object_types or "AcDb2dPolyline" in object_types:
            print("✅ POLYLINE으로 구성된 창호입니다.")
            print("   → 하나의 POLYLINE이 하나의 창호")
            print("   → 크기 비교가 쉬움")
            
        else:
            print("⚠️ 알 수 없는 구조입니다.")
            print("   다른 창호를 선택해서 다시 테스트해보세요.")
        
        # 정리
        selection.Delete()
        
    except Exception as e:
        print(f"\n오류: {e}")
        import traceback
        traceback.print_exc()
    finally:
        pythoncom.CoUninitialize()


def find_all_rectangles():
    """도면의 모든 사각형 찾기 (POLYLINE과 LINE 모두)"""
    
    print("\n" + "=" * 60)
    print("전체 도면에서 사각형 찾기")
    print("=" * 60)
    
    try:
        pythoncom.CoInitialize()
        
        acad = win32com.client.dynamic.Dispatch("AutoCAD.Application")
        doc = acad.ActiveDocument
        model_space = doc.ModelSpace
        
        polyline_rectangles = []
        line_count = 0
        
        print(f"전체 객체 수: {model_space.Count}개 분석 중...\n")
        
        # 빠른 스캔 (처음 1000개만)
        scan_count = min(1000, model_space.Count)
        
        for i in range(scan_count):
            try:
                obj = model_space.Item(i)
                obj = win32com.client.dynamic.Dispatch(obj)
                obj_type = str(obj.ObjectName)
                
                # POLYLINE 사각형
                if "Polyline" in obj_type:
                    try:
                        coords = obj.Coordinates
                        vertex_count = len(coords) // 2
                        
                        # 4-5개 점이면 사각형 가능성
                        if vertex_count in [4, 5]:
                            # 폐합 확인
                            closed = False
                            try:
                                closed = obj.Closed
                            except:
                                # 첫점과 끝점 비교
                                if vertex_count == 5:
                                    if coords[0] == coords[-2] and coords[1] == coords[-1]:
                                        closed = True
                            
                            if closed:
                                x_coords = [coords[j] for j in range(0, len(coords), 2)]
                                y_coords = [coords[j] for j in range(1, len(coords), 2)]
                                width = max(x_coords) - min(x_coords)
                                height = max(y_coords) - min(y_coords)
                                
                                if width > 100 and height > 100:  # 최소 크기
                                    polyline_rectangles.append({
                                        'width': width,
                                        'height': height,
                                        'obj': obj
                                    })
                    except:
                        pass
                
                # LINE 카운트
                elif "Line" in obj_type and "Polyline" not in obj_type:
                    line_count += 1
                    
            except:
                continue
        
        # 결과 출력
        print(f"발견된 POLYLINE 사각형: {len(polyline_rectangles)}개")
        print(f"발견된 LINE: {line_count}개\n")
        
        if polyline_rectangles:
            # 크기별 그룹화
            size_groups = {}
            for rect in polyline_rectangles:
                key = f"{int(rect['width'])}x{int(rect['height'])}"
                if key not in size_groups:
                    size_groups[key] = 0
                size_groups[key] += 1
            
            print("POLYLINE 사각형 크기별 분포:")
            for size, count in sorted(size_groups.items()):
                print(f"  {size}mm: {count}개")
        
        if line_count > 100:
            print(f"\n💡 LINE이 많습니다 ({line_count}개)")
            print("   창호가 LINE으로 구성되어 있을 가능성이 높습니다.")
            print("   → LINE 그룹화 방식을 사용해야 합니다.")
        
    except Exception as e:
        print(f"오류: {e}")
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    print("창호도 분석 도구")
    print("1. 선택한 창호 분석")
    print("2. 전체 도면 스캔")
    
    choice = input("\n선택 (1 or 2): ")
    
    if choice == "1":
        analyze_window_objects()
    elif choice == "2":
        find_all_rectangles()
    else:
        print("잘못된 선택입니다.")
    
    input("\n엔터를 눌러 종료...")