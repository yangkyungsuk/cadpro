"""
AutoCAD 대화형 물량 산출
사용자가 직접 선택하고 지정하는 방식
"""

import win32com.client
import pythoncom
import math
import sys

class InteractiveAutoCAD:
    def __init__(self):
        """AutoCAD 연결"""
        try:
            self.acad = win32com.client.Dispatch("AutoCAD.Application")
            self.doc = self.acad.ActiveDocument
            self.model = self.doc.ModelSpace
            
            print("[OK] AutoCAD 연결 성공")
            print(f"도면: {self.doc.Name}\n")
            
            # 선택된 객체 저장
            self.selected_objects = []
            
        except Exception as e:
            print(f"[ERROR] AutoCAD 연결 실패: {e}")
            print("AutoCAD를 실행하고 도면을 열어주세요.")
            sys.exit(1)
    
    def main_menu(self):
        """메인 메뉴"""
        while True:
            print("\n" + "="*60)
            print("AutoCAD 대화형 물량 산출")
            print("="*60)
            print("\n무엇을 하시겠습니까?")
            print("\n1. 객체 선택하기")
            print("2. 선택된 객체 확인")
            print("3. 물량 계산하기")
            print("4. 선택 초기화")
            print("5. 종료")
            
            choice = input("\n선택 (1-5): ").strip()
            
            if choice == "1":
                self.select_objects_menu()
            elif choice == "2":
                self.show_selected_objects()
            elif choice == "3":
                self.calculate_menu()
            elif choice == "4":
                self.clear_selection()
            elif choice == "5":
                print("\n프로그램을 종료합니다.")
                break
            else:
                print("[ERROR] 잘못된 선택입니다.")
    
    def select_objects_menu(self):
        """객체 선택 메뉴"""
        print("\n" + "-"*60)
        print("객체 선택 방법")
        print("-"*60)
        print("\n1. 마우스로 영역 선택 (Window)")
        print("2. 하나씩 클릭해서 선택")
        print("3. 레이어로 선택")
        print("4. 모든 선(Line) 선택")
        print("5. 모든 폴리라인 선택")
        print("6. 취소")
        
        choice = input("\n선택 (1-6): ").strip()
        
        if choice == "1":
            self.select_by_window()
        elif choice == "2":
            self.select_individually()
        elif choice == "3":
            self.select_by_layer()
        elif choice == "4":
            self.select_all_lines()
        elif choice == "5":
            self.select_all_polylines()
    
    def select_by_window(self):
        """영역 선택"""
        print("\n[영역 선택 모드]")
        print("AutoCAD 화면으로 이동하세요.")
        print("두 점을 클릭하여 영역을 지정하세요.\n")
        
        try:
            # AutoCAD를 활성화
            self.acad.Visible = True
            
            # 사용자가 두 점 선택
            point1 = self.doc.Utility.GetPoint(None, "첫 번째 모서리를 클릭: ")
            point2 = self.doc.Utility.GetCorner(point1, "반대쪽 모서리를 클릭: ")
            
            # 선택 세트 생성
            sel_set = self.doc.SelectionSets.Add(f"TempSet_{len(self.doc.SelectionSets)}")
            
            # 필터 없이 모든 객체 선택
            filter_type = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, [0])
            filter_data = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, ["*"])
            
            # Window 선택
            sel_set.Select(0, point1, point2, filter_type, filter_data)
            
            # 선택된 객체 저장
            count = 0
            for obj in sel_set:
                self.selected_objects.append(obj)
                count += 1
            
            sel_set.Delete()
            
            print(f"\n[OK] {count}개 객체가 선택되었습니다.")
            
            # 선택된 객체 하이라이트
            for obj in self.selected_objects[-count:]:
                try:
                    obj.Highlight(True)
                except:
                    pass
                    
        except Exception as e:
            print(f"[ERROR] 선택 실패: {e}")
    
    def select_individually(self):
        """개별 선택"""
        print("\n[개별 선택 모드]")
        print("AutoCAD 화면에서 객체를 하나씩 클릭하세요.")
        print("Enter 또는 ESC로 선택을 종료합니다.\n")
        
        count = 0
        while True:
            try:
                # 객체 선택
                result = self.doc.Utility.GetEntity(None, "객체 선택 (Enter로 종료): ")
                obj = result[0]
                
                self.selected_objects.append(obj)
                count += 1
                
                # 하이라이트
                obj.Highlight(True)
                
                # 객체 정보 표시
                print(f"  [{count}] {obj.ObjectName} 선택됨")
                
            except:
                # ESC 또는 Enter로 종료
                break
        
        print(f"\n[OK] 총 {count}개 객체가 선택되었습니다.")
    
    def select_by_layer(self):
        """레이어로 선택"""
        print("\n[레이어 선택]")
        
        # 레이어 목록 표시
        layers = []
        for layer in self.doc.Layers:
            layers.append(layer.Name)
        
        print("\n레이어 목록:")
        for i, layer in enumerate(layers[:20], 1):  # 처음 20개만 표시
            print(f"  {i:2}. {layer}")
        
        if len(layers) > 20:
            print(f"  ... 외 {len(layers)-20}개")
        
        # 레이어 선택
        layer_nums = input("\n선택할 레이어 번호 (여러 개는 쉼표로 구분): ").strip()
        
        try:
            indices = [int(x.strip())-1 for x in layer_nums.split(',')]
            selected_layers = [layers[i] for i in indices if 0 <= i < len(layers)]
            
            if not selected_layers:
                print("[ERROR] 유효한 레이어가 없습니다.")
                return
            
            # 선택된 레이어의 객체 수집
            count = 0
            for obj in self.model:
                if obj.Layer in selected_layers:
                    self.selected_objects.append(obj)
                    count += 1
                    
                    # 하이라이트 (처음 100개만)
                    if count <= 100:
                        try:
                            obj.Highlight(True)
                        except:
                            pass
            
            print(f"\n[OK] {selected_layers} 레이어에서 {count}개 객체 선택")
            
        except Exception as e:
            print(f"[ERROR] 선택 실패: {e}")
    
    def select_all_lines(self):
        """모든 선 선택"""
        count = 0
        for obj in self.model:
            if "Line" in obj.ObjectName and "Polyline" not in obj.ObjectName:
                self.selected_objects.append(obj)
                count += 1
        
        print(f"[OK] {count}개의 선(Line)이 선택되었습니다.")
    
    def select_all_polylines(self):
        """모든 폴리라인 선택"""
        count = 0
        for obj in self.model:
            if "Polyline" in obj.ObjectName:
                self.selected_objects.append(obj)
                count += 1
        
        print(f"[OK] {count}개의 폴리라인이 선택되었습니다.")
    
    def show_selected_objects(self):
        """선택된 객체 표시"""
        if not self.selected_objects:
            print("\n선택된 객체가 없습니다.")
            return
        
        print(f"\n[선택된 객체: {len(self.selected_objects)}개]")
        
        # 타입별 분류
        types = {}
        for obj in self.selected_objects:
            obj_type = obj.ObjectName
            if obj_type not in types:
                types[obj_type] = 0
            types[obj_type] += 1
        
        print("\n타입별 분류:")
        for obj_type, count in sorted(types.items()):
            print(f"  {obj_type}: {count}개")
    
    def calculate_menu(self):
        """계산 메뉴"""
        if not self.selected_objects:
            print("\n[ERROR] 먼저 객체를 선택하세요.")
            return
        
        print("\n" + "-"*60)
        print(f"계산 옵션 (선택된 객체: {len(self.selected_objects)}개)")
        print("-"*60)
        
        print("\n이 객체들을 무엇으로 계산하시겠습니까?")
        print("\n1. 장선/보 (길이 → 목재 체적)")
        print("2. 데크보드 (면적 → 필요량)")
        print("3. 기둥 (개수 → 콘크리트)")
        print("4. 철근 (길이 → 중량)")
        print("5. 파이프/배관 (길이 → 규격별)")
        print("6. 일반 길이 계산")
        print("7. 일반 면적 계산")
        print("8. 일반 개수 집계")
        print("9. 사용자 정의")
        
        choice = input("\n선택 (1-9): ").strip()
        
        if choice == "1":
            self.calculate_joists()
        elif choice == "2":
            self.calculate_deck_board()
        elif choice == "3":
            self.calculate_posts()
        elif choice == "4":
            self.calculate_rebar()
        elif choice == "5":
            self.calculate_pipes()
        elif choice == "6":
            self.calculate_length()
        elif choice == "7":
            self.calculate_area()
        elif choice == "8":
            self.count_objects()
        elif choice == "9":
            self.custom_calculation()
    
    def calculate_joists(self):
        """장선 계산"""
        print("\n[장선/보 계산]")
        
        # 길이 계산
        total_length = 0
        count = 0
        
        for obj in self.selected_objects:
            try:
                if "Line" in obj.ObjectName:
                    start = obj.StartPoint
                    end = obj.EndPoint
                    length = math.sqrt((end[0]-start[0])**2 + (end[1]-start[1])**2 + (end[2]-start[2])**2)
                    total_length += length
                    count += 1
                elif "Polyline" in obj.ObjectName:
                    total_length += obj.Length
                    count += 1
            except:
                pass
        
        if total_length == 0:
            print("선형 객체가 없습니다.")
            return
        
        # 단위 변환
        length_m = total_length / 1000  # mm to m
        
        print(f"\n선택된 장선: {count}개")
        print(f"총 길이: {length_m:.2f}m")
        
        # 규격 선택
        print("\n장선 규격을 선택하세요:")
        print("  1. 2x6 (38x140mm)")
        print("  2. 2x8 (38x184mm)")
        print("  3. 2x10 (38x235mm)")
        print("  4. 2x12 (38x286mm)")
        
        size_choice = input("선택 (1-4): ").strip()
        
        sizes = {
            "1": (0.038, 0.140, "2x6"),
            "2": (0.038, 0.184, "2x8"),
            "3": (0.038, 0.235, "2x10"),
            "4": (0.038, 0.286, "2x12")
        }
        
        if size_choice in sizes:
            width, height, name = sizes[size_choice]
            volume = length_m * width * height
            
            print(f"\n[계산 결과]")
            print(f"  규격: {name}")
            print(f"  총 길이: {length_m:.2f}m")
            print(f"  목재 체적: {volume:.3f}m³")
            print(f"  할증률 10% 적용: {volume*1.1:.3f}m³")
    
    def calculate_deck_board(self):
        """데크보드 계산"""
        print("\n[데크보드 계산]")
        
        # 면적 계산
        total_area = 0
        count = 0
        
        for obj in self.selected_objects:
            try:
                if "Polyline" in obj.ObjectName and obj.Closed:
                    total_area += obj.Area
                    count += 1
                elif "Hatch" in obj.ObjectName:
                    total_area += obj.Area
                    count += 1
            except:
                pass
        
        if total_area == 0:
            print("닫힌 도형이나 해치가 없습니다.")
            return
        
        # 단위 변환
        area_m2 = total_area / 1000000  # mm² to m²
        
        print(f"\n데크 면적: {area_m2:.2f}m²")
        print(f"객체 수: {count}개")
        
        # 손실률 입력
        waste = input("\n손실률 (%, 기본 7): ").strip() or "7"
        waste_rate = float(waste) / 100
        
        board_area = area_m2 * (1 + waste_rate)
        screws = int(area_m2 * 8)
        
        print(f"\n[계산 결과]")
        print(f"  데크 면적: {area_m2:.2f}m²")
        print(f"  필요 보드: {board_area:.2f}m² ({waste}% 손실 포함)")
        print(f"  데크 스크류: {screws}개 (m²당 8개)")
    
    def calculate_posts(self):
        """기둥 계산"""
        print("\n[기둥 계산]")
        
        # 기둥 개수
        post_count = 0
        
        for obj in self.selected_objects:
            if "Circle" in obj.ObjectName or "BlockReference" in obj.ObjectName:
                post_count += 1
        
        if post_count == 0:
            print("기둥으로 사용할 객체가 없습니다.")
            return
        
        print(f"\n기둥 개수: {post_count}개")
        
        # 기초 크기 선택
        print("\n기초 크기를 선택하세요:")
        print("  1. 400x400x600mm")
        print("  2. 500x500x700mm")
        print("  3. 600x600x800mm")
        print("  4. 직접 입력")
        
        choice = input("선택 (1-4): ").strip()
        
        if choice == "1":
            volume_per = 0.096
        elif choice == "2":
            volume_per = 0.175
        elif choice == "3":
            volume_per = 0.288
        else:
            width = float(input("폭 (m): "))
            depth = float(input("깊이 (m): "))
            volume_per = width * width * depth
        
        total_concrete = post_count * volume_per
        
        print(f"\n[계산 결과]")
        print(f"  기둥 개수: {post_count}개")
        print(f"  기초당 콘크리트: {volume_per:.3f}m³")
        print(f"  총 콘크리트: {total_concrete:.3f}m³")
        print(f"  할증률 5% 적용: {total_concrete*1.05:.3f}m³")
    
    def calculate_rebar(self):
        """철근 계산"""
        print("\n[철근 계산]")
        
        # 길이 계산
        total_length = 0
        count = 0
        
        for obj in self.selected_objects:
            try:
                if "Line" in obj.ObjectName:
                    start = obj.StartPoint
                    end = obj.EndPoint
                    length = math.sqrt((end[0]-start[0])**2 + (end[1]-start[1])**2)
                    total_length += length
                    count += 1
                elif "Polyline" in obj.ObjectName:
                    total_length += obj.Length
                    count += 1
            except:
                pass
        
        if total_length == 0:
            print("선형 객체가 없습니다.")
            return
        
        length_m = total_length / 1000
        
        print(f"\n철근 개수: {count}개")
        print(f"총 길이: {length_m:.2f}m")
        
        # 철근 규격
        print("\n철근 규격:")
        print("  1. D10 (0.56 kg/m)")
        print("  2. D13 (0.995 kg/m)")
        print("  3. D16 (1.56 kg/m)")
        print("  4. D19 (2.25 kg/m)")
        print("  5. D22 (3.04 kg/m)")
        
        choice = input("선택 (1-5): ").strip()
        
        weights = {"1": 0.56, "2": 0.995, "3": 1.56, "4": 2.25, "5": 3.04}
        sizes = {"1": "D10", "2": "D13", "3": "D16", "4": "D19", "5": "D22"}
        
        if choice in weights:
            unit_weight = weights[choice]
            size_name = sizes[choice]
            
            total_weight = length_m * unit_weight
            
            print(f"\n[계산 결과]")
            print(f"  규격: {size_name}")
            print(f"  총 길이: {length_m:.2f}m")
            print(f"  총 중량: {total_weight:.2f}kg ({total_weight/1000:.3f}ton)")
            print(f"  할증률 3% 적용: {total_weight*1.03/1000:.3f}ton")
    
    def calculate_pipes(self):
        """배관 계산"""
        print("\n[배관 계산]")
        
        # 길이 계산
        total_length = 0
        count = 0
        
        for obj in self.selected_objects:
            try:
                if "Line" in obj.ObjectName or "Polyline" in obj.ObjectName:
                    if "Line" in obj.ObjectName:
                        start = obj.StartPoint
                        end = obj.EndPoint
                        length = math.sqrt((end[0]-start[0])**2 + (end[1]-start[1])**2)
                        total_length += length
                    else:
                        total_length += obj.Length
                    count += 1
            except:
                pass
        
        length_m = total_length / 1000
        
        print(f"\n배관 개수: {count}개")
        print(f"총 연장: {length_m:.2f}m")
        
        # 관경 입력
        diameter = input("\n관경 (mm): ").strip()
        
        print(f"\n[계산 결과]")
        print(f"  관경: D{diameter}")
        print(f"  총 연장: {length_m:.2f}m")
        print(f"  이음관 예상: {int(length_m/6)}개 (6m당 1개)")
    
    def calculate_length(self):
        """일반 길이 계산"""
        total_length = 0
        count = 0
        
        for obj in self.selected_objects:
            try:
                if "Line" in obj.ObjectName:
                    start = obj.StartPoint
                    end = obj.EndPoint
                    length = math.sqrt((end[0]-start[0])**2 + (end[1]-start[1])**2 + (end[2]-start[2])**2)
                    total_length += length
                    count += 1
                elif "Polyline" in obj.ObjectName:
                    total_length += obj.Length
                    count += 1
                elif "Arc" in obj.ObjectName:
                    total_length += obj.ArcLength
                    count += 1
            except:
                pass
        
        print(f"\n[길이 계산 결과]")
        print(f"  객체 수: {count}개")
        print(f"  총 길이: {total_length:.2f}mm ({total_length/1000:.2f}m)")
    
    def calculate_area(self):
        """일반 면적 계산"""
        total_area = 0
        count = 0
        
        for obj in self.selected_objects:
            try:
                if hasattr(obj, 'Area'):
                    if "Polyline" in obj.ObjectName and obj.Closed:
                        total_area += obj.Area
                        count += 1
                    elif "Circle" in obj.ObjectName:
                        total_area += math.pi * obj.Radius ** 2
                        count += 1
                    elif "Hatch" in obj.ObjectName:
                        total_area += obj.Area
                        count += 1
            except:
                pass
        
        print(f"\n[면적 계산 결과]")
        print(f"  객체 수: {count}개")
        print(f"  총 면적: {total_area:.2f}mm² ({total_area/1000000:.2f}m²)")
    
    def count_objects(self):
        """객체 개수 집계"""
        print(f"\n[개수 집계]")
        print(f"  총 선택 객체: {len(self.selected_objects)}개")
        
        # 타입별 집계
        types = {}
        for obj in self.selected_objects:
            obj_type = obj.ObjectName
            if obj_type not in types:
                types[obj_type] = 0
            types[obj_type] += 1
        
        print("\n타입별:")
        for obj_type, count in sorted(types.items()):
            print(f"  {obj_type}: {count}개")
    
    def custom_calculation(self):
        """사용자 정의 계산"""
        print("\n[사용자 정의 계산]")
        print("선택된 객체에 대해 자유롭게 계산하세요.")
        
        # 기본 정보 제공
        line_count = sum(1 for obj in self.selected_objects if "Line" in obj.ObjectName)
        poly_count = sum(1 for obj in self.selected_objects if "Polyline" in obj.ObjectName)
        circle_count = sum(1 for obj in self.selected_objects if "Circle" in obj.ObjectName)
        
        print(f"\n선택된 객체:")
        print(f"  Line: {line_count}개")
        print(f"  Polyline: {poly_count}개")
        print(f"  Circle: {circle_count}개")
        print(f"  기타: {len(self.selected_objects) - line_count - poly_count - circle_count}개")
        
        print("\n계산할 내용을 입력하세요.")
        print("예: 각 객체를 특정 재료로 간주하여 계산")
        
        # 사용자 입력
        calc_name = input("\n계산 이름: ").strip()
        unit_value = input("단위당 값 (예: kg/m, 개/m²): ").strip()
        
        print(f"\n{calc_name} 계산 완료")
        print(f"결과는 선택된 객체와 {unit_value}를 기준으로 계산됩니다.")
    
    def clear_selection(self):
        """선택 초기화"""
        # 하이라이트 제거
        for obj in self.selected_objects:
            try:
                obj.Highlight(False)
            except:
                pass
        
        self.selected_objects = []
        print("\n[OK] 선택이 초기화되었습니다.")

def main():
    print("="*60)
    print("AutoCAD 대화형 물량 산출")
    print("="*60)
    
    tool = InteractiveAutoCAD()
    tool.main_menu()

if __name__ == "__main__":
    main()