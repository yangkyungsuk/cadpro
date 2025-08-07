"""
AutoCAD 플러그인 - Python COM Interface
AutoCAD 내에서 직접 실행되는 물량 산출 도구
"""

import win32com.client
import pythoncom
import math
from typing import List, Dict, Tuple
import tkinter as tk
from tkinter import ttk, messagebox
import json
import os
from datetime import datetime


class AutoCADQuantityTool:
    def __init__(self):
        """AutoCAD와 연결"""
        try:
            # AutoCAD 애플리케이션 연결
            self.acad = win32com.client.Dispatch("AutoCAD.Application")
            self.doc = self.acad.ActiveDocument
            self.model = self.doc.ModelSpace
            
            print(f"[OK] AutoCAD {self.acad.Version} 연결 성공")
            print(f"[OK] 현재 도면: {self.doc.Name}")
            
            # 선택된 객체들 저장
            self.selected_objects = []
            self.calculation_results = {}
            
        except Exception as e:
            print(f"[WARNING] AutoCAD 연결 실패: {e}")
            print("AutoCAD가 실행 중인지 확인하세요.")
            
    def select_objects_by_area(self):
        """영역 선택으로 객체 선택"""
        print("\n[영역 선택 모드]")
        print("AutoCAD에서 계산할 영역을 선택하세요...")
        
        try:
            # 사용자에게 두 점 선택 요청
            point1 = self.doc.Utility.GetPoint(None, "\n첫 번째 모서리를 클릭하세요: ")
            point2 = self.doc.Utility.GetCorner(point1, "\n반대쪽 모서리를 클릭하세요: ")
            
            # 선택 영역 내 객체 가져오기
            selection_set = self.doc.SelectionSets.Add("TempSet")
            
            # 선택 필터 설정 (모든 객체)
            filter_type = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, [0])
            filter_data = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, ["*"])
            
            # Window 선택 (완전히 포함된 객체만)
            selection_set.Select(0, point1, point2, filter_type, filter_data)
            
            self.selected_objects = []
            for obj in selection_set:
                self.selected_objects.append(obj)
            
            selection_set.Delete()
            
            print(f"[OK] {len(self.selected_objects)}개 객체 선택됨")
            
            # 선택된 객체 타입별 분류
            self.classify_selected_objects()
            
            return True
            
        except Exception as e:
            print(f"선택 실패: {e}")
            return False
    
    def select_objects_individually(self):
        """개별 객체 선택"""
        print("\n[개별 선택 모드]")
        print("AutoCAD에서 객체를 하나씩 선택하세요. (ESC로 종료)")
        
        self.selected_objects = []
        
        try:
            while True:
                try:
                    # 단일 객체 선택
                    result = self.doc.Utility.GetEntity(None, "\n객체를 선택하세요 (ESC로 종료): ")
                    obj = result[0]
                    point = result[1]
                    
                    self.selected_objects.append(obj)
                    
                    # 선택된 객체 하이라이트
                    obj.Highlight(True)
                    
                    print(f"  선택: {obj.ObjectName} (총 {len(self.selected_objects)}개)")
                    
                except:
                    # ESC 또는 Enter로 선택 종료
                    break
            
            print(f"\n[OK] 총 {len(self.selected_objects)}개 객체 선택 완료")
            
            # 선택된 객체 타입별 분류
            self.classify_selected_objects()
            
            return True
            
        except Exception as e:
            print(f"선택 실패: {e}")
            return False
    
    def select_by_layer(self):
        """레이어로 선택"""
        print("\n[레이어 선택 모드]")
        
        # 현재 도면의 레이어 목록
        layers = []
        for layer in self.doc.Layers:
            layers.append(layer.Name)
        
        print("\n레이어 목록:")
        for i, layer in enumerate(layers, 1):
            print(f"  {i}. {layer}")
        
        # 레이어 선택
        try:
            choice = input("\n레이어 번호 선택 (여러 개는 쉼표로 구분): ").strip()
            selected_indices = [int(x.strip())-1 for x in choice.split(',')]
            selected_layers = [layers[i] for i in selected_indices if 0 <= i < len(layers)]
            
            if not selected_layers:
                print("잘못된 선택")
                return False
            
            # 선택된 레이어의 객체 수집
            self.selected_objects = []
            for obj in self.model:
                if obj.Layer in selected_layers:
                    self.selected_objects.append(obj)
            
            print(f"[OK] {len(self.selected_objects)}개 객체 선택됨")
            
            # 선택된 객체 타입별 분류
            self.classify_selected_objects()
            
            return True
            
        except Exception as e:
            print(f"레이어 선택 실패: {e}")
            return False
    
    def classify_selected_objects(self):
        """선택된 객체 분류"""
        classification = {
            "Line": [],
            "Polyline": [],
            "Circle": [],
            "Arc": [],
            "Block": [],
            "Hatch": [],
            "Text": [],
            "Dimension": [],
            "기타": []
        }
        
        for obj in self.selected_objects:
            obj_type = obj.ObjectName
            
            if "Line" in obj_type:
                classification["Line"].append(obj)
            elif "Polyline" in obj_type:
                classification["Polyline"].append(obj)
            elif "Circle" in obj_type:
                classification["Circle"].append(obj)
            elif "Arc" in obj_type:
                classification["Arc"].append(obj)
            elif "BlockReference" in obj_type:
                classification["Block"].append(obj)
            elif "Hatch" in obj_type:
                classification["Hatch"].append(obj)
            elif "Text" in obj_type or "MText" in obj_type:
                classification["Text"].append(obj)
            elif "Dimension" in obj_type:
                classification["Dimension"].append(obj)
            else:
                classification["기타"].append(obj)
        
        print("\n[선택된 객체 분류]")
        for obj_type, objects in classification.items():
            if objects:
                print(f"  {obj_type}: {len(objects)}개")
        
        self.object_classification = classification
    
    def calculate_quantities(self):
        """물량 계산 메뉴"""
        print("\n" + "="*60)
        print("물량 계산 옵션")
        print("="*60)
        
        print("\n계산할 항목을 선택하세요:")
        print("  1. 길이 계산 (선, 폴리라인)")
        print("  2. 면적 계산 (닫힌 폴리라인, 해치)")
        print("  3. 체적 계산 (높이/두께 입력)")
        print("  4. 개수 집계 (블록, 객체)")
        print("  5. 철근 물량 (길이→중량)")
        print("  6. 콘크리트 물량")
        print("  7. 토공량 계산")
        print("  8. 사용자 정의 계산")
        
        choice = input("\n선택 (1-8): ").strip()
        
        if choice == "1":
            self.calculate_length()
        elif choice == "2":
            self.calculate_area()
        elif choice == "3":
            self.calculate_volume()
        elif choice == "4":
            self.count_objects()
        elif choice == "5":
            self.calculate_rebar()
        elif choice == "6":
            self.calculate_concrete()
        elif choice == "7":
            self.calculate_earthwork()
        elif choice == "8":
            self.custom_calculation()
    
    def calculate_length(self):
        """길이 계산"""
        print("\n[길이 계산]")
        
        total_length = 0
        count = 0
        
        # Line 객체
        for obj in self.object_classification.get("Line", []):
            try:
                start = obj.StartPoint
                end = obj.EndPoint
                length = math.sqrt((end[0]-start[0])**2 + (end[1]-start[1])**2 + (end[2]-start[2])**2)
                total_length += length
                count += 1
            except:
                pass
        
        # Polyline 객체
        for obj in self.object_classification.get("Polyline", []):
            try:
                length = obj.Length
                total_length += length
                count += 1
            except:
                pass
        
        # Arc 객체
        for obj in self.object_classification.get("Arc", []):
            try:
                length = obj.ArcLength
                total_length += length
                count += 1
            except:
                pass
        
        # 단위 변환 옵션
        print(f"\n계산 결과:")
        print(f"  객체 수: {count}개")
        print(f"  총 길이: {total_length:.3f} 도면단위")
        
        # 실제 단위로 변환
        scale = float(input("\n도면 축척 (1:n의 n값, 실척은 1): ") or "1")
        unit = input("실제 단위 (m/mm/cm, 기본 m): ").strip() or "m"
        
        if unit == "m":
            real_length = total_length / scale / 1000  # 도면이 mm 단위 가정
        elif unit == "cm":
            real_length = total_length / scale / 10
        else:  # mm
            real_length = total_length / scale
        
        print(f"\n최종 결과: {real_length:.3f} {unit}")
        
        self.calculation_results["길이"] = {
            "값": real_length,
            "단위": unit,
            "객체수": count
        }
    
    def calculate_area(self):
        """면적 계산"""
        print("\n[면적 계산]")
        
        total_area = 0
        count = 0
        
        # 닫힌 Polyline
        for obj in self.object_classification.get("Polyline", []):
            try:
                if obj.Closed:
                    area = obj.Area
                    total_area += area
                    count += 1
            except:
                pass
        
        # Circle
        for obj in self.object_classification.get("Circle", []):
            try:
                area = math.pi * obj.Radius ** 2
                total_area += area
                count += 1
            except:
                pass
        
        # Hatch
        for obj in self.object_classification.get("Hatch", []):
            try:
                area = obj.Area
                total_area += area
                count += 1
            except:
                pass
        
        print(f"\n계산 결과:")
        print(f"  객체 수: {count}개")
        print(f"  총 면적: {total_area:.3f} 도면단위²")
        
        # 실제 단위로 변환
        scale = float(input("\n도면 축척 (1:n의 n값, 실척은 1): ") or "1")
        unit = input("실제 단위 (m²/cm²/mm², 기본 m²): ").strip() or "m²"
        
        if unit == "m²":
            real_area = total_area / (scale ** 2) / 1000000  # 도면이 mm 단위 가정
        elif unit == "cm²":
            real_area = total_area / (scale ** 2) / 100
        else:  # mm²
            real_area = total_area / (scale ** 2)
        
        print(f"\n최종 결과: {real_area:.3f} {unit}")
        
        self.calculation_results["면적"] = {
            "값": real_area,
            "단위": unit,
            "객체수": count
        }
    
    def calculate_volume(self):
        """체적 계산"""
        print("\n[체적 계산]")
        
        # 먼저 면적 계산
        total_area = 0
        
        for obj in self.object_classification.get("Polyline", []):
            try:
                if obj.Closed:
                    total_area += obj.Area
            except:
                pass
        
        for obj in self.object_classification.get("Circle", []):
            try:
                total_area += math.pi * obj.Radius ** 2
            except:
                pass
        
        if total_area == 0:
            print("면적이 0입니다. 닫힌 도형을 선택하세요.")
            return
        
        print(f"기준 면적: {total_area:.3f} 도면단위²")
        
        # 높이/두께 입력
        height = float(input("\n높이/두께 입력 (도면단위): "))
        
        volume = total_area * height
        
        print(f"\n계산 결과:")
        print(f"  체적: {volume:.3f} 도면단위³")
        
        # 실제 단위로 변환
        scale = float(input("\n도면 축척 (1:n의 n값, 실척은 1): ") or "1")
        real_volume = volume / (scale ** 3) / 1000000000  # mm³ → m³
        
        print(f"\n최종 결과: {real_volume:.3f} m³")
        
        # 콘크리트 물량 계산 옵션
        if input("\n콘크리트 물량으로 변환? (y/n): ").lower() == 'y':
            waste_rate = float(input("할증률 (%, 기본 5): ") or "5")
            concrete_volume = real_volume * (1 + waste_rate/100)
            print(f"콘크리트 물량: {concrete_volume:.3f} m³")
            
            self.calculation_results["콘크리트"] = {
                "값": concrete_volume,
                "단위": "m³"
            }
        
        self.calculation_results["체적"] = {
            "값": real_volume,
            "단위": "m³"
        }
    
    def count_objects(self):
        """객체 개수 집계"""
        print("\n[객체 개수 집계]")
        
        # 블록 집계
        block_count = {}
        for obj in self.object_classification.get("Block", []):
            try:
                block_name = obj.Name
                if block_name not in block_count:
                    block_count[block_name] = 0
                block_count[block_name] += 1
            except:
                pass
        
        print("\n블록 집계:")
        for name, count in block_count.items():
            print(f"  {name}: {count}개")
        
        # 전체 객체 타입별 집계
        print("\n객체 타입별 집계:")
        total = 0
        for obj_type, objects in self.object_classification.items():
            if objects:
                print(f"  {obj_type}: {len(objects)}개")
                total += len(objects)
        
        print(f"\n총 객체 수: {total}개")
        
        self.calculation_results["개수"] = {
            "블록": block_count,
            "전체": total
        }
    
    def calculate_rebar(self):
        """철근 물량 계산"""
        print("\n[철근 물량 계산]")
        
        # 철근으로 간주할 선 선택
        total_length = 0
        
        for obj in self.object_classification.get("Line", []):
            try:
                start = obj.StartPoint
                end = obj.EndPoint
                length = math.sqrt((end[0]-start[0])**2 + (end[1]-start[1])**2 + (end[2]-start[2])**2)
                total_length += length
            except:
                pass
        
        for obj in self.object_classification.get("Polyline", []):
            try:
                total_length += obj.Length
            except:
                pass
        
        print(f"총 철근 길이: {total_length:.3f} 도면단위")
        
        # 철근 규격 선택
        print("\n철근 규격 선택:")
        rebar_weights = {
            "D10": 0.56,
            "D13": 0.995,
            "D16": 1.56,
            "D19": 2.25,
            "D22": 3.04,
            "D25": 3.98,
            "D29": 5.04,
            "D32": 6.23
        }
        
        for i, (size, weight) in enumerate(rebar_weights.items(), 1):
            print(f"  {i}. {size} ({weight} kg/m)")
        
        choice = int(input("\n선택 (1-8): ")) - 1
        rebar_size = list(rebar_weights.keys())[choice]
        unit_weight = rebar_weights[rebar_size]
        
        # 실제 길이로 변환
        scale = float(input("\n도면 축척 (1:n의 n값, 실척은 1): ") or "1")
        real_length = total_length / scale / 1000  # mm → m
        
        # 중량 계산
        weight = real_length * unit_weight
        waste_rate = float(input("할증률 (%, 기본 3): ") or "3")
        final_weight = weight * (1 + waste_rate/100)
        
        print(f"\n철근 물량:")
        print(f"  규격: {rebar_size}")
        print(f"  길이: {real_length:.2f} m")
        print(f"  중량: {final_weight:.3f} kg ({final_weight/1000:.3f} ton)")
        
        self.calculation_results["철근"] = {
            "규격": rebar_size,
            "길이": real_length,
            "중량": final_weight/1000,
            "단위": "ton"
        }
    
    def calculate_concrete(self):
        """콘크리트 물량 계산"""
        print("\n[콘크리트 물량 계산]")
        
        print("\n구조물 타입 선택:")
        print("  1. 기초 (독립/줄기초)")
        print("  2. 기둥")
        print("  3. 보")
        print("  4. 슬래브")
        print("  5. 벽체")
        print("  6. 계단")
        
        choice = input("\n선택 (1-6): ").strip()
        
        if choice == "1":  # 기초
            print("\n기초 타입:")
            print("  1. 독립기초")
            print("  2. 줄기초")
            
            if input("선택 (1-2): ").strip() == "1":
                # 독립기초 - 원 또는 사각형 선택
                count = len(self.object_classification.get("Circle", [])) + \
                       len([p for p in self.object_classification.get("Polyline", []) if p.Closed])
                
                width = float(input("기초 폭 (m): "))
                depth = float(input("기초 깊이 (m): "))
                
                volume_per = width * width * depth
                total_volume = volume_per * count
                
            else:
                # 줄기초 - 선 길이 기반
                total_length = sum(obj.Length for obj in self.object_classification.get("Polyline", []))
                scale = float(input("도면 축척 (1:n의 n값): ") or "1")
                real_length = total_length / scale / 1000
                
                width = float(input("기초 폭 (m): "))
                depth = float(input("기초 깊이 (m): "))
                
                total_volume = real_length * width * depth
        
        elif choice == "4":  # 슬래브
            total_area = 0
            for obj in self.object_classification.get("Polyline", []):
                if obj.Closed:
                    total_area += obj.Area
            
            scale = float(input("도면 축척 (1:n의 n값): ") or "1")
            real_area = total_area / (scale ** 2) / 1000000
            
            thickness = float(input("슬래브 두께 (m): "))
            total_volume = real_area * thickness
        
        else:
            print("개발 중...")
            return
        
        # 할증률 적용
        waste_rate = float(input("\n할증률 (%, 기본 5): ") or "5")
        final_volume = total_volume * (1 + waste_rate/100)
        
        print(f"\n콘크리트 물량: {final_volume:.3f} m³")
        
        # 콘크리트 강도 선택
        strength = input("콘크리트 강도 (예: 24, 기본 24): ") or "24"
        
        self.calculation_results["콘크리트"] = {
            "강도": f"{strength}MPa",
            "물량": final_volume,
            "단위": "m³"
        }
    
    def calculate_earthwork(self):
        """토공량 계산"""
        print("\n[토공량 계산]")
        print("개발 중...")
    
    def custom_calculation(self):
        """사용자 정의 계산"""
        print("\n[사용자 정의 계산]")
        
        print("계산식을 입력하세요.")
        print("사용 가능한 변수:")
        print("  L: 총 길이")
        print("  A: 총 면적")
        print("  N: 객체 개수")
        
        # 변수 계산
        L = sum(obj.Length for obj in self.object_classification.get("Polyline", []) 
               if hasattr(obj, 'Length'))
        
        A = sum(obj.Area for obj in self.object_classification.get("Polyline", []) 
               if hasattr(obj, 'Area') and obj.Closed)
        
        N = len(self.selected_objects)
        
        print(f"\n현재 값:")
        print(f"  L = {L:.3f}")
        print(f"  A = {A:.3f}")
        print(f"  N = {N}")
        
        # 계산식 입력
        formula = input("\n계산식 (예: L * 1.5 + A * 0.2): ")
        
        try:
            result = eval(formula)
            print(f"\n계산 결과: {result:.3f}")
            
            name = input("이 계산의 이름: ")
            unit = input("단위: ")
            
            self.calculation_results[name] = {
                "값": result,
                "단위": unit,
                "계산식": formula
            }
        except Exception as e:
            print(f"계산 오류: {e}")
    
    def export_results(self):
        """결과 내보내기"""
        if not self.calculation_results:
            print("계산된 결과가 없습니다.")
            return
        
        print("\n" + "="*60)
        print("결과 내보내기")
        print("="*60)
        
        # 결과 요약
        print("\n[계산 결과 요약]")
        for item, data in self.calculation_results.items():
            print(f"\n{item}:")
            for key, value in data.items():
                print(f"  {key}: {value}")
        
        # 저장 옵션
        print("\n저장 형식:")
        print("  1. Excel")
        print("  2. JSON")
        print("  3. AutoCAD 테이블")
        print("  4. 저장하지 않음")
        
        choice = input("\n선택 (1-4): ").strip()
        
        if choice == "2":
            # JSON 저장
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"autocad_quantity_{timestamp}.json"
            
            with open(filename, 'w', encoding='utf-8') as f:
                json.dump(self.calculation_results, f, ensure_ascii=False, indent=2)
            
            print(f"[OK] 저장 완료: {filename}")
        
        elif choice == "3":
            # AutoCAD 테이블 생성
            self.create_autocad_table()
    
    def create_autocad_table(self):
        """AutoCAD 내 테이블 생성"""
        try:
            # 테이블 삽입 위치
            point = self.doc.Utility.GetPoint(None, "\n테이블 삽입 위치를 클릭하세요: ")
            
            # 테이블 생성
            rows = len(self.calculation_results) + 2  # 제목 + 헤더 + 데이터
            cols = 3  # 항목, 값, 단위
            
            table = self.model.AddTable(point, rows, cols, 5, 20)  # 행높이 5, 열너비 20
            
            # 제목
            table.SetText(0, 0, "물량 산출 결과")
            table.MergeCells(0, 0, 0, 2)
            
            # 헤더
            table.SetText(1, 0, "항목")
            table.SetText(1, 1, "값")
            table.SetText(1, 2, "단위")
            
            # 데이터
            row = 2
            for item, data in self.calculation_results.items():
                table.SetText(row, 0, item)
                
                if "값" in data:
                    table.SetText(row, 1, f"{data['값']:.3f}")
                    table.SetText(row, 2, data.get("단위", ""))
                
                row += 1
            
            print("[OK] AutoCAD 테이블 생성 완료")
            
        except Exception as e:
            print(f"테이블 생성 실패: {e}")


def create_gui():
    """GUI 인터페이스"""
    root = tk.Tk()
    root.title("AutoCAD 물량 산출 도구")
    root.geometry("600x400")
    
    # 메인 프레임
    main_frame = ttk.Frame(root, padding="10")
    main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
    
    # 도구 인스턴스
    tool = AutoCADQuantityTool()
    
    # 선택 방법
    ttk.Label(main_frame, text="1. 객체 선택 방법:").grid(row=0, column=0, sticky=tk.W, pady=5)
    
    select_frame = ttk.Frame(main_frame)
    select_frame.grid(row=1, column=0, columnspan=2, pady=5)
    
    ttk.Button(select_frame, text="영역 선택", 
              command=lambda: tool.select_objects_by_area()).pack(side=tk.LEFT, padx=5)
    ttk.Button(select_frame, text="개별 선택", 
              command=lambda: tool.select_objects_individually()).pack(side=tk.LEFT, padx=5)
    ttk.Button(select_frame, text="레이어 선택", 
              command=lambda: tool.select_by_layer()).pack(side=tk.LEFT, padx=5)
    
    # 계산 옵션
    ttk.Label(main_frame, text="2. 계산 항목:").grid(row=2, column=0, sticky=tk.W, pady=5)
    
    calc_frame = ttk.Frame(main_frame)
    calc_frame.grid(row=3, column=0, columnspan=2, pady=5)
    
    ttk.Button(calc_frame, text="길이", 
              command=lambda: tool.calculate_length()).pack(side=tk.LEFT, padx=5)
    ttk.Button(calc_frame, text="면적", 
              command=lambda: tool.calculate_area()).pack(side=tk.LEFT, padx=5)
    ttk.Button(calc_frame, text="체적", 
              command=lambda: tool.calculate_volume()).pack(side=tk.LEFT, padx=5)
    ttk.Button(calc_frame, text="개수", 
              command=lambda: tool.count_objects()).pack(side=tk.LEFT, padx=5)
    
    calc_frame2 = ttk.Frame(main_frame)
    calc_frame2.grid(row=4, column=0, columnspan=2, pady=5)
    
    ttk.Button(calc_frame2, text="철근", 
              command=lambda: tool.calculate_rebar()).pack(side=tk.LEFT, padx=5)
    ttk.Button(calc_frame2, text="콘크리트", 
              command=lambda: tool.calculate_concrete()).pack(side=tk.LEFT, padx=5)
    ttk.Button(calc_frame2, text="사용자정의", 
              command=lambda: tool.custom_calculation()).pack(side=tk.LEFT, padx=5)
    
    # 결과 표시
    ttk.Label(main_frame, text="3. 계산 결과:").grid(row=5, column=0, sticky=tk.W, pady=5)
    
    result_text = tk.Text(main_frame, height=10, width=70)
    result_text.grid(row=6, column=0, columnspan=2, pady=5)
    
    # 내보내기
    ttk.Button(main_frame, text="결과 내보내기", 
              command=lambda: tool.export_results()).grid(row=7, column=0, pady=10)
    
    root.mainloop()


def main():
    print("="*60)
    print("AutoCAD 물량 산출 도구")
    print("="*60)
    
    print("\n실행 모드:")
    print("  1. 명령줄 인터페이스")
    print("  2. GUI 인터페이스")
    
    choice = input("\n선택 (1-2): ").strip()
    
    if choice == "2":
        create_gui()
    else:
        # CLI 모드
        tool = AutoCADQuantityTool()
        
        while True:
            print("\n" + "="*60)
            print("메인 메뉴")
            print("="*60)
            
            print("\n1. 객체 선택")
            print("2. 물량 계산")
            print("3. 결과 내보내기")
            print("4. 종료")
            
            choice = input("\n선택 (1-4): ").strip()
            
            if choice == "1":
                print("\n선택 방법:")
                print("  1. 영역 선택")
                print("  2. 개별 선택")
                print("  3. 레이어 선택")
                
                sub_choice = input("\n선택 (1-3): ").strip()
                
                if sub_choice == "1":
                    tool.select_objects_by_area()
                elif sub_choice == "2":
                    tool.select_objects_individually()
                elif sub_choice == "3":
                    tool.select_by_layer()
            
            elif choice == "2":
                if not tool.selected_objects:
                    print("먼저 객체를 선택하세요.")
                else:
                    tool.calculate_quantities()
            
            elif choice == "3":
                tool.export_results()
            
            elif choice == "4":
                break


if __name__ == "__main__":
    main()