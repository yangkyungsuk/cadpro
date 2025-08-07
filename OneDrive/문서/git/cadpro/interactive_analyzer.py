"""
대화형 CAD 물량 산출 시스템
사용자와 상호작용하며 도면 내용을 정확히 파악
"""

import os
import sys
from pathlib import Path
import ezdxf
from typing import Dict, List, Tuple
import json

sys.path.append(str(Path(__file__).parent / "src"))


class InteractiveCADAnalyzer:
    def __init__(self):
        self.doc = None
        self.filepath = None
        self.project_type = None
        self.layer_assignments = {}
        self.custom_materials = {}
        self.quantities = {}
        
    def start(self):
        """대화형 분석 시작"""
        print("="*60)
        print("대화형 CAD 물량 산출 시스템")
        print("="*60)
        
        # 1. 파일 로드
        self.load_file()
        
        # 2. 프로젝트 타입 확인
        self.identify_project_type()
        
        # 3. 레이어 정보 표시 및 할당
        self.analyze_layers()
        
        # 4. 재료 및 규격 설정
        self.setup_materials()
        
        # 5. 물량 계산
        self.calculate_quantities()
        
        # 6. 결과 저장
        self.save_results()
    
    def load_file(self):
        """CAD 파일 로드"""
        while True:
            filepath = input("\nCAD 파일 경로 (DXF/DWG): ").strip()
            
            if not os.path.exists(filepath):
                print("파일을 찾을 수 없습니다. 다시 입력해주세요.")
                continue
            
            try:
                if filepath.lower().endswith('.dxf'):
                    self.doc = ezdxf.readfile(filepath)
                    self.filepath = filepath
                    print(f"✓ 파일 로드 완료: {filepath}")
                    
                    # 기본 정보 표시
                    print(f"\n[파일 정보]")
                    print(f"  - CAD 버전: {self.doc.dxfversion}")
                    print(f"  - 레이어 수: {len(self.doc.layers)}")
                    
                    # 도면 단위 확인
                    units = self.doc.header.get('$INSUNITS', 0)
                    unit_names = {0: '없음', 1: '인치', 2: '피트', 4: '밀리미터', 5: '센티미터', 6: '미터'}
                    print(f"  - 도면 단위: {unit_names.get(units, '알 수 없음')}")
                    
                    break
                else:
                    print("\n⚠ DWG 파일은 먼저 DXF로 변환해주세요.")
                    print("  변환 도구: ODA File Converter (무료)")
                    print("  다운로드: https://www.opendesign.com/guestfiles/oda_file_converter")
                    
            except Exception as e:
                print(f"파일 로드 실패: {e}")
    
    def identify_project_type(self):
        """프로젝트 타입 식별"""
        print("\n" + "="*60)
        print("프로젝트 타입 선택")
        print("="*60)
        
        # 미리 정의된 타입들
        project_types = {
            "1": "건축 - 주택/빌딩",
            "2": "건축 - 인테리어",
            "3": "토목 - 도로/교량",
            "4": "토목 - 상하수도",
            "5": "조경 - 공원/정원",
            "6": "데크 공사",
            "7": "철골 구조물",
            "8": "전기 설비",
            "9": "기계 설비",
            "0": "사용자 정의"
        }
        
        print("\n도면의 프로젝트 타입을 선택하세요:")
        for key, value in project_types.items():
            print(f"  {key}. {value}")
        
        while True:
            choice = input("\n선택 (0-9): ").strip()
            
            if choice in project_types:
                if choice == "0":
                    self.project_type = input("프로젝트 타입 이름 입력: ").strip()
                else:
                    self.project_type = project_types[choice]
                
                print(f"\n✓ 선택된 프로젝트: {self.project_type}")
                break
            else:
                print("잘못된 입력입니다. 다시 선택해주세요.")
    
    def analyze_layers(self):
        """레이어 분석 및 할당"""
        print("\n" + "="*60)
        print("레이어 분석 및 할당")
        print("="*60)
        
        # 레이어별 엔티티 수 카운트
        layer_stats = {}
        msp = self.doc.modelspace()
        
        for entity in msp:
            layer_name = entity.dxf.layer
            if layer_name not in layer_stats:
                layer_stats[layer_name] = {
                    'count': 0,
                    'types': set()
                }
            layer_stats[layer_name]['count'] += 1
            layer_stats[layer_name]['types'].add(entity.dxftype())
        
        # 레이어 정보 표시
        print("\n[현재 도면의 레이어 목록]")
        print(f"{'번호':<5} {'레이어명':<20} {'엔티티수':<10} {'주요타입':<30}")
        print("-" * 65)
        
        layers = list(layer_stats.keys())
        for i, layer in enumerate(layers, 1):
            stats = layer_stats[layer]
            types = ', '.join(list(stats['types'])[:3])
            print(f"{i:<5} {layer:<20} {stats['count']:<10} {types:<30}")
        
        # 프로젝트별 기본 카테고리
        categories = self.get_categories_for_project()
        
        print(f"\n[{self.project_type} 물량 카테고리]")
        for i, cat in enumerate(categories, 1):
            print(f"  {i}. {cat}")
        
        # 레이어 할당
        print("\n각 레이어를 카테고리에 할당하세요.")
        print("(숫자 입력, 여러 개는 쉼표로 구분, 제외는 0, 전체 자동은 auto)")
        
        for i, layer in enumerate(layers, 1):
            while True:
                assignment = input(f"\n{layer} 레이어 → 카테고리 번호: ").strip()
                
                if assignment.lower() == 'auto':
                    # 자동 할당 로직
                    self.auto_assign_layers(layers[i-1:], categories)
                    break
                elif assignment == '0':
                    self.layer_assignments[layer] = "제외"
                    break
                else:
                    try:
                        cat_indices = [int(x.strip()) for x in assignment.split(',')]
                        assigned_cats = []
                        for idx in cat_indices:
                            if 1 <= idx <= len(categories):
                                assigned_cats.append(categories[idx-1])
                        
                        if assigned_cats:
                            self.layer_assignments[layer] = assigned_cats
                            print(f"  → {layer} = {', '.join(assigned_cats)}")
                            break
                    except:
                        print("잘못된 입력입니다. 다시 입력해주세요.")
            
            # auto 선택시 전체 완료
            if assignment.lower() == 'auto':
                break
        
        # 할당 결과 표시
        print("\n[레이어 할당 결과]")
        for layer, category in self.layer_assignments.items():
            if category != "제외":
                print(f"  {layer} → {category if isinstance(category, str) else ', '.join(category)}")
    
    def get_categories_for_project(self):
        """프로젝트 타입별 카테고리 반환"""
        categories_map = {
            "건축 - 주택/빌딩": ["기초", "기둥", "보", "슬래브", "벽체", "지붕", "창호", "계단", "마감재"],
            "건축 - 인테리어": ["바닥", "벽체", "천장", "가구", "조명", "창호", "설비"],
            "토목 - 도로/교량": ["토공", "포장", "구조물", "배수", "부대시설", "교량상판", "교각"],
            "토목 - 상하수도": ["관로", "맨홀", "밸브", "펌프", "처리시설"],
            "조경 - 공원/정원": ["수목", "잔디", "포장", "시설물", "수경시설", "조명"],
            "데크 공사": ["기초", "기둥", "장선", "데크보드", "난간", "계단", "부속철물"],
            "철골 구조물": ["기둥", "보", "브레이싱", "접합부", "데크플레이트"],
            "전기 설비": ["배관", "배선", "조명기구", "콘센트", "분전반", "접지"],
            "기계 설비": ["배관", "덕트", "장비", "밸브", "보온재", "지지대"]
        }
        
        # 사용자 정의 프로젝트인 경우
        if self.project_type not in categories_map:
            print("\n카테고리를 직접 입력하세요 (쉼표로 구분):")
            custom_cats = input("예) 구조물, 마감재, 설비: ").strip()
            return [c.strip() for c in custom_cats.split(',')]
        
        return categories_map.get(self.project_type, [])
    
    def auto_assign_layers(self, layers, categories):
        """레이어 자동 할당"""
        print("\n자동 할당 중...")
        
        # 키워드 매칭
        keywords = {
            "기초": ["FOUND", "FOOT", "기초", "BASE"],
            "기둥": ["COL", "PILLAR", "POST", "기둥"],
            "보": ["BEAM", "GIRDER", "보"],
            "슬래브": ["SLAB", "FLOOR", "슬라브", "바닥"],
            "벽체": ["WALL", "벽"],
            "지붕": ["ROOF", "지붕"],
            "창호": ["DOOR", "WINDOW", "문", "창"],
            "계단": ["STAIR", "계단"],
            "장선": ["JOIST", "장선"],
            "데크보드": ["DECK", "BOARD", "데크"],
            "난간": ["RAIL", "난간", "HANDRAIL"],
            "배관": ["PIPE", "PLUMB", "배관"],
            "수목": ["TREE", "PLANT", "수목", "나무"],
            "포장": ["PAVE", "포장", "ROAD", "도로"]
        }
        
        for layer in layers:
            layer_upper = layer.upper()
            assigned = False
            
            for category in categories:
                if category in keywords:
                    for keyword in keywords[category]:
                        if keyword in layer_upper:
                            self.layer_assignments[layer] = [category]
                            print(f"  {layer} → {category} (자동)")
                            assigned = True
                            break
                if assigned:
                    break
            
            if not assigned:
                self.layer_assignments[layer] = "제외"
                print(f"  {layer} → 제외 (매칭 없음)")
    
    def setup_materials(self):
        """재료 및 규격 설정"""
        print("\n" + "="*60)
        print("재료 및 규격 설정")
        print("="*60)
        
        print("\n주요 재료의 규격을 설정하시겠습니까?")
        print("  1. 기본값 사용")
        print("  2. 직접 입력")
        
        choice = input("\n선택 (1-2): ").strip()
        
        if choice == "2":
            # 카테고리별 재료 설정
            used_categories = set()
            for cats in self.layer_assignments.values():
                if cats != "제외":
                    if isinstance(cats, list):
                        used_categories.update(cats)
                    else:
                        used_categories.add(cats)
            
            for category in used_categories:
                print(f"\n[{category} 재료 설정]")
                
                # 카테고리별 기본 질문
                if "기둥" in category:
                    size = input("  기둥 크기 (예: 400x400, 기본값 Enter): ").strip()
                    if size:
                        self.custom_materials[category] = {"크기": size}
                
                elif "벽" in category:
                    thickness = input("  벽 두께 (mm, 기본값 200): ").strip()
                    if thickness:
                        self.custom_materials[category] = {"두께": float(thickness)/1000}
                
                elif "슬래브" in category or "바닥" in category:
                    thickness = input("  슬래브 두께 (mm, 기본값 200): ").strip()
                    if thickness:
                        self.custom_materials[category] = {"두께": float(thickness)/1000}
                
                elif "데크" in category:
                    material = input("  데크 재료 (1.방부목 2.합성재, 기본값 1): ").strip()
                    size = input("  데크보드 규격 (예: 2x6, 기본값 Enter): ").strip()
                    self.custom_materials[category] = {
                        "재료": "방부목" if material != "2" else "합성재",
                        "규격": size if size else "2x6"
                    }
                
                elif "철근" in category:
                    sizes = input("  철근 규격 (예: D16,D19,D22): ").strip()
                    if sizes:
                        self.custom_materials[category] = {"규격": sizes.split(',')}
    
    def calculate_quantities(self):
        """물량 계산"""
        print("\n" + "="*60)
        print("물량 계산 중...")
        print("="*60)
        
        msp = self.doc.modelspace()
        
        # 카테고리별 엔티티 수집
        category_entities = {}
        
        for entity in msp:
            layer = entity.dxf.layer
            if layer in self.layer_assignments:
                categories = self.layer_assignments[layer]
                if categories != "제외":
                    if not isinstance(categories, list):
                        categories = [categories]
                    
                    for cat in categories:
                        if cat not in category_entities:
                            category_entities[cat] = {
                                'lines': [],
                                'polylines': [],
                                'circles': [],
                                'blocks': [],
                                'hatches': []
                            }
                        
                        # 엔티티 타입별 분류
                        if entity.dxftype() == 'LINE':
                            start = entity.dxf.start
                            end = entity.dxf.end
                            length = ((end[0]-start[0])**2 + (end[1]-start[1])**2 + (end[2]-start[2])**2)**0.5
                            category_entities[cat]['lines'].append({'entity': entity, 'length': length})
                        
                        elif entity.dxftype() == 'LWPOLYLINE':
                            points = list(entity.get_points())
                            length = 0
                            area = 0
                            
                            for i in range(len(points)-1):
                                dx = points[i+1][0] - points[i][0]
                                dy = points[i+1][1] - points[i][1]
                                length += (dx*dx + dy*dy)**0.5
                            
                            if entity.is_closed and len(points) >= 3:
                                # 면적 계산
                                for i in range(len(points)):
                                    j = (i + 1) % len(points)
                                    area += points[i][0] * points[j][1]
                                    area -= points[j][0] * points[i][1]
                                area = abs(area) / 2
                            
                            category_entities[cat]['polylines'].append({
                                'entity': entity,
                                'length': length,
                                'area': area,
                                'is_closed': entity.is_closed
                            })
                        
                        elif entity.dxftype() == 'CIRCLE':
                            radius = entity.dxf.radius
                            category_entities[cat]['circles'].append({
                                'entity': entity,
                                'radius': radius,
                                'area': 3.14159 * radius * radius
                            })
                        
                        elif entity.dxftype() == 'INSERT':
                            category_entities[cat]['blocks'].append({
                                'entity': entity,
                                'name': entity.dxf.name
                            })
                        
                        elif entity.dxftype() == 'HATCH':
                            category_entities[cat]['hatches'].append({
                                'entity': entity,
                                'pattern': entity.dxf.pattern_name
                            })
        
        # 카테고리별 물량 계산
        for category, entities in category_entities.items():
            self.quantities[category] = self.calculate_category_quantity(category, entities)
        
        # 결과 표시
        print("\n[물량 산출 결과]")
        print("-" * 60)
        
        for category, quantity in self.quantities.items():
            print(f"\n{category}:")
            for key, value in quantity.items():
                if isinstance(value, dict):
                    for sub_key, sub_value in value.items():
                        print(f"  - {key} {sub_key}: {sub_value}")
                else:
                    print(f"  - {key}: {value}")
    
    def calculate_category_quantity(self, category, entities):
        """카테고리별 물량 계산 로직"""
        result = {}
        
        # 선형 요소 합계
        total_length = sum(e['length'] for e in entities['lines'])
        total_length += sum(e['length'] for e in entities['polylines'] if not e['is_closed'])
        
        # 면적 요소 합계
        total_area = sum(e['area'] for e in entities['polylines'] if e['is_closed'])
        total_area += sum(e['area'] for e in entities['circles'])
        
        # 개수 요소
        block_count = len(entities['blocks'])
        
        # 카테고리별 특수 계산
        if "기둥" in category or "POST" in category.upper():
            if entities['circles']:
                result['개수'] = len(entities['circles'])
                result['총단면적'] = f"{total_area:.3f} m²"
                
                # 기둥 높이 입력 받기
                height = input(f"\n{category} 높이 (m, 기본값 3.0): ").strip()
                height = float(height) if height else 3.0
                
                result['체적'] = f"{total_area * height:.3f} m³"
                result['콘크리트'] = f"{total_area * height * 1.05:.3f} m³"  # 5% 할증
        
        elif "벽" in category or "WALL" in category.upper():
            if total_length > 0:
                # 벽 높이 입력
                height = input(f"\n{category} 높이 (m, 기본값 3.0): ").strip()
                height = float(height) if height else 3.0
                
                thickness = self.custom_materials.get(category, {}).get('두께', 0.2)
                
                result['길이'] = f"{total_length:.2f} m"
                result['면적'] = f"{total_length * height:.2f} m²"
                result['체적'] = f"{total_length * height * thickness:.3f} m³"
        
        elif "슬래브" in category or "바닥" in category:
            if total_area > 0:
                thickness = self.custom_materials.get(category, {}).get('두께', 0.2)
                
                result['면적'] = f"{total_area:.2f} m²"
                result['체적'] = f"{total_area * thickness:.3f} m³"
                result['콘크리트'] = f"{total_area * thickness * 1.05:.3f} m³"
        
        elif "데크" in category:
            if total_area > 0:
                result['데크면적'] = f"{total_area:.2f} m²"
                result['데크보드'] = f"{total_area * 1.07:.2f} m²"  # 7% 손실
                result['스크류'] = f"{int(total_area * 8)} 개"  # m²당 8개
            
            if total_length > 0:
                result['장선길이'] = f"{total_length:.2f} m"
        
        elif "배관" in category or "PIPE" in category.upper():
            if total_length > 0:
                result['배관연장'] = f"{total_length:.2f} m"
                
                # 관경 입력
                diameter = input(f"\n{category} 관경 (mm, 기본값 100): ").strip()
                diameter = int(diameter) if diameter else 100
                result['관경'] = f"D{diameter}"
        
        elif "철근" in category:
            if total_length > 0:
                # 철근 규격 입력
                rebar_size = input(f"\n{category} 규격 (예: D16, 기본값 D16): ").strip()
                rebar_size = rebar_size if rebar_size else "D16"
                
                # 단위중량 (kg/m)
                weights = {"D10": 0.56, "D13": 0.995, "D16": 1.56, "D19": 2.25, "D22": 3.04, "D25": 3.98}
                unit_weight = weights.get(rebar_size, 1.56)
                
                result['길이'] = f"{total_length:.2f} m"
                result['중량'] = f"{total_length * unit_weight / 1000:.3f} ton"
                result['규격'] = rebar_size
        
        else:
            # 기본 계산
            if total_length > 0:
                result['길이'] = f"{total_length:.2f} m"
            if total_area > 0:
                result['면적'] = f"{total_area:.2f} m²"
            if block_count > 0:
                result['개수'] = f"{block_count} 개"
        
        return result
    
    def save_results(self):
        """결과 저장"""
        print("\n" + "="*60)
        print("결과 저장")
        print("="*60)
        
        # 저장 옵션
        print("\n저장 형식을 선택하세요:")
        print("  1. Excel")
        print("  2. JSON")
        print("  3. 둘 다")
        print("  4. 저장하지 않음")
        
        choice = input("\n선택 (1-4): ").strip()
        
        if choice in ["1", "2", "3"]:
            os.makedirs("output", exist_ok=True)
            timestamp = Path(self.filepath).stem
            
            output_data = {
                "파일": self.filepath,
                "프로젝트타입": self.project_type,
                "레이어할당": self.layer_assignments,
                "사용자설정": self.custom_materials,
                "물량산출": self.quantities
            }
            
            if choice in ["2", "3"]:
                # JSON 저장
                json_path = f"output/{timestamp}_물량산출.json"
                with open(json_path, 'w', encoding='utf-8') as f:
                    json.dump(output_data, f, ensure_ascii=False, indent=2)
                print(f"✓ JSON 저장: {json_path}")
            
            if choice in ["1", "3"]:
                # Excel 저장
                try:
                    import pandas as pd
                    
                    # 데이터프레임 생성
                    rows = []
                    for category, items in self.quantities.items():
                        for key, value in items.items():
                            rows.append({
                                "카테고리": category,
                                "항목": key,
                                "수량": value
                            })
                    
                    df = pd.DataFrame(rows)
                    excel_path = f"output/{timestamp}_물량산출.xlsx"
                    
                    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                        # 물량 시트
                        df.to_excel(writer, sheet_name='물량산출', index=False)
                        
                        # 설정 시트
                        config_df = pd.DataFrame([
                            ["프로젝트", self.project_type],
                            ["파일", self.filepath]
                        ], columns=["항목", "내용"])
                        config_df.to_excel(writer, sheet_name='설정', index=False)
                    
                    print(f"✓ Excel 저장: {excel_path}")
                    
                except ImportError:
                    print("⚠ pandas가 설치되지 않아 Excel 저장 실패")
        
        print("\n물량 산출 완료!")


def main():
    analyzer = InteractiveCADAnalyzer()
    analyzer.start()


if __name__ == "__main__":
    main()