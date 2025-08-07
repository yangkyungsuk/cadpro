"""
데크 공사 도면 전용 분석기
AI를 활용한 도면 내용 자동 인식
"""

import ezdxf
from typing import Dict, List, Tuple
import math
from config_deck import DECK_LAYER_MAPPING, DECK_MATERIAL_COEFFICIENTS, DECK_CALCULATION_STANDARDS


class DeckAnalyzer:
    def __init__(self, filepath: str):
        self.filepath = filepath
        self.doc = None
        self.deck_elements = {
            "구조재": {},
            "바닥재": {},
            "난간": {},
            "계단": {},
            "기초": {},
            "철물": {}
        }
        
    def analyze_drawing_type(self):
        """도면 타입 자동 분석"""
        if not self.doc:
            return "unknown"
        
        msp = self.doc.modelspace()
        
        # 키워드 카운트
        keywords = {
            "deck": 0,
            "건축": 0,
            "토목": 0,
            "조경": 0,
            "전기": 0,
            "기계": 0
        }
        
        # 레이어 이름 분석
        for layer in self.doc.layers:
            layer_name = layer.dxf.name.upper()
            
            # 데크 관련 키워드
            if any(kw in layer_name for kw in ["DECK", "데크", "BOARD", "JOIST", "장선", "난간", "RAIL"]):
                keywords["deck"] += 2
            
            # 건축 관련
            if any(kw in layer_name for kw in ["A-", "WALL", "DOOR", "WINDOW", "벽", "문", "창"]):
                keywords["건축"] += 1
                
            # 토목 관련
            if any(kw in layer_name for kw in ["C-", "ROAD", "PIPE", "도로", "배관"]):
                keywords["토목"] += 1
                
            # 조경 관련
            if any(kw in layer_name for kw in ["L-", "TREE", "PLANT", "수목", "잔디"]):
                keywords["조경"] += 1
        
        # 텍스트 내용 분석
        for entity in msp:
            if entity.dxftype() in ['TEXT', 'MTEXT']:
                text_content = entity.dxf.text if hasattr(entity.dxf, 'text') else str(entity.text)
                text_upper = text_content.upper()
                
                # 데크 관련 텍스트
                if any(kw in text_upper for kw in ["DECK", "데크", "DECKING", "목재", "방부목", "합성"]):
                    keywords["deck"] += 3
                    
        # 블록 이름 분석
        for entity in msp:
            if entity.dxftype() == 'INSERT':
                block_name = entity.dxf.name.upper()
                if any(kw in block_name for kw in ["POST", "RAIL", "BRACKET", "기둥", "난간"]):
                    keywords["deck"] += 2
        
        # 가장 높은 점수의 타입 반환
        max_type = max(keywords, key=keywords.get)
        max_score = keywords[max_type]
        
        print(f"\n도면 분석 결과:")
        for dtype, score in keywords.items():
            print(f"  - {dtype}: {score}점")
        
        if max_score > 0:
            return max_type
        return "unknown"
    
    def analyze_deck_elements(self):
        """데크 요소 자동 인식"""
        if not self.doc:
            return
        
        msp = self.doc.modelspace()
        
        # 1. 패턴 인식: 반복되는 선 = 장선 또는 데크보드
        self._detect_joists(msp)
        
        # 2. 기둥 위치 인식: 원 또는 사각형
        self._detect_posts(msp)
        
        # 3. 난간 인식: 외곽선 + 일정 높이
        self._detect_railings(msp)
        
        # 4. 계단 인식: 단계적 패턴
        self._detect_stairs(msp)
        
        # 5. 면적 계산: 폐곡선 영역
        self._calculate_deck_area(msp)
    
    def _detect_joists(self, msp):
        """장선 패턴 자동 감지"""
        parallel_lines = []
        
        # 평행선 그룹 찾기
        for entity in msp:
            if entity.dxftype() == 'LINE':
                line = entity
                start = line.dxf.start
                end = line.dxf.end
                
                # 수평선인지 확인 (Y 좌표 차이가 작음)
                if abs(end[1] - start[1]) < 0.1:
                    parallel_lines.append({
                        'y': (start[1] + end[1]) / 2,
                        'x_start': min(start[0], end[0]),
                        'x_end': max(start[0], end[0]),
                        'length': abs(end[0] - start[0])
                    })
        
        # 일정 간격으로 배치된 선들 찾기
        if len(parallel_lines) > 3:
            parallel_lines.sort(key=lambda x: x['y'])
            
            # 간격 계산
            spacings = []
            for i in range(len(parallel_lines) - 1):
                spacing = parallel_lines[i+1]['y'] - parallel_lines[i]['y']
                spacings.append(spacing)
            
            # 평균 간격
            if spacings:
                avg_spacing = sum(spacings) / len(spacings)
                
                # 장선으로 판단 (300~600mm 간격)
                if 0.3 <= avg_spacing <= 0.6:
                    total_length = sum(line['length'] for line in parallel_lines)
                    self.deck_elements["구조재"]["장선"] = {
                        "개수": len(parallel_lines),
                        "평균간격": round(avg_spacing, 3),
                        "총길이": round(total_length, 2)
                    }
                    print(f"  장선 감지: {len(parallel_lines)}개, 간격 {avg_spacing*1000:.0f}mm")
    
    def _detect_posts(self, msp):
        """기둥 위치 자동 감지"""
        posts = []
        
        for entity in msp:
            # 원형 기둥
            if entity.dxftype() == 'CIRCLE':
                radius = entity.dxf.radius
                # 일반적인 기둥 크기 (반경 40~150mm)
                if 0.04 <= radius <= 0.15:
                    center = entity.dxf.center
                    posts.append({
                        'type': 'circular',
                        'position': (center[0], center[1]),
                        'size': radius * 2
                    })
            
            # 사각형 기둥 (폐곡선)
            elif entity.dxftype() == 'LWPOLYLINE':
                if entity.is_closed:
                    points = list(entity.get_points())
                    if len(points) == 4:  # 사각형
                        # 크기 계산
                        x_coords = [p[0] for p in points]
                        y_coords = [p[1] for p in points]
                        width = max(x_coords) - min(x_coords)
                        height = max(y_coords) - min(y_coords)
                        
                        # 일반적인 기둥 크기 (80~200mm)
                        if 0.08 <= width <= 0.2 and 0.08 <= height <= 0.2:
                            center_x = sum(x_coords) / 4
                            center_y = sum(y_coords) / 4
                            posts.append({
                                'type': 'rectangular',
                                'position': (center_x, center_y),
                                'size': (width, height)
                            })
        
        if posts:
            self.deck_elements["구조재"]["기둥"] = {
                "개수": len(posts),
                "위치": posts
            }
            print(f"  기둥 감지: {len(posts)}개")
    
    def _detect_railings(self, msp):
        """난간 자동 감지"""
        potential_railings = []
        
        for entity in msp:
            if entity.dxftype() == 'LWPOLYLINE':
                if not entity.is_closed:
                    points = list(entity.get_points())
                    
                    # 난간은 보통 긴 선형 요소
                    if len(points) >= 2:
                        total_length = 0
                        for i in range(len(points) - 1):
                            dx = points[i+1][0] - points[i][0]
                            dy = points[i+1][1] - points[i][1]
                            total_length += math.sqrt(dx*dx + dy*dy)
                        
                        # 1m 이상의 선형 요소
                        if total_length > 1.0:
                            potential_railings.append({
                                'length': total_length,
                                'points': len(points)
                            })
        
        if potential_railings:
            total_railing = sum(r['length'] for r in potential_railings)
            self.deck_elements["난간"]["핸드레일"] = {
                "개수": len(potential_railings),
                "총길이": round(total_railing, 2)
            }
            print(f"  난간 감지: {total_railing:.2f}m")
    
    def _detect_stairs(self, msp):
        """계단 패턴 자동 감지"""
        # 계단은 일정한 간격의 평행선 또는 계단 모양 폴리라인
        step_lines = []
        
        for entity in msp:
            if entity.dxftype() == 'LINE':
                start = entity.dxf.start
                end = entity.dxf.end
                
                # 수평선이면서 길이가 일정 범위
                if abs(end[1] - start[1]) < 0.05:
                    length = abs(end[0] - start[0])
                    if 0.3 <= length <= 2.0:  # 계단 폭 범위
                        step_lines.append({
                            'y': start[1],
                            'length': length
                        })
        
        # 일정 간격으로 배치된 계단 찾기
        if len(step_lines) > 2:
            step_lines.sort(key=lambda x: x['y'])
            
            # 간격 확인
            uniform_spacing = True
            spacings = []
            for i in range(len(step_lines) - 1):
                spacing = step_lines[i+1]['y'] - step_lines[i]['y']
                spacings.append(spacing)
            
            if spacings:
                avg_spacing = sum(spacings) / len(spacings)
                # 계단 높이 범위 (150~200mm)
                if 0.15 <= avg_spacing <= 0.25:
                    self.deck_elements["계단"]["단수"] = len(step_lines)
                    self.deck_elements["계단"]["단높이"] = round(avg_spacing, 3)
                    self.deck_elements["계단"]["단너비"] = round(step_lines[0]['length'], 2)
                    print(f"  계단 감지: {len(step_lines)}단")
    
    def _calculate_deck_area(self, msp):
        """데크 면적 자동 계산"""
        deck_areas = []
        
        for entity in msp:
            if entity.dxftype() == 'LWPOLYLINE':
                if entity.is_closed:
                    points = list(entity.get_points())
                    
                    # 면적 계산 (Shoelace formula)
                    area = 0
                    for i in range(len(points)):
                        j = (i + 1) % len(points)
                        area += points[i][0] * points[j][1]
                        area -= points[j][0] * points[i][1]
                    area = abs(area) / 2
                    
                    # 데크 크기 범위 (5~500m²)
                    if 5 <= area <= 500:
                        deck_areas.append(area)
            
            elif entity.dxftype() == 'HATCH':
                # 해치 영역도 데크 면적으로 고려
                try:
                    # 해치 경계에서 면적 계산
                    for path in entity.paths:
                        if hasattr(path, 'vertices'):
                            vertices = path.vertices
                            if len(vertices) >= 3:
                                area = 0
                                for i in range(len(vertices)):
                                    j = (i + 1) % len(vertices)
                                    area += vertices[i][0] * vertices[j][1]
                                    area -= vertices[j][0] * vertices[i][1]
                                area = abs(area) / 2
                                
                                if 5 <= area <= 500:
                                    deck_areas.append(area)
                except:
                    pass
        
        if deck_areas:
            total_area = sum(deck_areas)
            self.deck_elements["바닥재"]["데크면적"] = round(total_area, 2)
            print(f"  데크 면적: {total_area:.2f}m²")
    
    def calculate_quantities(self):
        """데크 물량 자동 계산"""
        quantities = {}
        
        # 1. 바닥재 계산
        if "데크면적" in self.deck_elements.get("바닥재", {}):
            area = self.deck_elements["바닥재"]["데크면적"]
            
            # 데크보드 물량 (손실률 포함)
            board_area = area * DECK_CALCULATION_STANDARDS["데크보드"]["손실률"]
            quantities["데크보드"] = {
                "면적": round(board_area, 2),
                "단위": "m²"
            }
            
            # 데크 스크류 (m²당 8개)
            screws = area * DECK_MATERIAL_COEFFICIENTS["철물"]["데크스크류"]
            quantities["데크스크류"] = {
                "수량": round(screws),
                "단위": "개"
            }
        
        # 2. 장선 계산
        if "장선" in self.deck_elements.get("구조재", {}):
            joist_data = self.deck_elements["구조재"]["장선"]
            
            # 목재 체적 (2x10 기준)
            volume = joist_data["총길이"] * DECK_MATERIAL_COEFFICIENTS["목재"]["방부목"]["2x10"]
            quantities["장선"] = {
                "길이": joist_data["총길이"],
                "체적": round(volume, 3),
                "개수": joist_data["개수"],
                "단위": "m³"
            }
            
            # 조이스트 행거
            hangers = joist_data["개수"] * 2  # 양끝
            quantities["조이스트행거"] = {
                "수량": hangers,
                "단위": "개"
            }
        
        # 3. 기둥 계산
        if "기둥" in self.deck_elements.get("구조재", {}):
            post_data = self.deck_elements["구조재"]["기둥"]
            post_count = post_data["개수"]
            
            # 기둥 높이 추정 (기본 2.5m)
            post_height = 2.5
            post_volume = post_count * post_height * DECK_MATERIAL_COEFFICIENTS["목재"]["방부목"]["6x6"]
            
            quantities["기둥"] = {
                "개수": post_count,
                "체적": round(post_volume, 3),
                "단위": "m³"
            }
            
            # 포스트 베이스
            quantities["포스트베이스"] = {
                "수량": post_count,
                "단위": "개"
            }
            
            # 기초 콘크리트
            foundation_volume = post_count * DECK_MATERIAL_COEFFICIENTS["기초"]["독립기초"]["400x400x600"]
            quantities["기초콘크리트"] = {
                "체적": round(foundation_volume, 3),
                "단위": "m³"
            }
        
        # 4. 난간 계산
        if "핸드레일" in self.deck_elements.get("난간", {}):
            rail_data = self.deck_elements["난간"]["핸드레일"]
            rail_length = rail_data["총길이"]
            
            quantities["난간"] = {
                "길이": rail_length,
                "중량": round(rail_length * DECK_MATERIAL_COEFFICIENTS["난간"]["알루미늄"], 2),
                "단위": "kg"
            }
            
            # 난간 기둥 (1.2m 간격)
            rail_posts = int(rail_length / DECK_CALCULATION_STANDARDS["난간"]["기둥간격"]) + 1
            quantities["난간기둥"] = {
                "수량": rail_posts,
                "단위": "개"
            }
        
        # 5. 계단 계산
        if "단수" in self.deck_elements.get("계단", {}):
            stair_data = self.deck_elements["계단"]
            steps = stair_data["단수"]
            width = stair_data.get("단너비", 1.0)
            
            # 계단 면적
            stair_area = steps * width * 0.3  # 단너비 30cm
            quantities["계단재"] = {
                "면적": round(stair_area, 2),
                "단수": steps,
                "단위": "m²"
            }
        
        return quantities
    
    def load_and_analyze(self):
        """파일 로드 및 자동 분석"""
        try:
            # DXF 파일 로드
            self.doc = ezdxf.readfile(self.filepath)
            print(f"\n파일 로드 완료: {self.filepath}")
            
            # 도면 타입 분석
            drawing_type = self.analyze_drawing_type()
            
            if drawing_type == "deck":
                print("\n✓ 데크 공사 도면으로 확인됨")
                print("\n데크 요소 자동 분석 중...")
                
                # 데크 요소 분석
                self.analyze_deck_elements()
                
                # 물량 계산
                quantities = self.calculate_quantities()
                
                print("\n데크 물량 산출 결과:")
                print("-" * 40)
                for item, data in quantities.items():
                    if isinstance(data, dict):
                        unit = data.get("단위", "")
                        for key, value in data.items():
                            if key != "단위":
                                print(f"  {item} - {key}: {value} {unit if key != '개수' else '개'}")
                
                return True, drawing_type, quantities
            else:
                print(f"\n✗ 데크 도면이 아닌 {drawing_type} 도면으로 판단됨")
                return False, drawing_type, {}
                
        except Exception as e:
            print(f"분석 실패: {e}")
            import traceback
            traceback.print_exc()
            return False, "error", {}