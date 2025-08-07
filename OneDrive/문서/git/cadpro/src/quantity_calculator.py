"""
물량 계산 모듈
건축, 토목, 조경 분야별 물량 산출 알고리즘
"""

import math
from typing import Dict, List, Tuple
from config import MATERIAL_COEFFICIENTS, LANDSCAPE_COEFFICIENTS


class QuantityCalculator:
    def __init__(self, classified_entities: Dict):
        """
        물량 계산기 초기화
        
        Args:
            classified_entities: 분야별로 분류된 엔티티 딕셔너리
        """
        self.entities = classified_entities
        self.quantities = {
            "건축": {},
            "토목": {},
            "조경": {}
        }
    
    def calculate_all(self) -> Dict:
        """모든 분야의 물량 계산"""
        self.calculate_architecture()
        self.calculate_civil()
        self.calculate_landscape()
        return self.quantities
    
    def calculate_architecture(self):
        """건축 분야 물량 계산"""
        arch_entities = self.entities.get("건축", {})
        
        # 벽체 물량
        if "벽체" in arch_entities:
            wall_quantity = self._calculate_walls(arch_entities["벽체"])
            self.quantities["건축"]["벽체"] = wall_quantity
        
        # 기둥 물량
        if "기둥" in arch_entities:
            column_quantity = self._calculate_columns(arch_entities["기둥"])
            self.quantities["건축"]["기둥"] = column_quantity
        
        # 슬래브 물량
        if "슬래브" in arch_entities:
            slab_quantity = self._calculate_slabs(arch_entities["슬래브"])
            self.quantities["건축"]["슬래브"] = slab_quantity
        
        # 문/창문 개수
        if "문" in arch_entities:
            door_quantity = self._count_blocks(arch_entities["문"])
            self.quantities["건축"]["문"] = door_quantity
        
        if "창문" in arch_entities:
            window_quantity = self._count_blocks(arch_entities["창문"])
            self.quantities["건축"]["창문"] = window_quantity
    
    def _calculate_walls(self, wall_entities: Dict) -> Dict:
        """벽체 물량 계산"""
        total_length = 0
        total_area = 0
        wall_height = 3.0  # 기본 층고 3m (설정 가능)
        
        # 선형 벽체 (LINE, POLYLINE)
        for line in wall_entities.get("lines", []):
            total_length += line["length"]
        
        for poly in wall_entities.get("polylines", []):
            if not poly["is_closed"]:
                total_length += poly["length"]
            else:
                # 닫힌 폴리라인은 면적으로 계산
                total_area += poly["area"]
        
        # 해치로 표현된 벽체
        for hatch in wall_entities.get("hatches", []):
            total_area += hatch["area"]
        
        # 벽체 면적 = 길이 × 높이
        wall_area_from_length = total_length * wall_height
        total_wall_area = wall_area_from_length + total_area
        
        # 벽 두께 가정 (200mm)
        wall_thickness = 0.2
        wall_volume = total_wall_area * wall_thickness
        
        # 콘크리트 물량
        concrete_volume = wall_volume * MATERIAL_COEFFICIENTS["콘크리트"]["할증률"]
        
        # 벽돌 개수 (옵션)
        brick_count = total_wall_area * MATERIAL_COEFFICIENTS["벽돌"]["시멘트벽돌"] * \
                     MATERIAL_COEFFICIENTS["벽돌"]["할증률"]
        
        return {
            "벽체길이": round(total_length, 2),
            "벽체면적": round(total_wall_area, 2),
            "벽체체적": round(wall_volume, 2),
            "콘크리트": round(concrete_volume, 2),
            "벽돌": int(brick_count),
            "단위": {
                "벽체길이": "m",
                "벽체면적": "m²",
                "벽체체적": "m³",
                "콘크리트": "m³",
                "벽돌": "개"
            }
        }
    
    def _calculate_columns(self, column_entities: Dict) -> Dict:
        """기둥 물량 계산"""
        column_count = 0
        total_volume = 0
        column_height = 3.0  # 기본 층고
        
        # 원형 기둥
        for circle in column_entities.get("circles", []):
            column_count += 1
            volume = circle["area"] * column_height
            total_volume += volume
        
        # 사각 기둥 (닫힌 폴리라인)
        for poly in column_entities.get("polylines", []):
            if poly["is_closed"] and poly["area"] > 0:
                column_count += 1
                volume = poly["area"] * column_height
                total_volume += volume
        
        # 블록으로 표현된 기둥
        column_count += len(column_entities.get("blocks", []))
        
        # 콘크리트 물량
        concrete_volume = total_volume * MATERIAL_COEFFICIENTS["콘크리트"]["할증률"]
        
        # 철근 물량 (체적의 2% 가정)
        rebar_weight = total_volume * 0.02 * 7850  # kg (철근 밀도: 7850 kg/m³)
        
        return {
            "기둥개수": column_count,
            "기둥체적": round(total_volume, 2),
            "콘크리트": round(concrete_volume, 2),
            "철근": round(rebar_weight / 1000, 3),  # ton
            "단위": {
                "기둥개수": "개",
                "기둥체적": "m³",
                "콘크리트": "m³",
                "철근": "ton"
            }
        }
    
    def _calculate_slabs(self, slab_entities: Dict) -> Dict:
        """슬래브 물량 계산"""
        total_area = 0
        slab_thickness = 0.2  # 기본 슬래브 두께 200mm
        
        # 닫힌 폴리라인으로 표현된 슬래브
        for poly in slab_entities.get("polylines", []):
            if poly["is_closed"]:
                total_area += poly["area"]
        
        # 해치로 표현된 슬래브
        for hatch in slab_entities.get("hatches", []):
            total_area += hatch["area"]
        
        # 슬래브 체적
        slab_volume = total_area * slab_thickness
        
        # 콘크리트 물량
        concrete_volume = slab_volume * MATERIAL_COEFFICIENTS["콘크리트"]["할증률"]
        
        # 철근 물량 (m²당 10kg 가정)
        rebar_weight = total_area * 10  # kg
        
        return {
            "슬래브면적": round(total_area, 2),
            "슬래브체적": round(slab_volume, 2),
            "콘크리트": round(concrete_volume, 2),
            "철근": round(rebar_weight / 1000, 3),  # ton
            "단위": {
                "슬래브면적": "m²",
                "슬래브체적": "m³",
                "콘크리트": "m³",
                "철근": "ton"
            }
        }
    
    def _count_blocks(self, block_entities: Dict) -> Dict:
        """블록 개수 카운트"""
        block_count = len(block_entities.get("blocks", []))
        
        # 블록 타입별 분류
        block_types = {}
        for block in block_entities.get("blocks", []):
            block_name = block["name"]
            if block_name not in block_types:
                block_types[block_name] = 0
            block_types[block_name] += 1
        
        result = {
            "총개수": block_count,
            "타입별": block_types,
            "단위": {"총개수": "개", "타입별": "개"}
        }
        
        return result
    
    def calculate_civil(self):
        """토목 분야 물량 계산"""
        civil_entities = self.entities.get("토목", {})
        
        # 도로 물량
        if "도로" in civil_entities:
            road_quantity = self._calculate_roads(civil_entities["도로"])
            self.quantities["토목"]["도로"] = road_quantity
        
        # 배관 물량
        if "배관" in civil_entities:
            pipe_quantity = self._calculate_pipes(civil_entities["배관"])
            self.quantities["토목"]["배관"] = pipe_quantity
        
        # 철근 물량
        if "철근" in civil_entities:
            rebar_quantity = self._calculate_rebars(civil_entities["철근"])
            self.quantities["토목"]["철근"] = rebar_quantity
        
        # 토공 물량
        if "토공" in civil_entities:
            earthwork_quantity = self._calculate_earthwork(civil_entities["토공"])
            self.quantities["토목"]["토공"] = earthwork_quantity
    
    def _calculate_roads(self, road_entities: Dict) -> Dict:
        """도로 물량 계산"""
        total_area = 0
        total_length = 0
        road_width = 6.0  # 기본 도로 폭 6m
        asphalt_thickness = 0.1  # 아스팔트 두께 100mm
        
        # 선형 도로
        for line in road_entities.get("lines", []):
            total_length += line["length"]
        
        for poly in road_entities.get("polylines", []):
            if poly["is_closed"]:
                total_area += poly["area"]
            else:
                total_length += poly["length"]
        
        # 해치로 표현된 도로
        for hatch in road_entities.get("hatches", []):
            total_area += hatch["area"]
        
        # 도로 면적 계산
        area_from_length = total_length * road_width
        total_road_area = area_from_length + total_area
        
        # 아스팔트 체적
        asphalt_volume = total_road_area * asphalt_thickness
        
        # 기층 체적 (아스팔트의 1.5배)
        base_volume = asphalt_volume * 1.5
        
        return {
            "도로연장": round(total_length, 2),
            "도로면적": round(total_road_area, 2),
            "아스팔트": round(asphalt_volume, 2),
            "기층": round(base_volume, 2),
            "단위": {
                "도로연장": "m",
                "도로면적": "m²",
                "아스팔트": "m³",
                "기층": "m³"
            }
        }
    
    def _calculate_pipes(self, pipe_entities: Dict) -> Dict:
        """배관 물량 계산"""
        pipe_lengths = {}
        total_length = 0
        
        # 선형 배관
        for line in pipe_entities.get("lines", []):
            total_length += line["length"]
        
        for poly in pipe_entities.get("polylines", []):
            if not poly["is_closed"]:
                total_length += poly["length"]
        
        # 원형 배관 (단면)
        pipe_diameters = []
        for circle in pipe_entities.get("circles", []):
            diameter = circle["radius"] * 2 * 1000  # mm
            pipe_diameters.append(diameter)
        
        # 평균 관경 계산
        avg_diameter = sum(pipe_diameters) / len(pipe_diameters) if pipe_diameters else 200
        
        # 관경별 분류 (임시)
        if avg_diameter <= 100:
            pipe_type = "D100"
        elif avg_diameter <= 200:
            pipe_type = "D200"
        elif avg_diameter <= 300:
            pipe_type = "D300"
        else:
            pipe_type = "D400+"
        
        return {
            "배관연장": round(total_length, 2),
            "평균관경": round(avg_diameter, 0),
            "관종": pipe_type,
            "단위": {
                "배관연장": "m",
                "평균관경": "mm"
            }
        }
    
    def _calculate_rebars(self, rebar_entities: Dict) -> Dict:
        """철근 물량 계산"""
        rebar_by_size = {}
        total_weight = 0
        
        # 선형 철근
        for line in rebar_entities.get("lines", []):
            # 레이어명이나 색상으로 철근 규격 추정 (실제로는 속성 필요)
            rebar_size = "D16"  # 기본값
            
            if rebar_size not in rebar_by_size:
                rebar_by_size[rebar_size] = 0
            
            # 철근 중량 계산
            unit_weight = MATERIAL_COEFFICIENTS["철근"].get(rebar_size, 1.56)
            weight = line["length"] * unit_weight
            rebar_by_size[rebar_size] += line["length"]
            total_weight += weight
        
        for poly in rebar_entities.get("polylines", []):
            rebar_size = "D16"  # 기본값
            
            if rebar_size not in rebar_by_size:
                rebar_by_size[rebar_size] = 0
            
            unit_weight = MATERIAL_COEFFICIENTS["철근"].get(rebar_size, 1.56)
            weight = poly["length"] * unit_weight
            rebar_by_size[rebar_size] += poly["length"]
            total_weight += weight
        
        # 할증률 적용
        total_weight *= MATERIAL_COEFFICIENTS["철근"]["할증률"]
        
        # 규격별 중량 계산
        rebar_weights = {}
        for size, length in rebar_by_size.items():
            unit_weight = MATERIAL_COEFFICIENTS["철근"].get(size, 1.56)
            rebar_weights[size] = round(length * unit_weight * 
                                       MATERIAL_COEFFICIENTS["철근"]["할증률"] / 1000, 3)
        
        return {
            "총중량": round(total_weight / 1000, 3),  # ton
            "규격별길이": {k: round(v, 2) for k, v in rebar_by_size.items()},
            "규격별중량": rebar_weights,
            "단위": {
                "총중량": "ton",
                "규격별길이": "m",
                "규격별중량": "ton"
            }
        }
    
    def _calculate_earthwork(self, earthwork_entities: Dict) -> Dict:
        """토공 물량 계산"""
        cut_volume = 0
        fill_volume = 0
        
        # 해치로 표현된 토공 구역
        for hatch in earthwork_entities.get("hatches", []):
            # 재료 타입으로 절토/성토 구분 (실제로는 레벨 정보 필요)
            if hatch["material"] == "토사":
                # 임시로 면적 × 깊이로 계산
                volume = hatch["area"] * 2.0  # 기본 깊이 2m
                cut_volume += volume
        
        # 폴리라인으로 표현된 토공 구역
        for poly in earthwork_entities.get("polylines", []):
            if poly["is_closed"]:
                # 임시로 면적 × 깊이로 계산
                volume = poly["area"] * 2.0
                fill_volume += volume
        
        # 운반 거리 (임시)
        haul_distance = 5.0  # km
        
        return {
            "절토량": round(cut_volume, 2),
            "성토량": round(fill_volume, 2),
            "순토량": round(cut_volume - fill_volume, 2),
            "운반거리": haul_distance,
            "단위": {
                "절토량": "m³",
                "성토량": "m³",
                "순토량": "m³",
                "운반거리": "km"
            }
        }
    
    def calculate_landscape(self):
        """조경 분야 물량 계산"""
        landscape_entities = self.entities.get("조경", {})
        
        # 수목 물량
        if "수목" in landscape_entities:
            tree_quantity = self._calculate_trees(landscape_entities["수목"])
            self.quantities["조경"]["수목"] = tree_quantity
        
        # 잔디 물량
        if "잔디" in landscape_entities:
            grass_quantity = self._calculate_grass(landscape_entities["잔디"])
            self.quantities["조경"]["잔디"] = grass_quantity
        
        # 포장 물량
        if "포장" in landscape_entities:
            paving_quantity = self._calculate_paving(landscape_entities["포장"])
            self.quantities["조경"]["포장"] = paving_quantity
        
        # 시설물 물량
        if "시설물" in landscape_entities:
            furniture_quantity = self._count_landscape_furniture(landscape_entities["시설물"])
            self.quantities["조경"]["시설물"] = furniture_quantity
    
    def _calculate_trees(self, tree_entities: Dict) -> Dict:
        """수목 물량 계산"""
        tree_count = {"교목": 0, "관목": 0, "총계": 0}
        tree_by_size = {"대": 0, "중": 0, "소": 0}
        
        # 블록으로 표현된 수목
        for block in tree_entities.get("blocks", []):
            tree_count["총계"] += 1
            
            # 블록 이름으로 교목/관목 구분
            if "교목" in block["name"] or "TREE" in block["name"].upper():
                tree_count["교목"] += 1
            else:
                tree_count["관목"] += 1
            
            # 크기 구분 (스케일 기준)
            scale = block["scale"][0]
            if scale >= 2.0:
                tree_by_size["대"] += 1
            elif scale >= 1.0:
                tree_by_size["중"] += 1
            else:
                tree_by_size["소"] += 1
        
        # 원으로 표현된 수목
        for circle in tree_entities.get("circles", []):
            tree_count["총계"] += 1
            
            # 반경으로 교목/관목 구분
            if circle["radius"] >= 1.0:
                tree_count["교목"] += 1
                tree_by_size["대"] += 1
            elif circle["radius"] >= 0.5:
                tree_count["교목"] += 1
                tree_by_size["중"] += 1
            else:
                tree_count["관목"] += 1
                tree_by_size["소"] += 1
        
        return {
            "수목수량": tree_count,
            "규격별": tree_by_size,
            "단위": {"수목수량": "주", "규격별": "주"}
        }
    
    def _calculate_grass(self, grass_entities: Dict) -> Dict:
        """잔디 물량 계산"""
        total_area = 0
        
        # 해치로 표현된 잔디
        for hatch in grass_entities.get("hatches", []):
            if hatch["pattern"] in ["GRASS", "AR-GRASS"]:
                total_area += hatch["area"]
        
        # 닫힌 폴리라인으로 표현된 잔디
        for poly in grass_entities.get("polylines", []):
            if poly["is_closed"]:
                total_area += poly["area"]
        
        # 롤잔디 물량
        turf_area = total_area * LANDSCAPE_COEFFICIENTS["잔디"]["롤잔디"]
        
        # 잔디씨 물량 (선택적)
        seed_weight = total_area * LANDSCAPE_COEFFICIENTS["잔디"]["씨앗"]
        
        return {
            "잔디면적": round(total_area, 2),
            "롤잔디": round(turf_area, 2),
            "잔디씨": round(seed_weight, 2),
            "단위": {
                "잔디면적": "m²",
                "롤잔디": "m²",
                "잔디씨": "kg"
            }
        }
    
    def _calculate_paving(self, paving_entities: Dict) -> Dict:
        """포장 물량 계산"""
        paving_by_type = {
            "보도블록": 0,
            "아스팔트": 0,
            "자갈": 0,
            "기타": 0
        }
        
        # 해치 패턴으로 포장 타입 구분
        for hatch in paving_entities.get("hatches", []):
            area = hatch["area"]
            
            if "BRICK" in hatch["pattern"]:
                paving_by_type["보도블록"] += area
            elif "ASPHALT" in hatch["pattern"] or "AR-CONC" in hatch["pattern"]:
                paving_by_type["아스팔트"] += area
            elif "GRAVEL" in hatch["pattern"]:
                paving_by_type["자갈"] += area
            else:
                paving_by_type["기타"] += area
        
        # 닫힌 폴리라인
        for poly in paving_entities.get("polylines", []):
            if poly["is_closed"]:
                paving_by_type["보도블록"] += poly["area"]
        
        # 체적 계산
        volumes = {}
        if paving_by_type["아스팔트"] > 0:
            volumes["아스팔트"] = paving_by_type["아스팔트"] * \
                                 LANDSCAPE_COEFFICIENTS["포장재"]["아스팔트"]
        if paving_by_type["자갈"] > 0:
            volumes["자갈"] = paving_by_type["자갈"] * \
                             LANDSCAPE_COEFFICIENTS["포장재"]["자갈"]
        
        return {
            "포장면적": {k: round(v, 2) for k, v in paving_by_type.items() if v > 0},
            "포장체적": {k: round(v, 2) for k, v in volumes.items()},
            "단위": {
                "포장면적": "m²",
                "포장체적": "m³"
            }
        }
    
    def _count_landscape_furniture(self, furniture_entities: Dict) -> Dict:
        """조경 시설물 개수 카운트"""
        furniture_count = {}
        
        # 블록으로 표현된 시설물
        for block in furniture_entities.get("blocks", []):
            furniture_type = self._identify_furniture_type(block["name"])
            
            if furniture_type not in furniture_count:
                furniture_count[furniture_type] = 0
            furniture_count[furniture_type] += 1
        
        return {
            "시설물수량": furniture_count,
            "총개수": sum(furniture_count.values()),
            "단위": {"시설물수량": "개", "총개수": "개"}
        }
    
    def _identify_furniture_type(self, block_name: str) -> str:
        """블록 이름으로 시설물 타입 식별"""
        name_upper = block_name.upper()
        
        if "BENCH" in name_upper or "벤치" in block_name:
            return "벤치"
        elif "LAMP" in name_upper or "가로등" in block_name:
            return "가로등"
        elif "SIGN" in name_upper or "표지" in block_name:
            return "표지판"
        elif "TRASH" in name_upper or "쓰레기" in block_name:
            return "쓰레기통"
        elif "PERGOLA" in name_upper or "파고라" in block_name:
            return "파고라"
        else:
            return "기타시설물"