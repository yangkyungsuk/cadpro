"""
DWG/DXF 파일 파서
건축, 토목, 조경 도면의 엔티티를 추출하고 분류
"""

import ezdxf
from ezdxf.addons import odafc
import numpy as np
from typing import Dict, List, Tuple, Any
from shapely.geometry import Polygon, LineString, Point
import math
from config import LAYER_MAPPING, BLOCK_PATTERNS, HATCH_MATERIAL_MAPPING


class DWGParser:
    def __init__(self, filepath: str):
        """
        DWG/DXF 파일 파서 초기화
        
        Args:
            filepath: CAD 파일 경로
        """
        self.filepath = filepath
        self.doc = None
        self.entities = {
            "lines": [],
            "polylines": [],
            "circles": [],
            "arcs": [],
            "blocks": [],
            "hatches": [],
            "texts": [],
            "dimensions": []
        }
        self.layers_info = {}
        
    def load_file(self):
        """CAD 파일 로드"""
        try:
            if self.filepath.lower().endswith('.dwg'):
                # DWG 파일인 경우 ODA 변환기 사용 (설치 필요)
                print(f"DWG 파일 로드 중: {self.filepath}")
                # odafc 사용 시 ODA File Converter 설치 필요
                # self.doc = odafc.readfile(self.filepath)
                # 임시로 DXF로 변환된 파일을 읽는다고 가정
                temp_dxf = self.filepath.replace('.dwg', '.dxf')
                self.doc = ezdxf.readfile(temp_dxf)
            else:
                # DXF 파일 직접 읽기
                self.doc = ezdxf.readfile(self.filepath)
                
            print(f"CAD 파일 로드 완료: {self.doc.dxfversion}")
            self._extract_layers()
            return True
            
        except Exception as e:
            print(f"파일 로드 실패: {e}")
            return False
    
    def _extract_layers(self):
        """레이어 정보 추출"""
        for layer in self.doc.layers:
            self.layers_info[layer.dxf.name] = {
                "color": layer.dxf.color,
                "linetype": layer.dxf.linetype,
                "on": not layer.is_off(),
                "frozen": layer.is_frozen(),
                "locked": layer.is_locked()
            }
        print(f"총 {len(self.layers_info)}개 레이어 발견")
    
    def parse_entities(self):
        """모든 엔티티 파싱"""
        msp = self.doc.modelspace()
        
        for entity in msp:
            layer = entity.dxf.layer
            
            # 레이어가 꺼져있거나 동결된 경우 제외
            if layer in self.layers_info:
                if not self.layers_info[layer]["on"] or self.layers_info[layer]["frozen"]:
                    continue
            
            # 엔티티 타입별 분류
            if entity.dxftype() == 'LINE':
                self._parse_line(entity)
            elif entity.dxftype() == 'LWPOLYLINE':
                self._parse_polyline(entity)
            elif entity.dxftype() == 'POLYLINE':
                self._parse_polyline(entity)
            elif entity.dxftype() == 'CIRCLE':
                self._parse_circle(entity)
            elif entity.dxftype() == 'ARC':
                self._parse_arc(entity)
            elif entity.dxftype() == 'INSERT':
                self._parse_block(entity)
            elif entity.dxftype() == 'HATCH':
                self._parse_hatch(entity)
            elif entity.dxftype() in ['TEXT', 'MTEXT']:
                self._parse_text(entity)
            elif entity.dxftype() == 'DIMENSION':
                self._parse_dimension(entity)
    
    def _parse_line(self, entity):
        """선 엔티티 파싱"""
        start = entity.dxf.start
        end = entity.dxf.end
        length = math.sqrt((end[0]-start[0])**2 + (end[1]-start[1])**2 + (end[2]-start[2])**2)
        
        self.entities["lines"].append({
            "layer": entity.dxf.layer,
            "start": (start[0], start[1], start[2]),
            "end": (end[0], end[1], end[2]),
            "length": length,
            "color": entity.dxf.color
        })
    
    def _parse_polyline(self, entity):
        """폴리라인 엔티티 파싱"""
        points = []
        
        # LWPOLYLINE과 POLYLINE 처리
        if entity.dxftype() == 'LWPOLYLINE':
            points = [(p[0], p[1], 0) for p in entity.get_points()]
        else:
            points = [(p.dxf.location[0], p.dxf.location[1], p.dxf.location[2]) 
                     for p in entity.vertices]
        
        if len(points) < 2:
            return
            
        # 길이 계산
        total_length = 0
        for i in range(len(points) - 1):
            segment_length = math.sqrt(
                (points[i+1][0] - points[i][0])**2 + 
                (points[i+1][1] - points[i][1])**2 + 
                (points[i+1][2] - points[i][2])**2
            )
            total_length += segment_length
        
        # 닫힌 폴리라인인 경우 면적 계산
        area = 0
        if entity.is_closed:
            # 2D 폴리곤으로 변환하여 면적 계산
            poly_2d = [(p[0], p[1]) for p in points]
            if len(poly_2d) >= 3:
                polygon = Polygon(poly_2d)
                area = polygon.area
        
        self.entities["polylines"].append({
            "layer": entity.dxf.layer,
            "points": points,
            "length": total_length,
            "area": area,
            "is_closed": entity.is_closed,
            "color": entity.dxf.color
        })
    
    def _parse_circle(self, entity):
        """원 엔티티 파싱"""
        center = entity.dxf.center
        radius = entity.dxf.radius
        
        self.entities["circles"].append({
            "layer": entity.dxf.layer,
            "center": (center[0], center[1], center[2]),
            "radius": radius,
            "circumference": 2 * math.pi * radius,
            "area": math.pi * radius ** 2,
            "color": entity.dxf.color
        })
    
    def _parse_arc(self, entity):
        """호 엔티티 파싱"""
        center = entity.dxf.center
        radius = entity.dxf.radius
        start_angle = math.radians(entity.dxf.start_angle)
        end_angle = math.radians(entity.dxf.end_angle)
        
        # 호의 길이 계산
        if end_angle < start_angle:
            angle_diff = (2 * math.pi - start_angle) + end_angle
        else:
            angle_diff = end_angle - start_angle
        
        arc_length = radius * angle_diff
        
        self.entities["arcs"].append({
            "layer": entity.dxf.layer,
            "center": (center[0], center[1], center[2]),
            "radius": radius,
            "start_angle": entity.dxf.start_angle,
            "end_angle": entity.dxf.end_angle,
            "length": arc_length,
            "color": entity.dxf.color
        })
    
    def _parse_block(self, entity):
        """블록 참조 파싱"""
        insert_point = entity.dxf.insert
        block_name = entity.dxf.name
        
        # 블록 속성 추출
        attributes = {}
        if entity.has_attrib:
            for attrib in entity.attribs:
                attributes[attrib.dxf.tag] = attrib.dxf.text
        
        self.entities["blocks"].append({
            "layer": entity.dxf.layer,
            "name": block_name,
            "insert_point": (insert_point[0], insert_point[1], insert_point[2]),
            "scale": (entity.dxf.xscale, entity.dxf.yscale, entity.dxf.zscale),
            "rotation": entity.dxf.rotation,
            "attributes": attributes,
            "color": entity.dxf.color
        })
    
    def _parse_hatch(self, entity):
        """해치 엔티티 파싱"""
        pattern_name = entity.dxf.pattern_name
        
        # 해치 경계 추출 및 면적 계산
        area = 0
        boundaries = []
        
        try:
            paths = entity.paths
            for path in paths:
                boundary_points = []
                for edge in path.edges:
                    if hasattr(edge, 'start') and hasattr(edge, 'end'):
                        boundary_points.append((edge.start[0], edge.start[1]))
                        boundary_points.append((edge.end[0], edge.end[1]))
                
                if len(boundary_points) >= 3:
                    polygon = Polygon(boundary_points)
                    area += polygon.area
                    boundaries.append(boundary_points)
        except:
            pass
        
        # 재료 타입 매핑
        material = HATCH_MATERIAL_MAPPING.get(pattern_name, "기타")
        
        self.entities["hatches"].append({
            "layer": entity.dxf.layer,
            "pattern": pattern_name,
            "material": material,
            "area": area,
            "boundaries": boundaries,
            "color": entity.dxf.color
        })
    
    def _parse_text(self, entity):
        """텍스트 엔티티 파싱"""
        if entity.dxftype() == 'TEXT':
            text = entity.dxf.text
            insert = entity.dxf.insert
        else:  # MTEXT
            text = entity.text
            insert = entity.dxf.insert
        
        self.entities["texts"].append({
            "layer": entity.dxf.layer,
            "text": text,
            "position": (insert[0], insert[1], insert[2]),
            "height": entity.dxf.height if hasattr(entity.dxf, 'height') else 0,
            "color": entity.dxf.color
        })
    
    def _parse_dimension(self, entity):
        """치수 엔티티 파싱"""
        self.entities["dimensions"].append({
            "layer": entity.dxf.layer,
            "type": entity.dxftype(),
            "measurement": entity.get_measurement() if hasattr(entity, 'get_measurement') else 0,
            "color": entity.dxf.color
        })
    
    def classify_by_field(self) -> Dict[str, Dict]:
        """
        파싱된 엔티티를 건축/토목/조경 분야별로 분류
        
        Returns:
            분야별로 분류된 엔티티 딕셔너리
        """
        classified = {
            "건축": {},
            "토목": {},
            "조경": {},
            "기타": {}
        }
        
        # 각 분야별 레이어 매핑 확인
        for field, layer_map in LAYER_MAPPING.items():
            for category, layer_patterns in layer_map.items():
                classified[field][category] = {
                    "lines": [],
                    "polylines": [],
                    "circles": [],
                    "blocks": [],
                    "hatches": []
                }
                
                # 각 엔티티 타입별로 레이어 패턴 매칭
                for entity_type in ["lines", "polylines", "circles", "blocks", "hatches"]:
                    for entity in self.entities[entity_type]:
                        entity_layer = entity["layer"].upper()
                        
                        for pattern in layer_patterns:
                            pattern_upper = pattern.upper()
                            # 와일드카드 처리
                            if "*" in pattern_upper:
                                pattern_check = pattern_upper.replace("*", "")
                                if pattern_check in entity_layer:
                                    classified[field][category][entity_type].append(entity)
                                    break
                            elif pattern_upper == entity_layer:
                                classified[field][category][entity_type].append(entity)
                                break
        
        # 분류되지 않은 엔티티는 기타로 분류
        for entity_type, entities in self.entities.items():
            for entity in entities:
                found = False
                for field in ["건축", "토목", "조경"]:
                    for category in classified[field]:
                        if entity in classified[field][category].get(entity_type, []):
                            found = True
                            break
                    if found:
                        break
                
                if not found:
                    if entity_type not in classified["기타"]:
                        classified["기타"][entity_type] = []
                    classified["기타"][entity_type].append(entity)
        
        return classified
    
    def get_summary(self) -> Dict:
        """파싱 결과 요약"""
        summary = {
            "파일명": self.filepath,
            "CAD버전": self.doc.dxfversion if self.doc else "Unknown",
            "레이어수": len(self.layers_info),
            "엔티티수": {
                "선": len(self.entities["lines"]),
                "폴리라인": len(self.entities["polylines"]),
                "원": len(self.entities["circles"]),
                "호": len(self.entities["arcs"]),
                "블록": len(self.entities["blocks"]),
                "해치": len(self.entities["hatches"]),
                "텍스트": len(self.entities["texts"]),
                "치수": len(self.entities["dimensions"])
            },
            "총엔티티": sum(len(entities) for entities in self.entities.values())
        }
        return summary