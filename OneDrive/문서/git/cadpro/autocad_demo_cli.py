"""
AutoCAD 물량 산출 - 자동 데모
사용자 입력 없이 자동으로 실행
"""

import win32com.client
import pythoncom
import math
import json
from datetime import datetime

class AutoCADDemo:
    def __init__(self):
        """AutoCAD 연결"""
        try:
            self.acad = win32com.client.Dispatch("AutoCAD.Application")
            self.doc = self.acad.ActiveDocument
            self.model = self.doc.ModelSpace
            print(f"[OK] AutoCAD 연결 성공")
            print(f"[OK] 현재 도면: {self.doc.Name}\n")
        except Exception as e:
            print(f"[ERROR] AutoCAD 연결 실패: {e}")
            exit(1)
    
    def analyze_drawing(self):
        """도면 전체 분석"""
        print("="*60)
        print("1. 도면 분석")
        print("="*60)
        
        # 객체 타입별 분류
        object_types = {}
        for obj in self.model:
            obj_type = obj.ObjectName
            if obj_type not in object_types:
                object_types[obj_type] = []
            object_types[obj_type].append(obj)
        
        print("\n[객체 타입별 분류]")
        for obj_type, objects in sorted(object_types.items()):
            print(f"  {obj_type}: {len(objects)}개")
        
        return object_types
    
    def calculate_deck_quantities(self, object_types):
        """데크 물량 자동 계산"""
        print("\n"+"="*60)
        print("2. 데크 물량 자동 계산")
        print("="*60)
        
        results = {}
        
        # 1. 선 길이 계산 (장선/데크보드)
        print("\n[선형 요소 계산]")
        total_length = 0
        count = 0
        
        for obj in object_types.get("AcDbLine", []):
            try:
                start = obj.StartPoint
                end = obj.EndPoint
                length = math.sqrt((end[0]-start[0])**2 + (end[1]-start[1])**2 + (end[2]-start[2])**2)
                total_length += length
                count += 1
            except:
                pass
        
        if total_length > 0:
            length_m = total_length / 1000  # mm to m
            print(f"  선 개수: {count}개")
            print(f"  총 길이: {length_m:.2f}m")
            
            # 장선으로 가정 (2x10 방부목)
            volume = length_m * 0.038 * 0.235  # 2x10 inch = 38x235mm
            print(f"  목재 체적: {volume:.3f}m³ (2x10 기준)")
            
            results["장선"] = {
                "개수": count,
                "길이": f"{length_m:.2f}m",
                "체적": f"{volume:.3f}m³"
            }
        
        # 2. 폴리라인 계산
        print("\n[폴리라인 요소 계산]")
        poly_length = 0
        poly_area = 0
        poly_count = 0
        closed_count = 0
        
        for obj in object_types.get("AcDbPolyline", []):
            try:
                poly_count += 1
                poly_length += obj.Length
                
                if obj.Closed:
                    poly_area += obj.Area
                    closed_count += 1
            except:
                pass
        
        if poly_count > 0:
            length_m = poly_length / 1000
            area_m2 = poly_area / 1000000
            
            print(f"  폴리라인 개수: {poly_count}개")
            print(f"  - 닫힌 폴리라인: {closed_count}개")
            print(f"  총 길이: {length_m:.2f}m")
            print(f"  총 면적: {area_m2:.2f}m²")
            
            if area_m2 > 0:
                # 데크 면적으로 계산
                deck_area = area_m2 * 1.07  # 7% 손실
                screws = int(area_m2 * 8)  # m²당 8개
                
                print(f"\n  [데크보드 계산]")
                print(f"  - 데크 면적: {area_m2:.2f}m²")
                print(f"  - 필요 보드: {deck_area:.2f}m² (7% 손실 포함)")
                print(f"  - 데크 스크류: {screws}개")
                
                results["데크보드"] = {
                    "면적": f"{area_m2:.2f}m²",
                    "필요량": f"{deck_area:.2f}m²",
                    "스크류": f"{screws}개"
                }
        
        # 3. 블록 계산 (기둥, 브라켓 등)
        print("\n[블록 요소 계산]")
        block_types = {}
        
        for obj in object_types.get("AcDbBlockReference", []):
            try:
                block_name = obj.Name
                if block_name not in block_types:
                    block_types[block_name] = 0
                block_types[block_name] += 1
            except:
                pass
        
        if block_types:
            print("  블록 타입별 개수:")
            for name, count in sorted(block_types.items()):
                print(f"    {name}: {count}개")
            
            # 기둥으로 추정되는 블록 찾기
            post_count = 0
            for name, count in block_types.items():
                if any(kw in name.upper() for kw in ["POST", "기둥", "COLUMN"]):
                    post_count += count
            
            if post_count > 0:
                foundation = post_count * 0.096  # 400x400x600 기초
                print(f"\n  [기둥 계산]")
                print(f"  - 기둥 개수: {post_count}개")
                print(f"  - 기초 콘크리트: {foundation:.3f}m³")
                
                results["기둥"] = {
                    "개수": f"{post_count}개",
                    "기초": f"{foundation:.3f}m³"
                }
        
        # 4. 원형 객체 (기둥 단면 등)
        circles = object_types.get("AcDbCircle", [])
        if circles:
            print(f"\n[원형 객체]")
            print(f"  개수: {len(circles)}개")
            
            # 반경별 분류
            radius_groups = {}
            for obj in circles:
                try:
                    r = round(obj.Radius, 0)
                    if r not in radius_groups:
                        radius_groups[r] = 0
                    radius_groups[r] += 1
                except:
                    pass
            
            if radius_groups:
                print("  반경별 분포:")
                for r, count in sorted(radius_groups.items()):
                    print(f"    R{r}: {count}개")
        
        return results
    
    def calculate_by_layer(self):
        """레이어별 물량 계산"""
        print("\n"+"="*60)
        print("3. 레이어별 분석")
        print("="*60)
        
        # 레이어별 객체 수집
        layer_objects = {}
        
        for obj in self.model:
            layer = obj.Layer
            if layer not in layer_objects:
                layer_objects[layer] = {
                    'count': 0,
                    'length': 0,
                    'area': 0
                }
            
            layer_objects[layer]['count'] += 1
            
            # 길이 계산
            if "Line" in obj.ObjectName:
                try:
                    start = obj.StartPoint
                    end = obj.EndPoint
                    length = math.sqrt((end[0]-start[0])**2 + (end[1]-start[1])**2)
                    layer_objects[layer]['length'] += length
                except:
                    pass
            elif "Polyline" in obj.ObjectName:
                try:
                    layer_objects[layer]['length'] += obj.Length
                    if obj.Closed:
                        layer_objects[layer]['area'] += obj.Area
                except:
                    pass
        
        # 상위 5개 레이어 표시
        print("\n[주요 레이어 (객체 수 기준)]")
        sorted_layers = sorted(layer_objects.items(), key=lambda x: x[1]['count'], reverse=True)
        
        for layer, data in sorted_layers[:5]:
            print(f"\n  레이어: {layer}")
            print(f"    객체 수: {data['count']}개")
            if data['length'] > 0:
                print(f"    총 길이: {data['length']/1000:.2f}m")
            if data['area'] > 0:
                print(f"    총 면적: {data['area']/1000000:.2f}m²")
        
        return layer_objects
    
    def save_results(self, results):
        """결과 저장"""
        print("\n"+"="*60)
        print("4. 결과 저장")
        print("="*60)
        
        # JSON 형식으로 저장
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"autocad_result_{timestamp}.json"
        
        output = {
            "도면": self.doc.Name,
            "일시": datetime.now().isoformat(),
            "객체수": self.model.Count,
            "계산결과": results
        }
        
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(output, f, ensure_ascii=False, indent=2)
        
        print(f"\n[OK] 결과 저장: {filename}")
        
        # 결과 요약 출력
        print("\n[계산 결과 요약]")
        for category, data in results.items():
            print(f"\n{category}:")
            for key, value in data.items():
                print(f"  - {key}: {value}")
        
        return filename

def main():
    print("="*60)
    print("AutoCAD 물량 산출 자동 데모")
    print("="*60)
    
    # AutoCAD 연결
    demo = AutoCADDemo()
    
    # 도면 분석
    object_types = demo.analyze_drawing()
    
    # 데크 물량 계산
    results = demo.calculate_deck_quantities(object_types)
    
    # 레이어별 분석
    layer_info = demo.calculate_by_layer()
    
    # 결과 저장
    if results:
        demo.save_results(results)
    
    print("\n"+"="*60)
    print("데모 완료!")
    print("="*60)
    print("\n실제 사용 시에는:")
    print("  1. 특정 영역만 선택 가능")
    print("  2. 사용자가 계산 항목 선택")
    print("  3. AutoCAD 테이블로 직접 삽입")
    print("  4. Excel 리포트 생성")
    print("\n전체 기능: python autocad_plugin.py")

if __name__ == "__main__":
    main()