"""
CADPro Selector - AutoCAD 객체 선택 전문 모듈
SendCommand를 사용한 안정적인 선택 방식
"""

import win32com.client
import pythoncom
import time
import math


class AutoCADSelector:
    """AutoCAD 객체 선택 전문 클래스"""
    
    def __init__(self):
        self.acad = None
        self.doc = None
        self.model = None
        
    def connect(self):
        """AutoCAD 연결"""
        try:
            self.acad = win32com.client.Dispatch("AutoCAD.Application")
            if self.acad.Documents.Count == 0:
                return False, "열린 도면이 없습니다."
            
            self.doc = self.acad.ActiveDocument
            self.model = self.doc.ModelSpace
            
            # AutoCAD 활성화
            self.acad.Visible = True
            
            return True, f"도면: {self.doc.Name}"
            
        except Exception as e:
            return False, str(e)
    
    def get_pickfirst_selection(self):
        """현재 선택된 객체 가져오기 (PickfirstSelectionSet)"""
        try:
            sel_set = self.doc.PickfirstSelectionSet
            objects = []
            
            for obj in sel_set:
                objects.append(self._extract_object_info(obj))
            
            return objects, f"{len(objects)}개 객체"
            
        except Exception as e:
            return [], f"오류: {str(e)}"
    
    def select_all(self):
        """모든 객체 선택 (SendCommand 사용)"""
        try:
            # 모든 객체 선택
            self.doc.SendCommand("_SELECT\n_ALL\n\n")
            time.sleep(0.5)  # 명령 처리 대기
            
            # 선택된 객체 가져오기
            return self.get_pickfirst_selection()
            
        except Exception as e:
            return [], f"오류: {str(e)}"
    
    def select_by_window(self):
        """윈도우 선택 (SendCommand 사용)"""
        try:
            # 윈도우 선택 명령
            self.doc.SendCommand("_SELECT\n_W\n")
            
            # 사용자가 선택할 시간 제공
            time.sleep(5)  # 5초 대기
            
            # 선택된 객체 가져오기
            return self.get_pickfirst_selection()
            
        except Exception as e:
            return [], f"오류: {str(e)}"
    
    def select_by_crossing(self):
        """교차 선택 (SendCommand 사용)"""
        try:
            # 교차 선택 명령
            self.doc.SendCommand("_SELECT\n_C\n")
            
            # 사용자가 선택할 시간 제공
            time.sleep(5)  # 5초 대기
            
            # 선택된 객체 가져오기
            return self.get_pickfirst_selection()
            
        except Exception as e:
            return [], f"오류: {str(e)}"
    
    def select_by_fence(self):
        """울타리 선택 (SendCommand 사용)"""
        try:
            # 울타리 선택 명령
            self.doc.SendCommand("_SELECT\n_F\n")
            
            # 사용자가 선택할 시간 제공
            time.sleep(5)  # 5초 대기
            
            # 선택된 객체 가져오기
            return self.get_pickfirst_selection()
            
        except Exception as e:
            return [], f"오류: {str(e)}"
    
    def select_last(self):
        """마지막 객체 선택"""
        try:
            self.doc.SendCommand("_SELECT\n_L\n\n")
            time.sleep(0.5)
            
            return self.get_pickfirst_selection()
            
        except Exception as e:
            return [], f"오류: {str(e)}"
    
    def select_previous(self):
        """이전 선택 세트"""
        try:
            self.doc.SendCommand("_SELECT\n_P\n\n")
            time.sleep(0.5)
            
            return self.get_pickfirst_selection()
            
        except Exception as e:
            return [], f"오류: {str(e)}"
    
    def select_by_layer(self, layer_name):
        """레이어로 선택 (직접 순회)"""
        try:
            objects = []
            
            for obj in self.model:
                if obj.Layer == layer_name:
                    objects.append(self._extract_object_info(obj))
            
            return objects, f"{layer_name} 레이어: {len(objects)}개"
            
        except Exception as e:
            return [], f"오류: {str(e)}"
    
    def select_by_type(self, object_type):
        """타입으로 선택 (직접 순회)"""
        try:
            objects = []
            
            for obj in self.model:
                if object_type in obj.ObjectName:
                    objects.append(self._extract_object_info(obj))
            
            return objects, f"{object_type}: {len(objects)}개"
            
        except Exception as e:
            return [], f"오류: {str(e)}"
    
    def select_by_color(self, color_index):
        """색상으로 선택"""
        try:
            objects = []
            
            for obj in self.model:
                try:
                    if obj.Color == color_index:
                        objects.append(self._extract_object_info(obj))
                except:
                    pass
            
            return objects, f"색상 {color_index}: {len(objects)}개"
            
        except Exception as e:
            return [], f"오류: {str(e)}"
    
    def clear_selection(self):
        """선택 해제"""
        try:
            self.doc.SendCommand("_SELECT\n\n")
            time.sleep(0.2)
            return True, "선택 해제됨"
            
        except Exception as e:
            return False, str(e)
    
    def highlight_objects(self, objects, highlight=True):
        """객체 하이라이트"""
        count = 0
        for obj_data in objects:
            try:
                if 'com_object' in obj_data:
                    obj_data['com_object'].Highlight(highlight)
                    count += 1
            except:
                pass
        
        return count
    
    def _extract_object_info(self, obj):
        """객체 정보 추출"""
        info = {
            'type': obj.ObjectName,
            'layer': obj.Layer,
            'handle': obj.Handle,
            'com_object': obj  # COM 객체 참조 저장
        }
        
        try:
            info['color'] = obj.Color
        except:
            pass
        
        # 선
        if "Line" in obj.ObjectName and "Polyline" not in obj.ObjectName:
            try:
                start = obj.StartPoint
                end = obj.EndPoint
                length = math.sqrt(
                    (end[0]-start[0])**2 + 
                    (end[1]-start[1])**2 + 
                    (end[2]-start[2])**2
                )
                info['length'] = length
                info['start'] = start
                info['end'] = end
            except:
                pass
        
        # 폴리라인
        elif "Polyline" in obj.ObjectName:
            try:
                info['length'] = obj.Length
                info['closed'] = obj.Closed
                if obj.Closed:
                    info['area'] = obj.Area
            except:
                pass
        
        # 원
        elif "Circle" in obj.ObjectName:
            try:
                info['radius'] = obj.Radius
                info['center'] = obj.Center
                info['area'] = math.pi * obj.Radius ** 2
            except:
                pass
        
        # 호
        elif "Arc" in obj.ObjectName:
            try:
                info['radius'] = obj.Radius
                info['arc_length'] = obj.ArcLength
            except:
                pass
        
        # 블록
        elif "BlockReference" in obj.ObjectName:
            try:
                info['block_name'] = obj.Name
                info['position'] = obj.InsertionPoint
            except:
                pass
        
        # 텍스트
        elif "Text" in obj.ObjectName or "MText" in obj.ObjectName:
            try:
                info['text'] = obj.TextString
            except:
                pass
        
        # 해치
        elif "Hatch" in obj.ObjectName:
            try:
                info['pattern'] = obj.PatternName
                info['area'] = obj.Area
            except:
                pass
        
        return info
    
    def get_layers(self):
        """레이어 목록 가져오기"""
        layers = []
        try:
            for layer in self.doc.Layers:
                layers.append({
                    'name': layer.Name,
                    'on': layer.LayerOn,
                    'frozen': layer.Freeze,
                    'locked': layer.Lock,
                    'color': layer.Color
                })
        except:
            pass
        
        return layers
    
    def get_object_types(self):
        """객체 타입 목록 가져오기"""
        types = {}
        try:
            for obj in self.model:
                obj_type = obj.ObjectName
                if obj_type not in types:
                    types[obj_type] = 0
                types[obj_type] += 1
        except:
            pass
        
        return types
    
    def load_all_objects(self):
        """모든 객체 로드"""
        objects = []
        try:
            for obj in self.model:
                objects.append(self._extract_object_info(obj))
            
            return objects, f"총 {len(objects)}개 객체"
            
        except Exception as e:
            return [], f"오류: {str(e)}"


def test_selector():
    """선택 기능 테스트"""
    print("=== AutoCAD Selector 테스트 ===\n")
    
    selector = AutoCADSelector()
    
    # 연결
    success, msg = selector.connect()
    print(f"연결: {msg}")
    
    if not success:
        return
    
    # 현재 선택 가져오기
    objects, msg = selector.get_pickfirst_selection()
    print(f"현재 선택: {msg}")
    
    # 레이어 목록
    layers = selector.get_layers()
    print(f"레이어 수: {len(layers)}")
    
    # 객체 타입
    types = selector.get_object_types()
    print(f"객체 타입: {len(types)}종")
    
    for obj_type, count in sorted(types.items(), key=lambda x: x[1], reverse=True)[:5]:
        print(f"  {obj_type}: {count}개")
    
    # 모든 객체 로드
    objects, msg = selector.load_all_objects()
    print(f"\n전체 객체: {msg}")
    
    # 타입별 집계
    type_summary = {}
    for obj in objects:
        t = obj['type']
        if t not in type_summary:
            type_summary[t] = 0
        type_summary[t] += 1
    
    print("\n타입별 요약:")
    for t, count in sorted(type_summary.items()):
        print(f"  {t}: {count}개")


if __name__ == "__main__":
    test_selector()