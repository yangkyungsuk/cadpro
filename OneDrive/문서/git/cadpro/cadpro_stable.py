"""
CADPro - 안정적인 버전
AutoCAD의 선택 기능에 의존하지 않고 작동
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import win32com.client
import pythoncom
import math
import json
from datetime import datetime
import os


class CADProStable:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("CADPro - 안정적인 물량 산출")
        self.root.geometry("1000x600")
        
        # AutoCAD 연결
        self.acad = None
        self.doc = None
        self.model = None
        
        # 데이터 저장
        self.all_objects = []  # 모든 객체
        self.filtered_objects = []  # 필터링된 객체
        self.calculation_results = {}
        
        # GUI 구성
        self.setup_gui()
    
    def setup_gui(self):
        """GUI 구성"""
        
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 1단계: AutoCAD 연결
        self.create_step1(main_frame)
        
        # 2단계: 객체 필터링
        self.create_step2(main_frame)
        
        # 3단계: 계산
        self.create_step3(main_frame)
        
        # 결과 영역
        self.create_result_area(main_frame)
    
    def create_step1(self, parent):
        """1단계: AutoCAD 연결"""
        frame = ttk.LabelFrame(parent, text="1단계: AutoCAD 연결", padding="10")
        frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # 연결 버튼
        self.connect_btn = ttk.Button(frame, text="AutoCAD 연결", command=self.connect_autocad)
        self.connect_btn.grid(row=0, column=0, padx=5)
        
        # 상태 표시
        self.status_label = ttk.Label(frame, text="연결 대기 중...", foreground="gray")
        self.status_label.grid(row=0, column=1, padx=20)
        
        # 도면 정보
        self.drawing_info = ttk.Label(frame, text="도면: -")
        self.drawing_info.grid(row=0, column=2, padx=20)
        
        # 객체 수
        self.object_count = ttk.Label(frame, text="객체: 0개")
        self.object_count.grid(row=0, column=3, padx=20)
        
        # 새로고침 버튼
        ttk.Button(frame, text="새로고침", command=self.refresh_data).grid(row=0, column=4, padx=5)
    
    def create_step2(self, parent):
        """2단계: 객체 필터링"""
        frame = ttk.LabelFrame(parent, text="2단계: 객체 필터링 (선택적)", padding="10")
        frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # 레이어 필터
        layer_frame = ttk.Frame(frame)
        layer_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(layer_frame, text="레이어:").pack(side=tk.LEFT, padx=5)
        self.layer_combo = ttk.Combobox(layer_frame, width=30)
        self.layer_combo.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(layer_frame, text="레이어 필터 적용", 
                  command=self.filter_by_layer).pack(side=tk.LEFT, padx=5)
        
        # 타입 필터
        type_frame = ttk.Frame(frame)
        type_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(type_frame, text="객체 타입:").pack(side=tk.LEFT, padx=5)
        self.type_combo = ttk.Combobox(type_frame, width=30)
        self.type_combo.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(type_frame, text="타입 필터 적용", 
                  command=self.filter_by_type).pack(side=tk.LEFT, padx=5)
        
        # 필터 정보
        filter_info_frame = ttk.Frame(frame)
        filter_info_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=5)
        
        self.filter_info = ttk.Label(filter_info_frame, text="필터: 없음 (전체 객체)")
        self.filter_info.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(filter_info_frame, text="필터 초기화", 
                  command=self.clear_filter).pack(side=tk.LEFT, padx=20)
        
        self.filtered_count = ttk.Label(filter_info_frame, text="대상: 0개")
        self.filtered_count.pack(side=tk.LEFT, padx=20)
    
    def create_step3(self, parent):
        """3단계: 계산"""
        frame = ttk.LabelFrame(parent, text="3단계: 물량 계산", padding="10")
        frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        # 계산 버튼들
        calc_frame1 = ttk.Frame(frame)
        calc_frame1.grid(row=0, column=0, pady=5)
        
        buttons1 = [
            ("총 길이 계산", self.calc_total_length),
            ("총 면적 계산", self.calc_total_area),
            ("객체 개수", self.calc_count),
            ("장선/데크보드", self.calc_deck_joists)
        ]
        
        for i, (text, command) in enumerate(buttons1):
            ttk.Button(calc_frame1, text=text, command=command, width=15).grid(row=0, column=i, padx=3)
        
        calc_frame2 = ttk.Frame(frame)
        calc_frame2.grid(row=1, column=0, pady=5)
        
        buttons2 = [
            ("기둥 계산", self.calc_posts),
            ("벽체 계산", self.calc_walls),
            ("레이어별 분석", self.analyze_by_layer),
            ("타입별 분석", self.analyze_by_type)
        ]
        
        for i, (text, command) in enumerate(buttons2):
            ttk.Button(calc_frame2, text=text, command=command, width=15).grid(row=0, column=i, padx=3)
        
        # 옵션
        option_frame = ttk.Frame(frame)
        option_frame.grid(row=2, column=0, pady=10)
        
        ttk.Label(option_frame, text="단위:").grid(row=0, column=0, padx=5)
        self.unit_var = tk.StringVar(value="mm")
        ttk.Combobox(option_frame, textvariable=self.unit_var, 
                    values=["mm", "cm", "m"], width=8).grid(row=0, column=1, padx=5)
        
        ttk.Label(option_frame, text="할증률(%):").grid(row=0, column=2, padx=5)
        self.waste_var = tk.StringVar(value="5")
        ttk.Entry(option_frame, textvariable=self.waste_var, width=8).grid(row=0, column=3, padx=5)
        
        ttk.Label(option_frame, text="높이/두께(m):").grid(row=0, column=4, padx=5)
        self.height_var = tk.StringVar(value="3.0")
        ttk.Entry(option_frame, textvariable=self.height_var, width=8).grid(row=0, column=5, padx=5)
    
    def create_result_area(self, parent):
        """결과 영역"""
        frame = ttk.LabelFrame(parent, text="계산 결과", padding="10")
        frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        # 결과 텍스트
        self.result_text = scrolledtext.ScrolledText(frame, height=15, width=100)
        self.result_text.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 버튼들
        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=1, column=0, pady=5)
        
        ttk.Button(btn_frame, text="결과 지우기", 
                  command=lambda: self.result_text.delete(1.0, tk.END)).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="JSON 저장", 
                  command=self.save_json).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="클립보드 복사", 
                  command=self.copy_to_clipboard).pack(side=tk.LEFT, padx=5)
    
    # === AutoCAD 연결 ===
    
    def connect_autocad(self):
        """AutoCAD 연결 및 데이터 로드"""
        try:
            # AutoCAD 연결
            self.acad = win32com.client.Dispatch("AutoCAD.Application")
            
            if self.acad.Documents.Count == 0:
                messagebox.showwarning("경고", "열린 도면이 없습니다.")
                return
            
            self.doc = self.acad.ActiveDocument
            self.model = self.doc.ModelSpace
            
            # 모든 객체 로드
            self.load_all_objects()
            
            # UI 업데이트
            self.status_label.config(text="연결됨", foreground="green")
            self.drawing_info.config(text=f"도면: {self.doc.Name}")
            self.object_count.config(text=f"객체: {len(self.all_objects)}개")
            
            # 레이어와 타입 목록 업데이트
            self.update_combos()
            
            messagebox.showinfo("성공", 
                f"AutoCAD 연결 성공!\n"
                f"도면: {self.doc.Name}\n"
                f"객체 수: {len(self.all_objects)}개")
            
        except Exception as e:
            messagebox.showerror("오류", f"AutoCAD 연결 실패:\n{str(e)}")
    
    def load_all_objects(self):
        """모든 객체를 메모리에 로드"""
        self.all_objects = []
        self.filtered_objects = []
        
        try:
            for obj in self.model:
                obj_data = self.extract_object_data(obj)
                if obj_data:
                    self.all_objects.append(obj_data)
            
            # 초기에는 전체 객체가 필터링된 객체
            self.filtered_objects = self.all_objects.copy()
            self.filtered_count.config(text=f"대상: {len(self.filtered_objects)}개")
            
        except Exception as e:
            print(f"객체 로드 오류: {e}")
    
    def extract_object_data(self, obj):
        """객체에서 필요한 데이터 추출"""
        try:
            data = {
                'type': obj.ObjectName,
                'layer': obj.Layer,
                'handle': obj.Handle
            }
            
            # 길이 정보
            if hasattr(obj, 'Length'):
                data['length'] = obj.Length
            elif "Line" in obj.ObjectName:
                try:
                    start = obj.StartPoint
                    end = obj.EndPoint
                    length = math.sqrt((end[0]-start[0])**2 + (end[1]-start[1])**2 + (end[2]-start[2])**2)
                    data['length'] = length
                except:
                    data['length'] = 0
            
            # 면적 정보
            if hasattr(obj, 'Area'):
                try:
                    data['area'] = obj.Area
                except:
                    data['area'] = 0
            
            # 원의 경우
            if "Circle" in obj.ObjectName:
                try:
                    data['radius'] = obj.Radius
                    data['area'] = math.pi * obj.Radius ** 2
                except:
                    pass
            
            # 블록 이름
            if "BlockReference" in obj.ObjectName:
                try:
                    data['block_name'] = obj.Name
                except:
                    pass
            
            return data
            
        except:
            return None
    
    def update_combos(self):
        """콤보박스 업데이트"""
        # 레이어 목록
        layers = list(set(obj['layer'] for obj in self.all_objects if 'layer' in obj))
        layers.sort()
        layers.insert(0, "전체")
        self.layer_combo['values'] = layers
        self.layer_combo.set("전체")
        
        # 타입 목록
        types = list(set(obj['type'] for obj in self.all_objects if 'type' in obj))
        types.sort()
        types.insert(0, "전체")
        self.type_combo['values'] = types
        self.type_combo.set("전체")
    
    def refresh_data(self):
        """데이터 새로고침"""
        if self.acad:
            self.load_all_objects()
            self.update_combos()
            messagebox.showinfo("완료", "데이터를 새로고침했습니다.")
    
    # === 필터링 ===
    
    def filter_by_layer(self):
        """레이어로 필터링"""
        selected_layer = self.layer_combo.get()
        
        if selected_layer == "전체":
            self.filtered_objects = self.all_objects.copy()
            self.filter_info.config(text="필터: 없음 (전체 객체)")
        else:
            self.filtered_objects = [obj for obj in self.all_objects 
                                    if obj.get('layer') == selected_layer]
            self.filter_info.config(text=f"필터: 레이어 = {selected_layer}")
        
        self.filtered_count.config(text=f"대상: {len(self.filtered_objects)}개")
        
        # 타입 콤보박스 업데이트
        types = list(set(obj['type'] for obj in self.filtered_objects if 'type' in obj))
        types.sort()
        types.insert(0, "전체")
        self.type_combo['values'] = types
    
    def filter_by_type(self):
        """타입으로 필터링"""
        selected_type = self.type_combo.get()
        
        if selected_type == "전체":
            # 현재 레이어 필터 유지
            selected_layer = self.layer_combo.get()
            if selected_layer == "전체":
                self.filtered_objects = self.all_objects.copy()
                self.filter_info.config(text="필터: 없음 (전체 객체)")
            else:
                self.filtered_objects = [obj for obj in self.all_objects 
                                        if obj.get('layer') == selected_layer]
                self.filter_info.config(text=f"필터: 레이어 = {selected_layer}")
        else:
            selected_layer = self.layer_combo.get()
            if selected_layer == "전체":
                self.filtered_objects = [obj for obj in self.all_objects 
                                        if obj.get('type') == selected_type]
                self.filter_info.config(text=f"필터: 타입 = {selected_type}")
            else:
                self.filtered_objects = [obj for obj in self.all_objects 
                                        if obj.get('layer') == selected_layer 
                                        and obj.get('type') == selected_type]
                self.filter_info.config(text=f"필터: 레이어 = {selected_layer}, 타입 = {selected_type}")
        
        self.filtered_count.config(text=f"대상: {len(self.filtered_objects)}개")
    
    def clear_filter(self):
        """필터 초기화"""
        self.filtered_objects = self.all_objects.copy()
        self.filter_info.config(text="필터: 없음 (전체 객체)")
        self.filtered_count.config(text=f"대상: {len(self.filtered_objects)}개")
        self.layer_combo.set("전체")
        self.type_combo.set("전체")
    
    # === 계산 메서드 ===
    
    def calc_total_length(self):
        """총 길이 계산"""
        if not self.filtered_objects:
            messagebox.showwarning("경고", "계산할 객체가 없습니다.")
            return
        
        total_length = 0
        count = 0
        
        for obj in self.filtered_objects:
            if 'length' in obj and obj['length'] > 0:
                total_length += obj['length']
                count += 1
        
        # 단위 변환
        if self.unit_var.get() == "mm":
            length_m = total_length / 1000
            unit_str = "mm"
        elif self.unit_var.get() == "cm":
            length_m = total_length / 100
            unit_str = "cm"
        else:
            length_m = total_length
            unit_str = "m"
        
        result = f"=== 총 길이 계산 ===\n"
        result += f"대상 객체: {len(self.filtered_objects)}개\n"
        result += f"길이 있는 객체: {count}개\n"
        result += f"총 길이: {total_length:.2f} {unit_str}\n"
        result += f"미터 환산: {length_m:.2f} m\n"
        
        self.display_result(result)
        self.calculation_results['총길이'] = length_m
    
    def calc_total_area(self):
        """총 면적 계산"""
        if not self.filtered_objects:
            messagebox.showwarning("경고", "계산할 객체가 없습니다.")
            return
        
        total_area = 0
        count = 0
        
        for obj in self.filtered_objects:
            if 'area' in obj and obj['area'] > 0:
                total_area += obj['area']
                count += 1
        
        # 단위 변환
        if self.unit_var.get() == "mm":
            area_m2 = total_area / 1000000
            unit_str = "mm²"
        elif self.unit_var.get() == "cm":
            area_m2 = total_area / 10000
            unit_str = "cm²"
        else:
            area_m2 = total_area
            unit_str = "m²"
        
        result = f"=== 총 면적 계산 ===\n"
        result += f"대상 객체: {len(self.filtered_objects)}개\n"
        result += f"면적 있는 객체: {count}개\n"
        result += f"총 면적: {total_area:.2f} {unit_str}\n"
        result += f"평방미터 환산: {area_m2:.2f} m²\n"
        result += f"평 환산: {area_m2/3.3:.2f} 평\n"
        
        self.display_result(result)
        self.calculation_results['총면적'] = area_m2
    
    def calc_count(self):
        """객체 개수 집계"""
        if not self.filtered_objects:
            messagebox.showwarning("경고", "계산할 객체가 없습니다.")
            return
        
        # 타입별 집계
        type_count = {}
        for obj in self.filtered_objects:
            obj_type = obj.get('type', 'Unknown')
            if obj_type not in type_count:
                type_count[obj_type] = 0
            type_count[obj_type] += 1
        
        # 레이어별 집계
        layer_count = {}
        for obj in self.filtered_objects:
            layer = obj.get('layer', 'Unknown')
            if layer not in layer_count:
                layer_count[layer] = 0
            layer_count[layer] += 1
        
        result = f"=== 객체 개수 집계 ===\n"
        result += f"총 객체 수: {len(self.filtered_objects)}개\n\n"
        
        result += "타입별 집계:\n"
        for obj_type, count in sorted(type_count.items(), key=lambda x: x[1], reverse=True):
            result += f"  {obj_type}: {count}개\n"
        
        result += "\n레이어별 집계:\n"
        for layer, count in sorted(layer_count.items(), key=lambda x: x[1], reverse=True)[:10]:
            result += f"  {layer}: {count}개\n"
        
        self.display_result(result)
        self.calculation_results['개수'] = len(self.filtered_objects)
    
    def calc_deck_joists(self):
        """장선/데크보드 계산"""
        if not self.filtered_objects:
            messagebox.showwarning("경고", "계산할 객체가 없습니다.")
            return
        
        # 길이 계산 (장선)
        total_length = 0
        line_count = 0
        
        for obj in self.filtered_objects:
            if 'length' in obj and obj['length'] > 0:
                if 'Line' in obj.get('type', ''):
                    total_length += obj['length']
                    line_count += 1
        
        # 면적 계산 (데크보드)
        total_area = 0
        area_count = 0
        
        for obj in self.filtered_objects:
            if 'area' in obj and obj['area'] > 0:
                total_area += obj['area']
                area_count += 1
        
        # 단위 변환
        if self.unit_var.get() == "mm":
            length_m = total_length / 1000
            area_m2 = total_area / 1000000
        else:
            length_m = total_length
            area_m2 = total_area
        
        waste = float(self.waste_var.get()) / 100
        
        result = f"=== 장선/데크보드 계산 ===\n\n"
        
        if line_count > 0:
            volume = length_m * 0.038 * 0.235  # 2x10
            result += f"[장선]\n"
            result += f"  선 개수: {line_count}개\n"
            result += f"  총 길이: {length_m:.2f} m\n"
            result += f"  목재 체적: {volume:.3f} m³ (2x10)\n"
            result += f"  할증 {self.waste_var.get()}% 적용: {volume*(1+waste):.3f} m³\n\n"
        
        if area_count > 0:
            deck_area = area_m2 * (1 + waste)
            screws = int(area_m2 * 8)
            result += f"[데크보드]\n"
            result += f"  면적 객체: {area_count}개\n"
            result += f"  총 면적: {area_m2:.2f} m²\n"
            result += f"  필요 보드: {deck_area:.2f} m² (할증 포함)\n"
            result += f"  데크 스크류: {screws}개\n"
        
        self.display_result(result)
    
    def calc_posts(self):
        """기둥 계산"""
        if not self.filtered_objects:
            messagebox.showwarning("경고", "계산할 객체가 없습니다.")
            return
        
        # 원형 객체를 기둥으로 가정
        circle_count = 0
        for obj in self.filtered_objects:
            if 'Circle' in obj.get('type', ''):
                circle_count += 1
        
        # 블록도 포함
        block_count = 0
        for obj in self.filtered_objects:
            if 'Block' in obj.get('type', ''):
                block_count += 1
        
        total_posts = circle_count + block_count
        
        if total_posts == 0:
            result = "기둥으로 사용할 객체(원, 블록)가 없습니다."
        else:
            foundation = total_posts * 0.096  # 400x400x600
            
            result = f"=== 기둥 계산 ===\n\n"
            result += f"원형 객체: {circle_count}개\n"
            result += f"블록 객체: {block_count}개\n"
            result += f"총 기둥 수: {total_posts}개\n\n"
            result += f"기초 콘크리트:\n"
            result += f"  기초 크기: 400x400x600 mm\n"
            result += f"  개당 체적: 0.096 m³\n"
            result += f"  총 체적: {foundation:.3f} m³\n"
            result += f"  할증 5% 적용: {foundation*1.05:.3f} m³\n"
        
        self.display_result(result)
    
    def calc_walls(self):
        """벽체 계산"""
        if not self.filtered_objects:
            messagebox.showwarning("경고", "계산할 객체가 없습니다.")
            return
        
        # 선을 벽체로 가정
        total_length = 0
        count = 0
        
        for obj in self.filtered_objects:
            if 'length' in obj and obj['length'] > 0:
                total_length += obj['length']
                count += 1
        
        if count == 0:
            result = "벽체로 계산할 선형 객체가 없습니다."
        else:
            # 단위 변환
            if self.unit_var.get() == "mm":
                length_m = total_length / 1000
            else:
                length_m = total_length
            
            height = float(self.height_var.get())  # 높이
            thickness = 0.2  # 벽 두께 200mm
            
            wall_area = length_m * height
            wall_volume = wall_area * thickness
            
            result = f"=== 벽체 계산 ===\n\n"
            result += f"선형 객체: {count}개\n"
            result += f"총 길이: {length_m:.2f} m\n"
            result += f"벽 높이: {height} m\n"
            result += f"벽 두께: {thickness} m\n\n"
            result += f"벽체 면적: {wall_area:.2f} m²\n"
            result += f"벽체 체적: {wall_volume:.3f} m³\n"
            result += f"콘크리트: {wall_volume*1.05:.3f} m³ (5% 할증)\n"
        
        self.display_result(result)
    
    def analyze_by_layer(self):
        """레이어별 분석"""
        if not self.all_objects:
            messagebox.showwarning("경고", "분석할 데이터가 없습니다.")
            return
        
        layer_stats = {}
        
        for obj in self.all_objects:
            layer = obj.get('layer', 'Unknown')
            
            if layer not in layer_stats:
                layer_stats[layer] = {
                    'count': 0,
                    'types': set(),
                    'length': 0,
                    'area': 0
                }
            
            layer_stats[layer]['count'] += 1
            layer_stats[layer]['types'].add(obj.get('type', 'Unknown'))
            
            if 'length' in obj:
                layer_stats[layer]['length'] += obj['length']
            if 'area' in obj:
                layer_stats[layer]['area'] += obj['area']
        
        result = f"=== 레이어별 분석 ===\n\n"
        result += f"총 레이어 수: {len(layer_stats)}개\n\n"
        
        # 객체 수 기준 정렬
        sorted_layers = sorted(layer_stats.items(), key=lambda x: x[1]['count'], reverse=True)
        
        for layer, stats in sorted_layers[:15]:  # 상위 15개
            result += f"{layer}:\n"
            result += f"  객체 수: {stats['count']}개\n"
            result += f"  타입 수: {len(stats['types'])}종\n"
            
            if stats['length'] > 0:
                if self.unit_var.get() == "mm":
                    result += f"  총 길이: {stats['length']/1000:.2f} m\n"
                else:
                    result += f"  총 길이: {stats['length']:.2f} m\n"
            
            if stats['area'] > 0:
                if self.unit_var.get() == "mm":
                    result += f"  총 면적: {stats['area']/1000000:.2f} m²\n"
                else:
                    result += f"  총 면적: {stats['area']:.2f} m²\n"
            
            result += "\n"
        
        self.display_result(result)
    
    def analyze_by_type(self):
        """타입별 분석"""
        if not self.all_objects:
            messagebox.showwarning("경고", "분석할 데이터가 없습니다.")
            return
        
        type_stats = {}
        
        for obj in self.all_objects:
            obj_type = obj.get('type', 'Unknown')
            
            if obj_type not in type_stats:
                type_stats[obj_type] = {
                    'count': 0,
                    'layers': set(),
                    'length': 0,
                    'area': 0
                }
            
            type_stats[obj_type]['count'] += 1
            type_stats[obj_type]['layers'].add(obj.get('layer', 'Unknown'))
            
            if 'length' in obj:
                type_stats[obj_type]['length'] += obj['length']
            if 'area' in obj:
                type_stats[obj_type]['area'] += obj['area']
        
        result = f"=== 타입별 분석 ===\n\n"
        result += f"총 타입 수: {len(type_stats)}개\n\n"
        
        # 객체 수 기준 정렬
        sorted_types = sorted(type_stats.items(), key=lambda x: x[1]['count'], reverse=True)
        
        for obj_type, stats in sorted_types[:15]:  # 상위 15개
            result += f"{obj_type}:\n"
            result += f"  객체 수: {stats['count']}개\n"
            result += f"  레이어 수: {len(stats['layers'])}개\n"
            
            if stats['length'] > 0:
                if self.unit_var.get() == "mm":
                    result += f"  총 길이: {stats['length']/1000:.2f} m\n"
                else:
                    result += f"  총 길이: {stats['length']:.2f} m\n"
            
            if stats['area'] > 0:
                if self.unit_var.get() == "mm":
                    result += f"  총 면적: {stats['area']/1000000:.2f} m²\n"
                else:
                    result += f"  총 면적: {stats['area']:.2f} m²\n"
            
            result += "\n"
        
        self.display_result(result)
    
    # === 유틸리티 ===
    
    def display_result(self, text):
        """결과 표시"""
        # 기존 내용에 추가
        self.result_text.insert(tk.END, "\n" + text + "\n")
        self.result_text.insert(tk.END, f"계산 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        self.result_text.insert(tk.END, "-" * 50 + "\n")
        self.result_text.see(tk.END)
    
    def save_json(self):
        """JSON 저장"""
        if not self.calculation_results:
            messagebox.showwarning("경고", "저장할 결과가 없습니다.")
            return
        
        filename = f"cadpro_result_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        
        try:
            output = {
                "도면": self.doc.Name if self.doc else "Unknown",
                "일시": datetime.now().isoformat(),
                "필터": self.filter_info.cget("text"),
                "결과": self.calculation_results
            }
            
            with open(filename, 'w', encoding='utf-8') as f:
                json.dump(output, f, ensure_ascii=False, indent=2)
            
            messagebox.showinfo("성공", f"저장 완료: {filename}")
            
        except Exception as e:
            messagebox.showerror("오류", f"저장 실패: {str(e)}")
    
    def copy_to_clipboard(self):
        """클립보드 복사"""
        try:
            text = self.result_text.get(1.0, tk.END)
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            messagebox.showinfo("성공", "클립보드에 복사되었습니다.")
        except:
            pass
    
    def run(self):
        """프로그램 실행"""
        self.root.mainloop()


def main():
    app = CADProStable()
    app.run()


if __name__ == "__main__":
    main()