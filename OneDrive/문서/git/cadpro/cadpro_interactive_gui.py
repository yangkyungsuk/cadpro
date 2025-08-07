"""
CADPro Interactive GUI - 사용자 직접 선택 방식
AutoCAD에서 선택 → GUI에서 용도 지정 → 계산
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import win32com.client
import pythoncom
import math
import json
from datetime import datetime


class CADProInteractiveGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("CADPro - 대화형 물량 산출")
        self.root.geometry("900x700")
        
        # AutoCAD 연결
        self.acad = None
        self.doc = None
        self.model = None
        
        # 선택된 객체
        self.selected_objects = []
        
        # GUI 구성
        self.setup_gui()
    
    def setup_gui(self):
        """GUI 구성"""
        
        # 상단: AutoCAD 연결
        conn_frame = ttk.LabelFrame(self.root, text="1. AutoCAD 연결", padding="10")
        conn_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Button(conn_frame, text="AutoCAD 연결", 
                  command=self.connect_autocad, width=20).pack(side=tk.LEFT, padx=5)
        
        self.status_label = ttk.Label(conn_frame, text="연결 대기 중...", foreground="gray")
        self.status_label.pack(side=tk.LEFT, padx=20)
        
        self.drawing_label = ttk.Label(conn_frame, text="도면: -")
        self.drawing_label.pack(side=tk.LEFT, padx=20)
        
        # 선택 프레임
        select_frame = ttk.LabelFrame(self.root, text="2. AutoCAD에서 객체 선택", padding="10")
        select_frame.pack(fill='x', padx=10, pady=5)
        
        instruction = ttk.Label(select_frame, 
            text="① AutoCAD에서 원하는 객체를 선택하세요 (클릭, 윈도우, 교차 등)\n" +
                 "② 선택 완료 후 아래 버튼을 클릭하세요", 
            foreground="blue")
        instruction.pack(pady=5)
        
        btn_frame = ttk.Frame(select_frame)
        btn_frame.pack(pady=5)
        
        ttk.Button(btn_frame, text="선택한 객체 가져오기", 
                  command=self.get_selection, width=25).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(btn_frame, text="선택 초기화", 
                  command=self.clear_selection, width=15).pack(side=tk.LEFT, padx=5)
        
        self.selection_info = ttk.Label(select_frame, text="선택된 객체: 0개")
        self.selection_info.pack(pady=5)
        
        # 선택 정보 표시
        info_frame = ttk.Frame(select_frame)
        info_frame.pack(fill='x', pady=5)
        
        self.info_text = tk.Text(info_frame, height=6, width=80)
        self.info_text.pack(side=tk.LEFT, padx=5)
        
        info_scroll = ttk.Scrollbar(info_frame, command=self.info_text.yview)
        info_scroll.pack(side=tk.RIGHT, fill='y')
        self.info_text.config(yscrollcommand=info_scroll.set)
        
        # 용도 지정 및 계산
        calc_frame = ttk.LabelFrame(self.root, text="3. 선택한 객체의 용도 지정 및 계산", padding="10")
        calc_frame.pack(fill='x', padx=10, pady=5)
        
        usage_label = ttk.Label(calc_frame, 
            text="선택한 객체를 무엇으로 계산하시겠습니까?", 
            font=('', 10, 'bold'))
        usage_label.pack(pady=5)
        
        # 건축 계산
        arch_frame = ttk.Frame(calc_frame)
        arch_frame.pack(fill='x', pady=3)
        ttk.Label(arch_frame, text="건축:", width=10).pack(side=tk.LEFT)
        ttk.Button(arch_frame, text="벽체", command=lambda: self.calculate_as("벽체"), width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(arch_frame, text="기둥", command=lambda: self.calculate_as("기둥"), width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(arch_frame, text="보", command=lambda: self.calculate_as("보"), width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(arch_frame, text="슬래브", command=lambda: self.calculate_as("슬래브"), width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(arch_frame, text="철근", command=lambda: self.calculate_as("철근"), width=10).pack(side=tk.LEFT, padx=2)
        
        # 토목 계산
        civil_frame = ttk.Frame(calc_frame)
        civil_frame.pack(fill='x', pady=3)
        ttk.Label(civil_frame, text="토목:", width=10).pack(side=tk.LEFT)
        ttk.Button(civil_frame, text="도로", command=lambda: self.calculate_as("도로"), width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(civil_frame, text="배관", command=lambda: self.calculate_as("배관"), width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(civil_frame, text="맨홀", command=lambda: self.calculate_as("맨홀"), width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(civil_frame, text="U형측구", command=lambda: self.calculate_as("U형측구"), width=10).pack(side=tk.LEFT, padx=2)
        
        # 조경/데크 계산
        land_frame = ttk.Frame(calc_frame)
        land_frame.pack(fill='x', pady=3)
        ttk.Label(land_frame, text="조경/데크:", width=10).pack(side=tk.LEFT)
        ttk.Button(land_frame, text="데크장선", command=lambda: self.calculate_as("데크장선"), width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(land_frame, text="데크보드", command=lambda: self.calculate_as("데크보드"), width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(land_frame, text="데크기둥", command=lambda: self.calculate_as("데크기둥"), width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(land_frame, text="수목", command=lambda: self.calculate_as("수목"), width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(land_frame, text="포장", command=lambda: self.calculate_as("포장"), width=10).pack(side=tk.LEFT, padx=2)
        
        # 일반 계산
        general_frame = ttk.Frame(calc_frame)
        general_frame.pack(fill='x', pady=3)
        ttk.Label(general_frame, text="일반:", width=10).pack(side=tk.LEFT)
        ttk.Button(general_frame, text="길이", command=lambda: self.calculate_as("길이"), width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(general_frame, text="면적", command=lambda: self.calculate_as("면적"), width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(general_frame, text="개수", command=lambda: self.calculate_as("개수"), width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(general_frame, text="체적", command=lambda: self.calculate_as("체적"), width=10).pack(side=tk.LEFT, padx=2)
        
        # 옵션
        option_frame = ttk.Frame(calc_frame)
        option_frame.pack(fill='x', pady=5)
        
        ttk.Label(option_frame, text="할증률(%):").pack(side=tk.LEFT, padx=5)
        self.waste_var = tk.StringVar(value="5")
        ttk.Entry(option_frame, textvariable=self.waste_var, width=8).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(option_frame, text="높이/두께(m):").pack(side=tk.LEFT, padx=5)
        self.height_var = tk.StringVar(value="3.0")
        ttk.Entry(option_frame, textvariable=self.height_var, width=8).pack(side=tk.LEFT, padx=5)
        
        # 결과 표시
        result_frame = ttk.LabelFrame(self.root, text="4. 계산 결과", padding="10")
        result_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        self.result_text = scrolledtext.ScrolledText(result_frame, height=12, width=100)
        self.result_text.pack(fill='both', expand=True)
        
        # 하단 버튼
        bottom_frame = ttk.Frame(self.root)
        bottom_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Button(bottom_frame, text="결과 지우기", 
                  command=lambda: self.result_text.delete(1.0, tk.END)).pack(side=tk.LEFT, padx=5)
        ttk.Button(bottom_frame, text="JSON 저장", 
                  command=self.save_json).pack(side=tk.LEFT, padx=5)
        ttk.Button(bottom_frame, text="클립보드 복사", 
                  command=self.copy_to_clipboard).pack(side=tk.LEFT, padx=5)
        ttk.Button(bottom_frame, text="종료", 
                  command=self.on_closing).pack(side=tk.RIGHT, padx=5)
    
    def connect_autocad(self):
        """AutoCAD 연결"""
        try:
            self.acad = win32com.client.Dispatch("AutoCAD.Application")
            self.doc = self.acad.ActiveDocument
            self.model = self.doc.ModelSpace
            
            # AutoCAD 활성화
            self.acad.Visible = True
            
            # UI 업데이트
            self.status_label.config(text="연결됨", foreground="green")
            self.drawing_label.config(text=f"도면: {self.doc.Name}")
            
            messagebox.showinfo("성공", 
                f"AutoCAD 연결 성공!\n도면: {self.doc.Name}\n\n" +
                "이제 AutoCAD에서 객체를 선택하세요.")
            
        except Exception as e:
            messagebox.showerror("오류", f"AutoCAD 연결 실패:\n{str(e)}")
    
    def get_selection(self):
        """AutoCAD에서 선택한 객체 가져오기"""
        if not self.doc:
            messagebox.showwarning("경고", "먼저 AutoCAD에 연결하세요.")
            return
        
        try:
            # PickfirstSelectionSet 사용
            sel_set = self.doc.PickfirstSelectionSet
            
            if sel_set.Count == 0:
                messagebox.showinfo("정보", 
                    "선택된 객체가 없습니다.\n\n" +
                    "AutoCAD에서 객체를 선택한 후 다시 시도하세요.")
                return
            
            self.selected_objects = []
            
            # 선택된 객체 정보 수집
            for obj in sel_set:
                obj_data = self.extract_object_data(obj)
                if obj_data:
                    obj_data['com_object'] = obj  # COM 객체 참조 저장
                    self.selected_objects.append(obj_data)
            
            # 정보 표시
            self.show_selection_info()
            
            messagebox.showinfo("성공", 
                f"{len(self.selected_objects)}개 객체를 가져왔습니다.\n\n" +
                "이제 용도를 지정하고 계산하세요.")
            
        except Exception as e:
            messagebox.showerror("오류", f"선택 가져오기 실패:\n{str(e)}")
    
    def extract_object_data(self, obj):
        """객체 데이터 추출"""
        try:
            data = {
                'type': obj.ObjectName,
                'layer': obj.Layer,
                'handle': obj.Handle
            }
            
            # 선
            if "Line" in obj.ObjectName and "Polyline" not in obj.ObjectName:
                try:
                    start = obj.StartPoint
                    end = obj.EndPoint
                    length = math.sqrt((end[0]-start[0])**2 + (end[1]-start[1])**2 + (end[2]-start[2])**2)
                    data['length'] = length
                except:
                    pass
            
            # 폴리라인
            elif "Polyline" in obj.ObjectName:
                try:
                    data['length'] = obj.Length
                    data['closed'] = obj.Closed
                    if obj.Closed:
                        data['area'] = obj.Area
                except:
                    pass
            
            # 원
            elif "Circle" in obj.ObjectName:
                try:
                    data['radius'] = obj.Radius
                    data['area'] = math.pi * obj.Radius ** 2
                except:
                    pass
            
            # 호
            elif "Arc" in obj.ObjectName:
                try:
                    data['arc_length'] = obj.ArcLength
                except:
                    pass
            
            # 블록
            elif "BlockReference" in obj.ObjectName:
                try:
                    data['block_name'] = obj.Name
                except:
                    pass
            
            # 해치
            elif "Hatch" in obj.ObjectName:
                try:
                    data['area'] = obj.Area
                except:
                    pass
            
            return data
            
        except:
            return None
    
    def show_selection_info(self):
        """선택 정보 표시"""
        if not self.selected_objects:
            self.selection_info.config(text="선택된 객체: 0개")
            self.info_text.delete(1.0, tk.END)
            return
        
        # 타입별 집계
        type_count = {}
        total_length = 0
        total_area = 0
        
        for obj in self.selected_objects:
            # 타입
            obj_type = obj.get('type', 'Unknown')
            type_count[obj_type] = type_count.get(obj_type, 0) + 1
            
            # 길이
            if 'length' in obj:
                total_length += obj['length']
            
            # 면적
            if 'area' in obj:
                total_area += obj['area']
        
        # 정보 표시
        self.selection_info.config(text=f"선택된 객체: {len(self.selected_objects)}개")
        
        info = "타입별 분포:\n"
        for obj_type, count in sorted(type_count.items()):
            info += f"  {obj_type}: {count}개\n"
        
        if total_length > 0:
            info += f"\n총 길이: {total_length:.2f}mm ({total_length/1000:.2f}m)\n"
        
        if total_area > 0:
            info += f"총 면적: {total_area:.2f}mm² ({total_area/1000000:.2f}m²)\n"
        
        self.info_text.delete(1.0, tk.END)
        self.info_text.insert(tk.END, info)
    
    def clear_selection(self):
        """선택 초기화"""
        self.selected_objects = []
        self.show_selection_info()
        messagebox.showinfo("완료", "선택이 초기화되었습니다.")
    
    def calculate_as(self, usage):
        """선택한 객체를 지정된 용도로 계산"""
        if not self.selected_objects:
            messagebox.showwarning("경고", "먼저 객체를 선택하세요.")
            return
        
        waste = float(self.waste_var.get()) / 100
        height = float(self.height_var.get())
        
        result = f"\n=== {usage} 계산 ===\n"
        result += f"선택 객체: {len(self.selected_objects)}개\n"
        result += f"계산 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n"
        
        # 용도별 계산
        if usage == "벽체":
            result += self.calc_wall(waste, height)
        elif usage == "기둥":
            result += self.calc_column(waste)
        elif usage == "보":
            result += self.calc_beam(waste)
        elif usage == "슬래브":
            result += self.calc_slab(waste)
        elif usage == "철근":
            result += self.calc_rebar(waste)
        elif usage == "도로":
            result += self.calc_road(waste)
        elif usage == "배관":
            result += self.calc_pipe(waste)
        elif usage == "맨홀":
            result += self.calc_manhole()
        elif usage == "U형측구":
            result += self.calc_u_ditch()
        elif usage == "데크장선":
            result += self.calc_deck_joist(waste)
        elif usage == "데크보드":
            result += self.calc_deck_board(waste)
        elif usage == "데크기둥":
            result += self.calc_deck_post(waste)
        elif usage == "수목":
            result += self.calc_trees()
        elif usage == "포장":
            result += self.calc_pavement(waste)
        elif usage == "길이":
            result += self.calc_general_length()
        elif usage == "면적":
            result += self.calc_general_area()
        elif usage == "개수":
            result += self.calc_general_count()
        elif usage == "체적":
            result += self.calc_general_volume(height)
        
        result += "\n" + "-"*60 + "\n"
        
        # 결과 표시
        self.result_text.insert(tk.END, result)
        self.result_text.see(tk.END)
    
    # === 계산 메서드들 ===
    
    def calc_wall(self, waste, height):
        """벽체 계산"""
        total_length = sum(obj.get('length', 0) for obj in self.selected_objects)
        
        if total_length == 0:
            return "벽체로 계산할 선 객체가 없습니다.\n"
        
        length_m = total_length / 1000
        thickness = 0.2  # 200mm
        wall_area = length_m * height
        wall_volume = wall_area * thickness
        concrete = wall_volume * (1 + waste)
        
        result = f"벽체 길이: {length_m:.2f}m\n"
        result += f"벽체 높이: {height}m\n"
        result += f"벽체 두께: {thickness}m\n"
        result += f"벽체 면적: {wall_area:.2f}m²\n"
        result += f"벽체 체적: {wall_volume:.3f}m³\n"
        result += f"콘크리트: {concrete:.3f}m³ (할증 {waste*100:.0f}%)\n"
        
        return result
    
    def calc_column(self, waste):
        """기둥 계산"""
        count = len(self.selected_objects)
        foundation = count * 0.096  # 400x400x600
        concrete = foundation * (1 + waste)
        
        result = f"기둥 개수: {count}개\n"
        result += f"기초 크기: 400x400x600mm\n"
        result += f"기초 체적: {foundation:.3f}m³\n"
        result += f"콘크리트: {concrete:.3f}m³ (할증 {waste*100:.0f}%)\n"
        
        return result
    
    def calc_beam(self, waste):
        """보 계산"""
        total_length = sum(obj.get('length', 0) for obj in self.selected_objects)
        
        if total_length == 0:
            return "보로 계산할 선 객체가 없습니다.\n"
        
        length_m = total_length / 1000
        section = 0.3 * 0.5  # 300x500mm
        volume = length_m * section
        concrete = volume * (1 + waste)
        
        result = f"보 길이: {length_m:.2f}m\n"
        result += f"보 단면: 300x500mm\n"
        result += f"보 체적: {volume:.3f}m³\n"
        result += f"콘크리트: {concrete:.3f}m³ (할증 {waste*100:.0f}%)\n"
        
        return result
    
    def calc_slab(self, waste):
        """슬래브 계산"""
        total_area = sum(obj.get('area', 0) for obj in self.selected_objects)
        
        if total_area == 0:
            return "슬래브로 계산할 면적 객체가 없습니다.\n"
        
        area_m2 = total_area / 1000000
        thickness = 0.15  # 150mm
        volume = area_m2 * thickness
        concrete = volume * (1 + waste)
        rebar = area_m2 * 0.015  # ton/m²
        
        result = f"슬래브 면적: {area_m2:.2f}m²\n"
        result += f"슬래브 두께: {thickness*1000:.0f}mm\n"
        result += f"콘크리트: {concrete:.3f}m³ (할증 {waste*100:.0f}%)\n"
        result += f"철근: {rebar:.3f}ton\n"
        
        return result
    
    def calc_rebar(self, waste):
        """철근 계산"""
        total_length = sum(obj.get('length', 0) for obj in self.selected_objects)
        
        if total_length == 0:
            return "철근으로 계산할 선 객체가 없습니다.\n"
        
        length_m = total_length / 1000
        weight_d16 = length_m * 1.56  # kg (D16 기준)
        weight_ton = weight_d16 / 1000
        
        result = f"철근 길이: {length_m:.2f}m\n"
        result += f"D16 기준 중량: {weight_d16:.2f}kg\n"
        result += f"중량: {weight_ton:.3f}ton\n"
        result += f"할증 적용: {weight_ton*(1+waste):.3f}ton ({waste*100:.0f}%)\n"
        
        return result
    
    def calc_road(self, waste):
        """도로 계산"""
        total_area = sum(obj.get('area', 0) for obj in self.selected_objects)
        
        if total_area == 0:
            return "도로로 계산할 면적 객체가 없습니다.\n"
        
        area_m2 = total_area / 1000000
        asphalt = area_m2 * 0.1  # 100mm
        base = area_m2 * 0.15  # 150mm
        
        result = f"도로 면적: {area_m2:.2f}m²\n"
        result += f"아스팔트 (100mm): {asphalt*(1+waste):.3f}m³\n"
        result += f"기층 (150mm): {base*(1+waste):.3f}m³\n"
        
        return result
    
    def calc_pipe(self, waste):
        """배관 계산"""
        total_length = sum(obj.get('length', 0) for obj in self.selected_objects)
        
        if total_length == 0:
            return "배관으로 계산할 선 객체가 없습니다.\n"
        
        length_m = total_length / 1000
        joints = int(length_m / 6)  # 6m당 1개
        
        result = f"배관 연장: {length_m:.2f}m\n"
        result += f"이음관: {joints}개 (6m당 1개)\n"
        result += f"할증 포함: {length_m*(1+waste):.2f}m\n"
        
        return result
    
    def calc_manhole(self):
        """맨홀 계산"""
        count = len(self.selected_objects)
        
        result = f"맨홀 개수: {count}개\n"
        result += f"표준 규격: Φ900mm\n"
        
        return result
    
    def calc_u_ditch(self):
        """U형측구 계산"""
        total_length = sum(obj.get('length', 0) for obj in self.selected_objects)
        
        if total_length == 0:
            return "U형측구로 계산할 선 객체가 없습니다.\n"
        
        length_m = total_length / 1000
        pieces = int(length_m / 2)  # 2m 제품 기준
        
        result = f"U형측구 연장: {length_m:.2f}m\n"
        result += f"필요 개수: {pieces}개 (2m 제품)\n"
        
        return result
    
    def calc_deck_joist(self, waste):
        """데크 장선 계산"""
        total_length = sum(obj.get('length', 0) for obj in self.selected_objects)
        
        if total_length == 0:
            return "장선으로 계산할 선 객체가 없습니다.\n"
        
        length_m = total_length / 1000
        volume = length_m * 0.038 * 0.235  # 2x10 (38x235mm)
        
        result = f"장선 길이: {length_m:.2f}m\n"
        result += f"규격: 2x10 (38x235mm)\n"
        result += f"목재 체적: {volume:.3f}m³\n"
        result += f"할증 포함: {volume*(1+waste):.3f}m³ ({waste*100:.0f}%)\n"
        
        return result
    
    def calc_deck_board(self, waste):
        """데크보드 계산"""
        total_area = sum(obj.get('area', 0) for obj in self.selected_objects)
        
        if total_area == 0:
            return "데크보드로 계산할 면적 객체가 없습니다.\n"
        
        area_m2 = total_area / 1000000
        board_area = area_m2 * (1 + waste)
        screws = int(area_m2 * 8)
        
        result = f"데크 면적: {area_m2:.2f}m²\n"
        result += f"필요 보드: {board_area:.2f}m² (할증 {waste*100:.0f}%)\n"
        result += f"데크 스크류: {screws}개 (8개/m²)\n"
        
        return result
    
    def calc_deck_post(self, waste):
        """데크 기둥 계산"""
        count = len(self.selected_objects)
        foundation = count * 0.096  # 400x400x600
        
        result = f"데크 기둥: {count}개\n"
        result += f"기초 크기: 400x400x600mm\n"
        result += f"콘크리트: {foundation*(1+waste):.3f}m³\n"
        
        return result
    
    def calc_trees(self):
        """수목 계산"""
        count = len(self.selected_objects)
        
        result = f"수목 수량: {count}주\n"
        
        return result
    
    def calc_pavement(self, waste):
        """포장 계산"""
        total_area = sum(obj.get('area', 0) for obj in self.selected_objects)
        
        if total_area == 0:
            return "포장으로 계산할 면적 객체가 없습니다.\n"
        
        area_m2 = total_area / 1000000
        blocks = int(area_m2 * 50)  # 50개/m²
        sand = area_m2 * 0.03  # 30mm
        
        result = f"포장 면적: {area_m2:.2f}m²\n"
        result += f"보도블록: {blocks}개 (50개/m²)\n"
        result += f"모래: {sand*(1+waste):.3f}m³\n"
        
        return result
    
    def calc_general_length(self):
        """일반 길이 계산"""
        total = sum(obj.get('length', 0) for obj in self.selected_objects)
        
        result = f"총 길이: {total:.2f}mm\n"
        result += f"미터 환산: {total/1000:.2f}m\n"
        
        return result
    
    def calc_general_area(self):
        """일반 면적 계산"""
        total = sum(obj.get('area', 0) for obj in self.selected_objects)
        
        result = f"총 면적: {total:.2f}mm²\n"
        result += f"평방미터 환산: {total/1000000:.2f}m²\n"
        result += f"평 환산: {total/1000000/3.3:.2f}평\n"
        
        return result
    
    def calc_general_count(self):
        """일반 개수 계산"""
        # 타입별 집계
        type_count = {}
        for obj in self.selected_objects:
            obj_type = obj.get('type', 'Unknown')
            type_count[obj_type] = type_count.get(obj_type, 0) + 1
        
        result = f"총 개수: {len(self.selected_objects)}개\n\n"
        result += "타입별 개수:\n"
        for obj_type, count in sorted(type_count.items()):
            result += f"  {obj_type}: {count}개\n"
        
        return result
    
    def calc_general_volume(self, height):
        """일반 체적 계산"""
        total_area = sum(obj.get('area', 0) for obj in self.selected_objects)
        
        if total_area == 0:
            return "체적을 계산할 면적 객체가 없습니다.\n"
        
        area_m2 = total_area / 1000000
        volume = area_m2 * height
        
        result = f"면적: {area_m2:.2f}m²\n"
        result += f"높이/두께: {height}m\n"
        result += f"체적: {volume:.3f}m³\n"
        
        return result
    
    def save_json(self):
        """결과 JSON 저장"""
        text = self.result_text.get(1.0, tk.END)
        if not text.strip():
            messagebox.showwarning("경고", "저장할 결과가 없습니다.")
            return
        
        filename = f"cadpro_result_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        
        try:
            output = {
                "도면": self.doc.Name if self.doc else "Unknown",
                "일시": datetime.now().isoformat(),
                "결과": text
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
    
    def on_closing(self):
        """프로그램 종료"""
        try:
            self.root.quit()
            self.root.destroy()
        except:
            self.root.destroy()
    
    def run(self):
        """프로그램 실행"""
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.mainloop()


def main():
    app = CADProInteractiveGUI()
    app.run()


if __name__ == "__main__":
    main()