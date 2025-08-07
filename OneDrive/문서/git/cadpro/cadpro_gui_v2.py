"""
CAD 물량 산출 프로그램 - 개선된 GUI 버전
더 안정적인 AutoCAD 연동과 선택 기능
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import win32com.client
import pythoncom
import ezdxf
import math
import json
import os
from datetime import datetime
import pandas as pd
import time


class ImprovedCADProGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("CADPro v2 - 물량 산출 프로그램")
        self.root.geometry("1200x700")
        
        # AutoCAD 연결 상태
        self.acad = None
        self.doc = None
        self.model = None
        self.current_file = None
        self.selected_objects = []
        self.calculation_results = {}
        
        # 선택 모드 상태
        self.selection_mode = None
        
        # GUI 구성
        self.setup_gui()
        
        # 초기 상태 확인
        self.check_autocad_status()
    
    def setup_gui(self):
        """GUI 레이아웃 구성"""
        
        # 스타일 설정
        style = ttk.Style()
        style.theme_use('clam')
        
        # 메인 컨테이너
        main_container = ttk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 상단 툴바
        self.create_toolbar(main_container)
        
        # 중앙 영역을 3개 패널로 분할
        content_frame = ttk.Frame(main_container)
        content_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 왼쪽 패널
        self.create_left_panel(content_frame)
        
        # 중앙 패널
        self.create_center_panel(content_frame)
        
        # 오른쪽 패널
        self.create_right_panel(content_frame)
        
        # 하단 상태바
        self.create_statusbar()
    
    def create_toolbar(self, parent):
        """상단 툴바"""
        toolbar_frame = ttk.LabelFrame(parent, text="🔧 도구", relief=tk.RAISED)
        toolbar_frame.pack(fill=tk.X, pady=2)
        
        # AutoCAD 연결 섹션
        cad_frame = ttk.Frame(toolbar_frame)
        cad_frame.pack(side=tk.LEFT, padx=10, pady=5)
        
        self.connect_btn = ttk.Button(cad_frame, text="🔗 AutoCAD 연결", 
                                      command=self.connect_autocad, width=15)
        self.connect_btn.pack(side=tk.LEFT, padx=2)
        
        self.connection_status = ttk.Label(cad_frame, text="⚪ 미연결", foreground="gray")
        self.connection_status.pack(side=tk.LEFT, padx=10)
        
        ttk.Separator(toolbar_frame, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)
        
        # 파일 작업 섹션
        file_frame = ttk.Frame(toolbar_frame)
        file_frame.pack(side=tk.LEFT, padx=10, pady=5)
        
        ttk.Button(file_frame, text="📁 DXF 열기", 
                  command=self.open_dxf_file, width=12).pack(side=tk.LEFT, padx=2)
        
        ttk.Button(file_frame, text="💾 결과 저장", 
                  command=self.save_results, width=12).pack(side=tk.LEFT, padx=2)
        
        ttk.Separator(toolbar_frame, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)
        
        # 도움말
        help_frame = ttk.Frame(toolbar_frame)
        help_frame.pack(side=tk.RIGHT, padx=10, pady=5)
        
        ttk.Button(help_frame, text="❓ 도움말", 
                  command=self.show_help, width=10).pack(side=tk.LEFT, padx=2)
    
    def create_left_panel(self, parent):
        """왼쪽 패널 - 선택 도구"""
        left_frame = ttk.LabelFrame(parent, text="🎯 객체 선택", width=280)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, padx=2)
        left_frame.pack_propagate(False)
        
        # AutoCAD 선택 방법
        autocad_frame = ttk.LabelFrame(left_frame, text="AutoCAD 선택")
        autocad_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 선택 세트 사용
        ttk.Button(autocad_frame, text="현재 선택 세트 가져오기", 
                  command=self.get_current_selection, width=25).pack(pady=3)
        
        ttk.Label(autocad_frame, text="또는 AutoCAD에서:", foreground="gray").pack(pady=2)
        
        instruction_text = """1. 원하는 객체 선택
2. 위 버튼 클릭"""
        ttk.Label(autocad_frame, text=instruction_text, foreground="blue").pack(pady=2)
        
        ttk.Separator(left_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)
        
        # 프로그램 내 선택
        program_frame = ttk.LabelFrame(left_frame, text="프로그램 선택")
        program_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(program_frame, text="레이어로 선택", 
                  command=self.select_by_layer_dialog, width=25).pack(pady=3)
        
        ttk.Button(program_frame, text="타입으로 선택", 
                  command=self.select_by_type_dialog, width=25).pack(pady=3)
        
        ttk.Button(program_frame, text="모든 객체 선택", 
                  command=self.select_all, width=25).pack(pady=3)
        
        ttk.Separator(left_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)
        
        # 선택 정보
        info_frame = ttk.Frame(left_frame)
        info_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.selection_count = ttk.Label(info_frame, text="선택된 객체: 0개", 
                                        font=('Arial', 10, 'bold'))
        self.selection_count.pack(pady=5)
        
        # 선택 목록
        list_frame = ttk.Frame(left_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.selection_listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, height=8)
        self.selection_listbox.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.selection_listbox.yview)
        
        # 선택 관리 버튼
        btn_frame = ttk.Frame(left_frame)
        btn_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(btn_frame, text="선택 초기화", 
                  command=self.clear_selection, width=12).pack(side=tk.LEFT, padx=2)
        
        ttk.Button(btn_frame, text="선택 반전", 
                  command=self.invert_selection, width=12).pack(side=tk.LEFT, padx=2)
    
    def create_center_panel(self, parent):
        """중앙 패널 - 계산 도구"""
        center_frame = ttk.LabelFrame(parent, text="📊 물량 계산")
        center_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=2)
        
        # 계산 카테고리 탭
        self.calc_notebook = ttk.Notebook(center_frame)
        self.calc_notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 구조 계산 탭
        self.create_structure_tab()
        
        # 마감 계산 탭
        self.create_finish_tab()
        
        # 사용자 정의 탭
        self.create_custom_tab()
        
        # 분석 탭
        self.create_analysis_tab()
    
    def create_structure_tab(self):
        """구조 계산 탭"""
        structure_frame = ttk.Frame(self.calc_notebook)
        self.calc_notebook.add(structure_frame, text="🏗️ 구조")
        
        # 계산 버튼들
        buttons = [
            ("장선/보 계산", self.calc_joists, "선택한 선을 장선으로 계산\n길이 → 목재 체적"),
            ("기둥 계산", self.calc_posts, "선택한 객체를 기둥으로 계산\n개수 → 콘크리트"),
            ("슬래브 계산", self.calc_slab, "선택한 면적을 슬래브로 계산\n면적 → 콘크리트"),
            ("벽체 계산", self.calc_wall, "선택한 선을 벽체로 계산\n길이 → 면적 → 체적"),
            ("철근 계산", self.calc_rebar, "선택한 선을 철근으로 계산\n길이 → 중량")
        ]
        
        for i, (text, command, tooltip) in enumerate(buttons):
            frame = ttk.Frame(structure_frame)
            frame.pack(fill=tk.X, padx=10, pady=5)
            
            btn = ttk.Button(frame, text=text, command=command, width=20)
            btn.pack(side=tk.LEFT, padx=5)
            
            ttk.Label(frame, text=tooltip, foreground="gray").pack(side=tk.LEFT, padx=10)
    
    def create_finish_tab(self):
        """마감 계산 탭"""
        finish_frame = ttk.Frame(self.calc_notebook)
        self.calc_notebook.add(finish_frame, text="🎨 마감")
        
        buttons = [
            ("데크보드 계산", self.calc_deck, "선택한 면적을 데크로 계산"),
            ("타일 계산", self.calc_tile, "선택한 면적을 타일로 계산"),
            ("페인트 계산", self.calc_paint, "선택한 면적을 도장으로 계산"),
            ("방수 계산", self.calc_waterproof, "선택한 면적을 방수로 계산")
        ]
        
        for i, (text, command, tooltip) in enumerate(buttons):
            frame = ttk.Frame(finish_frame)
            frame.pack(fill=tk.X, padx=10, pady=5)
            
            btn = ttk.Button(frame, text=text, command=command, width=20)
            btn.pack(side=tk.LEFT, padx=5)
            
            ttk.Label(frame, text=tooltip, foreground="gray").pack(side=tk.LEFT, padx=10)
    
    def create_custom_tab(self):
        """사용자 정의 탭"""
        custom_frame = ttk.Frame(self.calc_notebook)
        self.calc_notebook.add(custom_frame, text="⚙️ 사용자 정의")
        
        # 기본 측정
        basic_frame = ttk.LabelFrame(custom_frame, text="기본 측정")
        basic_frame.pack(fill=tk.X, padx=10, pady=10)
        
        btn_frame = ttk.Frame(basic_frame)
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="길이 측정", command=self.measure_length, width=15).grid(row=0, column=0, padx=5, pady=3)
        ttk.Button(btn_frame, text="면적 측정", command=self.measure_area, width=15).grid(row=0, column=1, padx=5, pady=3)
        ttk.Button(btn_frame, text="개수 집계", command=self.count_objects, width=15).grid(row=0, column=2, padx=5, pady=3)
        
        # 계산식
        formula_frame = ttk.LabelFrame(custom_frame, text="계산식")
        formula_frame.pack(fill=tk.X, padx=10, pady=10)
        
        input_frame = ttk.Frame(formula_frame)
        input_frame.pack(pady=10)
        
        ttk.Label(input_frame, text="계산식:").pack(side=tk.LEFT, padx=5)
        self.formula_entry = ttk.Entry(input_frame, width=40)
        self.formula_entry.pack(side=tk.LEFT, padx=5)
        self.formula_entry.insert(0, "L * 1.5 + A * 0.2")
        
        ttk.Button(input_frame, text="계산", command=self.custom_calc).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(formula_frame, text="변수: L=길이, A=면적, N=개수", 
                 foreground="blue").pack(pady=5)
        
        # 단위 설정
        unit_frame = ttk.LabelFrame(custom_frame, text="단위 설정")
        unit_frame.pack(fill=tk.X, padx=10, pady=10)
        
        unit_grid = ttk.Frame(unit_frame)
        unit_grid.pack(pady=10)
        
        ttk.Label(unit_grid, text="도면 단위:").grid(row=0, column=0, padx=5, pady=3)
        self.unit_var = tk.StringVar(value="mm")
        ttk.Combobox(unit_grid, textvariable=self.unit_var, 
                    values=["mm", "cm", "m"], width=10).grid(row=0, column=1, padx=5, pady=3)
        
        ttk.Label(unit_grid, text="축척:").grid(row=0, column=2, padx=5, pady=3)
        self.scale_var = tk.StringVar(value="1")
        ttk.Entry(unit_grid, textvariable=self.scale_var, width=10).grid(row=0, column=3, padx=5, pady=3)
        
        ttk.Label(unit_grid, text="할증률(%):").grid(row=1, column=0, padx=5, pady=3)
        self.waste_var = tk.StringVar(value="5")
        ttk.Spinbox(unit_grid, from_=0, to=50, textvariable=self.waste_var, 
                   width=10).grid(row=1, column=1, padx=5, pady=3)
    
    def create_analysis_tab(self):
        """분석 탭"""
        analysis_frame = ttk.Frame(self.calc_notebook)
        self.calc_notebook.add(analysis_frame, text="📈 분석")
        
        # 분석 결과 텍스트
        self.analysis_text = scrolledtext.ScrolledText(analysis_frame, height=20, width=50)
        self.analysis_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 분석 버튼
        btn_frame = ttk.Frame(analysis_frame)
        btn_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Button(btn_frame, text="전체 분석", command=self.analyze_all, width=15).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="선택 분석", command=self.analyze_selection, width=15).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="레이어 분석", command=self.analyze_layers, width=15).pack(side=tk.LEFT, padx=5)
    
    def create_right_panel(self, parent):
        """오른쪽 패널 - 결과"""
        right_frame = ttk.LabelFrame(parent, text="📋 결과", width=350)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=2)
        right_frame.pack_propagate(False)
        
        # 결과 텍스트
        self.result_text = scrolledtext.ScrolledText(right_frame, height=25, width=40)
        self.result_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 내보내기 버튼
        export_frame = ttk.Frame(right_frame)
        export_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(export_frame, text="📊 Excel", command=self.export_excel, width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(export_frame, text="📄 JSON", command=self.export_json, width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(export_frame, text="📋 복사", command=self.copy_result, width=10).pack(side=tk.LEFT, padx=2)
    
    def create_statusbar(self):
        """하단 상태바"""
        statusbar = ttk.Frame(self.root)
        statusbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.status_label = ttk.Label(statusbar, text="준비", relief=tk.SUNKEN, anchor=tk.W)
        self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        self.progress = ttk.Progressbar(statusbar, length=100, mode='indeterminate')
        self.progress.pack(side=tk.RIGHT, padx=5)
    
    # === AutoCAD 연결 메서드 ===
    
    def connect_autocad(self):
        """AutoCAD 연결 (개선된 버전)"""
        try:
            self.update_status("AutoCAD 연결 중...")
            self.progress.start()
            
            # AutoCAD 연결
            self.acad = win32com.client.Dispatch("AutoCAD.Application")
            
            # 문서 확인
            if self.acad.Documents.Count == 0:
                self.progress.stop()
                messagebox.showwarning("경고", "열린 도면이 없습니다.\nAutoCAD에서 도면을 열어주세요.")
                return
            
            self.doc = self.acad.ActiveDocument
            self.model = self.doc.ModelSpace
            
            # 연결 성공
            self.connection_status.config(text=f"🟢 연결됨: {self.doc.Name}", foreground="green")
            self.connect_btn.config(text="✓ 연결됨")
            
            self.update_status(f"AutoCAD 연결 성공: {self.doc.Name}")
            self.progress.stop()
            
            # 자동 분석
            self.analyze_all()
            
            messagebox.showinfo("성공", f"AutoCAD와 연결되었습니다!\n도면: {self.doc.Name}")
            
        except Exception as e:
            self.progress.stop()
            self.connection_status.config(text="🔴 연결 실패", foreground="red")
            messagebox.showerror("오류", f"AutoCAD 연결 실패:\n{str(e)}")
            self.update_status("AutoCAD 연결 실패")
    
    def check_autocad_status(self):
        """AutoCAD 실행 상태 확인"""
        try:
            acad = win32com.client.Dispatch("AutoCAD.Application")
            self.update_status("AutoCAD 실행 중 (연결 필요)")
        except:
            self.update_status("AutoCAD 미실행")
    
    # === 선택 메서드 (개선된 버전) ===
    
    def get_current_selection(self):
        """현재 AutoCAD 선택 세트 가져오기 (안정적)"""
        if not self.acad:
            messagebox.showwarning("경고", "먼저 AutoCAD를 연결하세요.")
            return
        
        try:
            self.update_status("선택 세트 가져오는 중...")
            
            # AutoCAD 활성화
            self.acad.Visible = True
            
            # 기존 선택 초기화
            self.selected_objects = []
            self.selection_listbox.delete(0, tk.END)
            
            # PickFirst 선택 세트 가져오기
            selection_set = self.doc.PickfirstSelectionSet
            
            if selection_set.Count == 0:
                # 선택된 것이 없으면 사용자에게 선택 요청
                messagebox.showinfo("안내", 
                    "AutoCAD에서 객체를 선택한 후\n다시 이 버튼을 클릭하세요.")
                
                # AutoCAD로 포커스 이동
                self.acad.Visible = True
                return
            
            # 선택된 객체 수집
            count = 0
            object_types = {}
            
            for obj in selection_set:
                self.selected_objects.append(obj)
                obj_type = obj.ObjectName
                
                # 타입별 카운트
                if obj_type not in object_types:
                    object_types[obj_type] = 0
                object_types[obj_type] += 1
                count += 1
            
            # 선택 목록 업데이트
            for obj_type, type_count in object_types.items():
                self.selection_listbox.insert(tk.END, f"{obj_type}: {type_count}개")
            
            # 선택 정보 업데이트
            self.selection_count.config(text=f"선택된 객체: {count}개")
            self.update_status(f"{count}개 객체 선택됨")
            
            # 결과 메시지
            if count > 0:
                messagebox.showinfo("성공", f"{count}개 객체를 가져왔습니다.")
            
        except Exception as e:
            messagebox.showerror("오류", f"선택 가져오기 실패:\n{str(e)}")
            self.update_status("선택 실패")
    
    def select_by_layer_dialog(self):
        """레이어 선택 대화상자"""
        if not self.acad:
            messagebox.showwarning("경고", "먼저 AutoCAD를 연결하세요.")
            return
        
        # 레이어 목록 가져오기
        layers = []
        for layer in self.doc.Layers:
            layers.append(layer.Name)
        
        # 대화상자 생성
        dialog = tk.Toplevel(self.root)
        dialog.title("레이어 선택")
        dialog.geometry("400x500")
        
        ttk.Label(dialog, text="선택할 레이어를 체크하세요:").pack(pady=10)
        
        # 체크박스 프레임
        frame = ttk.Frame(dialog)
        frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # 스크롤바
        scrollbar = ttk.Scrollbar(frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 리스트박스
        listbox = tk.Listbox(frame, selectmode=tk.MULTIPLE, yscrollcommand=scrollbar.set)
        listbox.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=listbox.yview)
        
        # 레이어 추가
        for layer in layers:
            listbox.insert(tk.END, layer)
        
        # 버튼
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=10)
        
        def apply_selection():
            selected_indices = listbox.curselection()
            selected_layers = [layers[i] for i in selected_indices]
            
            if selected_layers:
                self.select_by_layers(selected_layers)
                dialog.destroy()
            else:
                messagebox.showwarning("경고", "레이어를 선택하세요.")
        
        ttk.Button(btn_frame, text="선택", command=apply_selection).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="취소", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
    
    def select_by_layers(self, layer_names):
        """레이어로 객체 선택"""
        try:
            self.update_status("레이어로 선택 중...")
            self.selected_objects = []
            self.selection_listbox.delete(0, tk.END)
            
            count = 0
            object_types = {}
            
            for obj in self.model:
                if obj.Layer in layer_names:
                    self.selected_objects.append(obj)
                    
                    obj_type = obj.ObjectName
                    if obj_type not in object_types:
                        object_types[obj_type] = 0
                    object_types[obj_type] += 1
                    count += 1
            
            # 선택 목록 업데이트
            for obj_type, type_count in object_types.items():
                self.selection_listbox.insert(tk.END, f"{obj_type}: {type_count}개")
            
            self.selection_count.config(text=f"선택된 객체: {count}개")
            self.update_status(f"{layer_names} 레이어에서 {count}개 선택")
            
        except Exception as e:
            messagebox.showerror("오류", f"레이어 선택 실패:\n{str(e)}")
    
    def select_by_type_dialog(self):
        """타입으로 선택 대화상자"""
        if not self.acad:
            messagebox.showwarning("경고", "먼저 AutoCAD를 연결하세요.")
            return
        
        # 타입 목록 수집
        types = set()
        for obj in self.model:
            types.add(obj.ObjectName)
        
        types = sorted(list(types))
        
        # 대화상자
        dialog = tk.Toplevel(self.root)
        dialog.title("객체 타입 선택")
        dialog.geometry("400x400")
        
        ttk.Label(dialog, text="선택할 객체 타입:").pack(pady=10)
        
        # 리스트박스
        frame = ttk.Frame(dialog)
        frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        scrollbar = ttk.Scrollbar(frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        listbox = tk.Listbox(frame, selectmode=tk.MULTIPLE, yscrollcommand=scrollbar.set)
        listbox.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=listbox.yview)
        
        for obj_type in types:
            listbox.insert(tk.END, obj_type)
        
        def apply_selection():
            selected_indices = listbox.curselection()
            selected_types = [types[i] for i in selected_indices]
            
            if selected_types:
                self.select_by_types(selected_types)
                dialog.destroy()
        
        ttk.Button(dialog, text="선택", command=apply_selection).pack(pady=10)
    
    def select_by_types(self, type_names):
        """타입으로 객체 선택"""
        try:
            self.update_status("타입으로 선택 중...")
            self.selected_objects = []
            self.selection_listbox.delete(0, tk.END)
            
            count = 0
            for obj in self.model:
                if obj.ObjectName in type_names:
                    self.selected_objects.append(obj)
                    count += 1
            
            for type_name in type_names:
                type_count = sum(1 for obj in self.selected_objects if obj.ObjectName == type_name)
                if type_count > 0:
                    self.selection_listbox.insert(tk.END, f"{type_name}: {type_count}개")
            
            self.selection_count.config(text=f"선택된 객체: {count}개")
            self.update_status(f"{count}개 객체 선택됨")
            
        except Exception as e:
            messagebox.showerror("오류", f"타입 선택 실패:\n{str(e)}")
    
    def select_all(self):
        """모든 객체 선택"""
        if not self.acad:
            messagebox.showwarning("경고", "먼저 AutoCAD를 연결하세요.")
            return
        
        try:
            self.selected_objects = []
            self.selection_listbox.delete(0, tk.END)
            
            object_types = {}
            for obj in self.model:
                self.selected_objects.append(obj)
                
                obj_type = obj.ObjectName
                if obj_type not in object_types:
                    object_types[obj_type] = 0
                object_types[obj_type] += 1
            
            for obj_type, count in object_types.items():
                self.selection_listbox.insert(tk.END, f"{obj_type}: {count}개")
            
            total = len(self.selected_objects)
            self.selection_count.config(text=f"선택된 객체: {total}개")
            self.update_status(f"전체 {total}개 객체 선택")
            
        except Exception as e:
            messagebox.showerror("오류", f"전체 선택 실패:\n{str(e)}")
    
    def clear_selection(self):
        """선택 초기화"""
        self.selected_objects = []
        self.selection_listbox.delete(0, tk.END)
        self.selection_count.config(text="선택된 객체: 0개")
        self.update_status("선택 초기화됨")
    
    def invert_selection(self):
        """선택 반전"""
        if not self.acad:
            return
        
        try:
            current_selection = set(self.selected_objects)
            self.selected_objects = []
            
            for obj in self.model:
                if obj not in current_selection:
                    self.selected_objects.append(obj)
            
            self.selection_count.config(text=f"선택된 객체: {len(self.selected_objects)}개")
            self.update_status("선택 반전됨")
            
        except Exception as e:
            messagebox.showerror("오류", f"선택 반전 실패:\n{str(e)}")
    
    # === 계산 메서드 ===
    
    def calc_joists(self):
        """장선/보 계산"""
        if not self.selected_objects:
            messagebox.showwarning("경고", "먼저 객체를 선택하세요.")
            return
        
        try:
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
            
            if count == 0:
                messagebox.showinfo("안내", "선형 객체가 없습니다.")
                return
            
            # 단위 변환
            if self.unit_var.get() == "mm":
                length_m = total_length / 1000
            else:
                length_m = total_length
            
            # 체적 계산 (2x10 기준)
            volume = length_m * 0.038 * 0.235
            waste = float(self.waste_var.get()) / 100
            
            result = f"=== 장선/보 계산 결과 ===\n\n"
            result += f"선택 객체: {count}개\n"
            result += f"총 길이: {length_m:.2f}m\n"
            result += f"목재 규격: 2x10 (38x235mm)\n"
            result += f"목재 체적: {volume:.3f}m³\n"
            result += f"할증 {self.waste_var.get()}% 적용: {volume*(1+waste):.3f}m³\n"
            
            self.display_result(result)
            self.calculation_results["장선"] = {"길이": length_m, "체적": volume}
            
        except Exception as e:
            messagebox.showerror("오류", f"계산 실패:\n{str(e)}")
    
    def calc_deck(self):
        """데크보드 계산"""
        if not self.selected_objects:
            messagebox.showwarning("경고", "먼저 객체를 선택하세요.")
            return
        
        try:
            total_area = 0
            count = 0
            
            for obj in self.selected_objects:
                try:
                    if hasattr(obj, 'Area'):
                        area = obj.Area
                        if area > 0:
                            total_area += area
                            count += 1
                except:
                    pass
            
            if count == 0:
                messagebox.showinfo("안내", "면적 객체가 없습니다.")
                return
            
            # 단위 변환
            if self.unit_var.get() == "mm":
                area_m2 = total_area / 1000000
            else:
                area_m2 = total_area
            
            waste = float(self.waste_var.get()) / 100
            deck_area = area_m2 * (1 + waste)
            screws = int(area_m2 * 8)
            
            result = f"=== 데크보드 계산 결과 ===\n\n"
            result += f"선택 객체: {count}개\n"
            result += f"데크 면적: {area_m2:.2f}m²\n"
            result += f"필요 보드: {deck_area:.2f}m²\n"
            result += f"(손실률 {self.waste_var.get()}% 포함)\n"
            result += f"데크 스크류: {screws}개\n"
            result += f"(m²당 8개 기준)\n"
            
            self.display_result(result)
            self.calculation_results["데크"] = {"면적": area_m2, "필요량": deck_area}
            
        except Exception as e:
            messagebox.showerror("오류", f"계산 실패:\n{str(e)}")
    
    def calc_posts(self):
        """기둥 계산"""
        if not self.selected_objects:
            messagebox.showwarning("경고", "먼저 객체를 선택하세요.")
            return
        
        count = len(self.selected_objects)
        foundation = count * 0.096  # 400x400x600 기준
        
        result = f"=== 기둥 계산 결과 ===\n\n"
        result += f"기둥 개수: {count}개\n"
        result += f"기초 크기: 400x400x600mm\n"
        result += f"기초당 콘크리트: 0.096m³\n"
        result += f"총 콘크리트: {foundation:.3f}m³\n"
        result += f"할증 5% 적용: {foundation*1.05:.3f}m³\n"
        
        self.display_result(result)
        self.calculation_results["기둥"] = {"개수": count, "콘크리트": foundation}
    
    def calc_slab(self):
        """슬래브 계산"""
        messagebox.showinfo("안내", "슬래브 계산 기능 개발 중")
    
    def calc_wall(self):
        """벽체 계산"""
        messagebox.showinfo("안내", "벽체 계산 기능 개발 중")
    
    def calc_rebar(self):
        """철근 계산"""
        messagebox.showinfo("안내", "철근 계산 기능 개발 중")
    
    def calc_tile(self):
        """타일 계산"""
        messagebox.showinfo("안내", "타일 계산 기능 개발 중")
    
    def calc_paint(self):
        """페인트 계산"""
        messagebox.showinfo("안내", "페인트 계산 기능 개발 중")
    
    def calc_waterproof(self):
        """방수 계산"""
        messagebox.showinfo("안내", "방수 계산 기능 개발 중")
    
    def measure_length(self):
        """길이 측정"""
        if not self.selected_objects:
            messagebox.showwarning("경고", "먼저 객체를 선택하세요.")
            return
        
        total_length = 0
        count = 0
        
        for obj in self.selected_objects:
            try:
                if hasattr(obj, 'Length'):
                    total_length += obj.Length
                    count += 1
                elif "Line" in obj.ObjectName:
                    start = obj.StartPoint
                    end = obj.EndPoint
                    length = math.sqrt((end[0]-start[0])**2 + (end[1]-start[1])**2)
                    total_length += length
                    count += 1
            except:
                pass
        
        if self.unit_var.get() == "mm":
            length_m = total_length / 1000
        else:
            length_m = total_length
        
        result = f"=== 길이 측정 ===\n\n"
        result += f"측정 객체: {count}개\n"
        result += f"총 길이: {total_length:.2f}{self.unit_var.get()}\n"
        result += f"미터 환산: {length_m:.2f}m\n"
        
        self.display_result(result)
    
    def measure_area(self):
        """면적 측정"""
        if not self.selected_objects:
            messagebox.showwarning("경고", "먼저 객체를 선택하세요.")
            return
        
        total_area = 0
        count = 0
        
        for obj in self.selected_objects:
            try:
                if hasattr(obj, 'Area'):
                    area = obj.Area
                    if area > 0:
                        total_area += area
                        count += 1
            except:
                pass
        
        if self.unit_var.get() == "mm":
            area_m2 = total_area / 1000000
        else:
            area_m2 = total_area
        
        result = f"=== 면적 측정 ===\n\n"
        result += f"측정 객체: {count}개\n"
        result += f"총 면적: {total_area:.2f}{self.unit_var.get()}²\n"
        result += f"평방미터 환산: {area_m2:.2f}m²\n"
        result += f"평 환산: {area_m2/3.3:.2f}평\n"
        
        self.display_result(result)
    
    def count_objects(self):
        """개수 집계"""
        if not self.selected_objects:
            messagebox.showwarning("경고", "먼저 객체를 선택하세요.")
            return
        
        types = {}
        for obj in self.selected_objects:
            obj_type = obj.ObjectName
            if obj_type not in types:
                types[obj_type] = 0
            types[obj_type] += 1
        
        result = f"=== 개수 집계 ===\n\n"
        result += f"총 객체 수: {len(self.selected_objects)}개\n\n"
        result += "타입별 분류:\n"
        
        for obj_type, count in sorted(types.items(), key=lambda x: x[1], reverse=True):
            result += f"  {obj_type}: {count}개\n"
        
        self.display_result(result)
    
    def custom_calc(self):
        """사용자 정의 계산"""
        if not self.selected_objects:
            messagebox.showwarning("경고", "먼저 객체를 선택하세요.")
            return
        
        # 변수 계산
        L = 0  # 길이
        A = 0  # 면적
        N = len(self.selected_objects)  # 개수
        
        for obj in self.selected_objects:
            try:
                if hasattr(obj, 'Length'):
                    L += obj.Length
                if hasattr(obj, 'Area'):
                    A += obj.Area
            except:
                pass
        
        # 단위 변환
        if self.unit_var.get() == "mm":
            L = L / 1000
            A = A / 1000000
        
        formula = self.formula_entry.get()
        
        try:
            result_value = eval(formula)
            
            result = f"=== 사용자 정의 계산 ===\n\n"
            result += f"계산식: {formula}\n\n"
            result += f"변수값:\n"
            result += f"  L (길이) = {L:.2f}m\n"
            result += f"  A (면적) = {A:.2f}m²\n"
            result += f"  N (개수) = {N}개\n\n"
            result += f"계산 결과: {result_value:.3f}\n"
            
            self.display_result(result)
            
        except Exception as e:
            messagebox.showerror("오류", f"계산 오류:\n{str(e)}")
    
    # === 분석 메서드 ===
    
    def analyze_all(self):
        """전체 분석"""
        if not self.acad:
            messagebox.showwarning("경고", "먼저 AutoCAD를 연결하세요.")
            return
        
        try:
            total = self.model.Count
            types = {}
            layers = {}
            
            for obj in self.model:
                # 타입별
                obj_type = obj.ObjectName
                if obj_type not in types:
                    types[obj_type] = 0
                types[obj_type] += 1
                
                # 레이어별
                layer = obj.Layer
                if layer not in layers:
                    layers[layer] = 0
                layers[layer] += 1
            
            result = f"=== 전체 도면 분석 ===\n\n"
            result += f"도면명: {self.doc.Name}\n"
            result += f"총 객체 수: {total}개\n"
            result += f"레이어 수: {len(layers)}개\n"
            result += f"객체 타입 수: {len(types)}개\n\n"
            
            result += "주요 객체 타입 (상위 10개):\n"
            sorted_types = sorted(types.items(), key=lambda x: x[1], reverse=True)[:10]
            for obj_type, count in sorted_types:
                result += f"  {obj_type}: {count}개 ({count/total*100:.1f}%)\n"
            
            result += "\n주요 레이어 (상위 10개):\n"
            sorted_layers = sorted(layers.items(), key=lambda x: x[1], reverse=True)[:10]
            for layer, count in sorted_layers:
                result += f"  {layer}: {count}개\n"
            
            self.analysis_text.delete(1.0, tk.END)
            self.analysis_text.insert(1.0, result)
            
            # 분석 탭으로 전환
            self.calc_notebook.select(3)
            
        except Exception as e:
            messagebox.showerror("오류", f"분석 실패:\n{str(e)}")
    
    def analyze_selection(self):
        """선택 분석"""
        if not self.selected_objects:
            messagebox.showwarning("경고", "먼저 객체를 선택하세요.")
            return
        
        types = {}
        layers = {}
        total_length = 0
        total_area = 0
        
        for obj in self.selected_objects:
            # 타입별
            obj_type = obj.ObjectName
            if obj_type not in types:
                types[obj_type] = 0
            types[obj_type] += 1
            
            # 레이어별
            try:
                layer = obj.Layer
                if layer not in layers:
                    layers[layer] = 0
                layers[layer] += 1
            except:
                pass
            
            # 측정
            try:
                if hasattr(obj, 'Length'):
                    total_length += obj.Length
                if hasattr(obj, 'Area'):
                    total_area += obj.Area
            except:
                pass
        
        result = f"=== 선택 객체 분석 ===\n\n"
        result += f"선택 객체 수: {len(self.selected_objects)}개\n\n"
        
        result += "타입별 분포:\n"
        for obj_type, count in sorted(types.items(), key=lambda x: x[1], reverse=True):
            result += f"  {obj_type}: {count}개\n"
        
        result += f"\n레이어별 분포:\n"
        for layer, count in sorted(layers.items(), key=lambda x: x[1], reverse=True):
            result += f"  {layer}: {count}개\n"
        
        if total_length > 0:
            result += f"\n총 길이: {total_length:.2f} {self.unit_var.get()}\n"
        if total_area > 0:
            result += f"총 면적: {total_area:.2f} {self.unit_var.get()}²\n"
        
        self.analysis_text.delete(1.0, tk.END)
        self.analysis_text.insert(1.0, result)
        
        self.calc_notebook.select(3)
    
    def analyze_layers(self):
        """레이어 분석"""
        if not self.acad:
            messagebox.showwarning("경고", "먼저 AutoCAD를 연결하세요.")
            return
        
        layer_info = {}
        
        for layer in self.doc.Layers:
            layer_info[layer.Name] = {
                'color': layer.color,
                'on': not layer.LayerOn,
                'locked': layer.Lock,
                'objects': 0
            }
        
        for obj in self.model:
            layer = obj.Layer
            if layer in layer_info:
                layer_info[layer]['objects'] += 1
        
        result = f"=== 레이어 분석 ===\n\n"
        result += f"총 레이어 수: {len(layer_info)}개\n\n"
        
        sorted_layers = sorted(layer_info.items(), key=lambda x: x[1]['objects'], reverse=True)
        
        for layer_name, info in sorted_layers[:20]:
            result += f"{layer_name}:\n"
            result += f"  객체 수: {info['objects']}개\n"
            result += f"  색상: {info['color']}\n"
            result += f"  상태: {'켜짐' if info['on'] else '꺼짐'}\n\n"
        
        self.analysis_text.delete(1.0, tk.END)
        self.analysis_text.insert(1.0, result)
        
        self.calc_notebook.select(3)
    
    # === 유틸리티 메서드 ===
    
    def display_result(self, text):
        """결과 표시"""
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(1.0, text)
        
        # 타임스탬프 추가
        timestamp = datetime.now().strftime("\n\n시간: %Y-%m-%d %H:%M:%S")
        self.result_text.insert(tk.END, timestamp)
    
    def update_status(self, message):
        """상태 업데이트"""
        self.status_label.config(text=message)
        self.root.update_idletasks()
    
    def open_dxf_file(self):
        """DXF 파일 열기"""
        messagebox.showinfo("안내", "DXF 파일 열기는 AutoCAD 연결 후 사용 가능합니다.")
    
    def save_results(self):
        """결과 저장"""
        if not self.calculation_results:
            messagebox.showwarning("경고", "저장할 결과가 없습니다.")
            return
        
        self.export_json()
    
    def export_excel(self):
        """Excel 내보내기"""
        if not self.calculation_results:
            messagebox.showwarning("경고", "저장할 결과가 없습니다.")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        
        if filename:
            try:
                data = []
                for category, values in self.calculation_results.items():
                    for key, value in values.items():
                        data.append({"카테고리": category, "항목": key, "값": value})
                
                df = pd.DataFrame(data)
                df.to_excel(filename, index=False)
                
                messagebox.showinfo("성공", f"저장 완료:\n{filename}")
                
            except Exception as e:
                messagebox.showerror("오류", f"저장 실패:\n{str(e)}")
    
    def export_json(self):
        """JSON 내보내기"""
        filename = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")]
        )
        
        if filename:
            try:
                output = {
                    "도면": self.doc.Name if self.doc else "Unknown",
                    "일시": datetime.now().isoformat(),
                    "결과": self.calculation_results
                }
                
                with open(filename, 'w', encoding='utf-8') as f:
                    json.dump(output, f, ensure_ascii=False, indent=2)
                
                messagebox.showinfo("성공", f"저장 완료:\n{filename}")
                
            except Exception as e:
                messagebox.showerror("오류", f"저장 실패:\n{str(e)}")
    
    def copy_result(self):
        """결과 복사"""
        try:
            text = self.result_text.get(1.0, tk.END)
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            messagebox.showinfo("성공", "결과가 클립보드에 복사되었습니다.")
        except:
            pass
    
    def show_help(self):
        """도움말"""
        help_text = """=== CADPro v2 사용법 ===

1. AutoCAD 연결
   - AutoCAD 실행 후 도면 열기
   - [AutoCAD 연결] 버튼 클릭

2. 객체 선택 (권장 방법)
   - AutoCAD에서 원하는 객체 선택
   - [현재 선택 세트 가져오기] 클릭

3. 물량 계산
   - 원하는 계산 버튼 클릭
   - 결과 확인

4. 결과 저장
   - Excel 또는 JSON 형식 저장
   - 복사 버튼으로 클립보드 복사

팁:
- AutoCAD에서 먼저 선택 후 가져오기가 가장 안정적
- 레이어/타입 선택은 전체 도면에서 필터링
- 계산 전 단위와 축척 확인
"""
        messagebox.showinfo("도움말", help_text)
    
    def run(self):
        """프로그램 실행"""
        self.root.mainloop()


def main():
    app = ImprovedCADProGUI()
    app.run()


if __name__ == "__main__":
    main()