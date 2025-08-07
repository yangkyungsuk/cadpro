"""
CAD 물량 산출 프로그램 - 전문 GUI 버전
사용자가 CAD 파일을 열고 직관적으로 작업할 수 있는 인터페이스
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
import threading
import pandas as pd


class CADProGUI:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("CADPro - 물량 산출 프로그램")
        self.root.geometry("1200x700")
        
        # 아이콘 색상 테마
        self.colors = {
            'primary': '#2196F3',
            'success': '#4CAF50',
            'warning': '#FF9800',
            'danger': '#F44336',
            'dark': '#424242',
            'light': '#F5F5F5'
        }
        
        # AutoCAD 연결 상태
        self.acad = None
        self.doc = None
        self.current_file = None
        self.selected_objects = []
        self.calculation_results = {}
        
        # GUI 구성
        self.setup_gui()
        
        # 초기 상태 확인
        self.check_autocad_status()
    
    def setup_gui(self):
        """GUI 레이아웃 구성"""
        
        # 스타일 설정
        style = ttk.Style()
        style.theme_use('clam')
        
        # 메뉴바
        self.create_menubar()
        
        # 툴바
        self.create_toolbar()
        
        # 메인 컨테이너
        main_container = ttk.Frame(self.root)
        main_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 왼쪽 패널 (파일 정보 & 선택 도구)
        self.create_left_panel(main_container)
        
        # 중앙 패널 (작업 영역)
        self.create_center_panel(main_container)
        
        # 오른쪽 패널 (계산 결과)
        self.create_right_panel(main_container)
        
        # 하단 상태바
        self.create_statusbar()
    
    def create_menubar(self):
        """메뉴바 생성"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # 파일 메뉴
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="파일", menu=file_menu)
        file_menu.add_command(label="DXF 파일 열기", command=self.open_dxf_file)
        file_menu.add_command(label="AutoCAD 연결", command=self.connect_autocad)
        file_menu.add_separator()
        file_menu.add_command(label="결과 저장", command=self.save_results)
        file_menu.add_separator()
        file_menu.add_command(label="종료", command=self.root.quit)
        
        # 도구 메뉴
        tools_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="도구", menu=tools_menu)
        tools_menu.add_command(label="DWG→DXF 변환 안내", command=self.show_conversion_guide)
        tools_menu.add_command(label="설정", command=self.show_settings)
        
        # 도움말 메뉴
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="도움말", menu=help_menu)
        help_menu.add_command(label="사용법", command=self.show_help)
        help_menu.add_command(label="정보", command=self.show_about)
    
    def create_toolbar(self):
        """툴바 생성"""
        toolbar = ttk.Frame(self.root)
        toolbar.pack(side=tk.TOP, fill=tk.X, padx=5, pady=2)
        
        # 툴바 버튼들
        ttk.Button(toolbar, text="📁 파일 열기", command=self.open_dxf_file, width=12).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="🔗 AutoCAD 연결", command=self.connect_autocad, width=15).pack(side=tk.LEFT, padx=2)
        ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, padx=5, fill=tk.Y)
        ttk.Button(toolbar, text="🔄 새로고침", command=self.refresh_data, width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="🗑️ 초기화", command=self.clear_all, width=10).pack(side=tk.LEFT, padx=2)
    
    def create_left_panel(self, parent):
        """왼쪽 패널 - 파일 정보 & 선택 도구"""
        left_frame = ttk.LabelFrame(parent, text="📋 도면 정보", width=300)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, padx=2, pady=2)
        left_frame.pack_propagate(False)
        
        # 파일 정보
        info_frame = ttk.Frame(left_frame)
        info_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.file_info_label = ttk.Label(info_frame, text="파일: 없음")
        self.file_info_label.pack(anchor=tk.W)
        
        self.status_label = ttk.Label(info_frame, text="상태: 대기 중", foreground="gray")
        self.status_label.pack(anchor=tk.W)
        
        self.object_count_label = ttk.Label(info_frame, text="객체 수: 0")
        self.object_count_label.pack(anchor=tk.W)
        
        # 구분선
        ttk.Separator(left_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)
        
        # 객체 선택 섹션
        select_frame = ttk.LabelFrame(left_frame, text="🎯 객체 선택")
        select_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(select_frame, text="영역 선택", command=self.select_by_area, width=25).pack(pady=2)
        ttk.Button(select_frame, text="개별 선택", command=self.select_individually, width=25).pack(pady=2)
        ttk.Button(select_frame, text="레이어로 선택", command=self.select_by_layer, width=25).pack(pady=2)
        
        # 자동 선택 옵션
        auto_frame = ttk.Frame(select_frame)
        auto_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(auto_frame, text="자동 선택:").pack(anchor=tk.W)
        ttk.Button(auto_frame, text="모든 선", command=lambda: self.auto_select("lines"), width=11).pack(side=tk.LEFT, padx=1)
        ttk.Button(auto_frame, text="모든 폴리라인", command=lambda: self.auto_select("polylines"), width=11).pack(side=tk.LEFT, padx=1)
        
        # 선택 정보
        self.selection_info = ttk.Label(select_frame, text="선택된 객체: 0개", foreground="blue")
        self.selection_info.pack(pady=5)
        
        ttk.Button(select_frame, text="선택 초기화", command=self.clear_selection, width=25).pack(pady=2)
        
        # 레이어 목록
        ttk.Separator(left_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)
        
        layer_frame = ttk.LabelFrame(left_frame, text="📑 레이어 목록")
        layer_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 레이어 리스트박스
        scrollbar = ttk.Scrollbar(layer_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.layer_listbox = tk.Listbox(layer_frame, yscrollcommand=scrollbar.set, height=10)
        self.layer_listbox.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.layer_listbox.yview)
    
    def create_center_panel(self, parent):
        """중앙 패널 - 작업 영역"""
        center_frame = ttk.Frame(parent)
        center_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=2, pady=2)
        
        # 탭 위젯
        self.notebook = ttk.Notebook(center_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # 물량 계산 탭
        self.create_calculation_tab()
        
        # 분석 탭
        self.create_analysis_tab()
        
        # 로그 탭
        self.create_log_tab()
    
    def create_calculation_tab(self):
        """물량 계산 탭"""
        calc_frame = ttk.Frame(self.notebook)
        self.notebook.add(calc_frame, text="📊 물량 계산")
        
        # 계산 타입 선택
        type_frame = ttk.LabelFrame(calc_frame, text="계산 타입 선택")
        type_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # 계산 버튼들을 그리드로 배치
        calc_buttons = [
            ("🪵 장선/보", self.calc_joists, 0, 0),
            ("📐 데크보드", self.calc_deck, 0, 1),
            ("🏗️ 기둥", self.calc_posts, 0, 2),
            ("🔩 철근", self.calc_rebar, 1, 0),
            ("🚰 배관", self.calc_pipes, 1, 1),
            ("🧱 콘크리트", self.calc_concrete, 1, 2),
            ("📏 길이 계산", self.calc_length, 2, 0),
            ("📐 면적 계산", self.calc_area, 2, 1),
            ("🔢 개수 집계", self.calc_count, 2, 2)
        ]
        
        for text, command, row, col in calc_buttons:
            btn = ttk.Button(type_frame, text=text, command=command, width=15)
            btn.grid(row=row, column=col, padx=5, pady=5)
        
        # 계산 옵션
        option_frame = ttk.LabelFrame(calc_frame, text="계산 옵션")
        option_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # 단위 설정
        unit_frame = ttk.Frame(option_frame)
        unit_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(unit_frame, text="도면 단위:").pack(side=tk.LEFT, padx=5)
        self.unit_var = tk.StringVar(value="mm")
        unit_combo = ttk.Combobox(unit_frame, textvariable=self.unit_var, values=["mm", "cm", "m", "inch"], width=10)
        unit_combo.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(unit_frame, text="축척:").pack(side=tk.LEFT, padx=20)
        self.scale_var = tk.StringVar(value="1")
        scale_entry = ttk.Entry(unit_frame, textvariable=self.scale_var, width=10)
        scale_entry.pack(side=tk.LEFT, padx=5)
        
        # 할증률 설정
        waste_frame = ttk.Frame(option_frame)
        waste_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(waste_frame, text="할증률 (%):").pack(side=tk.LEFT, padx=5)
        self.waste_var = tk.StringVar(value="5")
        waste_spin = ttk.Spinbox(waste_frame, from_=0, to=50, textvariable=self.waste_var, width=10)
        waste_spin.pack(side=tk.LEFT, padx=5)
        
        # 사용자 정의 계산
        custom_frame = ttk.LabelFrame(calc_frame, text="사용자 정의 계산")
        custom_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        formula_frame = ttk.Frame(custom_frame)
        formula_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(formula_frame, text="계산식:").pack(side=tk.LEFT, padx=5)
        self.formula_entry = ttk.Entry(formula_frame, width=40)
        self.formula_entry.pack(side=tk.LEFT, padx=5)
        self.formula_entry.insert(0, "L * 1.5 + A * 0.2")
        
        ttk.Button(formula_frame, text="계산", command=self.custom_calculation).pack(side=tk.LEFT, padx=5)
        
        # 계산식 설명
        info_text = "L: 총 길이, A: 총 면적, N: 객체 개수"
        ttk.Label(custom_frame, text=info_text, foreground="gray").pack(padx=10, pady=2)
    
    def create_analysis_tab(self):
        """분석 탭"""
        analysis_frame = ttk.Frame(self.notebook)
        self.notebook.add(analysis_frame, text="📈 분석")
        
        # 통계 정보
        stats_frame = ttk.LabelFrame(analysis_frame, text="통계")
        stats_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.stats_text = scrolledtext.ScrolledText(stats_frame, height=20, width=60)
        self.stats_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 분석 버튼
        btn_frame = ttk.Frame(analysis_frame)
        btn_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Button(btn_frame, text="전체 분석", command=self.analyze_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="레이어별 분석", command=self.analyze_by_layer).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="타입별 분석", command=self.analyze_by_type).pack(side=tk.LEFT, padx=5)
    
    def create_log_tab(self):
        """로그 탭"""
        log_frame = ttk.Frame(self.notebook)
        self.notebook.add(log_frame, text="📝 로그")
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=20, width=60)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 로그 제어 버튼
        btn_frame = ttk.Frame(log_frame)
        btn_frame.pack(fill=tk.X, padx=5, pady=2)
        
        ttk.Button(btn_frame, text="로그 지우기", command=lambda: self.log_text.delete(1.0, tk.END)).pack(side=tk.RIGHT, padx=5)
    
    def create_right_panel(self, parent):
        """오른쪽 패널 - 계산 결과"""
        right_frame = ttk.LabelFrame(parent, text="📋 계산 결과", width=350)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=2, pady=2)
        right_frame.pack_propagate(False)
        
        # 결과 표시 영역
        self.result_text = scrolledtext.ScrolledText(right_frame, height=25, width=40)
        self.result_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 결과 내보내기 버튼
        export_frame = ttk.Frame(right_frame)
        export_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(export_frame, text="Excel 저장", command=self.export_excel, width=15).pack(side=tk.LEFT, padx=2)
        ttk.Button(export_frame, text="JSON 저장", command=self.export_json, width=15).pack(side=tk.LEFT, padx=2)
        
        if self.acad:
            ttk.Button(export_frame, text="CAD 테이블", command=self.export_to_cad, width=15).pack(side=tk.LEFT, padx=2)
    
    def create_statusbar(self):
        """상태바 생성"""
        self.statusbar = ttk.Frame(self.root)
        self.statusbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.status_text = ttk.Label(self.statusbar, text="준비", relief=tk.SUNKEN, anchor=tk.W)
        self.status_text.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        self.progress = ttk.Progressbar(self.statusbar, length=100, mode='indeterminate')
        self.progress.pack(side=tk.RIGHT, padx=5)
    
    # === 파일 작업 메서드 ===
    
    def open_dxf_file(self):
        """DXF 파일 열기"""
        filename = filedialog.askopenfilename(
            title="DXF 파일 선택",
            filetypes=[("DXF files", "*.dxf"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                self.log("DXF 파일 로딩 중...")
                self.progress.start()
                
                # DXF 파일 읽기
                self.doc = ezdxf.readfile(filename)
                self.current_file = filename
                
                # 파일 정보 업데이트
                self.file_info_label.config(text=f"파일: {os.path.basename(filename)}")
                self.status_label.config(text="상태: DXF 로드 완료", foreground="green")
                
                # 레이어 목록 업데이트
                self.update_layer_list()
                
                # 객체 수 업데이트
                msp = self.doc.modelspace()
                self.object_count_label.config(text=f"객체 수: {len(msp)}")
                
                self.log(f"DXF 파일 로드 완료: {os.path.basename(filename)}")
                self.progress.stop()
                
                # 자동 분석
                self.analyze_all()
                
            except Exception as e:
                self.progress.stop()
                messagebox.showerror("오류", f"파일 로드 실패: {str(e)}")
                self.log(f"오류: {str(e)}")
    
    def connect_autocad(self):
        """AutoCAD 연결"""
        try:
            self.log("AutoCAD 연결 시도 중...")
            self.progress.start()
            
            self.acad = win32com.client.Dispatch("AutoCAD.Application")
            self.doc = self.acad.ActiveDocument
            
            self.file_info_label.config(text=f"파일: {self.doc.Name}")
            self.status_label.config(text="상태: AutoCAD 연결됨", foreground="green")
            self.object_count_label.config(text=f"객체 수: {self.doc.ModelSpace.Count}")
            
            # 레이어 목록 업데이트
            self.update_layer_list_autocad()
            
            self.log(f"AutoCAD 연결 성공: {self.doc.Name}")
            self.progress.stop()
            
            messagebox.showinfo("성공", "AutoCAD와 연결되었습니다!")
            
        except Exception as e:
            self.progress.stop()
            messagebox.showerror("오류", f"AutoCAD 연결 실패: {str(e)}")
            self.log(f"AutoCAD 연결 실패: {str(e)}")
    
    def check_autocad_status(self):
        """AutoCAD 상태 확인"""
        try:
            acad = win32com.client.Dispatch("AutoCAD.Application")
            self.status_label.config(text="상태: AutoCAD 실행 중", foreground="blue")
        except:
            self.status_label.config(text="상태: AutoCAD 미실행", foreground="gray")
    
    # === 선택 메서드 ===
    
    def select_by_area(self):
        """영역 선택"""
        if not self.acad:
            messagebox.showwarning("경고", "먼저 AutoCAD를 연결하세요.")
            return
        
        try:
            self.log("영역 선택 모드...")
            
            # AutoCAD 활성화
            self.acad.Visible = True
            
            # 두 점 선택
            point1 = self.doc.Utility.GetPoint(None, "첫 번째 점: ")
            point2 = self.doc.Utility.GetCorner(point1, "두 번째 점: ")
            
            # 선택 세트 생성
            sel_set = self.doc.SelectionSets.Add(f"TempSet_{datetime.now().timestamp()}")
            
            filter_type = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, [0])
            filter_data = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, ["*"])
            
            sel_set.Select(0, point1, point2, filter_type, filter_data)
            
            # 선택된 객체 저장
            for obj in sel_set:
                self.selected_objects.append(obj)
            
            sel_set.Delete()
            
            self.selection_info.config(text=f"선택된 객체: {len(self.selected_objects)}개")
            self.log(f"{len(self.selected_objects)}개 객체 선택됨")
            
        except Exception as e:
            self.log(f"선택 오류: {str(e)}")
    
    def select_individually(self):
        """개별 선택"""
        if not self.acad:
            messagebox.showwarning("경고", "먼저 AutoCAD를 연결하세요.")
            return
        
        messagebox.showinfo("안내", "AutoCAD에서 객체를 하나씩 클릭하세요.\nESC로 종료합니다.")
        
        count = 0
        while True:
            try:
                result = self.doc.Utility.GetEntity(None, "객체 선택 (ESC로 종료): ")
                obj = result[0]
                self.selected_objects.append(obj)
                obj.Highlight(True)
                count += 1
            except:
                break
        
        self.selection_info.config(text=f"선택된 객체: {len(self.selected_objects)}개")
        self.log(f"{count}개 객체 개별 선택됨")
    
    def select_by_layer(self):
        """레이어로 선택"""
        if not self.layer_listbox.curselection():
            messagebox.showwarning("경고", "레이어를 먼저 선택하세요.")
            return
        
        selected_layers = [self.layer_listbox.get(i) for i in self.layer_listbox.curselection()]
        
        if self.acad:
            # AutoCAD 모드
            count = 0
            for obj in self.doc.ModelSpace:
                if obj.Layer in selected_layers:
                    self.selected_objects.append(obj)
                    count += 1
        else:
            # DXF 모드
            count = 0
            msp = self.doc.modelspace()
            for entity in msp:
                if entity.dxf.layer in selected_layers:
                    self.selected_objects.append(entity)
                    count += 1
        
        self.selection_info.config(text=f"선택된 객체: {len(self.selected_objects)}개")
        self.log(f"{selected_layers} 레이어에서 {count}개 객체 선택")
    
    def auto_select(self, obj_type):
        """자동 선택"""
        count = 0
        
        if self.acad:
            # AutoCAD 모드
            for obj in self.doc.ModelSpace:
                if obj_type == "lines" and "Line" in obj.ObjectName:
                    self.selected_objects.append(obj)
                    count += 1
                elif obj_type == "polylines" and "Polyline" in obj.ObjectName:
                    self.selected_objects.append(obj)
                    count += 1
        else:
            # DXF 모드
            msp = self.doc.modelspace()
            for entity in msp:
                if obj_type == "lines" and entity.dxftype() == "LINE":
                    self.selected_objects.append(entity)
                    count += 1
                elif obj_type == "polylines" and entity.dxftype() in ["LWPOLYLINE", "POLYLINE"]:
                    self.selected_objects.append(entity)
                    count += 1
        
        self.selection_info.config(text=f"선택된 객체: {len(self.selected_objects)}개")
        self.log(f"{obj_type} {count}개 자동 선택")
    
    def clear_selection(self):
        """선택 초기화"""
        if self.acad:
            for obj in self.selected_objects:
                try:
                    obj.Highlight(False)
                except:
                    pass
        
        self.selected_objects = []
        self.selection_info.config(text="선택된 객체: 0개")
        self.log("선택 초기화됨")
    
    # === 계산 메서드 ===
    
    def calc_joists(self):
        """장선/보 계산"""
        if not self.selected_objects:
            messagebox.showwarning("경고", "먼저 객체를 선택하세요.")
            return
        
        self.log("장선/보 계산 시작...")
        
        total_length = 0
        count = 0
        
        for obj in self.selected_objects:
            try:
                if self.acad:
                    # AutoCAD 객체
                    if "Line" in obj.ObjectName:
                        start = obj.StartPoint
                        end = obj.EndPoint
                        length = math.sqrt((end[0]-start[0])**2 + (end[1]-start[1])**2 + (end[2]-start[2])**2)
                        total_length += length
                        count += 1
                else:
                    # DXF 엔티티
                    if obj.dxftype() == "LINE":
                        start = obj.dxf.start
                        end = obj.dxf.end
                        length = start.distance(end)
                        total_length += length
                        count += 1
            except:
                pass
        
        # 단위 변환
        if self.unit_var.get() == "mm":
            length_m = total_length / 1000
        else:
            length_m = total_length
        
        # 2x10 기준 체적 계산
        volume = length_m * 0.038 * 0.235
        
        result = f"=== 장선/보 계산 결과 ===\n"
        result += f"개수: {count}개\n"
        result += f"총 길이: {length_m:.2f}m\n"
        result += f"목재 체적: {volume:.3f}m³ (2x10)\n"
        result += f"할증 {self.waste_var.get()}% 적용: {volume * (1 + float(self.waste_var.get())/100):.3f}m³\n"
        
        self.display_result(result)
        self.calculation_results["장선"] = {"길이": length_m, "체적": volume}
    
    def calc_deck(self):
        """데크보드 계산"""
        if not self.selected_objects:
            messagebox.showwarning("경고", "먼저 객체를 선택하세요.")
            return
        
        self.log("데크보드 계산 시작...")
        
        total_area = 0
        count = 0
        
        for obj in self.selected_objects:
            try:
                if self.acad:
                    if "Polyline" in obj.ObjectName and obj.Closed:
                        total_area += obj.Area
                        count += 1
                else:
                    if obj.dxftype() in ["LWPOLYLINE", "POLYLINE"] and obj.is_closed:
                        # Shapely로 면적 계산
                        points = obj.get_points()
                        if len(points) >= 3:
                            from shapely.geometry import Polygon
                            poly = Polygon(points)
                            total_area += poly.area
                            count += 1
            except:
                pass
        
        # 단위 변환
        if self.unit_var.get() == "mm":
            area_m2 = total_area / 1000000
        else:
            area_m2 = total_area
        
        waste_rate = float(self.waste_var.get()) / 100
        board_area = area_m2 * (1 + waste_rate)
        screws = int(area_m2 * 8)
        
        result = f"=== 데크보드 계산 결과 ===\n"
        result += f"폐곡선 개수: {count}개\n"
        result += f"데크 면적: {area_m2:.2f}m²\n"
        result += f"필요 보드: {board_area:.2f}m² ({self.waste_var.get()}% 손실)\n"
        result += f"데크 스크류: {screws}개\n"
        
        self.display_result(result)
        self.calculation_results["데크보드"] = {"면적": area_m2, "필요량": board_area}
    
    def calc_posts(self):
        """기둥 계산"""
        # 간단히 선택된 원형 객체를 기둥으로 가정
        post_count = len(self.selected_objects)
        
        if post_count == 0:
            messagebox.showwarning("경고", "먼저 객체를 선택하세요.")
            return
        
        # 기초 콘크리트 계산
        foundation = post_count * 0.096  # 400x400x600 기준
        
        result = f"=== 기둥 계산 결과 ===\n"
        result += f"기둥 개수: {post_count}개\n"
        result += f"기초 콘크리트: {foundation:.3f}m³\n"
        result += f"할증 5% 적용: {foundation * 1.05:.3f}m³\n"
        
        self.display_result(result)
        self.calculation_results["기둥"] = {"개수": post_count, "콘크리트": foundation}
    
    def calc_rebar(self):
        """철근 계산"""
        messagebox.showinfo("안내", "철근 계산 기능은 개발 중입니다.")
    
    def calc_pipes(self):
        """배관 계산"""
        messagebox.showinfo("안내", "배관 계산 기능은 개발 중입니다.")
    
    def calc_concrete(self):
        """콘크리트 계산"""
        messagebox.showinfo("안내", "콘크리트 계산 기능은 개발 중입니다.")
    
    def calc_length(self):
        """길이 계산"""
        total_length = 0
        count = 0
        
        for obj in self.selected_objects:
            try:
                if self.acad:
                    if "Line" in obj.ObjectName:
                        start = obj.StartPoint
                        end = obj.EndPoint
                        length = math.sqrt((end[0]-start[0])**2 + (end[1]-start[1])**2)
                        total_length += length
                        count += 1
                    elif "Polyline" in obj.ObjectName:
                        total_length += obj.Length
                        count += 1
                else:
                    if obj.dxftype() == "LINE":
                        start = obj.dxf.start
                        end = obj.dxf.end
                        length = start.distance(end)
                        total_length += length
                        count += 1
            except:
                pass
        
        if self.unit_var.get() == "mm":
            length_m = total_length / 1000
        else:
            length_m = total_length
        
        result = f"=== 길이 계산 결과 ===\n"
        result += f"객체 수: {count}개\n"
        result += f"총 길이: {total_length:.2f} {self.unit_var.get()}\n"
        result += f"총 길이: {length_m:.2f}m\n"
        
        self.display_result(result)
    
    def calc_area(self):
        """면적 계산"""
        total_area = 0
        count = 0
        
        for obj in self.selected_objects:
            try:
                if self.acad:
                    if hasattr(obj, 'Area'):
                        total_area += obj.Area
                        count += 1
                else:
                    if obj.dxftype() in ["LWPOLYLINE", "POLYLINE"] and obj.is_closed:
                        points = obj.get_points()
                        if len(points) >= 3:
                            from shapely.geometry import Polygon
                            poly = Polygon(points)
                            total_area += poly.area
                            count += 1
            except:
                pass
        
        if self.unit_var.get() == "mm":
            area_m2 = total_area / 1000000
        else:
            area_m2 = total_area
        
        result = f"=== 면적 계산 결과 ===\n"
        result += f"객체 수: {count}개\n"
        result += f"총 면적: {total_area:.2f} {self.unit_var.get()}²\n"
        result += f"총 면적: {area_m2:.2f}m²\n"
        
        self.display_result(result)
    
    def calc_count(self):
        """개수 집계"""
        result = f"=== 개수 집계 결과 ===\n"
        result += f"총 선택 객체: {len(self.selected_objects)}개\n\n"
        
        # 타입별 집계
        types = {}
        for obj in self.selected_objects:
            if self.acad:
                obj_type = obj.ObjectName
            else:
                obj_type = obj.dxftype()
            
            if obj_type not in types:
                types[obj_type] = 0
            types[obj_type] += 1
        
        result += "타입별 분류:\n"
        for obj_type, count in sorted(types.items()):
            result += f"  {obj_type}: {count}개\n"
        
        self.display_result(result)
    
    def custom_calculation(self):
        """사용자 정의 계산"""
        formula = self.formula_entry.get()
        
        # 변수 계산
        L = 0  # 총 길이
        A = 0  # 총 면적  
        N = len(self.selected_objects)  # 객체 수
        
        # L과 A 계산 (간단히)
        for obj in self.selected_objects:
            try:
                if self.acad:
                    if hasattr(obj, 'Length'):
                        L += obj.Length
                    if hasattr(obj, 'Area'):
                        A += obj.Area
            except:
                pass
        
        try:
            result = eval(formula)
            
            output = f"=== 사용자 정의 계산 ===\n"
            output += f"계산식: {formula}\n"
            output += f"L (길이): {L:.2f}\n"
            output += f"A (면적): {A:.2f}\n"
            output += f"N (개수): {N}\n"
            output += f"결과: {result:.3f}\n"
            
            self.display_result(output)
            
        except Exception as e:
            messagebox.showerror("오류", f"계산 오류: {str(e)}")
    
    # === 분석 메서드 ===
    
    def analyze_all(self):
        """전체 분석"""
        self.log("전체 도면 분석 중...")
        
        if self.acad:
            # AutoCAD 분석
            total = self.doc.ModelSpace.Count
            
            types = {}
            for obj in self.doc.ModelSpace:
                obj_type = obj.ObjectName
                if obj_type not in types:
                    types[obj_type] = 0
                types[obj_type] += 1
            
            result = f"=== 전체 도면 분석 ===\n"
            result += f"총 객체 수: {total}개\n\n"
            result += "객체 타입별 분포:\n"
            
            for obj_type, count in sorted(types.items(), key=lambda x: x[1], reverse=True):
                result += f"  {obj_type}: {count}개 ({count/total*100:.1f}%)\n"
            
        else:
            # DXF 분석
            msp = self.doc.modelspace()
            total = len(msp)
            
            types = {}
            for entity in msp:
                entity_type = entity.dxftype()
                if entity_type not in types:
                    types[entity_type] = 0
                types[entity_type] += 1
            
            result = f"=== 전체 도면 분석 ===\n"
            result += f"DXF 버전: {self.doc.dxfversion}\n"
            result += f"총 객체 수: {total}개\n\n"
            result += "엔티티 타입별 분포:\n"
            
            for entity_type, count in sorted(types.items(), key=lambda x: x[1], reverse=True):
                result += f"  {entity_type}: {count}개 ({count/total*100:.1f}%)\n"
        
        self.stats_text.delete(1.0, tk.END)
        self.stats_text.insert(1.0, result)
        self.notebook.select(1)  # 분석 탭으로 전환
    
    def analyze_by_layer(self):
        """레이어별 분석"""
        self.log("레이어별 분석 중...")
        
        layer_stats = {}
        
        if self.acad:
            for obj in self.doc.ModelSpace:
                layer = obj.Layer
                if layer not in layer_stats:
                    layer_stats[layer] = {'count': 0, 'types': set()}
                layer_stats[layer]['count'] += 1
                layer_stats[layer]['types'].add(obj.ObjectName)
        else:
            msp = self.doc.modelspace()
            for entity in msp:
                layer = entity.dxf.layer
                if layer not in layer_stats:
                    layer_stats[layer] = {'count': 0, 'types': set()}
                layer_stats[layer]['count'] += 1
                layer_stats[layer]['types'].add(entity.dxftype())
        
        result = "=== 레이어별 분석 ===\n\n"
        
        for layer, stats in sorted(layer_stats.items(), key=lambda x: x[1]['count'], reverse=True):
            result += f"레이어: {layer}\n"
            result += f"  객체 수: {stats['count']}개\n"
            result += f"  타입: {', '.join(list(stats['types'])[:5])}\n\n"
        
        self.stats_text.delete(1.0, tk.END)
        self.stats_text.insert(1.0, result)
        self.notebook.select(1)
    
    def analyze_by_type(self):
        """타입별 분석"""
        messagebox.showinfo("안내", "타입별 분석 기능은 개발 중입니다.")
    
    # === 유틸리티 메서드 ===
    
    def update_layer_list(self):
        """레이어 목록 업데이트 (DXF)"""
        self.layer_listbox.delete(0, tk.END)
        
        for layer in self.doc.layers:
            self.layer_listbox.insert(tk.END, layer.dxf.name)
    
    def update_layer_list_autocad(self):
        """레이어 목록 업데이트 (AutoCAD)"""
        self.layer_listbox.delete(0, tk.END)
        
        for layer in self.doc.Layers:
            self.layer_listbox.insert(tk.END, layer.Name)
    
    def display_result(self, text):
        """결과 표시"""
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(1.0, text)
        
        # 현재 시간 추가
        timestamp = datetime.now().strftime("\n\n계산 시간: %Y-%m-%d %H:%M:%S")
        self.result_text.insert(tk.END, timestamp)
    
    def log(self, message):
        """로그 메시지 추가"""
        timestamp = datetime.now().strftime("[%H:%M:%S]")
        self.log_text.insert(tk.END, f"{timestamp} {message}\n")
        self.log_text.see(tk.END)
        
        # 상태바 업데이트
        self.status_text.config(text=message)
    
    def refresh_data(self):
        """데이터 새로고침"""
        self.log("데이터 새로고침...")
        
        if self.acad:
            self.object_count_label.config(text=f"객체 수: {self.doc.ModelSpace.Count}")
            self.update_layer_list_autocad()
        elif self.doc:
            msp = self.doc.modelspace()
            self.object_count_label.config(text=f"객체 수: {len(msp)}")
            self.update_layer_list()
    
    def clear_all(self):
        """모두 초기화"""
        self.clear_selection()
        self.result_text.delete(1.0, tk.END)
        self.stats_text.delete(1.0, tk.END)
        self.calculation_results = {}
        self.log("모든 데이터 초기화")
    
    def save_results(self):
        """결과 저장"""
        if not self.calculation_results:
            messagebox.showwarning("경고", "저장할 계산 결과가 없습니다.")
            return
        
        self.export_json()
    
    def export_excel(self):
        """Excel로 내보내기"""
        if not self.calculation_results:
            messagebox.showwarning("경고", "저장할 계산 결과가 없습니다.")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                # pandas DataFrame 생성
                data = []
                for category, values in self.calculation_results.items():
                    for key, value in values.items():
                        data.append({"카테고리": category, "항목": key, "값": value})
                
                df = pd.DataFrame(data)
                df.to_excel(filename, index=False)
                
                messagebox.showinfo("성공", f"Excel 파일이 저장되었습니다:\n{filename}")
                self.log(f"Excel 저장: {filename}")
                
            except Exception as e:
                messagebox.showerror("오류", f"저장 실패: {str(e)}")
    
    def export_json(self):
        """JSON으로 내보내기"""
        if not self.calculation_results:
            messagebox.showwarning("경고", "저장할 계산 결과가 없습니다.")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if filename:
            try:
                output = {
                    "파일": self.current_file or "Unknown",
                    "계산일시": datetime.now().isoformat(),
                    "결과": self.calculation_results
                }
                
                with open(filename, 'w', encoding='utf-8') as f:
                    json.dump(output, f, ensure_ascii=False, indent=2)
                
                messagebox.showinfo("성공", f"JSON 파일이 저장되었습니다:\n{filename}")
                self.log(f"JSON 저장: {filename}")
                
            except Exception as e:
                messagebox.showerror("오류", f"저장 실패: {str(e)}")
    
    def export_to_cad(self):
        """AutoCAD 테이블로 내보내기"""
        if not self.acad:
            messagebox.showwarning("경고", "AutoCAD가 연결되지 않았습니다.")
            return
        
        messagebox.showinfo("안내", "AutoCAD 테이블 내보내기는 개발 중입니다.")
    
    def show_conversion_guide(self):
        """DWG→DXF 변환 안내"""
        guide = """DWG를 DXF로 변환하는 방법:

1. ODA File Converter (무료)
   - https://www.opendesign.com/guestfiles/oda_file_converter
   - 다운로드 후 설치
   - Input folder: DWG 파일 위치
   - Output folder: DXF 저장 위치
   - Output version: ACAD2018 DXF
   - Run 클릭

2. AutoCAD 사용
   - DWG 파일 열기
   - 다른 이름으로 저장
   - 파일 형식: DXF 선택

3. 온라인 변환 서비스
   - CloudConvert
   - Convertio
   """
        messagebox.showinfo("DWG→DXF 변환 안내", guide)
    
    def show_settings(self):
        """설정 창"""
        messagebox.showinfo("설정", "설정 기능은 개발 중입니다.")
    
    def show_help(self):
        """도움말"""
        help_text = """CADPro 사용법:

1. 파일 열기
   - DXF 파일: 파일 → DXF 파일 열기
   - AutoCAD: 파일 → AutoCAD 연결

2. 객체 선택
   - 영역 선택: 마우스로 범위 지정
   - 개별 선택: 하나씩 클릭
   - 레이어 선택: 레이어 목록에서 선택

3. 물량 계산
   - 원하는 계산 버튼 클릭
   - 옵션 설정 (단위, 할증률)
   - 결과 확인

4. 결과 저장
   - Excel 또는 JSON 형식
   - AutoCAD 테이블 (개발 중)
"""
        messagebox.showinfo("도움말", help_text)
    
    def show_about(self):
        """프로그램 정보"""
        about_text = """CADPro - CAD 물량 산출 프로그램
버전: 1.0.0

AutoCAD 도면에서 자동으로 물량을 계산하는
전문 프로그램입니다.

지원 분야:
- 건축 (벽체, 기둥, 슬래브)
- 토목 (도로, 배관, 철근)
- 조경 (수목, 잔디, 포장)
- 데크 공사

© 2024 CADPro
"""
        messagebox.showinfo("CADPro 정보", about_text)
    
    def run(self):
        """프로그램 실행"""
        self.root.mainloop()


def main():
    app = CADProGUI()
    app.run()


if __name__ == "__main__":
    main()