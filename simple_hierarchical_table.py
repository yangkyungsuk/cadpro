"""
Simple Hierarchical Table - 평면 테이블에 계층구조 추가
QTableWidget을 사용한 단순한 계층구조 구현
"""

from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
import json
from typing import Dict, List, Any, Optional
from enum import Enum
import win32com.client
import pythoncom
import time


class RowType(Enum):
    """행 타입"""
    CATEGORY = "category"      # 대분류
    SUBCATEGORY = "subcategory"  # 중분류  
    ITEM = "item"              # 일반 항목


class SimpleHierarchicalTable(QTableWidget):
    """평면 테이블에 간단한 계층구조를 추가한 테이블"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.acad = None
        self.doc = None
        self.parent_widget = parent
        self.setup_table()
        self.init_context_menu()
        
        # 행 타입 추적
        self.row_types = {}  # {row_index: RowType}
        self.row_levels = {}  # {row_index: level_string} e.g., "1", "1-1", "1-1-1"
        
    def setup_table(self):
        """테이블 설정"""
        # 컬럼 설정 (번호 컬럼 추가)
        columns = [
            "번호",        # 0 - 계층 번호 (1, 1-1 등)
            "구분",        # 1 - 카테고리/서브카테고리/일반
            "품명",        # 2
            "규격",        # 3
            "수량",        # 4
            "단위",        # 5
            "가로",        # 6
            "세로",        # 7
            "면적",        # 8
            "둘레",        # 9
            "두께",        # 10
            "층고",        # 11
            "계산식",      # 12
            "결과",        # 13
            "추출모드",    # 14
            "레이어",      # 15
            "비고",        # 16
            "선택",        # 17 - 선택 버튼
            "돋보기"       # 18 - 돋보기 버튼
        ]
        
        self.setColumnCount(len(columns))
        self.setHorizontalHeaderLabels(columns)
        
        # 편집 트리거를 더블클릭으로 변경 (기본 클릭으로 편집 시작하지 않도록)
        self.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.EditKeyPressed)
        
        # 셀 값 변경 시 수식 계산
        self.itemChanged.connect(self.calculate_formula)
        
        # 컬럼 너비 설정
        self.setColumnWidth(0, 60)   # 번호
        self.setColumnWidth(1, 80)   # 구분
        self.setColumnWidth(2, 150)  # 품명
        self.setColumnWidth(3, 100)  # 규격
        self.setColumnWidth(4, 60)   # 수량
        self.setColumnWidth(5, 50)   # 단위
        self.setColumnWidth(6, 60)   # 가로
        self.setColumnWidth(7, 60)   # 세로
        self.setColumnWidth(8, 80)   # 면적
        self.setColumnWidth(9, 80)   # 둘레
        self.setColumnWidth(10, 60)  # 두께
        self.setColumnWidth(11, 60)  # 층고
        self.setColumnWidth(12, 120) # 계산식
        self.setColumnWidth(13, 80)  # 결과
        self.setColumnWidth(14, 80)  # 추출모드
        self.setColumnWidth(15, 100) # 레이어
        self.setColumnWidth(16, 150) # 비고
        self.setColumnWidth(17, 50)  # 선택 버튼
        self.setColumnWidth(18, 50)  # 돋보기 버튼
        
        # 테이블 설정
        self.setAlternatingRowColors(True)
        self.setSelectionBehavior(QAbstractItemView.SelectRows)
        
        # 헤더 설정
        header = self.horizontalHeader()
        header.setStretchLastSection(False)
        
    def init_context_menu(self):
        """컨텍스트 메뉴 초기화"""
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self.show_context_menu)
        
    def show_context_menu(self, position):
        """컨텍스트 메뉴 표시"""
        menu = QMenu(self)
        current_row = self.currentRow()
        
        # 대분류 추가
        add_category = menu.addAction("📊 대분류 추가")
        add_category.triggered.connect(self.add_category)
        
        if current_row >= 0:
            row_type = self.row_types.get(current_row, RowType.ITEM)
            
            if row_type == RowType.CATEGORY:
                # 대분류 선택 시
                add_subcategory = menu.addAction("📁 중분류 추가")
                add_subcategory.triggered.connect(lambda: self.add_subcategory(current_row))
                
            elif row_type == RowType.SUBCATEGORY:
                # 중분류 선택 시
                add_item = menu.addAction("➕ 항목 추가")
                add_item.triggered.connect(lambda: self.add_item(current_row))
                
            menu.addSeparator()
            
            # 행 삭제
            delete_row = menu.addAction("🗑️ 행 삭제")
            delete_row.triggered.connect(lambda: self.delete_row(current_row))
            
        menu.addSeparator()
        
        # 일반 행 추가
        add_row = menu.addAction("➕ 행 추가")
        add_row.triggered.connect(self.add_row)
        
        menu.exec_(self.mapToGlobal(position))
        
    def add_category(self):
        """대분류 추가"""
        text, ok = QInputDialog.getText(self, "대분류 추가", "대분류 이름:")
        if ok and text:
            row = self.rowCount()
            self.insertRow(row)
            
            # 대분류 번호 생성 (1, 2, 3...)
            category_count = sum(1 for t in self.row_types.values() if t == RowType.CATEGORY)
            level_num = str(category_count + 1)
            
            # 번호 설정
            num_item = QTableWidgetItem(level_num)
            num_item.setFlags(num_item.flags() & ~Qt.ItemIsEditable)
            self.setItem(row, 0, num_item)
            
            # 구분 설정
            type_item = QTableWidgetItem("대분류")
            type_item.setBackground(QColor(200, 200, 255))
            type_item.setFlags(type_item.flags() & ~Qt.ItemIsEditable)
            self.setItem(row, 1, type_item)
            
            # 품명 설정 - 편집 불가능하게 설정
            name_item = QTableWidgetItem(text)
            name_item.setFont(QFont("", 10, QFont.Bold))
            name_item.setFlags(name_item.flags() & ~Qt.ItemIsEditable)  # 편집 불가
            self.setItem(row, 2, name_item)
            
            # 나머지 컬럼 비활성화
            for col in range(3, self.columnCount()-2):  # 버튼 컬럼 제외
                empty_item = QTableWidgetItem("")
                empty_item.setFlags(empty_item.flags() & ~Qt.ItemIsEditable)
                empty_item.setBackground(QColor(230, 230, 230))
                self.setItem(row, col, empty_item)
                
            # 행 타입 저장
            self.row_types[row] = RowType.CATEGORY
            self.row_levels[row] = level_num
            
    def add_subcategory(self, parent_row):
        """중분류 추가"""
        if parent_row < 0 or self.row_types.get(parent_row) != RowType.CATEGORY:
            return
            
        text, ok = QInputDialog.getText(self, "중분류 추가", "중분류 이름:")
        if ok and text:
            # 부모 대분류의 번호
            parent_level = self.row_levels.get(parent_row, "1")
            
            # 같은 대분류 아래의 중분류 개수 계산
            subcategory_count = 0
            for i in range(parent_row + 1, self.rowCount()):
                if self.row_types.get(i) == RowType.CATEGORY:
                    break  # 다음 대분류를 만나면 중단
                if self.row_types.get(i) == RowType.SUBCATEGORY:
                    if self.row_levels.get(i, "").startswith(parent_level + "-"):
                        subcategory_count += 1
                        
            # 중분류 번호 생성 (1-1, 1-2, ...)
            level_num = f"{parent_level}-{subcategory_count + 1}"
            
            # 삽입 위치 찾기 (현재 대분류의 마지막 항목 다음)
            insert_row = parent_row + 1
            for i in range(parent_row + 1, self.rowCount()):
                if self.row_types.get(i) == RowType.CATEGORY:
                    break
                insert_row = i + 1
                
            self.insertRow(insert_row)
            
            # 번호 설정 - 편집 불가
            num_item = QTableWidgetItem(level_num)
            num_item.setFlags(num_item.flags() & ~Qt.ItemIsEditable)
            self.setItem(insert_row, 0, num_item)
            
            # 구분 설정
            type_item = QTableWidgetItem("중분류")
            type_item.setBackground(QColor(220, 220, 255))
            type_item.setFlags(type_item.flags() & ~Qt.ItemIsEditable)
            self.setItem(insert_row, 1, type_item)
            
            # 품명 설정 - 편집 불가
            name_item = QTableWidgetItem("  " + text)  # 들여쓰기
            name_item.setFont(QFont("", 9, QFont.Bold))
            name_item.setFlags(name_item.flags() & ~Qt.ItemIsEditable)  # 편집 불가
            self.setItem(insert_row, 2, name_item)
            
            # 나머지 컬럼 비활성화
            for col in range(3, self.columnCount()-2):  # 버튼 컬럼 제외
                empty_item = QTableWidgetItem("")
                empty_item.setFlags(empty_item.flags() & ~Qt.ItemIsEditable)
                empty_item.setBackground(QColor(240, 240, 240))
                self.setItem(insert_row, col, empty_item)
                
            # 행 타입 저장
            self.row_types[insert_row] = RowType.SUBCATEGORY
            self.row_levels[insert_row] = level_num
            
    def add_item(self, parent_row):
        """중분류 아래에 일반 항목 추가"""
        if parent_row < 0 or self.row_types.get(parent_row) != RowType.SUBCATEGORY:
            return
            
        # 부모 중분류의 번호
        parent_level = self.row_levels.get(parent_row, "1-1")
        
        # 같은 중분류 아래의 항목 개수 계산
        item_count = 0
        for i in range(parent_row + 1, self.rowCount()):
            row_type = self.row_types.get(i)
            if row_type in [RowType.CATEGORY, RowType.SUBCATEGORY]:
                break  # 다른 분류를 만나면 중단
            if row_type == RowType.ITEM:
                if self.row_levels.get(i, "").startswith(parent_level + "-"):
                    item_count += 1
                    
        # 항목 번호 생성 (1-1-1, 1-1-2, ...)
        level_num = f"{parent_level}-{item_count + 1}"
        
        # 삽입 위치 찾기 (현재 중분류의 마지막 항목 다음)
        insert_row = parent_row + 1
        for i in range(parent_row + 1, self.rowCount()):
            row_type = self.row_types.get(i)
            if row_type in [RowType.CATEGORY, RowType.SUBCATEGORY]:
                break
            insert_row = i + 1
            
        self.insertRow(insert_row)
        
        # 번호 설정
        self.setItem(insert_row, 0, QTableWidgetItem(level_num))
        
        # 구분 설정
        type_item = QTableWidgetItem("항목")
        self.setItem(insert_row, 1, type_item)
        
        # 기본값 설정 - 필요한 필드만 초기화
        self.setItem(insert_row, 2, QTableWidgetItem(""))  # 품명
        self.setItem(insert_row, 3, QTableWidgetItem(""))  # 규격
        self.setItem(insert_row, 4, QTableWidgetItem(""))  # 수량 - 빈값
        self.setItem(insert_row, 5, QTableWidgetItem(""))  # 단위 - 빈값
        self.setItem(insert_row, 6, QTableWidgetItem(""))  # 가로
        self.setItem(insert_row, 7, QTableWidgetItem(""))  # 세로
        self.setItem(insert_row, 8, QTableWidgetItem(""))  # 면적
        self.setItem(insert_row, 9, QTableWidgetItem(""))  # 둘레
        self.setItem(insert_row, 10, QTableWidgetItem(""))  # 두께
        self.setItem(insert_row, 11, QTableWidgetItem(""))  # 층고
        self.setItem(insert_row, 12, QTableWidgetItem(""))  # 계산식 - 빈값
        
        # 결과 컬럼은 편집 불가
        result_item = QTableWidgetItem("")
        result_item.setFlags(result_item.flags() & ~Qt.ItemIsEditable)
        result_item.setBackground(QColor(245, 245, 245))
        self.setItem(insert_row, 13, result_item)  # 결과
        
        self.setItem(insert_row, 14, QTableWidgetItem(""))  # 추출모드
        self.setItem(insert_row, 15, QTableWidgetItem(""))  # 레이어
        self.setItem(insert_row, 16, QTableWidgetItem(""))  # 비고
        
        # 버튼 추가
        self.add_buttons(insert_row)
        
        # 행 타입 저장
        self.row_types[insert_row] = RowType.ITEM
        self.row_levels[insert_row] = level_num
        
    def add_row(self):
        """일반 행 추가 (기존 평면 테이블처럼)"""
        # 현재 선택된 행 찾기
        current_row = self.currentRow()
        
        # 적절한 중분류 찾기
        target_subcategory_row = -1
        
        if current_row >= 0:
            # 현재 행의 타입 확인
            row_type = self.row_types.get(current_row, RowType.ITEM)
            
            if row_type == RowType.SUBCATEGORY:
                target_subcategory_row = current_row
            elif row_type == RowType.ITEM:
                # 현재 항목이 속한 중분류 찾기
                for i in range(current_row - 1, -1, -1):
                    if self.row_types.get(i) == RowType.SUBCATEGORY:
                        target_subcategory_row = i
                        break
                        
        # 중분류가 없으면 기본 구조 생성
        if target_subcategory_row < 0:
            # 대분류와 중분류 자동 생성
            self.insertRow(0)
            
            # 대분류
            num_item = QTableWidgetItem("1")
            num_item.setFlags(num_item.flags() & ~Qt.ItemIsEditable)
            self.setItem(0, 0, num_item)
            
            type_item = QTableWidgetItem("대분류")
            type_item.setBackground(QColor(200, 200, 255))
            type_item.setFlags(type_item.flags() & ~Qt.ItemIsEditable)
            self.setItem(0, 1, type_item)
            
            name_item = QTableWidgetItem("일반")
            name_item.setFlags(name_item.flags() & ~Qt.ItemIsEditable)
            self.setItem(0, 2, name_item)
            
            self.row_types[0] = RowType.CATEGORY
            self.row_levels[0] = "1"
            
            # 나머지 컬럼 비활성화
            for col in range(3, self.columnCount()-2):
                empty_item = QTableWidgetItem("")
                empty_item.setFlags(empty_item.flags() & ~Qt.ItemIsEditable)
                empty_item.setBackground(QColor(230, 230, 230))
                self.setItem(0, col, empty_item)
            
            # 중분류
            self.insertRow(1)
            
            num_item = QTableWidgetItem("1-1")
            num_item.setFlags(num_item.flags() & ~Qt.ItemIsEditable)
            self.setItem(1, 0, num_item)
            
            type_item = QTableWidgetItem("중분류")
            type_item.setBackground(QColor(220, 220, 255))
            type_item.setFlags(type_item.flags() & ~Qt.ItemIsEditable)
            self.setItem(1, 1, type_item)
            
            name_item = QTableWidgetItem("  항목")
            name_item.setFlags(name_item.flags() & ~Qt.ItemIsEditable)
            self.setItem(1, 2, name_item)
            
            self.row_types[1] = RowType.SUBCATEGORY
            self.row_levels[1] = "1-1"
            
            # 나머지 컬럼 비활성화
            for col in range(3, self.columnCount()-2):
                empty_item = QTableWidgetItem("")
                empty_item.setFlags(empty_item.flags() & ~Qt.ItemIsEditable)
                empty_item.setBackground(QColor(240, 240, 240))
                self.setItem(1, col, empty_item)
            
            target_subcategory_row = 1
            
        # 중분류 아래에 항목 추가
        self.add_item(target_subcategory_row)
        
    def add_buttons(self, row):
        """선택 및 돋보기 버튼 추가"""
        # 추출모드 드롭다운 추가
        extract_combo = QComboBox()
        extract_combo.addItems(["선택", "전체", "면적", "둘레", "길이", "체적"])
        self.setCellWidget(row, 14, extract_combo)
        
        # 선택 버튼
        select_btn = QPushButton("🎯")
        select_btn.setMaximumWidth(40)
        # 람다에서 row 값을 캡처하도록 수정
        select_btn.clicked.connect(lambda checked, r=row: self.select_from_cad(r))
        self.setCellWidget(row, 17, select_btn)
        
        # 돋보기 버튼
        magnifier_btn = QPushButton("🔍")
        magnifier_btn.setMaximumWidth(40)
        # 람다에서 row 값을 캡처하도록 수정
        magnifier_btn.clicked.connect(lambda checked, r=row: self.show_selection_helper(r))
        self.setCellWidget(row, 18, magnifier_btn)
        
    def select_from_cad(self, row):
        """CAD에서 객체 선택 - 평면 테이블과 동일한 로직"""
        if not self.doc:
            QMessageBox.warning(self, "경고", "먼저 AutoCAD를 연결하세요")
            return
            
        try:
            pythoncom.CoInitialize()
            
            # 선택 세트 생성
            sel_name = f"Sel_{int(time.time())}"
            selection = self.doc.SelectionSets.Add(sel_name)
            
            print("\n🎯 CAD 객체 선택 모드")
            selection.SelectOnScreen()
            
            if selection.Count == 0:
                print("  선택된 객체 없음")
                selection.Delete()
                return
                
            # 선택 객체 처리
            selected_objects = []
            total_length = 0
            total_area = 0
            total_perimeter = 0
            
            # 사각형 감지를 위한 변수
            rectangles = []
            
            for i in range(selection.Count):
                obj = selection.Item(i)
                selected_objects.append(obj)
                
                # 치수 추출
                obj_type = str(obj.ObjectName)
                
                if "Polyline" in obj_type:
                    try:
                        # 면적
                        if hasattr(obj, 'Area'):
                            total_area += obj.Area
                            
                        # 폴리라인이 사각형인지 확인
                        if hasattr(obj, 'Closed') and obj.Closed:
                            try:
                                coords = obj.Coordinates
                                # 4개 점으로 이루어진 폐합 폴리라인 = 사각형
                                if len(coords) >= 8:  # 최소 4개 점 (8개 좌표값)
                                    # 사각형의 가로, 세로 계산
                                    x_coords = [coords[i] for i in range(0, len(coords), 2)]
                                    y_coords = [coords[i] for i in range(1, len(coords), 2)]
                                    
                                    width = max(x_coords) - min(x_coords)
                                    height = max(y_coords) - min(y_coords)
                                    
                                    # 사각형인지 확인 (4점 또는 5점)
                                    if len(coords) in [8, 10]:
                                        rectangles.append({
                                            'width': width,
                                            'height': height,
                                            'area': width * height,
                                            'perimeter': 2 * (width + height)
                                        })
                                    
                                    total_perimeter += 2 * (width + height)
                            except Exception as e:
                                print(f"폴리라인 처리 오류: {e}")
                                
                        # 일반 폴리라인 둘레
                        try:
                            coords = obj.Coordinates
                            perimeter = 0
                            for j in range(0, len(coords)-2, 2):
                                x1, y1 = coords[j], coords[j+1]
                                x2, y2 = coords[j+2], coords[j+3]
                                import math
                                perimeter += math.sqrt((x2-x1)**2 + (y2-y1)**2)
                            # 폐합된 경우 마지막 변도 추가
                            if hasattr(obj, 'Closed') and obj.Closed and len(coords) >= 4:
                                x1, y1 = coords[-2], coords[-1]
                                x2, y2 = coords[0], coords[1]
                                perimeter += math.sqrt((x2-x1)**2 + (y2-y1)**2)
                            if perimeter > 0 and not rectangles:  # 사각형이 아닌 경우만
                                total_perimeter += perimeter
                        except:
                            pass
                            
                    except:
                        pass
                        
                elif "Line" in obj_type and "Polyline" not in obj_type:
                    try:
                        start = obj.StartPoint
                        end = obj.EndPoint
                        import math
                        length = math.sqrt((end[0]-start[0])**2 + (end[1]-start[1])**2)
                        total_length += length
                    except:
                        pass
                        
                elif "Circle" in obj_type:
                    try:
                        import math
                        radius = obj.Radius
                        total_area += math.pi * radius * radius
                        total_perimeter += 2 * math.pi * radius
                    except:
                        pass
                        
            # 테이블 업데이트 - 항목이 없으면 생성
            # 개수
            if not self.item(row, 4):
                self.setItem(row, 4, QTableWidgetItem(""))
            self.item(row, 4).setText(str(selection.Count))
            
            # 사각형이 감지된 경우
            if rectangles:
                print(f"   사각형 {len(rectangles)}개 감지됨")
                # 첫 번째 사각형의 치수 사용
                rect = rectangles[0]
                
                # 가로
                if not self.item(row, 6):
                    self.setItem(row, 6, QTableWidgetItem(""))
                self.item(row, 6).setText(f"{rect['width']:.1f}")
                print(f"   가로: {rect['width']:.1f}")
                    
                # 세로
                if not self.item(row, 7):
                    self.setItem(row, 7, QTableWidgetItem(""))
                self.item(row, 7).setText(f"{rect['height']:.1f}")
                print(f"   세로: {rect['height']:.1f}")
                
                # 여러 사각형인 경우 합계
                if len(rectangles) > 1:
                    total_rect_area = sum(r['area'] for r in rectangles)
                    total_rect_perimeter = sum(r['perimeter'] for r in rectangles)
                    
                    if not self.item(row, 8):
                        self.setItem(row, 8, QTableWidgetItem(""))
                    self.item(row, 8).setText(f"{total_rect_area:.1f}")  # 면적
                    
                    if not self.item(row, 9):
                        self.setItem(row, 9, QTableWidgetItem(""))
                    self.item(row, 9).setText(f"{total_rect_perimeter:.1f}")  # 둘레
                else:
                    if not self.item(row, 8):
                        self.setItem(row, 8, QTableWidgetItem(""))
                    self.item(row, 8).setText(f"{rect['area']:.1f}")  # 면적
                    
                    if not self.item(row, 9):
                        self.setItem(row, 9, QTableWidgetItem(""))
                    self.item(row, 9).setText(f"{rect['perimeter']:.1f}")  # 둘레
                    print(f"   둘레: {rect['perimeter']:.1f}")
                    
            else:
                # 사각형이 아닌 경우
                if total_area > 0:
                    if not self.item(row, 8):
                        self.setItem(row, 8, QTableWidgetItem(""))
                    self.item(row, 8).setText(f"{total_area:.1f}")  # 면적
                    
                if total_length > 0:
                    if not self.item(row, 6):
                        self.setItem(row, 6, QTableWidgetItem(""))
                    self.item(row, 6).setText(f"{total_length:.1f}")  # 가로(길이)
                    
                if total_perimeter > 0:
                    if not self.item(row, 9):
                        self.setItem(row, 9, QTableWidgetItem(""))
                    self.item(row, 9).setText(f"{total_perimeter:.1f}")  # 둘레
                    
            # 레이어
            if selected_objects:
                if not self.item(row, 17):
                    self.setItem(row, 17, QTableWidgetItem(""))
                self.item(row, 17).setText(str(selected_objects[0].Layer))
                
            # 선택 객체 저장 (행 데이터로)
            if not hasattr(self, 'row_selections'):
                self.row_selections = {}
            self.row_selections[row] = selected_objects
            
            # Line 객체가 여러 개인 경우 길이별 그룹화
            if len(selected_objects) > 1:
                all_lines = True
                for obj in selected_objects:
                    obj_type = str(obj.ObjectName)
                    if not ("Line" in obj_type and "Polyline" not in obj_type):
                        all_lines = False
                        break
                
                if all_lines:
                    print(f"\n📊 Line 객체 {len(selected_objects)}개 - 길이별 그룹화")
                    
                    # 길이별로 그룹화
                    groups = {}
                    for obj in selected_objects:
                        try:
                            start = obj.StartPoint
                            end = obj.EndPoint
                            length = math.sqrt((end[0]-start[0])**2 + (end[1]-start[1])**2)
                            length_key = round(length, 1)
                            if length_key not in groups:
                                groups[length_key] = []
                            groups[length_key].append(obj)
                        except:
                            pass
                    
                    if len(groups) > 1:
                        print(f"  {len(groups)}개 그룹으로 분할 필요")
                        
                        # 현재 행의 부모 찾기 (중분류)
                        parent_row = -1
                        for i in range(row - 1, -1, -1):
                            if self.row_types.get(i) == RowType.SUBCATEGORY:
                                parent_row = i
                                break
                        
                        # 첫 번째 그룹은 현재 행에
                        sorted_groups = sorted(groups.items())
                        first_key, first_objects = sorted_groups[0]
                        self.row_selections[row] = first_objects
                        self.setItem(row, 4, QTableWidgetItem(str(len(first_objects))))
                        self.setItem(row, 6, QTableWidgetItem(f"{first_key:.1f}"))
                        print(f"  행 {row}: 길이={first_key:.1f}, 수량={len(first_objects)}")
                        
                        # 품명 가져오기
                        original_name = self.item(row, 2).text() if self.item(row, 2) else ""
                        
                        # 나머지 그룹은 새 행에 추가
                        for length_key, objects in sorted_groups[1:]:
                            if parent_row >= 0:
                                self.add_item(parent_row)
                                new_row = self.rowCount() - 1
                                
                                # 데이터 설정
                                self.setItem(new_row, 2, QTableWidgetItem(original_name))  # 품명
                                self.setItem(new_row, 4, QTableWidgetItem(str(len(objects))))  # 수량
                                self.setItem(new_row, 5, QTableWidgetItem("개"))  # 단위
                                self.setItem(new_row, 6, QTableWidgetItem(f"{length_key:.1f}"))  # 가로에 길이
                                
                                # 선택 객체 저장
                                self.row_selections[new_row] = objects
                                
                                print(f"  행 {new_row}: 길이={length_key:.1f}, 수량={len(objects)}")
            
            # 결과 메시지
            print(f"✅ {selection.Count}개 객체 선택됨")
            if rectangles:
                print(f"   사각형 {len(rectangles)}개 감지")
                for i, rect in enumerate(rectangles[:3]):  # 최대 3개만 표시
                    print(f"   사각형{i+1}: {rect['width']:.1f} x {rect['height']:.1f}")
            if total_length > 0:
                print(f"   총 길이: {total_length:.3f}mm")
            if total_area > 0:
                print(f"   총 면적: {total_area:.3f}mm²")
                
            selection.Delete()
            
        except Exception as e:
            print(f"❌ CAD 선택 오류: {e}")
        finally:
            pythoncom.CoUninitialize()
            
    def show_selection_helper(self, row):
        """선택 도우미 표시"""
        if not hasattr(self, 'row_selections') or row not in self.row_selections:
            QMessageBox.information(self, "안내", 
                "먼저 선택 버튼을 눌러 CAD 객체를 선택하세요.")
            return
            
        selected_objects = self.row_selections[row]
        
        # doc이 있는지 확인
        if not self.doc:
            QMessageBox.warning(self, "경고", "AutoCAD 연결이 필요합니다.")
            return
        
        # parent_widget이 있고 SelectionHelperDialog가 있는지 확인
        if hasattr(self, 'parent_widget') and self.parent_widget:
            # parent_widget에서 SelectionHelperDialog 가져오기
            from CAD_Quantity_Pro_Simple import SelectionHelperDialog
            
            # 선택 도우미 대화상자 표시
            dialog = SelectionHelperDialog(self, self.doc, selected_objects, row)
            if dialog.exec_():
                # 선택 결과 업데이트
                new_selection = dialog.get_final_selection()
                if new_selection:
                    print(f"\n📊 선택 도우미 결과: {len(new_selection)}개 객체")
                    
                    # Line 객체인 경우 길이별 그룹화 확인
                    all_lines = True
                    for obj in new_selection:
                        obj_type = str(obj.ObjectName)
                        if not ("Line" in obj_type and "Polyline" not in obj_type):
                            all_lines = False
                            break
                    
                    if all_lines and len(new_selection) > 1:
                        print(f"  Line 객체 {len(new_selection)}개 - 길이별 그룹화 시도")
                        
                        # 길이별로 그룹화
                        import math
                        groups = {}
                        for obj in new_selection:
                            try:
                                start = obj.StartPoint
                                end = obj.EndPoint
                                length = math.sqrt((end[0]-start[0])**2 + (end[1]-start[1])**2)
                                length_key = round(length, 1)  # 0.1 단위로 반올림
                                if length_key not in groups:
                                    groups[length_key] = []
                                groups[length_key].append(obj)
                            except:
                                pass
                        
                        print(f"  그룹화 결과: {len(groups)}개 그룹")
                        for key in sorted(groups.keys()):
                            print(f"    길이 {key:.1f}: {len(groups[key])}개")
                        
                        if len(groups) > 1:
                            print(f"  ✅ {len(groups)}개 그룹으로 분할하여 행 추가")
                            
                            # 첫 번째 그룹은 현재 행에
                            sorted_groups = sorted(groups.items())
                            first_key, first_objects = sorted_groups[0]
                            
                            self.row_selections[row] = first_objects
                            # Line 길이는 가로(6번)에 넣기
                            self.setItem(row, 6, QTableWidgetItem(f"{first_key:.1f}"))  # 가로에 길이
                            self.setItem(row, 4, QTableWidgetItem(str(len(first_objects))))  # 수량
                            print(f"    행 {row}: 길이={first_key:.1f}, 수량={len(first_objects)}")
                            
                            # 품명 가져오기
                            original_name = self.item(row, 2).text() if self.item(row, 2) else ""
                            
                            # 현재 행의 부모 찾기 (중분류)
                            parent_row = -1
                            for i in range(row - 1, -1, -1):
                                if self.row_types.get(i) == RowType.SUBCATEGORY:
                                    parent_row = i
                                    break
                            
                            # 나머지 그룹은 새 행에 추가
                            for length_key, objects in sorted_groups[1:]:
                                if parent_row >= 0:
                                    # 중분류 아래에 새 항목 추가
                                    self.add_item(parent_row)
                                    new_row = self.rowCount() - 1
                                    
                                    # 데이터 설정
                                    self.setItem(new_row, 2, QTableWidgetItem(original_name))  # 품명
                                    # Line 길이는 가로(6번)에 넣기, 규격은 비워두기
                                    self.setItem(new_row, 3, QTableWidgetItem(""))  # 규격 비움
                                    self.setItem(new_row, 4, QTableWidgetItem(str(len(objects))))  # 수량
                                    self.setItem(new_row, 5, QTableWidgetItem("개"))  # 단위
                                    self.setItem(new_row, 6, QTableWidgetItem(f"{length_key:.1f}"))  # 가로에 길이
                                    
                                    # 선택 객체 저장
                                    self.row_selections[new_row] = objects
                                    
                                    print(f"    행 {new_row}: 길이={length_key:.1f}, 수량={len(objects)}")
                            
                            print(f"  완료: {len(new_selection)}개 객체를 {len(groups)}개 행으로 분할")
                        else:
                            # 단일 그룹
                            self.row_selections[row] = new_selection
                            self.setItem(row, 4, QTableWidgetItem(str(len(new_selection))))
                            print(f"  단일 그룹 (길이 동일)")
                    else:
                        # Line이 아니거나 단일 객체
                        self.row_selections[row] = new_selection
                        self.setItem(row, 4, QTableWidgetItem(str(len(new_selection))))
                        print(f"✅ 선택 업데이트: {len(new_selection)}개")
        else:
            # 기본 정보 표시 (폴백)
            info = f"선택된 객체: {len(selected_objects)}개\n\n"
            for i, obj in enumerate(selected_objects[:10]):
                try:
                    obj_type = str(obj.ObjectName).replace("AcDb", "")
                    layer = str(obj.Layer)
                    info += f"{i+1}. {obj_type} (레이어: {layer})\n"
                except:
                    pass
            if len(selected_objects) > 10:
                info += f"... 외 {len(selected_objects)-10}개"
            QMessageBox.information(self, "선택 정보", info)
        
    def calculate_formula(self, item):
        """수식 계산"""
        if not item:
            return
            
        row = item.row()
        col = item.column()
        
        # 항목 행에서만 계산 (대분류/중분류 제외)
        if self.row_types.get(row) != RowType.ITEM:
            return
            
        # 계산식 컬럼(12)이 변경되었거나, 다른 값이 변경되었을 때
        if col == 12 or col in [4, 6, 7, 8, 9, 10, 11]:  # 계산식 또는 수량, 치수 등 (단위 제외)
            formula_item = self.item(row, 12)
            if not formula_item:
                return
                
            formula = formula_item.text().strip()
            if not formula:
                # 수식이 없으면 결과 지우기
                if self.item(row, 13):
                    self.item(row, 13).setText("")
                return
                
            try:
                # 변수 값 가져오기
                variables = {
                    '수량': self.get_float_value(row, 4),
                    '가로': self.get_float_value(row, 6),
                    '세로': self.get_float_value(row, 7),
                    '면적': self.get_float_value(row, 8),
                    '둘레': self.get_float_value(row, 9),
                    '두께': self.get_float_value(row, 10),
                    '층고': self.get_float_value(row, 11),
                }
                
                # 영문 변수명도 지원
                variables.update({
                    'qty': variables['수량'],
                    'width': variables['가로'],
                    'height': variables['세로'],
                    'area': variables['면적'],
                    'perimeter': variables['둘레'],
                    'thickness': variables['두께'],
                    'floor': variables['층고'],
                })
                
                # 수식 평가
                result = eval(formula, {"__builtins__": {}}, variables)
                
                # 결과 표시
                if self.item(row, 13):
                    self.item(row, 13).setText(f"{result:.2f}")
                else:
                    self.setItem(row, 13, QTableWidgetItem(f"{result:.2f}"))
                    
            except Exception as e:
                # 오류 시 결과를 빈값으로
                if self.item(row, 13):
                    self.item(row, 13).setText("")
                print(f"  수식 계산 오류 (행 {row}): {e}")
    
    def get_float_value(self, row, col):
        """셀 값을 float로 변환"""
        item = self.item(row, col)
        if not item:
            return 0.0
        try:
            text = item.text().strip()
            if not text:
                return 0.0
            # 숫자만 추출 (단위 제거)
            import re
            match = re.search(r'[\d.]+', text)
            if match:
                return float(match.group())
            return float(text)
        except:
            return 0.0
    
    def delete_row(self, row):
        """행 삭제"""
        row_type = self.row_types.get(row, RowType.ITEM)
        
        if row_type == RowType.CATEGORY:
            reply = QMessageBox.question(self, "확인", 
                "대분류를 삭제하면 하위 항목도 모두 삭제됩니다.\n계속하시겠습니까?",
                QMessageBox.Yes | QMessageBox.No)
            if reply != QMessageBox.Yes:
                return
                
            # 하위 항목 모두 삭제
            rows_to_delete = [row]
            for i in range(row + 1, self.rowCount()):
                if self.row_types.get(i) == RowType.CATEGORY:
                    break
                rows_to_delete.append(i)
                
            # 역순으로 삭제
            for r in reversed(rows_to_delete):
                self.removeRow(r)
                if r in self.row_types:
                    del self.row_types[r]
                if r in self.row_levels:
                    del self.row_levels[r]
                    
        elif row_type == RowType.SUBCATEGORY:
            reply = QMessageBox.question(self, "확인", 
                "중분류를 삭제하면 하위 항목도 모두 삭제됩니다.\n계속하시겠습니까?",
                QMessageBox.Yes | QMessageBox.No)
            if reply != QMessageBox.Yes:
                return
                
            # 하위 항목 모두 삭제
            rows_to_delete = [row]
            for i in range(row + 1, self.rowCount()):
                if self.row_types.get(i) in [RowType.CATEGORY, RowType.SUBCATEGORY]:
                    break
                rows_to_delete.append(i)
                
            # 역순으로 삭제
            for r in reversed(rows_to_delete):
                self.removeRow(r)
                if r in self.row_types:
                    del self.row_types[r]
                if r in self.row_levels:
                    del self.row_levels[r]
                    
        else:
            # 일반 항목 삭제
            self.removeRow(row)
            if row in self.row_types:
                del self.row_types[row]
            if row in self.row_levels:
                del self.row_levels[row]
                
    def set_cad_connection(self, acad, doc):
        """CAD 연결 설정"""
        self.acad = acad
        self.doc = doc
        
    def get_data(self):
        """테이블 데이터 가져오기"""
        data = []
        for row in range(self.rowCount()):
            row_data = {
                'level': self.row_levels.get(row, ""),
                'type': self.row_types.get(row, RowType.ITEM).value,
                'items': []
            }
            
            for col in range(self.columnCount() - 2):  # 버튼 컬럼 제외
                item = self.item(row, col)
                if item:
                    row_data['items'].append(item.text())
                else:
                    row_data['items'].append("")
                    
            data.append(row_data)
            
        return data
        
    def load_data(self, data):
        """데이터 로드"""
        self.setRowCount(0)
        self.row_types.clear()
        self.row_levels.clear()
        
        for row_data in data:
            row = self.rowCount()
            self.insertRow(row)
            
            # 행 타입과 레벨 복원
            row_type = RowType(row_data.get('type', 'item'))
            self.row_types[row] = row_type
            self.row_levels[row] = row_data.get('level', "")
            
            # 데이터 복원
            items = row_data.get('items', [])
            for col, value in enumerate(items[:self.columnCount()-2]):
                item = QTableWidgetItem(str(value))
                
                # 카테고리/서브카테고리 스타일 적용
                if col == 1:  # 구분 컬럼
                    if row_type == RowType.CATEGORY:
                        item.setBackground(QColor(200, 200, 255))
                        item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    elif row_type == RowType.SUBCATEGORY:
                        item.setBackground(QColor(220, 220, 255))
                        item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                        
                self.setItem(row, col, item)
                
            # 일반 항목인 경우 버튼 추가
            if row_type == RowType.ITEM:
                self.add_buttons(row)