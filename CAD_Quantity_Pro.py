"""
CAD Quantity Pro - 간단한 계층구조 버전
Simple Hierarchical Table 통합 버전
"""

import win32com.client
import win32com.client.dynamic
import pythoncom
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
import sys
import json
import os
from typing import Dict, List, Any
import math
import pandas as pd
from datetime import datetime
import io


# ==================== 콘솔 리디렉션 ====================

class ConsoleRedirect:
    """콘솔 출력을 위젯으로 리디렉션"""
    def __init__(self, text_widget):
        self.text_widget = text_widget
        self.buffer = io.StringIO()
        
    def write(self, text):
        # 위젯에 텍스트 추가
        if self.text_widget:
            cursor = self.text_widget.textCursor()
            cursor.movePosition(QTextCursor.End)
            cursor.insertText(text)
            self.text_widget.setTextCursor(cursor)
            self.text_widget.ensureCursorVisible()
        # 버퍼에도 저장
        self.buffer.write(text)
        
    def flush(self):
        pass

# 간단한 계층구조 테이블 임포트
try:
    from simple_hierarchical_table import SimpleHierarchicalTable, RowType
    HIERARCHICAL_TABLE_AVAILABLE = True
except ImportError:
    HIERARCHICAL_TABLE_AVAILABLE = False
    print("⚠️ 계층구조 테이블 모듈 없음 - 평면 테이블만 사용 가능")


# ==================== 평면 테이블 ====================

class FlatQuantityTable(QTableWidget):
    """평면 물량 테이블"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.acad = None
        self.doc = None
        self.parent_widget = parent
        self.setup_table()
        
    def setup_table(self):
        """테이블 설정"""
        columns = [
            "품명", "규격", "수량", "단위", "가로", "세로", 
            "면적", "둘레", "두께", "층고", "계산식",
            "결과", "추출모드",
            "레이어", "비고", "선택", "돋보기"
        ]
        
        self.setColumnCount(len(columns))
        self.setHorizontalHeaderLabels(columns)
        
        # 컬럼 너비
        for i in range(len(columns) - 2):
            self.setColumnWidth(i, 100)
        self.setColumnWidth(len(columns) - 2, 50)  # 선택 버튼
        self.setColumnWidth(len(columns) - 1, 50)  # 돋보기 버튼
        
        self.setAlternatingRowColors(True)
        self.setSelectionBehavior(QAbstractItemView.SelectRows)
        
    def add_row(self):
        """행 추가"""
        row = self.rowCount()
        self.insertRow(row)
        
        # 기본값
        self.setItem(row, 0, QTableWidgetItem(""))  # 품명
        self.setItem(row, 1, QTableWidgetItem(""))  # 규격
        self.setItem(row, 2, QTableWidgetItem("0"))  # 수량
        self.setItem(row, 3, QTableWidgetItem("EA"))  # 단위
        self.setItem(row, 10, QTableWidgetItem("{수량}"))  # 계산식
        
        # 추출모드 드롭다운
        extract_combo = QComboBox()
        extract_combo.addItems(["선택", "전체", "면적", "둘레", "길이", "체적"])
        self.setCellWidget(row, 12, extract_combo)
        
        # 버튼 (lambda에서 row 값을 캡처하도록 수정)
        select_btn = QPushButton("🎯")
        select_btn.clicked.connect(lambda checked, r=row: self.select_from_cad(r))
        self.setCellWidget(row, 15, select_btn)
        
        magnifier_btn = QPushButton("🔍")
        magnifier_btn.clicked.connect(lambda checked, r=row: self.show_selection_helper(r))
        self.setCellWidget(row, 16, magnifier_btn)
        
    def select_from_cad(self, row):
        """CAD 선택"""
        # parent_widget(메인 윈도우)의 doc 사용
        if not self.parent_widget or not hasattr(self.parent_widget, 'doc') or not self.parent_widget.doc:
            QMessageBox.warning(self, "경고", "먼저 AutoCAD를 연결하세요")
            return
        
        self.doc = self.parent_widget.doc  # parent의 doc 사용
            
        try:
            import pythoncom
            import time
            pythoncom.CoInitialize()
            
            # 기존 선택 세트 정리
            try:
                for i in range(self.doc.SelectionSets.Count - 1, -1, -1):
                    try:
                        sel_set = self.doc.SelectionSets.Item(i)
                        if "Sel_" in sel_set.Name:
                            sel_set.Delete()
                    except:
                        pass
            except:
                pass
            
            # 선택 세트 생성
            sel_name = f"Sel_{int(time.time())}"
            
            # 선택 세트 추가 시도
            try:
                selection = self.doc.SelectionSets.Add(sel_name)
            except:
                # 이미 존재하는 경우 삭제 후 재생성
                try:
                    self.doc.SelectionSets.Item(sel_name).Delete()
                except:
                    pass
                selection = self.doc.SelectionSets.Add(sel_name)
            
            print("\n🎯 CAD 객체 선택 모드")
            print("  AutoCAD에서 객체를 선택하세요...")
            
            # AutoCAD 활성화
            try:
                if hasattr(self, 'acad') and self.acad:
                    self.acad.Visible = True
                    # self.acad.WindowState = 1  # 이 부분이 오류를 일으킬 수 있음
            except:
                pass
            
            # 사용자 선택 대기
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
            
            # 테이블 업데이트
            # 수량 (2번 컬럼)
            if not self.item(row, 2):
                self.setItem(row, 2, QTableWidgetItem(""))
            self.item(row, 2).setText(str(selection.Count))
            print(f"  테이블 업데이트: 행 {row}, 수량 = {selection.Count}")
            
            # 사각형이 감지된 경우
            if rectangles:
                print(f"   사각형 {len(rectangles)}개 감지됨")
                # 첫 번째 사각형의 치수 사용
                rect = rectangles[0]
                print(f"   첫 번째 사각형 정보: {rect}")
                
                # 가로 (4번 컬럼)
                if not self.item(row, 4):
                    self.setItem(row, 4, QTableWidgetItem(""))
                self.item(row, 4).setText(f"{rect['width']:.1f}")
                print(f"   가로 설정: 행 {row}, 컬럼 4 = {rect['width']:.1f}")
                
                # 세로 (5번 컬럼)
                if not self.item(row, 5):
                    self.setItem(row, 5, QTableWidgetItem(""))
                self.item(row, 5).setText(f"{rect['height']:.1f}")
                print(f"   세로 설정: 행 {row}, 컬럼 5 = {rect['height']:.1f}")
                
                # 면적 (6번 컬럼)
                if not self.item(row, 6):
                    self.setItem(row, 6, QTableWidgetItem(""))
                self.item(row, 6).setText(f"{rect['area']:.1f}")
                print(f"   면적 설정: 행 {row}, 컬럼 6 = {rect['area']:.1f}")
                
                # 둘레 (7번 컬럼)
                if not self.item(row, 7):
                    self.setItem(row, 7, QTableWidgetItem(""))
                self.item(row, 7).setText(f"{rect['perimeter']:.1f}")
                print(f"   둘레 설정: 행 {row}, 컬럼 7 = {rect['perimeter']:.1f}")
            else:
                # 사각형이 아닌 경우
                if total_area > 0:
                    if not self.item(row, 6):
                        self.setItem(row, 6, QTableWidgetItem(""))
                    self.item(row, 6).setText(f"{total_area:.1f}")  # 면적
                    
                if total_perimeter > 0:
                    if not self.item(row, 7):
                        self.setItem(row, 7, QTableWidgetItem(""))
                    self.item(row, 7).setText(f"{total_perimeter:.1f}")  # 둘레
            
            # 선택된 객체 정보 저장
            if not hasattr(self, 'row_selections'):
                self.row_selections = {}
            
            # 선택된 객체 저장 (길이별 그룹화 전에 먼저 저장)
            self.row_selections[row] = selected_objects
            
            # Line 객체이고 여러 개인 경우 길이별로 그룹화
            if selected_objects and len(selected_objects) > 1:
                # 모든 객체가 Line인지 확인
                all_lines = True
                for obj in selected_objects:
                    obj_type = str(obj.ObjectName)
                    if not ("Line" in obj_type and "Polyline" not in obj_type):
                        all_lines = False
                        break
                
                if all_lines:
                    print(f"\n📊 Line 객체 {len(selected_objects)}개 - 길이별 그룹화 시도")
                    
                    # 길이별로 그룹화
                    groups = {}
                    for obj in selected_objects:
                        try:
                            start = obj.StartPoint
                            end = obj.EndPoint
                            length = math.sqrt((end[0]-start[0])**2 + (end[1]-start[1])**2)
                            length_key = round(length, 1)  # 0.1 단위로 반올림 (너무 세밀하면 그룹이 많아짐)
                            if length_key not in groups:
                                groups[length_key] = []
                            groups[length_key].append(obj)
                        except Exception as e:
                            print(f"    길이 계산 오류: {e}")
                    
                    print(f"  그룹화 결과: {len(groups)}개 그룹")
                    for key in sorted(groups.keys()):
                        print(f"    길이 {key:.1f}: {len(groups[key])}개")
                    
                    if len(groups) > 1:
                        print(f"  ✅ {len(groups)}개 그룹으로 분할하여 테이블에 추가")
                        
                        # 첫 번째 그룹은 현재 행에 업데이트
                        sorted_groups = sorted(groups.items())
                        first_key, first_objects = sorted_groups[0]
                        
                        self.row_selections[row] = first_objects
                        self.setItem(row, 1, QTableWidgetItem(f"L={first_key:.1f}"))  # 규격
                        self.setItem(row, 2, QTableWidgetItem(str(len(first_objects))))  # 수량
                        print(f"    행 {row} 업데이트: 길이={first_key:.1f}, 수량={len(first_objects)}")
                        
                        # 나머지 그룹은 새 행에 추가
                        original_name = self.item(row, 0).text() if self.item(row, 0) else ""
                        
                        for length_key, objects in sorted_groups[1:]:
                            new_row = self.rowCount()
                            self.add_row()
                            
                            # 데이터 설정
                            self.setItem(new_row, 0, QTableWidgetItem(original_name))  # 품명
                            self.setItem(new_row, 1, QTableWidgetItem(f"L={length_key:.1f}"))  # 규격
                            self.setItem(new_row, 2, QTableWidgetItem(str(len(objects))))  # 수량
                            self.setItem(new_row, 3, QTableWidgetItem("개"))  # 단위
                            
                            # 선택 객체 저장
                            self.row_selections[new_row] = objects
                            
                            # 버튼 다시 추가 (add_row에서 이미 추가되지만 확실히 하기 위해)
                            if not self.cellWidget(new_row, self.columnCount() - 1):
                                search_btn = QPushButton("🔍")
                                search_btn.setMaximumWidth(30)
                                search_btn.clicked.connect(lambda checked, r=new_row: self.parent.show_selection_helper(r))
                                self.setCellWidget(new_row, self.columnCount() - 1, search_btn)
                            
                            print(f"    행 {new_row} 추가: 길이={length_key:.1f}, 수량={len(objects)}")
                        
                        print(f"  완료: 총 {len(selected_objects)}개 객체를 {len(groups)}개 행으로 분할")
                    else:
                        print(f"  단일 그룹 (길이가 모두 동일)")
                else:
                    print(f"  Line이 아닌 객체 포함 (그룹화 안 함)")
            
            print(f"✅ {selection.Count}개 객체 선택됨")
            if rectangles:
                print(f"   사각형 {len(rectangles)}개 감지")
                for i, rect in enumerate(rectangles[:3]):  # 최대 3개만 표시
                    print(f"   사각형{i+1}: {rect['width']:.1f} x {rect['height']:.1f}")
            
            selection.Delete()
            
        except Exception as e:
            print(f"❌ CAD 선택 오류: {e}")
        finally:
            pythoncom.CoUninitialize()
        
    def show_selection_helper(self, row):
        """선택 도우미"""
        print(f"\n🔍 선택 도우미 호출 - 행: {row}")
        print(f"  row_selections 존재: {hasattr(self, 'row_selections')}")
        
        if hasattr(self, 'row_selections'):
            print(f"  row_selections 내용: {list(self.row_selections.keys())}")
            print(f"  현재 행({row})이 저장되어 있는가: {row in self.row_selections}")
        
        if not hasattr(self, 'row_selections'):
            QMessageBox.information(self, "안내", 
                "선택된 객체가 없습니다.\n먼저 선택 버튼을 눌러 CAD 객체를 선택하세요.")
            return
            
        if row not in self.row_selections:
            QMessageBox.information(self, "안내", 
                f"행 {row}에 선택된 객체가 없습니다.\n먼저 선택 버튼을 눌러 CAD 객체를 선택하세요.")
            return
            
        selected_objects = self.row_selections[row]
        
        print(f"  선택된 객체 수: {len(selected_objects)}")
        
        # parent_widget(메인 윈도우)의 doc 사용
        if not self.parent_widget or not hasattr(self.parent_widget, 'doc') or not self.parent_widget.doc:
            QMessageBox.warning(self, "경고", "AutoCAD 연결이 필요합니다.")
            return
        
        self.doc = self.parent_widget.doc  # parent의 doc 사용
        
        print("  선택 도우미 대화상자 생성 중...")
        
        try:
            # 선택 도우미 대화상자 표시
            dialog = SelectionHelperDialog(self, self.doc, selected_objects, row)
            print("  대화상자 생성 완료, 표시 중...")
            
            dialog_result = dialog.exec_()
            print(f"  대화상자 결과: {dialog_result}")
            
            if dialog_result:
                # 선택 결과 업데이트
                new_selection = dialog.get_final_selection()
                print(f"  new_selection 크기: {len(new_selection) if new_selection else 0}")
                
                if new_selection:
                    print(f"\n📊 최종 선택: {len(new_selection)}개 객체")
                    
                    # 선(Line) 객체인지 확인
                    is_line_type = False
                    if new_selection:
                        first_obj_type = str(new_selection[0].ObjectName)
                        if "Line" in first_obj_type and "Polyline" not in first_obj_type:
                            is_line_type = True
                            print(f"  타입: Line 객체 - 길이별 그룹화 적용")
                    
                    # Line 객체인 경우 길이별로 그룹화
                    if is_line_type:
                        # dialog의 current_selection이 이미 업데이트되었으므로 그대로 사용
                        groups = dialog.group_by_length()
                        
                        # 그룹 정보 출력
                        print(f"\n📦 그룹화 결과: {len(groups)}개 그룹")
                        for key, objs in sorted(groups.items(), key=lambda x: (x[0] if isinstance(x[0], (int, float)) else float('inf'))):
                            if isinstance(key, (int, float)):
                                print(f"  - 길이 {key:.2f}: {len(objs)}개")
                            else:
                                print(f"  - {key}: {len(objs)}개")
                    else:
                        # Line이 아닌 경우 전체를 하나의 그룹으로
                        groups = {'all': new_selection}
                        print(f"  타입: {first_obj_type if new_selection else 'Unknown'} - 단일 그룹")
                    
                    # 그룹이 여러 개인 경우 행 분할
                    if len(groups) > 1:
                        # 여러 그룹이 있으면 각각 다른 행에 추가
                        print(f"\n✅ {len(groups)}개 그룹으로 분할하여 각 행에 추가")
                        
                        # 현재 행 업데이트 (첫 번째 그룹)
                        first_group = True
                        current_row = row
                        
                        # 길이로 정렬하여 처리
                        sorted_groups = sorted(groups.items(), key=lambda x: (x[0] if isinstance(x[0], (int, float)) else float('inf')))
                        
                        for idx, (length_key, objects) in enumerate(sorted_groups):
                            if first_group:
                                # 첫 번째 그룹은 현재 행에
                                self.row_selections[current_row] = objects
                                if self.item(current_row, 2):
                                    self.item(current_row, 2).setText(str(len(objects)))
                                
                                # 규격 컬럼에 길이 표시
                                if isinstance(length_key, (int, float)):
                                    if self.item(current_row, 1):
                                        self.item(current_row, 1).setText(f"L={length_key:.2f}")
                                
                                print(f"  행 {current_row}: 길이={length_key}, 수량={len(objects)}개")
                                first_group = False
                            else:
                                # 나머지 그룹은 새로운 행에 추가
                                new_row = self.rowCount()
                                self.add_row()  # 새 행 추가
                                
                                # 품명 복사 (원래 행에서)
                                if self.item(row, 0):
                                    original_name = self.item(row, 0).text()
                                    self.setItem(new_row, 0, QTableWidgetItem(original_name))
                                
                                # 규격에 길이 표시
                                if isinstance(length_key, (int, float)):
                                    self.setItem(new_row, 1, QTableWidgetItem(f"L={length_key:.2f}"))
                                else:
                                    self.setItem(new_row, 1, QTableWidgetItem(str(length_key)))
                                
                                # 수량 설정
                                self.setItem(new_row, 2, QTableWidgetItem(str(len(objects))))
                                
                                # 단위 설정
                                self.setItem(new_row, 3, QTableWidgetItem("개"))
                                
                                # 선택 객체 저장
                                self.row_selections[new_row] = objects
                                
                                # 돋보기 버튼 추가
                                search_btn = QPushButton("🔍")
                                search_btn.setMaximumWidth(30)
                                # parent는 CADQuantityProWindow 인스턴스
                                parent_window = self.parent
                                search_btn.clicked.connect(lambda checked, r=new_row: parent_window.show_selection_helper(r))
                                self.setCellWidget(new_row, self.columnCount() - 1, search_btn)
                                
                                # 행 색상 구분 (번갈아가면서)
                                if new_row % 2 == 1:
                                    for col in range(self.columnCount()):
                                        if self.item(new_row, col):
                                            self.item(new_row, col).setBackground(QColor(240, 240, 255))
                                
                                print(f"  행 {new_row}: 길이={length_key}, 수량={len(objects)}개")
                        
                        print(f"✅ 총 {len(new_selection)}개 객체가 {len(groups)}개 행으로 분할됨")
                    else:
                        # 단일 그룹인 경우 기존 방식대로
                        self.row_selections[row] = new_selection
                        if self.item(row, 2):
                            self.item(row, 2).setText(str(len(new_selection)))
                        print(f"✅ 선택 업데이트: {len(new_selection)}개")
            else:
                print("  대화상자 취소됨")
        except Exception as e:
            print(f"❌ 선택 도우미 오류: {e}")
            import traceback
            traceback.print_exc()
            QMessageBox.critical(self, "오류", f"선택 도우미 오류:\n{str(e)}")


# ==================== 선택 도우미 대화상자 ====================

class SelectionHelperDialog(QDialog):
    """선택 도우미 대화상자"""
    
    def __init__(self, parent, doc, current_selection, row):
        super().__init__(parent)
        self.doc = doc
        self.current_selection = current_selection
        self.row = row
        self.found_objects = []
        self.setup_ui()
        
    def setup_ui(self):
        """UI 설정"""
        self.setWindowTitle("🔍 선택 도우미")
        self.setModal(True)
        
        # 화면 크기에 맞춰 대화상자 크기 조정
        screen = QApplication.desktop().screenGeometry()
        if screen.height() <= 1080:  # FHD 이하
            dialog_height = min(600, int(screen.height() * 0.7))
            self.setGeometry(200, 50, 500, dialog_height)
        else:
            self.setGeometry(200, 200, 500, 600)
        
        # 메인 레이아웃
        main_layout = QVBoxLayout(self)
        
        # 메인 스크롤 영역 설정
        main_scroll_area = QScrollArea()
        main_scroll_area.setWidgetResizable(True)
        main_scroll_widget = QWidget()
        layout = QVBoxLayout(main_scroll_widget)
        
        # 현재 선택 정보
        info_group = QGroupBox("현재 선택")
        info_layout = QVBoxLayout()
        
        self.info_label = QLabel(f"선택된 객체: {len(self.current_selection)}개")
        info_layout.addWidget(self.info_label)
        
        if self.current_selection:
            obj = self.current_selection[0]
            try:
                obj_type = str(obj.ObjectName).replace("AcDb", "")
                layer = str(obj.Layer)
                self.info_label.setText(
                    f"선택된 객체: {len(self.current_selection)}개\n"
                    f"기준 객체: {obj_type}\n"
                    f"레이어: {layer}"
                )
            except:
                pass
                
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)
        
        # 찾기 옵션
        option_group = QGroupBox("찾기 옵션")
        option_layout = QVBoxLayout()
        
        self.same_layer = QCheckBox("같은 레이어")
        self.same_layer.setChecked(True)
        option_layout.addWidget(self.same_layer)
        
        self.same_type = QCheckBox("같은 타입")
        self.same_type.setChecked(True)
        option_layout.addWidget(self.same_type)
        
        self.same_color = QCheckBox("같은 색상")
        option_layout.addWidget(self.same_color)
        
        self.same_size = QCheckBox("같은 크기 (±10%)")
        option_layout.addWidget(self.same_size)
        
        # 블록인 경우 같은 블록 찾기
        if self.current_selection and "BlockReference" in str(self.current_selection[0].ObjectName):
            self.same_block = QCheckBox("같은 블록 이름")
            self.same_block.setChecked(True)
            option_layout.addWidget(self.same_block)
        else:
            self.same_block = None
            
        option_group.setLayout(option_layout)
        layout.addWidget(option_group)
        
        # 영역 선택 옵션
        area_group = QGroupBox("🗺️ 영역 선택 (선택사항)")
        area_layout = QVBoxLayout()
        
        self.use_area = QCheckBox("특정 영역 내에서만 검색")
        area_layout.addWidget(self.use_area)
        
        # 영역 설정 버튼
        area_btn_layout = QHBoxLayout()
        
        self.set_area_btn = QPushButton("📏 영역 설정 (AutoCAD에서 두 점 선택)")
        self.set_area_btn.clicked.connect(self.set_search_area)
        self.set_area_btn.setEnabled(False)
        area_btn_layout.addWidget(self.set_area_btn)
        
        area_layout.addLayout(area_btn_layout)
        
        # 영역 정보 표시
        self.area_info_label = QLabel("영역이 설정되지 않음")
        self.area_info_label.setStyleSheet("color: gray; font-style: italic;")
        area_layout.addWidget(self.area_info_label)
        
        area_group.setLayout(area_layout)
        layout.addWidget(area_group)
        
        # 체크박스 상태 변경 시 버튼 활성화
        self.use_area.stateChanged.connect(lambda state: self.set_area_btn.setEnabled(state == Qt.Checked))
        
        # 찾기 버튼
        find_btn = QPushButton("🔍 유사 객체 찾기")
        find_btn.clicked.connect(self.find_similar)
        layout.addWidget(find_btn)
        
        # 결과
        result_group = QGroupBox("찾기 결과")
        result_layout = QVBoxLayout()
        
        self.result_label = QLabel("찾기를 클릭하세요")
        result_layout.addWidget(self.result_label)
        
        # 결과 스크롤 영역에 체크박스 리스트 추가
        result_scroll_area = QScrollArea()
        result_scroll_widget = QWidget()
        self.result_layout = QVBoxLayout(result_scroll_widget)
        self.checkboxes = []
        
        result_scroll_area.setWidget(result_scroll_widget)
        result_scroll_area.setWidgetResizable(True)
        result_layout.addWidget(result_scroll_area)
        
        result_group.setLayout(result_layout)
        layout.addWidget(result_group)
        
        # 선택 제어 버튼
        select_control_layout = QHBoxLayout()
        
        select_all_btn = QPushButton("☑️ 전체 선택")
        select_all_btn.clicked.connect(self.select_all)
        select_control_layout.addWidget(select_all_btn)
        
        deselect_all_btn = QPushButton("⬜ 전체 해제")
        deselect_all_btn.clicked.connect(self.deselect_all)
        select_control_layout.addWidget(deselect_all_btn)
        
        invert_btn = QPushButton("🔄 선택 반전")
        invert_btn.clicked.connect(self.invert_selection)
        select_control_layout.addWidget(invert_btn)
        
        layout.addLayout(select_control_layout)
        
        # 선택 모드 옵션
        mode_group = QGroupBox("선택 모드")
        mode_layout = QHBoxLayout()
        
        self.replace_mode = QRadioButton("현재 선택 대체")
        self.replace_mode.setChecked(True)  # 대체 모드를 기본값으로
        mode_layout.addWidget(self.replace_mode)
        
        self.add_mode = QRadioButton("현재 선택에 추가")
        mode_layout.addWidget(self.add_mode)
        
        mode_group.setLayout(mode_layout)
        layout.addWidget(mode_group)
        
        # 메인 스크롤 위젯 설정 및 메인 레이아웃에 추가
        main_scroll_area.setWidget(main_scroll_widget)
        main_layout.addWidget(main_scroll_area)
        
        # 다이얼로그 버튼 (스크롤 영역 밖에 배치)
        dialog_btn_layout = QHBoxLayout()
        
        ok_btn = QPushButton("✅ 확인")
        ok_btn.clicked.connect(self.accept_and_apply)
        dialog_btn_layout.addWidget(ok_btn)
        
        cancel_btn = QPushButton("❌ 취소")
        cancel_btn.clicked.connect(self.reject)
        dialog_btn_layout.addWidget(cancel_btn)
        
        main_layout.addLayout(dialog_btn_layout)
        
    def set_search_area(self):
        """검색 영역 설정"""
        try:
            import pythoncom
            pythoncom.CoInitialize()
            
            print("\n📏 영역 설정 모드")
            print("  AutoCAD에서 대각선 모서리 두 점을 선택하세요...")
            
            # AutoCAD에서 두 점 선택
            # 첫 번째 점 선택
            print("  첫 번째 모서리 점을 클릭하세요...")
            point1 = self.doc.Utility.GetPoint()
            print(f"  첫 번째 점: {point1[0]:.1f}, {point1[1]:.1f}")
            
            # 두 번째 점 선택 - GetPoint를 다시 사용
            print("  두 번째 모서리 점을 클릭하세요...")
            point2 = self.doc.Utility.GetPoint()
            print(f"  두 번째 점: {point2[0]:.1f}, {point2[1]:.1f}")
            
            # 영역 좌표 저장
            self.search_area = {
                'x1': min(point1[0], point2[0]),
                'y1': min(point1[1], point2[1]),
                'x2': max(point1[0], point2[0]),
                'y2': max(point1[1], point2[1])
            }
            
            # 영역 정보 표시
            width = self.search_area['x2'] - self.search_area['x1']
            height = self.search_area['y2'] - self.search_area['y1']
            self.area_info_label.setText(
                f"영역 설정됨: {width:.1f} x {height:.1f}\n"
                f"X: {self.search_area['x1']:.1f} ~ {self.search_area['x2']:.1f}\n"
                f"Y: {self.search_area['y1']:.1f} ~ {self.search_area['y2']:.1f}"
            )
            self.area_info_label.setStyleSheet("color: green; font-weight: bold;")
            
            print(f"  ✅ 영역 설정: {width:.1f} x {height:.1f}")
            
            # 시각적 표시를 위한 사각형 그리기 (선택사항)
            try:
                # 임시 폴리라인으로 영역 표시
                points = [
                    self.search_area['x1'], self.search_area['y1'], 0,
                    self.search_area['x2'], self.search_area['y1'], 0,
                    self.search_area['x2'], self.search_area['y2'], 0,
                    self.search_area['x1'], self.search_area['y2'], 0,
                    self.search_area['x1'], self.search_area['y1'], 0
                ]
                import win32com.client
                points_var = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, points)
                temp_rect = self.doc.ModelSpace.AddPolyline(points_var)
                temp_rect.Color = 1  # 빨간색
                temp_rect.LineWeight = 30  # 두께
                
                # 잠시 후 삭제
                import time
                time.sleep(1)
                temp_rect.Delete()
            except:
                pass
                
        except Exception as e:
            print(f"  ❌ 영역 설정 오류: {e}")
            QMessageBox.warning(self, "오류", f"영역 설정 중 오류:\n{str(e)}")
        finally:
            pythoncom.CoUninitialize()
    
    def is_in_area(self, obj):
        """객체가 설정된 영역 내에 있는지 확인"""
        if not hasattr(self, 'search_area'):
            return True  # 영역이 설정되지 않으면 모든 객체 포함
            
        try:
            obj_type = str(obj.ObjectName)
            x1, y1 = self.search_area['x1'], self.search_area['y1']
            x2, y2 = self.search_area['x2'], self.search_area['y2']
            
            # 점 기반 객체
            if hasattr(obj, 'InsertionPoint'):
                pt = obj.InsertionPoint
                return x1 <= pt[0] <= x2 and y1 <= pt[1] <= y2
                
            # LINE
            elif "Line" in obj_type and "Polyline" not in obj_type:
                start = obj.StartPoint
                end = obj.EndPoint
                # 양 끝점이 영역 내에 있는지
                return (x1 <= start[0] <= x2 and y1 <= start[1] <= y2 and
                       x1 <= end[0] <= x2 and y1 <= end[1] <= y2)
                
            # CIRCLE
            elif "Circle" in obj_type:
                center = obj.Center
                return x1 <= center[0] <= x2 and y1 <= center[1] <= y2
                
            # POLYLINE
            elif "Polyline" in obj_type:
                coords = obj.Coordinates
                # 모든 점이 영역 내에 있는지
                for i in range(0, len(coords), 2):
                    x, y = coords[i], coords[i+1]
                    if not (x1 <= x <= x2 and y1 <= y <= y2):
                        return False
                return True
                
            # BLOCK
            elif "BlockReference" in obj_type:
                pt = obj.InsertionPoint
                return x1 <= pt[0] <= x2 and y1 <= pt[1] <= y2
                
        except:
            pass
            
        return True  # 확인할 수 없는 객체는 포함
    
    def find_similar(self):
        """유사 객체 찾기"""
        if not self.current_selection:
            return
            
        base_obj = self.current_selection[0]
        self.found_objects = []
        
        try:
            pythoncom.CoInitialize()
            
            # 기준 객체 속성
            base_type = str(base_obj.ObjectName)
            base_layer = str(base_obj.Layer)
            base_color = base_obj.color if hasattr(base_obj, 'color') else None
            
            print(f"\n🔍 유사 객체 찾기 시작")
            print(f"  기준 타입: {base_type}")
            print(f"  기준 레이어: {base_layer}")
            print(f"  기준 색상: {base_color}")
            
            # 블록인 경우
            base_block_name = None
            if "BlockReference" in base_type:
                base_block_name = str(base_obj.Name)
                print(f"  기준 블록 이름: {base_block_name}")
                
            # 크기 정보
            base_size = None
            try:
                # 폴리라인의 경우 면적 계산
                if "Polyline" in base_type:
                    if hasattr(base_obj, 'Closed') and base_obj.Closed:
                        coords = base_obj.Coordinates
                        if len(coords) >= 8:
                            x_coords = [coords[i] for i in range(0, len(coords), 2)]
                            y_coords = [coords[i] for i in range(1, len(coords), 2)]
                            width = max(x_coords) - min(x_coords)
                            height = max(y_coords) - min(y_coords)
                            base_size = width * height
                            print(f"  기준 폴리라인 크기: {width:.2f} x {height:.2f} = {base_size:.2f}")
                elif hasattr(base_obj, 'Radius'):
                    # 원의 경우
                    base_size = 3.14159 * base_obj.Radius * base_obj.Radius
                    print(f"  기준 원 크기: 반지름 {base_obj.Radius:.2f}, 면적 {base_size:.2f}")
                elif hasattr(base_obj, 'Length'):
                    base_size = base_obj.Length
                    print(f"  기준 길이: {base_size:.2f}")
                elif hasattr(base_obj, 'Area'):
                    base_size = base_obj.Area
                    print(f"  기준 면적: {base_size:.2f}")
            except Exception as e:
                print(f"  크기 계산 오류: {e}")
                
            # 기준 객체의 Handle 저장
            base_handle = base_obj.Handle
            print(f"  기준 객체 Handle: {base_handle}")
            
            # 모델 공간 검색
            count = 0
            
            # 먼저 기준 객체 자체를 추가 (대체 모드의 경우 포함되어야 함)
            self.found_objects.append(base_obj)
            count = 1
            
            for i in range(self.doc.ModelSpace.Count):
                obj = self.doc.ModelSpace.Item(i)
                
                # 기준 객체 자신은 이미 추가했으므로 제외
                if obj.Handle == base_handle:
                    continue
                
                # 각 조건을 순차적으로 확인하고 하나라도 만족하지 않으면 건너뛰기
                should_include = True
                
                # 영역 체크 (영역이 설정된 경우만)
                if should_include and self.use_area.isChecked() and hasattr(self, 'search_area'):
                    if not self.is_in_area(obj):
                        should_include = False
                
                # 같은 타입 체크
                if self.same_type.isChecked():
                    if str(obj.ObjectName) != base_type:
                        should_include = False
                        
                # 같은 레이어 체크
                if should_include and self.same_layer.isChecked():
                    if str(obj.Layer) != base_layer:
                        should_include = False
                        
                # 같은 색상 체크
                if should_include and self.same_color.isChecked():
                    if base_color is not None:
                        obj_color = obj.color if hasattr(obj, 'color') else None
                        if obj_color != base_color:
                            should_include = False
                    else:
                        # 기준 객체에 색상이 없으면 색상 비교 건너뛰기
                        pass
                            
                # 같은 블록 체크 (블록일 때만)
                if should_include and self.same_block and self.same_block.isChecked():
                    # 블록이 아닌 객체는 제외
                    if "BlockReference" not in str(obj.ObjectName):
                        should_include = False
                    else:
                        # 블록 이름 비교
                        if str(obj.Name) != base_block_name:
                            should_include = False
                            
                # 같은 크기 체크
                if should_include and self.same_size.isChecked():
                    if base_size is not None:
                        obj_size = None
                        try:
                            # 폴리라인의 경우 면적 계산
                            if "Polyline" in str(obj.ObjectName):
                                if hasattr(obj, 'Closed') and obj.Closed:
                                    coords = obj.Coordinates
                                    if len(coords) >= 8:
                                        x_coords = [coords[i] for i in range(0, len(coords), 2)]
                                        y_coords = [coords[i] for i in range(1, len(coords), 2)]
                                        width = max(x_coords) - min(x_coords)
                                        height = max(y_coords) - min(y_coords)
                                        obj_size = width * height
                            elif "Circle" in str(obj.ObjectName) and hasattr(obj, 'Radius'):
                                # 원의 경우 면적 계산
                                obj_size = 3.14159 * obj.Radius * obj.Radius
                            elif hasattr(obj, 'Length'):
                                obj_size = obj.Length
                            elif hasattr(obj, 'Area'):
                                obj_size = obj.Area
                        except:
                            pass
                        
                        if obj_size is None:
                            # 크기를 측정할 수 없는 객체는 제외
                            should_include = False
                        else:
                            # ±10% 허용
                            if abs(obj_size - base_size) / base_size > 0.1:
                                should_include = False
                    else:
                        # 기준 객체에 크기 정보가 없으면 크기 비교 건너뛰기
                        pass
                
                # 모든 조건을 만족하면 추가
                if should_include:
                    self.found_objects.append(obj)
                    count += 1
                
            # 결과 표시
            print(f"\n✅ 찾기 완료: 총 {count}개 객체")
            print(f"  선택된 옵션:")
            if self.same_type.isChecked():
                print(f"    - 같은 타입: {base_type}")
            if self.same_layer.isChecked():
                print(f"    - 같은 레이어: {base_layer}")
            if self.same_color.isChecked():
                print(f"    - 같은 색상: {base_color}")
            if self.same_block and self.same_block.isChecked():
                print(f"    - 같은 블록: {base_block_name}")
            if self.same_size.isChecked():
                print(f"    - 같은 크기 (±10%): {base_size}")
            if self.use_area.isChecked() and hasattr(self, 'search_area'):
                print(f"    - 영역 제한: X({self.search_area['x1']:.1f}~{self.search_area['x2']:.1f}), Y({self.search_area['y1']:.1f}~{self.search_area['y2']:.1f})")
            
            self.result_label.setText(f"찾은 객체: {count}개")
            
            # 기존 체크박스 제거
            for checkbox in self.checkboxes:
                checkbox.setParent(None)
            self.checkboxes.clear()
            
            # 새 체크박스 생성
            for i, obj in enumerate(self.found_objects[:100]):  # 최대 100개 표시
                try:
                    obj_type = str(obj.ObjectName).replace("AcDb", "")
                    layer = str(obj.Layer)
                    
                    # 치수 정보 추출
                    size_info = ""
                    
                    # 폴리라인인 경우
                    if "Polyline" in obj_type:
                        try:
                            # 폐합된 폴리라인이면 사각형 검사
                            if hasattr(obj, 'Closed') and obj.Closed:
                                coords = obj.Coordinates
                                if len(coords) >= 8:
                                    x_coords = [coords[j] for j in range(0, len(coords), 2)]
                                    y_coords = [coords[j] for j in range(1, len(coords), 2)]
                                    width = max(x_coords) - min(x_coords)
                                    height = max(y_coords) - min(y_coords)
                                    
                                    if len(coords) in [8, 10]:  # 사각형
                                        size_info = f" | 📐 {width:.1f} x {height:.1f}"
                                    else:
                                        area = obj.Area if hasattr(obj, 'Area') else 0
                                        size_info = f" | 면적: {area:.1f}"
                            else:
                                # 열린 폴리라인
                                length = obj.Length if hasattr(obj, 'Length') else 0
                                size_info = f" | 길이: {length:.1f}"
                        except:
                            pass
                    
                    # 원인 경우
                    elif "Circle" in obj_type:
                        try:
                            radius = obj.Radius
                            size_info = f" | 반지름: {radius:.1f}"
                        except:
                            pass
                    
                    # 선인 경우
                    elif "Line" in obj_type and "Polyline" not in obj_type:
                        try:
                            start = obj.StartPoint
                            end = obj.EndPoint
                            import math
                            length = math.sqrt((end[0]-start[0])**2 + (end[1]-start[1])**2)
                            size_info = f" | 길이: {length:.1f}"
                        except:
                            pass
                    
                    # 블록인 경우
                    elif "BlockReference" in obj_type:
                        try:
                            block_name = str(obj.Name)
                            size_info = f" | 블록: {block_name}"
                            # 스케일 정보가 있으면 추가
                            if hasattr(obj, 'XScaleFactor'):
                                scale = obj.XScaleFactor
                                if scale != 1.0:
                                    size_info += f" (스케일: {scale:.2f})"
                        except:
                            pass
                    
                    # 체크박스 생성
                    checkbox_text = f"{i+1}. {obj_type} [{layer}]{size_info}"
                    checkbox = QCheckBox(checkbox_text)
                    checkbox.setChecked(True)  # 기본적으로 체크
                    self.checkboxes.append(checkbox)
                    self.result_layout.addWidget(checkbox)
                except Exception as e:
                    print(f"항목 생성 오류: {e}")
                    
            if count > 100:
                label = QLabel(f"... 외 {count-100}개 (최대 100개만 표시)")
                self.result_layout.addWidget(label)
                
        except Exception as e:
            QMessageBox.warning(self, "오류", f"검색 중 오류: {str(e)}")
        finally:
            pythoncom.CoUninitialize()
            
    def accept_and_apply(self):
        """확인 버튼 - 체크된 항목 적용 후 닫기"""
        # 체크박스가 없으면 (유사 객체 찾기를 하지 않은 경우) found_objects를 그대로 사용
        if not self.checkboxes and self.found_objects:
            print(f"  체크박스 없음 - found_objects 그대로 사용: {len(self.found_objects)}개")
            if self.replace_mode.isChecked():
                self.current_selection = self.found_objects[:]
            else:
                # 추가 모드: 중복 제거하여 추가
                existing_handles = set()
                for obj in self.current_selection:
                    try:
                        existing_handles.add(obj.Handle)
                    except:
                        pass
                for obj in self.found_objects:
                    try:
                        if obj.Handle not in existing_handles:
                            self.current_selection.append(obj)
                    except:
                        self.current_selection.append(obj)
        elif self.checkboxes:
            # 체크박스가 있는 경우 기존 로직 사용
            if self.replace_mode.isChecked():
                # 대체 모드: 체크된 항목으로 선택 대체
                self.current_selection = []
                for i, checkbox in enumerate(self.checkboxes):
                    if checkbox.isChecked() and i < len(self.found_objects):
                        self.current_selection.append(self.found_objects[i])
            else:
                # 추가 모드: 체크된 항목을 현재 선택에 추가 (중복 제거)
                # 현재 선택된 객체의 Handle 목록 생성
                existing_handles = set()
                for obj in self.current_selection:
                    try:
                        existing_handles.add(obj.Handle)
                    except:
                        pass
                
                # 체크된 항목 추가 (중복 제거)
                for i, checkbox in enumerate(self.checkboxes):
                    if checkbox.isChecked() and i < len(self.found_objects):
                        obj = self.found_objects[i]
                        try:
                            if obj.Handle not in existing_handles:
                                self.current_selection.append(obj)
                                existing_handles.add(obj.Handle)
                        except:
                            # Handle이 없는 경우 그냥 추가
                            self.current_selection.append(obj)
        
        print(f"  최종 current_selection: {len(self.current_selection)}개")
        
        # 다이얼로그 닫기
        self.accept()
            
    def select_all(self):
        """전체 선택"""
        for checkbox in self.checkboxes:
            checkbox.setChecked(True)
            
    def deselect_all(self):
        """전체 해제"""
        for checkbox in self.checkboxes:
            checkbox.setChecked(False)
            
    def invert_selection(self):
        """선택 반전"""
        for checkbox in self.checkboxes:
            checkbox.setChecked(not checkbox.isChecked())
        
    def get_final_selection(self):
        """최종 선택 반환"""
        return self.current_selection
    
    def group_by_length(self):
        """길이별로 객체 그룹화"""
        import math
        groups = {}
        
        print(f"\n📏 길이별 그룹화 시작 (객체 수: {len(self.current_selection)})")
        
        # 길이 목록 수집 (디버깅용)
        lengths = []
        
        for i, obj in enumerate(self.current_selection):
            try:
                obj_type = str(obj.ObjectName)
                length = None
                
                # LINE 객체의 길이 계산
                if "Line" in obj_type and "Polyline" not in obj_type:
                    start = obj.StartPoint
                    end = obj.EndPoint
                    length = math.sqrt((end[0]-start[0])**2 + (end[1]-start[1])**2)
                    lengths.append(length)
                    print(f"  [{i+1}] Line 길이: {length:.3f}")
                # 폴리라인의 길이
                elif "Polyline" in obj_type:
                    if hasattr(obj, 'Length'):
                        length = obj.Length
                        lengths.append(length)
                        print(f"  [{i+1}] Polyline 길이: {length:.3f}")
                # 원의 둘레
                elif "Circle" in obj_type:
                    if hasattr(obj, 'Radius'):
                        length = 2 * math.pi * obj.Radius
                        lengths.append(length)
                        print(f"  [{i+1}] Circle 둘레: {length:.3f}")
                
                if length is not None:
                    # 길이를 0.01 단위로 반올림하여 그룹화 (더 정밀한 그룹화)
                    length_key = round(length, 2)
                    if length_key not in groups:
                        groups[length_key] = []
                        print(f"    → 새 그룹 생성: {length_key:.2f}")
                    groups[length_key].append(obj)
                else:
                    # 길이가 없는 객체는 'other' 그룹으로
                    if 'other' not in groups:
                        groups['other'] = []
                    groups['other'].append(obj)
                    print(f"  [{i+1}] {obj_type} - 길이 없음")
                    
            except Exception as e:
                print(f"  [{i+1}] 그룹화 오류: {e}")
                if 'error' not in groups:
                    groups['error'] = []
                groups['error'].append(obj)
        
        # 유니크한 길이 확인
        if lengths:
            unique_lengths = len(set([round(l, 2) for l in lengths]))
            print(f"\n📊 길이 분석: 총 {len(lengths)}개 중 유니크한 길이 {unique_lengths}개")
        
        print(f"\n📦 그룹화 완료: {len(groups)}개 그룹")
        for key, objs in sorted(groups.items(), key=lambda x: (x[0] if isinstance(x[0], (int, float)) else float('inf'))):
            if isinstance(key, (int, float)):
                print(f"  - 길이 {key:.2f}: {len(objs)}개 객체")
            else:
                print(f"  - {key}: {len(objs)}개 객체")
        
        return groups


# ==================== 메인 윈도우 ====================

class CADQuantityProWindow(QMainWindow):
    """메인 윈도우"""
    
    def __init__(self):
        super().__init__()
        self.acad = None
        self.doc = None
        # 계층구조 모드를 기본으로 설정
        self.current_mode = "hierarchical" if HIERARCHICAL_TABLE_AVAILABLE else "flat"
        self.init_ui()
        
        # 콘솔 리디렉션
        sys.stdout = ConsoleRedirect(self.console_widget)
        
        print("=" * 60)
        print("CAD Quantity Pro - Hierarchical Version")
        print("=" * 60)
        print("\n기능:")
        print("- 계층구조 테이블 (대분류/중분류/항목)")
        print("- CAD 객체 선택 (🎯 버튼)")
        print("- 유사 객체 찾기 (🔍 버튼)")
        print("- 길이별 자동 그룹화")
        print("- 사각형 자동 감지")
        print("=" * 60)
        
        # 초기 모드가 계층구조면 전환
        if self.current_mode == "hierarchical" and HIERARCHICAL_TABLE_AVAILABLE:
            self.switch_to_hierarchical()
        
    def init_ui(self):
        """UI 초기화"""
        self.setWindowTitle("CAD Quantity Pro - 간단한 계층구조")
        
        # 화면 크기에 따라 윈도우 크기 조정
        screen = QApplication.desktop().screenGeometry()
        if screen.height() <= 1080:  # FHD 이하
            # 화면의 90% 크기로 설정
            width = int(screen.width() * 0.9)
            height = int(screen.height() * 0.85)
            self.setGeometry(50, 30, width, height)
        else:  # QHD 이상
            self.setGeometry(100, 100, 1400, 900)
        
        # 중앙 위젯
        central = QWidget()
        self.setCentralWidget(central)
        
        # 메인 레이아웃 (수평 분할)
        main_layout = QHBoxLayout(central)
        
        # 왼쪽: 테이블 영역
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        
        # 상단 툴바
        toolbar = self.create_toolbar()
        left_layout.addLayout(toolbar)
        
        # 모드 전환 버튼
        mode_layout = QHBoxLayout()
        
        self.flat_btn = QPushButton("📊 평면 테이블")
        self.flat_btn.setCheckable(True)
        self.flat_btn.setChecked(True)
        self.flat_btn.clicked.connect(self.switch_to_flat)
        mode_layout.addWidget(self.flat_btn)
        
        if HIERARCHICAL_TABLE_AVAILABLE:
            self.hierarchical_btn = QPushButton("🔢 계층구조 테이블")
            self.hierarchical_btn.setCheckable(True)
            self.hierarchical_btn.clicked.connect(self.switch_to_hierarchical)
            mode_layout.addWidget(self.hierarchical_btn)
        
        mode_layout.addStretch()
        left_layout.addLayout(mode_layout)
        
        # 스택 위젯 (테이블 전환용) - 스크롤 영역 추가
        self.stacked_widget = QStackedWidget()
        
        # 평면 테이블 - 스크롤 영역에 감싸기
        flat_scroll = QScrollArea()
        self.flat_table = FlatQuantityTable(self)
        
        # FHD 화면용 크기 설정
        if screen.height() <= 1080:
            self.flat_table.setMinimumHeight(400)
            self.flat_table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        
        flat_scroll.setWidget(self.flat_table)
        flat_scroll.setWidgetResizable(True)
        flat_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        flat_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.stacked_widget.addWidget(flat_scroll)
        
        # 계층구조 테이블 - 스크롤 영역에 감싸기
        if HIERARCHICAL_TABLE_AVAILABLE:
            hierarchical_scroll = QScrollArea()
            self.hierarchical_table = SimpleHierarchicalTable(self)
            self.hierarchical_table.parent_widget = self
            
            # FHD 화면용 크기 설정
            if screen.height() <= 1080:
                self.hierarchical_table.setMinimumHeight(400)
                self.hierarchical_table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
            
            hierarchical_scroll.setWidget(self.hierarchical_table)
            hierarchical_scroll.setWidgetResizable(True)
            hierarchical_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
            hierarchical_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
            self.stacked_widget.addWidget(hierarchical_scroll)
        
        left_layout.addWidget(self.stacked_widget)
        
        # 오른쪽: 콘솔 영역
        right_widget = QWidget()
        right_widget.setMaximumWidth(400)
        right_layout = QVBoxLayout(right_widget)
        
        # 콘솔 헤더
        console_header = QHBoxLayout()
        console_header.addWidget(QLabel("📝 디버그 콘솔"))
        console_header.addStretch()
        
        clear_btn = QPushButton("🗑 지우기")
        clear_btn.clicked.connect(lambda: self.console_widget.clear())
        console_header.addWidget(clear_btn)
        
        right_layout.addLayout(console_header)
        
        # 콘솔 위젯
        self.console_widget = QTextEdit()
        self.console_widget.setReadOnly(True)
        self.console_widget.setStyleSheet("""
            QTextEdit {
                background-color: #1e1e1e;
                color: #00ff00;
                font-family: 'Consolas', 'Courier New', monospace;
                font-size: 10pt;
                border: 1px solid #555;
            }
        """)
        right_layout.addWidget(self.console_widget)
        
        # 레이아웃 추가 (왼쪽 3, 오른쪽 1 비율)
        main_layout.addWidget(left_widget, 3)
        main_layout.addWidget(right_widget, 1)
        
        # 하단 상태바
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_label = QLabel("대기 중...")
        self.status_bar.addWidget(self.status_label)
        
    def create_toolbar(self):
        """툴바 생성"""
        toolbar = QHBoxLayout()
        
        # AutoCAD 연결
        self.connect_btn = QPushButton("🔗 AutoCAD 연결")
        self.connect_btn.clicked.connect(self.connect_autocad)
        toolbar.addWidget(self.connect_btn)
        
        # 파일 작업
        toolbar.addWidget(QLabel(" | "))
        
        new_btn = QPushButton("📄 새 프로젝트")
        new_btn.clicked.connect(self.new_project)
        toolbar.addWidget(new_btn)
        
        save_btn = QPushButton("💾 저장")
        save_btn.clicked.connect(self.save_file)
        toolbar.addWidget(save_btn)
        
        load_btn = QPushButton("📂 열기")
        load_btn.clicked.connect(self.load_file)
        toolbar.addWidget(load_btn)
        
        # 행 추가
        toolbar.addWidget(QLabel(" | "))
        
        add_row_btn = QPushButton("➕ 행 추가")
        add_row_btn.clicked.connect(self.add_row)
        toolbar.addWidget(add_row_btn)
        
        toolbar.addStretch()
        
        return toolbar
        
    def connect_autocad(self):
        """AutoCAD 연결"""
        try:
            # 기존 연결 확인
            if self.acad and self.doc:
                try:
                    obj_count = self.doc.ModelSpace.Count
                    QMessageBox.information(self, "연결 상태", 
                        f"이미 연결됨\n\n도면: {self.doc.Name}\n객체: {obj_count}개")
                    return
                except:
                    self.acad = None
                    self.doc = None
            
            # AutoCAD 연결
            print("AutoCAD 연결 시도...")
            self.acad = win32com.client.Dispatch("AutoCAD.Application")
            self.doc = self.acad.ActiveDocument
            
            # 테스트
            obj_count = self.doc.ModelSpace.Count
            
            # UI 업데이트
            self.status_label.setText(f"✅ 연결 성공 - {self.doc.Name}")
            self.status_label.setStyleSheet("color: green;")
            
            # 테이블에 연결 정보 전달
            self.flat_table.acad = self.acad
            self.flat_table.doc = self.doc
            
            if HIERARCHICAL_TABLE_AVAILABLE:
                self.hierarchical_table.set_cad_connection(self.acad, self.doc)
            
            # 성공 메시지 - 명확하게 표시
            msg = QMessageBox(self)
            msg.setWindowTitle("연결 성공")
            msg.setIcon(QMessageBox.Information)
            msg.setText("AutoCAD 연결 성공!")
            msg.setInformativeText(f"도면: {self.doc.Name}\n객체 수: {obj_count}개")
            msg.exec_()
            
            print(f"✅ AutoCAD 연결 성공: {self.doc.Name}")
            
        except Exception as e:
            self.status_label.setText("❌ 연결 실패")
            self.status_label.setStyleSheet("color: red;")
            
            error_msg = str(e)
            if "ActiveDocument" in error_msg:
                QMessageBox.warning(self, "연결 오류", 
                    "AutoCAD에서 도면을 열어주세요.\n\n"
                    "1. AutoCAD에서 새 도면 생성 (Ctrl+N)\n"
                    "2. 또는 기존 도면 열기 (Ctrl+O)")
            else:
                QMessageBox.critical(self, "연결 오류", 
                    f"AutoCAD 연결 실패:\n{e}\n\n"
                    "AutoCAD가 실행 중인지 확인하세요.")
                    
    def switch_to_flat(self):
        """평면 테이블로 전환"""
        self.current_mode = "flat"
        self.stacked_widget.setCurrentIndex(0)
        self.flat_btn.setChecked(True)
        if HIERARCHICAL_TABLE_AVAILABLE:
            self.hierarchical_btn.setChecked(False)
            
    def switch_to_hierarchical(self):
        """계층구조 테이블로 전환"""
        if HIERARCHICAL_TABLE_AVAILABLE:
            self.current_mode = "hierarchical"
            self.stacked_widget.setCurrentIndex(1)
            self.hierarchical_btn.setChecked(True)
            self.flat_btn.setChecked(False)
            
    def add_row(self):
        """현재 테이블에 행 추가"""
        if self.current_mode == "flat":
            self.flat_table.add_row()
        elif HIERARCHICAL_TABLE_AVAILABLE:
            self.hierarchical_table.add_row()
            
    def new_project(self):
        """새 프로젝트"""
        reply = QMessageBox.question(self, "새 프로젝트", 
            "현재 작업을 모두 삭제하고 새로 시작하시겠습니까?",
            QMessageBox.Yes | QMessageBox.No)
            
        if reply == QMessageBox.Yes:
            self.flat_table.setRowCount(0)
            if HIERARCHICAL_TABLE_AVAILABLE:
                self.hierarchical_table.setRowCount(0)
                self.hierarchical_table.row_types.clear()
                self.hierarchical_table.row_levels.clear()
                
    def save_file(self):
        """파일 저장"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "프로젝트 저장", "", "CAD Quantity Files (*.cqp)")
            
        if file_path:
            data = {
                'version': '2.0',
                'mode': self.current_mode
            }
            
            # 평면 테이블 데이터
            flat_data = []
            for row in range(self.flat_table.rowCount()):
                row_data = []
                for col in range(self.flat_table.columnCount() - 2):
                    item = self.flat_table.item(row, col)
                    row_data.append(item.text() if item else "")
                flat_data.append(row_data)
            data['flat_table'] = flat_data
            
            # 계층구조 테이블 데이터
            if HIERARCHICAL_TABLE_AVAILABLE:
                data['hierarchical_table'] = self.hierarchical_table.get_data()
            
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
                
            QMessageBox.information(self, "저장 완료", "프로젝트가 저장되었습니다.")
            
    def load_file(self):
        """파일 불러오기"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "프로젝트 열기", "", "CAD Quantity Files (*.cqp)")
            
        if file_path:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            # 평면 테이블 로드
            if 'flat_table' in data:
                self.flat_table.setRowCount(0)
                for row_data in data['flat_table']:
                    row = self.flat_table.rowCount()
                    self.flat_table.insertRow(row)
                    for col, value in enumerate(row_data[:self.flat_table.columnCount()-2]):
                        self.flat_table.setItem(row, col, QTableWidgetItem(str(value)))
            
            # 계층구조 테이블 로드
            if HIERARCHICAL_TABLE_AVAILABLE and 'hierarchical_table' in data:
                self.hierarchical_table.load_data(data['hierarchical_table'])
            
            # 모드 설정
            if 'mode' in data:
                if data['mode'] == 'hierarchical' and HIERARCHICAL_TABLE_AVAILABLE:
                    self.switch_to_hierarchical()
                else:
                    self.switch_to_flat()
            
            QMessageBox.information(self, "불러오기 완료", "프로젝트를 불러왔습니다.")


def main():
    """메인 함수"""
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    
    window = CADQuantityProWindow()
    window.show()
    
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()