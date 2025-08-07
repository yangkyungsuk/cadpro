"""
AutoCAD 연결 및 선택 간단 테스트
"""

import win32com.client
import pythoncom
import sys


def test_autocad():
    """AutoCAD 기본 테스트"""
    print("=== AutoCAD 연결 테스트 ===\n")
    
    try:
        # AutoCAD 연결
        print("1. AutoCAD 연결 시도...")
        acad = win32com.client.Dispatch("AutoCAD.Application")
        print("   [OK] AutoCAD 연결 성공")
        
        # 도면 확인
        print("\n2. 도면 확인...")
        try:
            doc = acad.ActiveDocument
            print(f"   [OK] 활성 도면: {doc.Name}")
        except:
            print("   [경고] 열린 도면이 없습니다.")
            print("   AutoCAD에서 도면을 열어주세요.")
            return
        
        # ModelSpace 접근
        print("\n3. ModelSpace 접근...")
        model = doc.ModelSpace
        print(f"   [OK] 객체 수: {model.Count}개")
        
        # 선택 세트 테스트
        print("\n4. 선택 세트 테스트...")
        
        # PickfirstSelectionSet 테스트
        try:
            sel_set = doc.PickfirstSelectionSet
            print(f"   현재 선택: {sel_set.Count}개")
        except Exception as e:
            print(f"   PickfirstSelectionSet 오류: {e}")
        
        # 객체 순회 테스트
        print("\n5. 객체 타입 분석...")
        types = {}
        count = 0
        
        for obj in model:
            try:
                obj_type = obj.ObjectName
                if obj_type not in types:
                    types[obj_type] = 0
                types[obj_type] += 1
                count += 1
                
                if count >= 100:  # 처음 100개만
                    break
            except:
                pass
        
        print(f"   분석된 객체: {count}개")
        print("   타입별 분포:")
        for obj_type, cnt in sorted(types.items()):
            print(f"     {obj_type}: {cnt}개")
        
        # SendCommand 테스트
        print("\n6. SendCommand 테스트...")
        try:
            # 줌 익스텐트
            doc.SendCommand("_ZOOM\n_E\n")
            print("   [OK] Zoom Extents 실행")
        except Exception as e:
            print(f"   SendCommand 오류: {e}")
        
        print("\n=== 테스트 완료 ===")
        print("\n이제 cadpro_advanced.py를 실행할 수 있습니다.")
        print("선택 방법:")
        print("1. '모든 객체 로드' - 전체 도면 객체를 메모리에 로드")
        print("2. '현재 선택 가져오기' - AutoCAD에서 미리 선택한 객체 가져오기")
        print("3. '레이어로 선택' - 특정 레이어의 객체만 선택")
        print("4. '타입으로 선택' - 특정 타입의 객체만 선택")
        
    except Exception as e:
        print(f"\n[오류] {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    test_autocad()