"""
AutoCAD 연결 테스트
"""

import win32com.client
import pythoncom

def test_autocad_connection():
    """AutoCAD 연결 테스트"""
    print("="*60)
    print("AutoCAD 연결 테스트")
    print("="*60)
    
    try:
        # AutoCAD 연결 시도
        print("\nAutoCAD 연결 시도 중...")
        acad = win32com.client.Dispatch("AutoCAD.Application")
        
        print(f"[OK] AutoCAD 연결 성공!")
        print(f"  - 버전: {acad.Version}")
        print(f"  - 프로그램: {acad.Name}")
        print(f"  - 경로: {acad.Path}")
        
        # 현재 도면 정보
        if acad.Documents.Count > 0:
            doc = acad.ActiveDocument
            print(f"\n현재 도면:")
            print(f"  - 이름: {doc.Name}")
            print(f"  - 경로: {doc.Path if doc.Path else '저장되지 않음'}")
            
            # 레이어 정보
            print(f"\n레이어 정보:")
            print(f"  - 레이어 수: {doc.Layers.Count}")
            
            # 처음 5개 레이어 표시
            print("  - 레이어 목록 (처음 5개):")
            count = 0
            for layer in doc.Layers:
                if count >= 5:
                    break
                print(f"    {count+1}. {layer.Name}")
                count += 1
            
            # 모델스페이스 객체 수
            model = doc.ModelSpace
            print(f"\n모델스페이스:")
            print(f"  - 객체 수: {model.Count}")
            
            # 객체 타입별 집계
            object_types = {}
            for obj in model:
                obj_type = obj.ObjectName
                if obj_type not in object_types:
                    object_types[obj_type] = 0
                object_types[obj_type] += 1
            
            print(f"  - 객체 타입별 집계:")
            for obj_type, count in sorted(object_types.items())[:10]:
                print(f"    {obj_type}: {count}개")
            
        else:
            print("\n열린 도면이 없습니다.")
            print("AutoCAD에서 도면을 열어주세요.")
        
        return True
        
    except Exception as e:
        print(f"\n[ERROR] AutoCAD 연결 실패!")
        print(f"  오류: {e}")
        print("\n가능한 원인:")
        print("  1. AutoCAD가 실행되지 않음")
        print("  2. AutoCAD가 설치되지 않음")
        print("  3. COM 인터페이스 문제")
        print("\n해결 방법:")
        print("  1. AutoCAD를 먼저 실행하세요")
        print("  2. 도면을 하나 열어두세요 (test.dwg)")
        return False

def main():
    # AutoCAD 연결 테스트
    if test_autocad_connection():
        print("\n" + "="*60)
        print("테스트 성공! AutoCAD 플러그인을 사용할 수 있습니다.")
        print("="*60)
        
        print("\n다음 명령으로 플러그인 실행:")
        print("  python autocad_plugin.py")
    else:
        print("\n" + "="*60)
        print("테스트 실패! AutoCAD를 먼저 실행하세요.")
        print("="*60)

if __name__ == "__main__":
    main()