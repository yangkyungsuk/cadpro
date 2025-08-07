# CADPro 사용 가이드

## 문제 해결 완료

### 1. GetPoint 오류 수정
- AutoCAD COM API의 GetPoint, GetEntity 메서드 호출 시 None 대신 빈 배열 전달 필요
- SendCommand를 사용한 더 안정적인 선택 방식으로 대체

### 2. 선택 인식 문제 해결
- 대화형 선택 시 COM 인터페이스가 멈추는 문제 해결
- 세 가지 안정적인 선택 방법 제공:
  1. **모든 객체 로드** - 전체 도면을 메모리에 로드
  2. **현재 선택 가져오기** - AutoCAD에서 미리 선택한 객체 사용
  3. **필터링 선택** - 레이어/타입별 선택

### 3. GUI 종료 문제 해결
- WM_DELETE_WINDOW 이벤트 핸들러 추가
- root.quit()과 root.destroy() 순차 호출

## 사용 방법

### 1. 준비 사항
```bash
# 필수 라이브러리 설치
pip install pywin32
pip install numpy==1.24.3
```

### 2. AutoCAD 준비
1. AutoCAD 실행
2. DWG 또는 DXF 파일 열기
3. 필요한 경우 DXF로 저장 (SAVEAS 명령)

### 3. 프로그램 실행

#### 안정적인 버전 (권장)
```bash
python cadpro_stable.py
```
- 모든 객체를 메모리에 로드
- 필터링으로 선택
- 선택 문제 없음

#### 고급 버전
```bash
python cadpro_advanced.py
```
- 다양한 선택 방법
- 상세한 계산 옵션
- 탭 인터페이스

### 4. 사용 순서

1. **AutoCAD 연결**
   - "AutoCAD 연결" 버튼 클릭
   - 도면이 열려있는지 확인

2. **객체 선택**
   - 방법 1: "모든 객체 로드" (가장 안정적)
   - 방법 2: AutoCAD에서 먼저 선택 → "현재 선택 가져오기"
   - 방법 3: "레이어로 선택" 또는 "타입으로 선택"

3. **물량 계산**
   - 원하는 계산 버튼 클릭
   - 옵션 설정 (단위, 할증률, 높이)

4. **결과 저장**
   - JSON 저장
   - 클립보드 복사

## 주요 개선 사항

### cadpro_stable.py
```python
def load_all_objects(self):
    """모든 객체를 메모리에 로드"""
    self.all_objects = []
    for obj in self.model:
        obj_data = self.extract_object_data(obj)
        if obj_data:
            self.all_objects.append(obj_data)
```
- 선택 문제 완전 해결
- 필터링 기반 작업

### cadpro_advanced.py
```python
def get_current_selection(self):
    """현재 AutoCAD에서 선택된 객체 가져오기"""
    sel_set = self.doc.PickfirstSelectionSet
    # AutoCAD에서 미리 선택한 객체 사용
```
- PickfirstSelectionSet 사용
- SendCommand 기반 선택

## 계산 기능

### 건축
- 벽체: 길이 → 면적 → 체적 → 콘크리트
- 기둥: 개수 → 기초 콘크리트
- 슬래브: 면적 → 콘크리트, 철근

### 토목
- 도로: 면적 → 아스팔트, 기층
- 배관: 길이 → 이음관
- 철근: 길이 → 중량

### 조경
- 데크: 장선 길이, 데크보드 면적
- 수목: 개수 집계
- 포장: 면적 → 블록, 모래

## 문제 발생 시

1. **AutoCAD 연결 실패**
   - AutoCAD가 실행 중인지 확인
   - 도면이 열려있는지 확인
   - 관리자 권한으로 실행

2. **선택 인식 안됨**
   - "모든 객체 로드" 사용 (가장 안정적)
   - AutoCAD에서 먼저 선택 후 "현재 선택 가져오기"

3. **GUI 종료 안됨**
   - 작업 관리자에서 python.exe 종료
   - cadpro_advanced.py 사용 (종료 핸들러 포함)

## 테스트 파일

```bash
# AutoCAD 연결 테스트
python test_autocad_simple.py

# 선택 기능 테스트
python cadpro_selector.py
```

## 참고 사항

- DWG 파일은 직접 읽을 수 없음 (AutoCAD COM 필요)
- DXF 파일은 ezdxf로 직접 읽기 가능
- 한글 Windows에서 인코딩 문제 주의 (cp949)
- 대용량 도면은 메모리 사용량 주의

## 지원

문제 발생 시:
1. test_autocad_simple.py 실행하여 연결 확인
2. cadpro_stable.py 사용 (가장 안정적)
3. 오류 메시지 확인