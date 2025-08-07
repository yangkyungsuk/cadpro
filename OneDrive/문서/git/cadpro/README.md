# CAD 물량 산출 시스템

건축, 토목, 조경 분야의 AutoCAD 도면에서 자동으로 물량을 산출하는 프로그램입니다.

## 주요 기능

### 지원 분야
- **건축**: 벽체, 기둥, 슬래브, 문/창문 등
- **토목**: 도로, 배관, 철근, 토공 등  
- **조경**: 수목, 잔디, 포장, 시설물 등

### 주요 특징
- DWG/DXF 파일 직접 처리
- 레이어 기반 자동 분류
- 분야별 맞춤 물량 계산
- Excel/JSON 리포트 생성

## 설치 방법

### 1. 필수 요구사항
- Python 3.9 이상
- Windows/Linux/Mac OS

### 2. 패키지 설치
```bash
pip install -r requirements.txt
```

### 3. DWG 파일 처리 (선택사항)
DWG 파일을 직접 처리하려면 다음 중 하나를 설치:
- ODA File Converter (무료)
- Aspose.CAD (상용)

## 사용 방법

### 기본 실행
```bash
python main.py
```

실행 후 CAD 파일 경로를 입력하면 자동으로 물량 산출이 진행됩니다.

### 출력 결과
- `output/물량산출서_YYYYMMDD_HHMMSS.xlsx` - Excel 리포트
- `output/물량산출_YYYYMMDD_HHMMSS.json` - JSON 데이터

## 프로젝트 구조

```
cadpro/
├── main.py                 # 메인 실행 파일
├── requirements.txt        # 필요 패키지
├── src/
│   ├── config.py          # 설정 및 계수 정의
│   ├── dwg_parser.py      # CAD 파일 파싱
│   ├── quantity_calculator.py  # 물량 계산
│   └── report_generator.py     # 리포트 생성
├── data/                  # 샘플 CAD 파일
├── output/                # 출력 리포트
└── tests/                 # 테스트 코드
```

## 레이어 명명 규칙

### 건축 레이어
- 벽체: A-WALL, WALL, 벽
- 기둥: A-COLS, COLUMN, 기둥
- 슬래브: A-SLAB, SLAB, 슬라브
- 문: A-DOOR, DOOR, 문
- 창문: A-GLAZ, WINDOW, 창

### 토목 레이어
- 도로: C-ROAD, ROAD, 도로
- 배관: C-PIPE, PIPE, 배관
- 철근: C-RBAR, REBAR, 철근
- 토공: C-TOPO, EARTHWORK, 토공

### 조경 레이어
- 수목: L-PLNT, TREE, 수목
- 잔디: L-PLNT-TURF, GRASS, 잔디
- 포장: L-PAVE, PAVING, 포장
- 시설물: L-SITE, FURNITURE, 시설물

## 물량 계산 방식

### 건축
- **벽체**: 길이 × 높이 × 두께 = 체적
- **기둥**: 단면적 × 높이 = 체적
- **슬래브**: 면적 × 두께 = 체적
- **철근**: 길이 × 단위중량 = 중량

### 토목
- **도로**: 연장 × 폭 × 두께 = 아스팔트량
- **배관**: 연장 × 관경별 분류
- **토공**: 절토량, 성토량 계산

### 조경
- **수목**: 교목/관목 개수 집계
- **잔디**: 면적 × 할증률
- **포장**: 재료별 면적 집계

## 주의사항

1. **DWG 파일**: 직접 읽기 위해서는 추가 도구 필요
2. **레이어 명명**: 표준 레이어명 사용 권장
3. **단위**: 도면 단위는 미터(m) 기준
4. **블록**: 표준 블록명 사용시 자동 인식

## 라이선스 및 법적 고지

- 이 프로그램은 교육/연구 목적으로 개발되었습니다
- DWG 형식은 Autodesk의 독점 형식입니다
- 상업적 사용시 적절한 라이선스 확인 필요

## 향후 개발 계획

- [ ] BIM(IFC) 파일 지원
- [ ] AI 기반 자동 인식 강화
- [ ] 웹 인터페이스 개발
- [ ] 3D 시각화 기능
- [ ] 클라우드 연동

## 문의 및 지원

문제 발생시 Issues 탭에서 문의해주세요.