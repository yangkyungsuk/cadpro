"""
물량 산출 시스템 설정 파일
건축, 토목, 조경 분야별 설정 및 계수 정의
"""

# 분야별 레이어 매핑
LAYER_MAPPING = {
    "건축": {
        "벽체": ["A-WALL", "WALL", "벽", "A-WALL-*"],
        "기둥": ["A-COLS", "COLUMN", "기둥", "A-COLS-*"],
        "슬래브": ["A-SLAB", "SLAB", "슬라브", "A-FLOR-*"],
        "문": ["A-DOOR", "DOOR", "문", "A-DOOR-*"],
        "창문": ["A-GLAZ", "WINDOW", "창", "A-GLAZ-*"],
        "계단": ["A-STRS", "STAIR", "계단", "A-STRS-*"],
        "지붕": ["A-ROOF", "ROOF", "지붕", "A-ROOF-*"]
    },
    "토목": {
        "도로": ["C-ROAD", "ROAD", "도로", "C-ROAD-*"],
        "배관": ["C-PIPE", "PIPE", "배관", "C-SSWR-*", "C-WATR-*"],
        "철근": ["C-RBAR", "REBAR", "철근", "S-RBAR-*"],
        "옹벽": ["C-WALL", "RETAIN", "옹벽", "C-WALL-*"],
        "교량": ["C-BRDG", "BRIDGE", "교량", "C-BRDG-*"],
        "터널": ["C-TUNL", "TUNNEL", "터널", "C-TUNL-*"],
        "토공": ["C-TOPO", "EARTHWORK", "토공", "C-GRAD-*"]
    },
    "조경": {
        "수목": ["L-PLNT", "TREE", "수목", "L-PLNT-TREE"],
        "잔디": ["L-PLNT-TURF", "GRASS", "잔디", "L-TURF"],
        "포장": ["L-PAVE", "PAVING", "포장", "L-WALK", "L-PAVE-*"],
        "시설물": ["L-SITE", "FURNITURE", "시설물", "L-FURN-*"],
        "연못": ["L-WATR", "POND", "연못", "L-WATR-*"],
        "화단": ["L-PLNT-BED", "FLOWERBED", "화단", "L-PLBD-*"],
        "울타리": ["L-SITE-FENC", "FENCE", "울타리", "L-FENC-*"]
    }
}

# 재료별 단위 중량 및 계수
MATERIAL_COEFFICIENTS = {
    "콘크리트": {
        "단위중량": 2400,  # kg/m³
        "할증률": 1.05,  # 5% 손실
        "단위": "m³"
    },
    "철근": {
        "D10": 0.56,  # kg/m
        "D13": 0.995,
        "D16": 1.56,
        "D19": 2.25,
        "D22": 3.04,
        "D25": 3.98,
        "D29": 5.04,
        "D32": 6.23,
        "할증률": 1.03,
        "단위": "ton"
    },
    "벽돌": {
        "일반벽돌": 512,  # 개/m³
        "시멘트벽돌": 62.5,  # 개/m²
        "할증률": 1.03,
        "단위": "개"
    },
    "모르타르": {
        "두께": 0.02,  # m (20mm)
        "단위중량": 2100,  # kg/m³
        "할증률": 1.1,
        "단위": "m³"
    }
}

# 조경 재료 계수
LANDSCAPE_COEFFICIENTS = {
    "잔디": {
        "롤잔디": 1.05,  # m²당 할증률
        "씨앗": 0.02,  # kg/m²
        "단위": "m²"
    },
    "수목": {
        "교목": {"소": 1, "중": 1, "대": 1},  # 주
        "관목": {"소": 4, "중": 2, "대": 1},  # 주/m²
        "단위": "주"
    },
    "포장재": {
        "아스팔트": 0.1,  # m³/m² (두께 100mm)
        "보도블록": 1.02,  # m²당 할증률
        "자갈": 0.05,  # m³/m² (두께 50mm)
        "단위": "m²"
    }
}

# 블록 인식 패턴
BLOCK_PATTERNS = {
    "문": ["DOOR*", "DR-*", "*문*"],
    "창문": ["WINDOW*", "WIN-*", "*창*"],
    "나무": ["TREE*", "TR-*", "*수목*"],
    "가로등": ["LAMP*", "LP-*", "*가로등*"],
    "벤치": ["BENCH*", "BN-*", "*벤치*"],
    "표지판": ["SIGN*", "SN-*", "*표지*"]
}

# 해치 패턴별 재료 매핑
HATCH_MATERIAL_MAPPING = {
    "CONCRETE": "콘크리트",
    "AR-CONC": "콘크리트",
    "BRICK": "벽돌",
    "AR-BRSTD": "벽돌",
    "EARTH": "토사",
    "GRASS": "잔디",
    "GRAVEL": "자갈",
    "SOLID": "일반"
}

# 출력 설정
OUTPUT_CONFIG = {
    "excel_template": "templates/quantity_template.xlsx",
    "pdf_font": "malgun.ttf",  # 한글 폰트
    "decimal_places": 2,
    "include_formula": True,
    "group_by_layer": True
}