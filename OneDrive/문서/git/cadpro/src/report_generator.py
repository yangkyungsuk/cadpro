"""
리포트 생성 모듈
Excel 및 PDF 형식으로 물량 산출 결과 출력
"""

import pandas as pd
from datetime import datetime
import os
from typing import Dict
import json


class ReportGenerator:
    def __init__(self, quantities: Dict, file_info: Dict):
        """
        리포트 생성기 초기화
        
        Args:
            quantities: 물량 계산 결과
            file_info: CAD 파일 정보
        """
        self.quantities = quantities
        self.file_info = file_info
        self.timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    def generate_excel(self, output_path: str = None) -> str:
        """Excel 리포트 생성"""
        if not output_path:
            output_path = f"output/물량산출서_{self.timestamp}.xlsx"
        
        # Excel Writer 생성
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # 요약 시트
            self._create_summary_sheet(writer)
            
            # 건축 물량 시트
            if "건축" in self.quantities and self.quantities["건축"]:
                self._create_architecture_sheet(writer)
            
            # 토목 물량 시트
            if "토목" in self.quantities and self.quantities["토목"]:
                self._create_civil_sheet(writer)
            
            # 조경 물량 시트
            if "조경" in self.quantities and self.quantities["조경"]:
                self._create_landscape_sheet(writer)
        
        print(f"Excel 리포트 생성 완료: {output_path}")
        return output_path
    
    def _create_summary_sheet(self, writer):
        """요약 시트 생성"""
        summary_data = {
            "항목": ["파일명", "CAD버전", "작성일시", "총 레이어수", "총 엔티티수"],
            "내용": [
                self.file_info.get("파일명", ""),
                self.file_info.get("CAD버전", ""),
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                self.file_info.get("레이어수", 0),
                self.file_info.get("총엔티티", 0)
            ]
        }
        
        df_summary = pd.DataFrame(summary_data)
        df_summary.to_excel(writer, sheet_name="요약", index=False)
        
        # 분야별 요약
        field_summary = []
        for field in ["건축", "토목", "조경"]:
            if field in self.quantities:
                field_summary.append({
                    "분야": field,
                    "항목수": len(self.quantities[field]),
                    "주요항목": ", ".join(list(self.quantities[field].keys())[:3])
                })
        
        if field_summary:
            df_field = pd.DataFrame(field_summary)
            df_field.to_excel(writer, sheet_name="요약", startrow=8, index=False)
    
    def _create_architecture_sheet(self, writer):
        """건축 물량 시트 생성"""
        arch_data = []
        
        for category, values in self.quantities["건축"].items():
            if isinstance(values, dict):
                # 단위 정보 분리
                units = values.get("단위", {})
                
                for key, value in values.items():
                    if key != "단위":
                        if isinstance(value, dict):
                            # 세부 항목이 있는 경우
                            for sub_key, sub_value in value.items():
                                arch_data.append({
                                    "구분": category,
                                    "항목": f"{key}-{sub_key}",
                                    "수량": sub_value,
                                    "단위": units.get(key, "")
                                })
                        else:
                            arch_data.append({
                                "구분": category,
                                "항목": key,
                                "수량": value,
                                "단위": units.get(key, "")
                            })
        
        if arch_data:
            df_arch = pd.DataFrame(arch_data)
            df_arch.to_excel(writer, sheet_name="건축물량", index=False)
    
    def _create_civil_sheet(self, writer):
        """토목 물량 시트 생성"""
        civil_data = []
        
        for category, values in self.quantities["토목"].items():
            if isinstance(values, dict):
                units = values.get("단위", {})
                
                for key, value in values.items():
                    if key != "단위":
                        if isinstance(value, dict):
                            for sub_key, sub_value in value.items():
                                civil_data.append({
                                    "구분": category,
                                    "항목": f"{key}-{sub_key}",
                                    "수량": sub_value,
                                    "단위": units.get(key, "")
                                })
                        else:
                            civil_data.append({
                                "구분": category,
                                "항목": key,
                                "수량": value,
                                "단위": units.get(key, "")
                            })
        
        if civil_data:
            df_civil = pd.DataFrame(civil_data)
            df_civil.to_excel(writer, sheet_name="토목물량", index=False)
    
    def _create_landscape_sheet(self, writer):
        """조경 물량 시트 생성"""
        landscape_data = []
        
        for category, values in self.quantities["조경"].items():
            if isinstance(values, dict):
                units = values.get("단위", {})
                
                for key, value in values.items():
                    if key != "단위":
                        if isinstance(value, dict):
                            for sub_key, sub_value in value.items():
                                landscape_data.append({
                                    "구분": category,
                                    "항목": f"{key}-{sub_key}",
                                    "수량": sub_value,
                                    "단위": units.get(key, "")
                                })
                        else:
                            landscape_data.append({
                                "구분": category,
                                "항목": key,
                                "수량": value,
                                "단위": units.get(key, "")
                            })
        
        if landscape_data:
            df_landscape = pd.DataFrame(landscape_data)
            df_landscape.to_excel(writer, sheet_name="조경물량", index=False)
    
    def generate_json(self, output_path: str = None) -> str:
        """JSON 형식으로 결과 저장"""
        if not output_path:
            output_path = f"output/물량산출_{ self.timestamp}.json"
        
        output_data = {
            "file_info": self.file_info,
            "timestamp": datetime.now().isoformat(),
            "quantities": self.quantities
        }
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(output_data, f, ensure_ascii=False, indent=2)
        
        print(f"JSON 리포트 생성 완료: {output_path}")
        return output_path
    
    def print_summary(self):
        """콘솔에 요약 출력"""
        print("\n" + "="*60)
        print("물량 산출 결과 요약")
        print("="*60)
        
        for field in ["건축", "토목", "조경"]:
            if field in self.quantities and self.quantities[field]:
                print(f"\n[{field} 분야]")
                
                for category, values in self.quantities[field].items():
                    print(f"\n  {category}:")
                    
                    if isinstance(values, dict):
                        for key, value in values.items():
                            if key != "단위":
                                if isinstance(value, dict):
                                    print(f"    {key}:")
                                    for sub_key, sub_value in value.items():
                                        print(f"      - {sub_key}: {sub_value}")
                                else:
                                    unit = values.get("단위", {}).get(key, "")
                                    print(f"    - {key}: {value} {unit}")
        
        print("\n" + "="*60)