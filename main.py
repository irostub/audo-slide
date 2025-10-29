"""
2025-10-29
AutoSlide - 엑셀 기반 PowerPoint 자동 생성 도구

엑셀 데이터와 PowerPoint 템플릿을 결합하여
자동으로 여러 슬라이드를 생성하는 프로그램입니다.

사용법:
    1. template.pptx 파일 준비 (첫 슬라이드에 엑셀 컬럼명을 텍스트 박스로 배치)
    2. data.xlsx 파일 준비 (첫 행: 컬럼명, 나머지 행: 데이터)
    3. 이 파일 실행: python main.py
    4. output.pptx 파일 생성 확인

주요 특징:
    - 텍스트, 이미지 자동 삽입
    - 기존 스타일(폰트, 색상, 크기) 완벽 보존
    - 배경, 그룹 Shape 등 복잡한 템플릿 지원
    - 빠른 처리 속도 (XML 레벨 최적화)
    - 다중 시트 지원 (export 시트 우선 사용)
"""
import traceback
import pandas as pd
from core.ppt_builder import build_ppt_from_excel


# 설정값
TEMPLATE_FILE = 'template.pptx'  # 템플릿 PPT 파일명
EXCEL_FILE = 'data.xlsx'          # 엑셀 데이터 파일명
OUTPUT_FILE = 'output.pptx'       # 출력 PPT 파일명
PREFERRED_SHEET = 'export'         # 우선 사용할 시트명


def detect_sheet_name(excel_path, preferred_sheet='export'):
    """
    엑셀 파일의 시트를 자동으로 감지합니다.

    Args:
        excel_path (str): 엑셀 파일 경로
        preferred_sheet (str): 우선 사용할 시트명 (기본값: 'export')

    Returns:
        str or int: 사용할 시트명 또는 인덱스

    로직:
        1. 엑셀 파일의 모든 시트명 확인
        2. 시트가 1개만 있으면 첫 번째 시트(0) 사용
        3. 시트가 여러 개 있으면:
           - preferred_sheet가 존재하면 해당 시트 사용
           - 없으면 첫 번째 시트(0) 사용
    """
    try:
        # 엑셀 파일의 모든 시트명 확인
        excel_file = pd.ExcelFile(excel_path)
        sheet_names = excel_file.sheet_names

        print(f"엑셀 시트 감지: {sheet_names}")

        # 시트가 1개만 있는 경우
        if len(sheet_names) == 1:
            print(f"→ 단일 시트 '{sheet_names[0]}' 사용")
            return 0

        # 시트가 여러 개 있는 경우
        if preferred_sheet in sheet_names:
            print(f"→ '{preferred_sheet}' 시트 사용")
            return preferred_sheet
        else:
            print(f"→ '{preferred_sheet}' 시트 없음, 첫 번째 시트 '{sheet_names[0]}' 사용")
            return 0

    except Exception as e:
        print(f"경고: 시트 감지 실패 ({e}), 첫 번째 시트 사용")
        return 0


def main():
    """
    메인 실행 함수

    설정된 파일 경로를 사용하여 PowerPoint를 생성합니다.
    엑셀 파일의 시트를 자동으로 감지하여 사용합니다.
    """
    print("=" * 60)
    print("AutoSlide - 엑셀 기반 PPT 자동 생성")
    print("=" * 60)

    try:
        # 엑셀 시트 자동 감지
        sheet_name = detect_sheet_name(EXCEL_FILE, PREFERRED_SHEET)

        # PPT 생성 실행
        result_path = build_ppt_from_excel(
            template_path=TEMPLATE_FILE,
            excel_path=EXCEL_FILE,
            out_path=OUTPUT_FILE,
            sheet_name=sheet_name
        )

        print("=" * 60)
        print(f"✓ 생성 완료: {result_path}")
        print("=" * 60)

    except FileNotFoundError as e:
        print("\n[오류] 파일을 찾을 수 없습니다:")
        print(f"  {e}")
        print("\n다음을 확인하세요:")
        print(f"  1. {TEMPLATE_FILE} 파일이 존재하는지")
        print(f"  2. {EXCEL_FILE} 파일이 존재하는지")

    except Exception as e:
        print("\n[오류] 실행 중 문제가 발생했습니다:")
        print(f"  {e}")
        print("\n상세 오류:")
        traceback.print_exc()


if __name__ == '__main__':
    main()
