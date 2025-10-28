"""
이미지 처리 유틸리티 모듈

PowerPoint Shape에 이미지를 삽입하는 기능을 제공합니다.
"""
import os


def fill_picture(shape, img_path):
    """
    Shape에 이미지를 채웁니다.
    
    이미지 삽입 방식:
        1. Placeholder Shape인 경우: insert_picture() 메서드 사용 (권장)
        2. 일반 Shape인 경우: 새 이미지 Shape를 동일 위치/크기로 생성하고 기존 Shape 제거
    
    Args:
        shape: python-pptx Shape 객체
        img_path: 이미지 파일 경로 (str 또는 Path 객체)
        
    처리 과정:
        1. 이미지 경로 유효성 검증
        2. Placeholder 여부 확인
           - Placeholder인 경우: insert_picture() 직접 사용
           - 일반 Shape인 경우: 위치/크기 복사 후 새 이미지 Shape 생성
        3. 기존 Shape 제거 (일반 Shape의 경우)
        
    Note:
        - 이미지 파일이 없으면 조용히 무시 (에러 발생 안 함)
        - Placeholder가 아닌 Shape는 XML 레벨에서 교체되므로 복잡함
        - 이미지 삽입 실패 시 에러 메시지 출력 후 계속 진행
        
    Example:
        >>> fill_picture(text_shape, "images/logo.png")
        # text_shape가 있던 위치에 logo.png 이미지가 표시됨
    """
    # 이미지 경로 검증
    if not img_path:
        return
    
    img_path = str(img_path)
    
    if not os.path.isfile(img_path):
        return

    # 방법 1: Placeholder Shape인 경우 직접 insert_picture 사용
    try:
        _ = shape.placeholder_format  # placeholder인지 확인
        shape.insert_picture(img_path)
        return
    except Exception:
        # Placeholder가 아니면 방법 2로 진행
        pass

    # 방법 2: 일반 Shape인 경우 - 새 이미지 Shape 생성 후 기존 Shape 제거
    try:
        # 슬라이드 참조 및 Shape 위치/크기 정보 저장
        sl = shape.part.slide
        left, top, width, height = shape.left, shape.top, shape.width, shape.height
        
        # 동일한 위치/크기로 새 이미지 Shape 추가
        sl.shapes.add_picture(img_path, left, top, width=width, height=height)
        
        # 기존 Shape 제거 (XML 레벨에서 제거)
        shape._element.getparent().remove(shape._element)
        
    except Exception as e:
        print(f'이미지 채우기 실패: {e}')


# 지원하는 이미지 파일 확장자 목록
IMAGE_EXTENSIONS = (
    '.png', '.jpg', '.jpeg', '.gif', '.bmp', 
    '.tif', '.tiff', '.webp'
)
