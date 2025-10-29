"""
텍스트 처리 유틸리티 모듈

텍스트 정규화, 텍스트 프레임 읽기/쓰기 등의 기능을 제공합니다.
"""
import pandas as pd


def normalize_key(s: str) -> str:
    """
    텍스트 키를 정규화합니다.
    
    엑셀 데이터와 PowerPoint 텍스트 박스의 텍스트를 매칭하기 위해
    공백 문자, 특수 문자 등을 제거하고 일관된 형식으로 변환합니다.
    
    Args:
        s (str): 정규화할 문자열
        
    Returns:
        str: 정규화된 문자열
        
    처리 과정:
        1. None 값은 빈 문자열로 변환
        2. 보이지 않는 특수 공백 문자 제거 (Zero Width Space, BOM 등)
        3. Non-breaking space를 일반 공백으로 치환
        4. 모든 줄바꿈을 공백으로 치환
        5. 앞뒤 공백 제거
        6. 연속된 공백을 하나로 축약
        
    Example:
        >>> normalize_key("  고객명\\n  ")
        "고객명"
        >>> normalize_key("제품\\u200b이름")
        "제품이름"
    """
    if s is None:
        return ''

    s = str(s)

    # 보이지 않는 특수 공백 문자 제거
    # \u200b: Zero Width Space
    # \ufeff: Byte Order Mark (BOM)
    # \u2060: Word Joiner
    for ch in ['\u200b', '\ufeff', '\u2060']:
        s = s.replace(ch, '')

    # Non-breaking space를 일반 공백으로 변환
    s = s.replace('\xa0', ' ')

    # 탭을 일반 공백으로 변환
    s = s.replace('\t', ' ')

    # 모든 형태의 줄바꿈 제거 (공백 대체 안함)
    s = s.replace('\r\n', '').replace('\n', '').replace('\r', '')

    # 앞뒤 공백 제거
    s = s.strip()

    # 연속된 공백을 하나로 축약
    while '  ' in s:
        s = s.replace('  ', ' ')

    return s


def get_text_content(shape) -> str:
    """
    Shape 객체에서 텍스트 내용을 추출합니다.
    
    Args:
        shape: python-pptx Shape 객체
        
    Returns:
        str: Shape의 텍스트 내용. 텍스트가 없으면 빈 문자열 반환
        
    Note:
        - has_text_frame 속성이 없는 Shape(예: 이미지, 도형)는 빈 문자열 반환
        - text_frame이 있어도 텍스트가 None인 경우 빈 문자열 반환
    """
    if not getattr(shape, 'has_text_frame', False):
        return ''
    return shape.text or ''


def set_text_preserve_style(shape, value):
    """
    텍스트만 교체하고 기존 스타일(폰트, 색상, 크기 등)은 보존합니다.
    
    PowerPoint 텍스트 구조:
        TextFrame
        └── Paragraph (여러 개 가능)
            └── Run (여러 개 가능, 각각 독립적인 스타일)
    
    이 함수는 첫 번째 Paragraph의 첫 번째 Run만 남기고 나머지는 제거하여
    원본 스타일을 유지하면서 텍스트만 교체합니다.
    
    Args:
        shape: python-pptx Shape 객체
        value: 설정할 텍스트 값 (None, NaN, 일반 문자열 등 모두 처리)
        
    처리 과정:
        1. text_frame이 없으면 무시
        2. 값이 None이나 NaN이면 빈 문자열로 변환
        3. 첫 Paragraph의 첫 Run에 텍스트 설정 (스타일 유지)
        4. 나머지 Run 및 Paragraph 제거 (깔끔한 단일 텍스트 유지)
        
    Example:
        >>> # 기존: "홍길동" (굵게, 빨간색, 24pt)
        >>> set_text_preserve_style(shape, "김철수")
        >>> # 결과: "김철수" (굵게, 빨간색, 24pt) - 스타일 유지
    """
    if not getattr(shape, 'has_text_frame', False):
        return

    tf = shape.text_frame

    # None, NaN 등을 빈 문자열로 처리
    text = '' if value is None or (isinstance(value, float) and pd.isna(value)) else str(value)

    # Paragraph/Run 구조 유지하면서 텍스트만 교체
    if tf.paragraphs:
        para = tf.paragraphs[0]

        if para.runs:
            # 첫 번째 Run의 스타일을 유지하며 텍스트만 교체
            para.runs[0].text = text

            # 나머지 Run 처리 (제거 시도, 실패 시 빈 텍스트로)
            for i in range(len(para.runs) - 1, 0, -1):
                run = para.runs[i]
                try:
                    # _element 또는 element 속성 확인
                    if hasattr(run, '_element'):
                        r_elem = run._element
                    elif hasattr(run, 'element'):
                        r_elem = run.element
                    else:
                        # element를 찾을 수 없으면 텍스트만 비움
                        run.text = ''
                        continue
                    r_elem.getparent().remove(r_elem)
                except Exception:
                    # Run 제거 실패 시 텍스트를 빈 문자열로 설정
                    try:
                        run.text = ''
                    except Exception:
                        pass
        else:
            # Run이 없으면 새로 생성
            run = para.add_run()
            run.text = text

        # 추가 Paragraph 처리 (제거 시도, 실패 시 빈 텍스트로)
        for i in range(len(tf.paragraphs) - 1, 0, -1):
            p = tf.paragraphs[i]
            try:
                if hasattr(p, '_element'):
                    p_elem = p._element
                elif hasattr(p, 'element'):
                    p_elem = p.element
                else:
                    # element를 찾을 수 없으면 텍스트만 비움
                    if p.runs:
                        for r in p.runs:
                            r.text = ''
                    continue
                p_elem.getparent().remove(p_elem)
            except Exception:
                # Paragraph 제거 실패 시 모든 Run을 빈 문자열로 설정
                try:
                    if p.runs:
                        for r in p.runs:
                            r.text = ''
                except Exception:
                    pass
