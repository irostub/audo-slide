"""
슬라이드 조작 유틸리티 모듈

슬라이드 복제, 관계(Relationship) 리매핑 등의 기능을 제공합니다.
"""
import copy

# PowerPoint OOXML 네임스페이스 상수
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main'


def _remap_relationships_in_element(src_slide, dst_slide, el):
    """
    XML 요소 내의 관계(Relationship) ID를 새 슬라이드용으로 리매핑합니다.
    
    PowerPoint는 내부적으로 이미지, 차트 등의 리소스를 관계(Relationship)로 관리합니다.
    각 관계는 고유한 rId (예: rId1, rId2...)를 가지며, 슬라이드마다 별도로 관리됩니다.
    
    슬라이드를 복제할 때, 기존 슬라이드의 rId를 그대로 사용하면 잘못된 리소스를 참조하게 되므로
    새 슬라이드에 관계를 다시 생성하고 XML의 rId 참조를 업데이트해야 합니다.
    
    Args:
        src_slide: 원본 슬라이드 객체
        dst_slide: 대상 슬라이드 객체
        el: 리매핑할 XML 요소
        
    처리 과정:
        1. XML 트리의 모든 노드를 순회
        2. r:embed, r:id 속성을 찾음 (이미지, 비디오, 하이퍼링크 등)
        3. 원본 슬라이드의 관계를 확인
        4. 대상 슬라이드에 동일한 관계를 생성 (새 rId 부여)
        5. XML 속성을 새 rId로 업데이트
        
    관계 타입:
        - 내부 관계: 이미지, 차트 등 프레젠테이션 내부 리소스
        - 외부 관계: 외부 URL, 하이퍼링크 등
        
    Note:
        - 이 함수는 슬라이드 복제의 핵심 로직입니다
        - 리매핑하지 않으면 이미지가 깨지거나 잘못된 차트가 표시될 수 있습니다
    """
    # 관계 ID를 참조하는 XML 속성들
    attr_qnames = [
        f'{{{R_NS}}}embed',  # 이미지, 비디오 등 임베디드 리소스
        f'{{{R_NS}}}id'      # 하이퍼링크, 차트 등 일반 관계
    ]

    # XML 트리의 모든 노드 순회
    for node in el.iter():
        for attr in attr_qnames:
            old_rid = node.get(attr)
            
            # 속성이 없거나 원본 슬라이드에 해당 관계가 없으면 스킵
            if not old_rid:
                continue
            if old_rid not in src_slide.part.rels:
                continue

            # 원본 관계 정보 가져오기
            old_rel = src_slide.part.rels[old_rid]
            
            # 새 슬라이드에 동일한 관계 생성
            if getattr(old_rel, 'is_external', False):
                # 외부 관계 (URL 등)
                new_rid = dst_slide.part.relate_to(
                    old_rel.target_ref, 
                    old_rel.reltype, 
                    is_external=True
                )
            else:
                # 내부 관계 (이미지, 차트 등)
                new_rid = dst_slide.part.relate_to(
                    old_rel.target_part, 
                    old_rel.reltype
                )
            
            # XML 속성을 새 rId로 업데이트
            node.set(attr, new_rid)


def duplicate_slide(prs, src_slide):
    """
    슬라이드를 완전히 복제합니다 (배경, 이미지, 모든 관계 포함).
    
    단순히 python-pptx의 add_slide()를 사용하면 빈 슬라이드가 생성되므로
    원본 슬라이드의 모든 요소를 깊은 복사(deep copy)하여 복제합니다.
    
    Args:
        prs: Presentation 객체
        src_slide: 복제할 원본 슬라이드
        
    Returns:
        복제된 새 슬라이드 객체
        
    처리 과정:
        1. 동일한 레이아웃으로 빈 슬라이드 생성
        2. 원본의 모든 Shape를 deep copy하여 추가
        3. 각 Shape의 관계(이미지 등) 리매핑
        4. 슬라이드 배경 복제 및 관계 리매핑
        
    복제되는 요소:
        - 모든 텍스트 박스 및 내용
        - 모든 이미지 (관계 포함)
        - 모든 도형 및 그룹
        - 슬라이드 배경 (색상, 그라데이션, 이미지 등)
        - 차트, 표, SmartArt 등
        
    Note:
        - deep copy를 사용하여 원본과 완전히 독립적인 복사본 생성
        - 관계 리매핑을 통해 이미지 등이 올바르게 참조되도록 보장
        
    Example:
        >>> slide1 = prs.slides[0]
        >>> slide2 = duplicate_slide(prs, slide1)
        # slide2는 slide1의 완전한 복사본
    """
    # 1. 동일한 레이아웃으로 빈 슬라이드 생성
    dst_slide = prs.slides.add_slide(src_slide.slide_layout)

    # 2. 모든 Shape 복제 및 관계 리매핑
    for shape in src_slide.shapes:
        # Shape의 XML 요소를 deep copy
        new_el = copy.deepcopy(shape.element)
        
        # 새 슬라이드의 shape tree에 추가
        dst_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')
        
        # 관계 ID 리매핑 (이미지 등이 올바르게 참조되도록)
        _remap_relationships_in_element(src_slide, dst_slide, new_el)

    # 3. 배경 복제 및 관계 리매핑
    try:
        if src_slide.background is not None:
            # 배경 XML을 deep copy
            bg_el = copy.deepcopy(src_slide.background.element)
            
            # 배경의 관계 리매핑 (배경 이미지 등)
            _remap_relationships_in_element(src_slide, dst_slide, bg_el)
            
            # 새 슬라이드의 배경을 복제된 배경으로 교체
            dst_bg = dst_slide.background._element
            dst_bg.getparent().replace(dst_bg, bg_el)
            
    except Exception as e:
        print(f'배경 복사/리매핑 실패: {e}')

    return dst_slide
