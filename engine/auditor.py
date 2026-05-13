"""
engine/auditor.py — 역산 판정 핵심 로직

핵심 공식 (변경 불가):
    실질 가공비 = 협력사 단가 - 재료비 - 후처리비
    적정 가공비 = 임률(원/h) × 예상시간(h) × 난이도계수
    오차율(%)  = 실질 가공비 / 적정 가공비 × 100

판정 기준 (변경 불가):
    > 120%       경고 · 과다 청구   red
    110 ~ 120%   주의 · 소폭 과다   orange
    90  ~ 110%   신뢰 · 적정        green
    80  ~  90%   관찰 · 소폭 저가   yellow
    < 80%        경고 · 저가 수주   red
"""

from dataclasses import dataclass


@dataclass
class AuditResult:
    """역산 감사 결과 데이터 클래스."""
    vendor_price:       float
    material_cost:      float
    postprocess_cost:   float
    actual_machining:   float
    standard_machining: float
    variance_pct:       float
    verdict:            str
    color:              str
    action:             str
    hourly_rate:        float
    estimated_hours:    float
    difficulty_coeff:   float


def _get_verdict(variance_pct: float) -> tuple:
    """
    오차율로 (판정, 색상, 권장액션) 튜플을 반환한다.
    경계값: 120 초과 / 110 초과 / 90 이상 / 80 이상 / 미만
    """
    if variance_pct > 120:
        return "경고 · 과다 청구", "red",    "항목별 소명 요청 후 재견적"
    elif variance_pct > 110:
        return "주의 · 소폭 과다", "orange", "단가 재협의 검토 권장"
    elif variance_pct >= 90:
        return "신뢰 · 적정",      "green",  "발주 진행"
    elif variance_pct >= 80:
        return "관찰 · 소폭 저가", "yellow", "원자재 규격 및 품질 조건 확인"
    else:
        return "경고 · 저가 수주", "red",    "품질 조건 명기 필수"


def audit(
    vendor_price: float,
    material_cost: float,
    hourly_rate: float,
    estimated_hours: float,
    difficulty_coeff: float,
    postprocess_cost: float = 0.0,
) -> AuditResult:
    """
    역산 감사를 실행하고 AuditResult를 반환한다.

    Args:
        vendor_price:      협력사 제출 단가 (원)
        material_cost:     계산된 재료비 (원)
        hourly_rate:       임률 (원/h)
        estimated_hours:   예상 가공시간 (h)
        difficulty_coeff:  난이도 계수 (1.0 / 1.3 / 1.7 / 2.2)
        postprocess_cost:  후처리비 (원, 기본값 0)
    """
    actual_machining   = vendor_price - material_cost - postprocess_cost
    standard_machining = hourly_rate * estimated_hours * difficulty_coeff

    if standard_machining == 0:
        raise ValueError("적정 가공비가 0입니다. 임률·가공시간을 확인해주세요.")

    variance_pct = (actual_machining / standard_machining) * 100
    verdict, color, action = _get_verdict(variance_pct)

    return AuditResult(
        vendor_price=vendor_price,
        material_cost=material_cost,
        postprocess_cost=postprocess_cost,
        actual_machining=actual_machining,
        standard_machining=standard_machining,
        variance_pct=variance_pct,
        verdict=verdict,
        color=color,
        action=action,
        hourly_rate=hourly_rate,
        estimated_hours=estimated_hours,
        difficulty_coeff=difficulty_coeff,
    )


def audit_scenario(base: AuditResult, material_adj_pct: float) -> AuditResult:
    """
    재료비를 material_adj_pct(%) 만큼 조정한 시나리오 AuditResult를 반환한다.

    Args:
        base:             기준 AuditResult
        material_adj_pct: 재료비 조정률 (예: +10 → 10% 인상, -5 → 5% 인하)
    """
    new_material = base.material_cost * (1 + material_adj_pct / 100)
    return audit(
        vendor_price=base.vendor_price,
        material_cost=new_material,
        hourly_rate=base.hourly_rate,
        estimated_hours=base.estimated_hours,
        difficulty_coeff=base.difficulty_coeff,
        postprocess_cost=base.postprocess_cost,
    )
