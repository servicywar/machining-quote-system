"""
engine/difficulty.py — 난이도 등급 정의 및 계수 반환
"""

DIFFICULTY_TABLE = {
    1: {
        "label":  "단순",
        "coeff":  1.0,
        "desc":   "평면 가공, 홀 5개 이하, 셋업 1회",
    },
    2: {
        "label":  "보통",
        "coeff":  1.3,
        "desc":   "복합 가공, 셋업 2회, 홀 6~15개",
    },
    3: {
        "label":  "복잡",
        "coeff":  1.7,
        "desc":   "5축·심공·IT7 이하, 셋업 3회 이상",
    },
    4: {
        "label":  "특수",
        "coeff":  2.2,
        "desc":   "난삭재·초정밀 (인코넬, 티타늄 등)",
    },
}


def get_coefficient(level: int) -> float:
    """난이도 등급(1~4)에 해당하는 계수를 반환한다."""
    return DIFFICULTY_TABLE.get(level, DIFFICULTY_TABLE[1])["coeff"]


def get_label(level: int) -> str:
    """난이도 등급(1~4)에 해당하는 한글 레이블을 반환한다."""
    return DIFFICULTY_TABLE.get(level, DIFFICULTY_TABLE[1])["label"]


def get_all_options() -> list[str]:
    """selectbox용 옵션 문자열 리스트를 반환한다."""
    return [
        f"{lv} · {info['label']}  (×{info['coeff']})"
        for lv, info in DIFFICULTY_TABLE.items()
    ]


def parse_level_from_option(option: str) -> int:
    """
    get_all_options() 결과 문자열에서 등급 번호를 파싱하여 반환한다.
    예: "2 · 보통  (×1.3)" → 2
    """
    try:
        return int(option.split("·")[0].strip())
    except (ValueError, IndexError):
        return 1


def auto_difficulty(hole_count: int, setup_count: int,
                    has_deep_hole: bool = False,
                    material_label: str = "") -> dict:
    """
    홀 개수·셋업 횟수·심공 여부·소재명으로 난이도를 자동 추정한다.
    반환값: {"level": int, "reason": str}
    """
    hard_materials = ["인코넬", "티타늄", "하스텔로이", "inconel", "titanium"]
    is_hard = any(m in material_label.lower() for m in hard_materials)

    if is_hard:
        return {"level": 4, "reason": "난삭재 소재"}

    if has_deep_hole or setup_count >= 3:
        return {"level": 3, "reason": "심공 또는 셋업 3회 이상"}

    if setup_count == 2 or 6 <= hole_count <= 15:
        return {"level": 2, "reason": "셋업 2회 또는 홀 6~15개"}

    return {"level": 1, "reason": "평면 가공 / 홀 5개 이하 / 셋업 1회"}
