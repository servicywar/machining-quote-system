"""
engine/analyzer.py — STEP 파일 형상 분석

pythonOCC 설치 시: 바운딩박스·홀·셋업 자동 추출
pythonOCC 미설치 시: is_step_available() == False → 수동 입력 모드로 전환
                     역산 감사 기능은 그대로 동작

설치 방법 (conda 전용):
    conda install -c conda-forge pythonocc-core
"""

from __future__ import annotations

import math
from pathlib import Path

from engine.difficulty import auto_difficulty

# pythonOCC 임포트 시도 — 없으면 수동 모드
try:
    from OCC.Core.BRep import BRep_Tool
    from OCC.Core.BRepBndLib import brepbndlib
    from OCC.Core.BRepMesh import BRepMesh_IncrementalMesh
    from OCC.Core.BRepPrimAPI import BRepPrimAPI_MakeBox
    from OCC.Core.Bnd import Bnd_Box
    from OCC.Core.STEPControl import STEPControl_Reader
    from OCC.Core.TopAbs import TopAbs_FACE, TopAbs_EDGE
    from OCC.Core.TopExp import TopExp_Explorer
    from OCC.Core.TopoDS import topods
    from OCC.Core.BRepAdaptor import BRepAdaptor_Surface
    from OCC.Core.GeomAbs import GeomAbs_Cylinder
    _OCC_AVAILABLE = True
except ImportError:
    _OCC_AVAILABLE = False


def is_step_available() -> bool:
    """pythonOCC 설치 여부를 반환한다."""
    return _OCC_AVAILABLE


def analyze_step_file(filepath: str, machinability: float = 1.0) -> dict:
    """
    STEP 파일을 분석하여 형상 정보를 반환한다.

    Returns:
        {
            "bounding_box": {"x_mm", "y_mm", "z_mm", "volume_cm3"},
            "holes":        [{"diameter_mm", "depth_mm", "is_deep"}, ...],
            "setups":       int,          # 추정 셋업 횟수
            "difficulty":   {"level", "reason"},
            "error":        str | None,
        }
    """
    result = {
        "bounding_box": {},
        "holes":        [],
        "setups":       1,
        "difficulty":   {"level": 1, "reason": ""},
        "error":        None,
    }

    if not _OCC_AVAILABLE:
        result["error"] = "pythonOCC 미설치"
        return result

    try:
        reader = STEPControl_Reader()
        status = reader.ReadFile(str(filepath))
        if status != 1:
            result["error"] = "STEP 파일 읽기 실패"
            return result

        reader.TransferRoots()
        shape = reader.OneShape()

        # ── 바운딩박스 ─────────────────────────────────────
        bbox = Bnd_Box()
        brepbndlib.Add(shape, bbox)
        xmin, ymin, zmin, xmax, ymax, zmax = bbox.Get()

        x_mm = (xmax - xmin) * 1000
        y_mm = (ymax - ymin) * 1000
        z_mm = (zmax - zmin) * 1000
        volume_cm3 = (x_mm * y_mm * z_mm) / 1000.0  # mm³ → cm³

        result["bounding_box"] = {
            "x_mm":      round(x_mm, 2),
            "y_mm":      round(y_mm, 2),
            "z_mm":      round(z_mm, 2),
            "volume_cm3": round(volume_cm3, 3),
        }

        # ── 홀 탐색 (원통형 면 기준) ───────────────────────
        holes = []
        exp = TopExp_Explorer(shape, TopAbs_FACE)
        while exp.More():
            face    = topods.Face(exp.Current())
            adaptor = BRepAdaptor_Surface(face)
            if adaptor.GetType() == GeomAbs_Cylinder:
                cyl      = adaptor.Cylinder()
                diameter = cyl.Radius() * 2 * 1000  # m → mm
                # 깊이: z_mm 기준 간이 추정
                depth    = z_mm
                is_deep  = depth > diameter * 5
                holes.append({
                    "diameter_mm": round(diameter, 2),
                    "depth_mm":    round(depth, 2),
                    "is_deep":     is_deep,
                })
            exp.Next()

        result["holes"] = holes

        # ── 셋업 횟수 추정 ─────────────────────────────────
        # 단순 휴리스틱: 홀 방향이 여러 면에 분산되어 있으면 셋업 증가
        hole_count = len(holes)
        has_deep   = any(h["is_deep"] for h in holes)
        if hole_count > 15 or has_deep:
            setups = 3
        elif hole_count > 5:
            setups = 2
        else:
            setups = 1
        result["setups"] = setups

        # ── 난이도 자동 추정 ───────────────────────────────
        result["difficulty"] = auto_difficulty(
            hole_count=hole_count,
            setup_count=setups,
            has_deep_hole=has_deep,
        )

    except Exception as e:
        result["error"] = str(e)

    return result
