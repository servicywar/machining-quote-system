"""
가공 견적 AI 엔진 PRO v2.2
main.py — Streamlit 메인 UI

변경사항 (v2.1 → v2.2):
- [버그] 업체 추가 탭: rerun 후 탭 초기화 문제 수정
- [강화] 감사 이력 검색: 날짜 범위 / 오차율 범위 / 통합 텍스트 검색
- [신규] 유사 부품 매칭: 감사 결과 하단 자동 표시 + 협상 문구 생성
- [신규] 업체별 예상 단가 시뮬레이션 페이지
- [신규] 협상 근거 엑셀 리포트 다운로드
"""

import io
import sqlite3
import tempfile
import os
from datetime import datetime, date
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter

from engine import (
    audit, audit_scenario, AuditResult,
    calc_material_cost, load_material_master, get_market_prices,
    get_coefficient, get_label, get_all_options, parse_level_from_option,
    DIFFICULTY_TABLE,
    analyze_step_file, is_step_available,
)

BASE_DIR = Path(__file__).parent
DB_PATH  = BASE_DIR / "data" / "history_log.db"


# ══════════════════════════════════════════════════════════════
# DB 초기화
# ══════════════════════════════════════════════════════════════
def init_db():
    """DB 파일 및 테이블 초기화. 앱 시작 시 1회 실행."""
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    try:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS vendors (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                name       TEXT NOT NULL,
                contact    TEXT,
                phone      TEXT,
                email      TEXT,
                note       TEXT,
                is_active  INTEGER DEFAULT 1
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS history_log (
                id                  INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at          DATETIME DEFAULT CURRENT_TIMESTAMP,
                part_name           TEXT,
                part_no             TEXT,
                step_file           TEXT,
                material_code       TEXT,
                vendor_id           INTEGER REFERENCES vendors(id),
                vendor_name         TEXT,
                vendor_price        REAL,
                material_cost       REAL,
                actual_machining    REAL,
                standard_machining  REAL,
                variance_pct        REAL,
                verdict             TEXT,
                difficulty_level    INTEGER,
                hourly_rate         REAL,
                estimated_hours     REAL,
                postprocess_cost    REAL,
                volume_cm3          REAL,
                hole_count          INTEGER,
                setup_count         INTEGER,
                price_snapshot      REAL,
                ordered             INTEGER DEFAULT 0,
                note                TEXT
            )
        """)
        conn.commit()
    finally:
        conn.close()


# ══════════════════════════════════════════════════════════════
# 업체 CRUD
# ══════════════════════════════════════════════════════════════
def load_vendors(active_only=True) -> pd.DataFrame:
    conn = sqlite3.connect(DB_PATH)
    try:
        q = "SELECT * FROM vendors"
        if active_only:
            q += " WHERE is_active = 1"
        return pd.read_sql(q + " ORDER BY name", conn)
    finally:
        conn.close()


def add_vendor(name, contact="", phone="", email="", note=""):
    conn = sqlite3.connect(DB_PATH)
    try:
        conn.execute(
            "INSERT INTO vendors (name,contact,phone,email,note) VALUES (?,?,?,?,?)",
            (name, contact, phone, email, note)
        )
        conn.commit()
    finally:
        conn.close()


def update_vendor(vid, name, contact, phone, email, note):
    conn = sqlite3.connect(DB_PATH)
    try:
        conn.execute(
            "UPDATE vendors SET name=?,contact=?,phone=?,email=?,note=? WHERE id=?",
            (name, contact, phone, email, note, vid)
        )
        conn.commit()
    finally:
        conn.close()


def deactivate_vendor(vid):
    conn = sqlite3.connect(DB_PATH)
    try:
        conn.execute("UPDATE vendors SET is_active=0 WHERE id=?", (vid,))
        conn.commit()
    finally:
        conn.close()


def vendor_has_history(vid) -> bool:
    conn = sqlite3.connect(DB_PATH)
    try:
        return conn.execute(
            "SELECT COUNT(*) FROM history_log WHERE vendor_id=?", (vid,)
        ).fetchone()[0] > 0
    finally:
        conn.close()


# ══════════════════════════════════════════════════════════════
# 이력 저장·조회
# ══════════════════════════════════════════════════════════════
def save_to_db(part_name, part_no, step_file, material_code,
               vendor_id, vendor_name, result: AuditResult,
               difficulty_level, estimated_hours,
               volume_cm3=None, hole_count=None, setup_count=None,
               price_snapshot=None, note=""):
    conn = sqlite3.connect(DB_PATH)
    try:
        conn.execute("""
            INSERT INTO history_log (
                created_at,part_name,part_no,step_file,material_code,
                vendor_id,vendor_name,vendor_price,material_cost,
                actual_machining,standard_machining,variance_pct,verdict,
                difficulty_level,hourly_rate,estimated_hours,postprocess_cost,
                volume_cm3,hole_count,setup_count,price_snapshot,note
            ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
        """, (
            datetime.now().isoformat(),
            part_name, part_no, step_file, material_code,
            vendor_id, vendor_name, result.vendor_price, result.material_cost,
            result.actual_machining, result.standard_machining, result.variance_pct,
            result.verdict, difficulty_level, result.hourly_rate, estimated_hours,
            result.postprocess_cost, volume_cm3, hole_count, setup_count,
            price_snapshot, note
        ))
        conn.commit()
    finally:
        conn.close()


def load_history() -> pd.DataFrame:
    if not DB_PATH.exists():
        return pd.DataFrame()
    conn = sqlite3.connect(DB_PATH)
    try:
        return pd.read_sql("SELECT * FROM history_log ORDER BY created_at DESC", conn)
    finally:
        conn.close()


def load_vendor_history(vendor_id) -> pd.DataFrame:
    conn = sqlite3.connect(DB_PATH)
    try:
        return pd.read_sql(
            "SELECT * FROM history_log WHERE vendor_id=? ORDER BY created_at DESC",
            conn, params=(vendor_id,)
        )
    finally:
        conn.close()


def load_vendor_stats() -> pd.DataFrame:
    if not DB_PATH.exists():
        return pd.DataFrame()
    conn = sqlite3.connect(DB_PATH)
    try:
        try:
            return pd.read_sql("""
            SELECT
                v.id, v.name,
                COUNT(h.id)                                                    AS total,
                ROUND(AVG(h.variance_pct), 1)                                  AS avg_variance,
                ROUND(MIN(h.variance_pct), 1)                                  AS min_variance,
                ROUND(MAX(h.variance_pct), 1)                                  AS max_variance,
                SUM(COALESCE(h.ordered, 0))                                    AS ordered_count,
                ROUND(AVG(h.material_cost / NULLIF(h.vendor_price,0)*100), 1)  AS mat_ratio,
                ROUND(AVG(h.actual_machining / NULLIF(h.standard_machining,0)), 3) AS avg_margin_ratio
            FROM vendors v
            LEFT JOIN history_log h ON v.id = h.vendor_id
            WHERE v.is_active = 1
            GROUP BY v.id
        """, conn)
        except Exception:
            return pd.DataFrame()
    finally:
        conn.close()


def search_similar_parts(volume_cm3, hole_count, setup_count,
                         exclude_id=None) -> pd.DataFrame:
    """부피 ±20% / 홀 ±2개 / 셋업 동일 조건으로 전 업체 이력 검색."""
    if not DB_PATH.exists():
        return pd.DataFrame()
    conn = sqlite3.connect(DB_PATH)
    try:
        exclude_clause = f"AND h.id != {exclude_id}" if exclude_id else ""
        return pd.read_sql(f"""
            SELECT
                h.id, h.created_at, h.part_name, h.part_no,
                h.vendor_name, h.vendor_price, h.material_cost,
                h.variance_pct, h.verdict,
                h.volume_cm3, h.hole_count, h.setup_count, h.difficulty_level,
                ABS(h.volume_cm3 - {volume_cm3}) AS vol_diff
            FROM history_log h
            WHERE
                h.volume_cm3  BETWEEN {volume_cm3 * 0.8} AND {volume_cm3 * 1.2}
                AND h.hole_count  BETWEEN {hole_count - 2} AND {hole_count + 2}
                AND h.setup_count = {setup_count}
                AND h.volume_cm3 IS NOT NULL
                {exclude_clause}
            ORDER BY vol_diff ASC
            LIMIT 8
        """, conn)
    finally:
        conn.close()


# ══════════════════════════════════════════════════════════════
# 엑셀 리포트 생성
# ══════════════════════════════════════════════════════════════
def _cell_style(ws, row, col, value, bold=False, bg=None, align="left",
                num_format=None, font_color="000000"):
    """셀 하나에 값·스타일을 한번에 적용하는 헬퍼."""
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = Font(bold=bold, color=font_color,
                     name="맑은 고딕", size=10)
    if bg:
        cell.fill = PatternFill("solid", fgColor=bg)
    thin = Side(style="thin", color="CCCCCC")
    cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
    cell.alignment = Alignment(
        horizontal=align, vertical="center", wrap_text=True
    )
    if num_format:
        cell.number_format = num_format
    return cell


def build_report_excel(
    part_name: str,
    part_no: str,
    vendor_name: str,
    material_code: str,
    result: AuditResult,
    difficulty_level: int,
    estimated_hours: float,
    similar_df: pd.DataFrame,
    sim_df: pd.DataFrame | None,
    audit_date: str,
) -> bytes:
    """
    협상 근거 엑셀 리포트를 생성하고 bytes로 반환한다.

    시트 구성:
      1. 감사 결과   — 역산 수치 + 판정 + 협상 문구
      2. 유사 부품   — 유사 이력 비교표
      3. 업체 시뮬   — 업체별 예상 단가 비교 (sim_df 있을 때만)
    """
    wb = Workbook()

    # ── 공통 색상 ──────────────────────────────────────────
    C_HEADER  = "2E4057"   # 네이비
    C_SUB     = "4A90D9"   # 파랑
    C_GREEN   = "E8F5E9"
    C_RED     = "FFEBEE"
    C_ORANGE  = "FFF3E0"
    C_YELLOW  = "FFFDE7"
    C_GRAY    = "F5F5F5"
    C_WHITE   = "FFFFFF"

    verdict_bg = {
        "경고 · 과다 청구": C_RED,
        "주의 · 소폭 과다": C_ORANGE,
        "신뢰 · 적정":      C_GREEN,
        "관찰 · 소폭 저가": C_YELLOW,
        "경고 · 저가 수주": C_RED,
    }

    # ══════════════════════════════════════════════════════
    # 시트 1: 감사 결과
    # ══════════════════════════════════════════════════════
    ws1 = wb.active
    ws1.title = "감사 결과"
    ws1.column_dimensions["A"].width = 22
    ws1.column_dimensions["B"].width = 22
    ws1.column_dimensions["C"].width = 18
    ws1.column_dimensions["D"].width = 18
    ws1.row_dimensions[1].height = 30

    # 제목
    ws1.merge_cells("A1:D1")
    _cell_style(ws1, 1, 1, "가공 견적 역산 감사 리포트",
                bold=True, bg=C_HEADER, align="center", font_color="FFFFFF")
    ws1.cell(1, 1).font = Font(bold=True, color="FFFFFF",
                                name="맑은 고딕", size=14)

    # 기본 정보
    ws1.merge_cells("A2:D2")
    _cell_style(ws1, 2, 1, f"감사 일시: {audit_date}",
                bg=C_GRAY, align="left")

    row = 3
    info_rows = [
        ("부품명",   part_name),
        ("형번",     part_no or "—"),
        ("협력사",   vendor_name or "—"),
        ("소재",     material_code),
        ("난이도",   f"{difficulty_level}등급  ×{get_coefficient(difficulty_level)}"),
        ("예상 가공시간", f"{estimated_hours:.1f} h"),
    ]
    for label, val in info_rows:
        _cell_style(ws1, row, 1, label, bold=True, bg=C_GRAY)
        ws1.merge_cells(f"B{row}:D{row}")
        _cell_style(ws1, row, 2, val)
        row += 1

    row += 1
    # 역산 수치 헤더
    ws1.merge_cells(f"A{row}:D{row}")
    _cell_style(ws1, row, 1, "역산 수치",
                bold=True, bg=C_SUB, align="center", font_color="FFFFFF")
    row += 1

    num_rows = [
        ("협력사 단가",   result.vendor_price,       "#,##0 원"),
        ("재료비",        result.material_cost,       "#,##0 원"),
        ("후처리비",      result.postprocess_cost,    "#,##0 원"),
        ("실질 가공비",   result.actual_machining,    "#,##0 원"),
        ("적정 가공비",   result.standard_machining,  "#,##0 원"),
        ("오차율",        result.variance_pct,        "0.0 \"%\""),
        ("임률",          result.hourly_rate,         "#,##0 원/h"),
    ]
    for label, val, fmt in num_rows:
        _cell_style(ws1, row, 1, label, bold=True, bg=C_GRAY)
        ws1.merge_cells(f"B{row}:D{row}")
        _cell_style(ws1, row, 2, val, num_format=fmt)
        row += 1

    row += 1
    # 판정
    vbg = verdict_bg.get(result.verdict, C_WHITE)
    ws1.merge_cells(f"A{row}:D{row}")
    ws1.row_dimensions[row].height = 24
    _cell_style(ws1, row, 1,
                f"판정:  {result.verdict}  ({result.variance_pct:.1f}%)",
                bold=True, bg=vbg, align="center")

    row += 1
    ws1.merge_cells(f"A{row}:D{row}")
    _cell_style(ws1, row, 1, f"권장 액션:  {result.action}",
                bg=vbg, align="center")

    row += 2
    # 협상 근거 문구
    ws1.merge_cells(f"A{row}:D{row}")
    _cell_style(ws1, row, 1, "협상 근거",
                bold=True, bg=C_SUB, align="center", font_color="FFFFFF")
    row += 1

    # 과다/저가 여부에 따라 문구 생성
    diff   = result.actual_machining - result.standard_machining
    pct    = result.variance_pct
    if pct > 110:
        nego = (
            f"협력사 단가 {result.vendor_price:,.0f}원 기준 실질 가공비가 "
            f"적정 가공비({result.standard_machining:,.0f}원) 대비 "
            f"{pct - 100:.1f}% 초과합니다.\n"
            f"과다 청구 금액 추정: 약 {diff:,.0f}원\n"
            f"적정 단가 제안: {result.vendor_price - diff:,.0f}원 수준"
        )
    elif pct < 80:
        nego = (
            f"협력사 단가 {result.vendor_price:,.0f}원은 "
            f"적정 가공비({result.standard_machining:,.0f}원) 대비 "
            f"{100 - pct:.1f}% 낮은 수준입니다.\n"
            f"품질 조건 및 원자재 규격을 명기하고 진행을 권장합니다."
        )
    else:
        nego = (
            f"협력사 단가 {result.vendor_price:,.0f}원은 "
            f"적정 가공비 기준 오차율 {pct:.1f}%로 적정 범위입니다.\n"
            f"현재 단가로 발주를 진행해도 무방합니다."
        )

    ws1.merge_cells(f"A{row}:D{row + 2}")
    c = ws1.cell(row=row, column=1, value=nego)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    c.border = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC"),
    )
    ws1.row_dimensions[row].height = 60

    # ══════════════════════════════════════════════════════
    # 시트 2: 유사 부품 비교
    # ══════════════════════════════════════════════════════
    ws2 = wb.create_sheet("유사 부품 비교")
    ws2.column_dimensions["A"].width = 14
    ws2.column_dimensions["B"].width = 18
    ws2.column_dimensions["C"].width = 14
    ws2.column_dimensions["D"].width = 16
    ws2.column_dimensions["E"].width = 14
    ws2.column_dimensions["F"].width = 12
    ws2.column_dimensions["G"].width = 12

    ws2.merge_cells("A1:G1")
    _cell_style(ws2, 1, 1, "유사 부품 이력 비교",
                bold=True, bg=C_HEADER, align="center", font_color="FFFFFF")
    ws2.cell(1, 1).font = Font(bold=True, color="FFFFFF", name="맑은 고딕", size=12)

    headers2 = ["일시", "부품명", "형번", "업체", "단가(원)", "오차율(%)", "판정"]
    for ci, h in enumerate(headers2, 1):
        _cell_style(ws2, 2, ci, h, bold=True, bg="4A90D9", font_color="FFFFFF")

    if similar_df is not None and not similar_df.empty:
        for ri, (_, row_data) in enumerate(similar_df.iterrows(), 3):
            vbg2 = verdict_bg.get(str(row_data.get("verdict", "")), C_WHITE)
            vals = [
                str(row_data.get("created_at", ""))[:10],
                row_data.get("part_name", ""),
                row_data.get("part_no", "") or "—",
                row_data.get("vendor_name", "") or "—",
                row_data.get("vendor_price", 0),
                row_data.get("variance_pct", 0),
                row_data.get("verdict", ""),
            ]
            fmts = [None, None, None, None, "#,##0", "0.0", None]
            for ci, (v, f) in enumerate(zip(vals, fmts), 1):
                bg_use = vbg2 if ci == 7 else C_WHITE
                _cell_style(ws2, ri, ci, v, num_format=f, bg=bg_use)
    else:
        ws2.merge_cells("A3:G3")
        _cell_style(ws2, 3, 1, "유사 부품 이력이 없습니다.", bg=C_GRAY, align="center")

    # ══════════════════════════════════════════════════════
    # 시트 3: 업체별 예상 단가 시뮬레이션 (선택)
    # ══════════════════════════════════════════════════════
    if sim_df is not None and not sim_df.empty:
        ws3 = wb.create_sheet("업체별 시뮬레이션")
        ws3.column_dimensions["A"].width = 20
        ws3.column_dimensions["B"].width = 16
        ws3.column_dimensions["C"].width = 16
        ws3.column_dimensions["D"].width = 16
        ws3.column_dimensions["E"].width = 14
        ws3.column_dimensions["F"].width = 18

        ws3.merge_cells("A1:F1")
        _cell_style(ws3, 1, 1, f"업체별 예상 단가 시뮬레이션 — {part_name}",
                    bold=True, bg=C_HEADER, align="center", font_color="FFFFFF")
        ws3.cell(1, 1).font = Font(bold=True, color="FFFFFF", name="맑은 고딕", size=12)

        headers3 = ["업체명", "예상 단가(원)", "예상 가공비(원)", "재료비(원)", "예상 오차율(%)", "비고"]
        for ci, h in enumerate(headers3, 1):
            _cell_style(ws3, 2, ci, h, bold=True, bg="4A90D9", font_color="FFFFFF")

        for ri, (_, row_data) in enumerate(sim_df.iterrows(), 3):
            pct_v = row_data.get("예상오차율", 0)
            if pct_v > 110:
                row_bg = C_ORANGE
            elif pct_v < 80:
                row_bg = C_YELLOW
            else:
                row_bg = C_GREEN
            vals3 = [
                row_data.get("업체명", ""),
                row_data.get("예상단가", 0),
                row_data.get("예상가공비", 0),
                row_data.get("재료비", 0),
                pct_v,
                row_data.get("비고", ""),
            ]
            fmts3 = [None, "#,##0", "#,##0", "#,##0", "0.0", None]
            for ci, (v, f) in enumerate(zip(vals3, fmts3), 1):
                bg_use = row_bg if ci in (2, 5) else C_WHITE
                _cell_style(ws3, ri, ci, v, num_format=f, bg=bg_use)

    # bytes 반환
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ══════════════════════════════════════════════════════════════
# 판정 UI
# ══════════════════════════════════════════════════════════════
def render_verdict(result: AuditResult):
    fn = {"green": st.success, "orange": st.warning,
          "yellow": st.warning, "red": st.error}.get(result.color, st.info)
    fn(f"**{result.verdict}** ({result.variance_pct:.1f}%) — {result.action}")


def render_cost_cards(result: AuditResult):
    c1, c2, c3 = st.columns(3)
    c1.metric("재료비", f"{result.material_cost:,.0f} 원")
    c2.metric("실질 가공비", f"{result.actual_machining:,.0f} 원")
    delta = result.actual_machining - result.standard_machining
    c3.metric("적정 가공비", f"{result.standard_machining:,.0f} 원",
              delta=f"{delta:+,.0f} 원", delta_color="inverse")


def render_similar_parts(volume_cm3, hole_count, setup_count,
                         current_vendor_price, saved_id=None):
    """유사 부품 이력 표시 + 협상 문구 자동 생성."""
    if not (volume_cm3 and hole_count is not None and setup_count):
        return
    similar = search_similar_parts(volume_cm3, hole_count, setup_count, saved_id)
    if similar.empty:
        return

    st.subheader("🔍 유사 부품 이력")
    st.caption(
        f"부피 {volume_cm3:.0f}cm³ ±20% / 홀 {hole_count}개 ±2 / "
        f"셋업 {setup_count}회 기준 — 전 업체 대상"
    )
    col_map = {
        "created_at": "일시", "part_name": "부품명", "part_no": "형번",
        "vendor_name": "업체", "vendor_price": "단가",
        "variance_pct": "오차율(%)", "verdict": "판정",
        "volume_cm3": "부피(cm³)",
    }
    show_cols = [c for c in col_map if c in similar.columns]
    st.dataframe(
        similar[show_cols].rename(columns=col_map),
        hide_index=True, use_container_width=True,
        column_config={
            "오차율(%)": st.column_config.NumberColumn(format="%.1f %%"),
            "단가":      st.column_config.NumberColumn(format="%,.0f 원"),
            "부피(cm³)": st.column_config.NumberColumn(format="%.1f"),
        }
    )
    valid = similar[similar["vendor_price"].notna() & (similar["vendor_price"] > 0)]
    if not valid.empty:
        avg_price = valid["vendor_price"].mean()
        min_price = valid["vendor_price"].min()
        n         = len(valid)
        diff_pct  = (current_vendor_price - avg_price) / avg_price * 100
        if abs(diff_pct) >= 3:
            direction = "높습니다" if diff_pct > 0 else "낮습니다"
            msg = (
                f"유사 부품 {n}건 평균 단가 **{avg_price:,.0f}원** 대비 "
                f"현재 견적 **{current_vendor_price:,.0f}원**으로 "
                f"**{abs(diff_pct):.1f}% {direction}**. "
                f"(이력 최저가: {min_price:,.0f}원)"
            )
            if diff_pct > 10:
                st.error(f"💬 협상 근거: {msg}")
            elif diff_pct > 3:
                st.warning(f"💬 협상 근거: {msg}")
            else:
                st.info(f"💬 참고: {msg}")
    return similar


# ══════════════════════════════════════════════════════════════
# 업체별 예상 단가 시뮬레이션 계산
# ══════════════════════════════════════════════════════════════
def calc_vendor_simulation(
    volume_cm3: float,
    material_code: str,
    form_key: str,
    loss_rate: float,
    hourly_rate: float,
    estimated_hours: float,
    difficulty_coeff: float,
    postprocess_cost: float,
    price_override=None,
) -> pd.DataFrame:
    """
    등록된 업체별로 과거 평균 마진 비율을 적용해 예상 단가를 시뮬레이션한다.

    공식:
        재료비      = 부피 × 밀도 × 시세 × (1+로스율)
        적정 가공비 = 임률 × 시간 × 난이도계수
        예상 가공비 = 적정 가공비 × 업체 평균 마진 비율 (이력 없으면 1.0)
        예상 단가   = 재료비 + 예상 가공비 + 후처리비
    """
    try:
        mat_result = calc_material_cost(
            volume_cm3=volume_cm3,
            material_code=material_code,
            form=form_key,
            loss_rate_override=loss_rate,
            price_override=price_override,
        )
    except Exception:
        return pd.DataFrame()

    material_cost    = mat_result["material_cost"]
    standard_maching = hourly_rate * estimated_hours * difficulty_coeff

    stats_df = load_vendor_stats()
    if stats_df.empty:
        return pd.DataFrame()

    rows = []
    for _, row in stats_df.iterrows():
        name         = row["name"]
        total        = int(row["total"]) if row["total"] else 0
        avg_margin   = float(row["avg_margin_ratio"]) if row["avg_margin_ratio"] else 1.0

        if avg_margin <= 0:
            avg_margin = 1.0

        est_machining = standard_maching * avg_margin
        est_price     = material_cost + est_machining + postprocess_cost
        est_variance  = (est_machining / standard_maching * 100) if standard_maching > 0 else 0

        if total >= 3:
            note = f"이력 {total}건 기반 (평균 오차율 {row['avg_variance']}%)"
        else:
            note = f"이력 {total}건 — 참고용 (데이터 부족)"

        rows.append({
            "업체명":    name,
            "예상단가":  round(est_price),
            "예상가공비": round(est_machining),
            "재료비":    round(material_cost),
            "예상오차율": round(est_variance, 1),
            "이력건수":  total,
            "비고":      note,
        })

    if not rows:
        return pd.DataFrame()

    df = pd.DataFrame(rows).sort_values("예상단가")
    return df


# ══════════════════════════════════════════════════════════════
# 메인
# ══════════════════════════════════════════════════════════════
def main():
    st.set_page_config(
        page_title="가공 견적 AI 엔진 PRO v2.2",
        page_icon="🔧",
        layout="wide"
    )
    init_db()

    if "vendor_tab" not in st.session_state:
        st.session_state.vendor_tab = 0

    with st.sidebar:
        st.markdown("### 🔧 Quote AI v2.2")
        st.caption("가공 견적 역산 감사 시스템")
        st.divider()

        menu = st.radio(
            "메뉴",
            ["📋 역산 감사", "📈 업체 시뮬레이션", "🏢 업체 관리",
             "📊 업체 비교", "📁 감사 이력"],
            label_visibility="collapsed"
        )

        st.divider()

        mat_df       = load_material_master()
        mat_labels   = [f"{r['material_code']} — {r['material_name']}"
                        for _, r in mat_df.iterrows()]
        sel_mat_lbl  = st.selectbox("소재", mat_labels)
        sel_mat_code = mat_df["material_code"].tolist()[mat_labels.index(sel_mat_lbl)]
        mat_info     = mat_df[mat_df["material_code"] == sel_mat_code].iloc[0]

        form_type    = st.radio("원자재 형태", ["봉재", "판재"], horizontal=True)
        form_key     = "bar" if form_type == "봉재" else "plate"
        default_loss = float(mat_info[f"default_loss_{form_key}"])
        loss_rate    = st.slider("로스율 (%)", 0, 30, int(default_loss * 100)) / 100

        st.divider()
        hourly_rate   = st.number_input("임률 (원/h)", min_value=10000, max_value=200000,
                                        value=45000, step=1000)
        st.caption("시세 설정")
        use_override   = st.toggle("수동 입력")
        price_override = None
        if use_override:
            price_override = st.number_input("직접 입력 (원/kg)", min_value=0,
                                             value=4800, step=100)

        prices = get_market_prices()
        source = prices.get("source", "알 수 없음")
        if source == "KOMIS API":
            st.success(f"출처: {source}")
        else:
            st.warning(f"⚠ 출처: {source}")
        if st.button("🔄 시세 새로고침"):
            get_market_prices(force_refresh=True)
            st.rerun()

    if menu == "📋 역산 감사":
        page_audit(mat_df, sel_mat_code, mat_info, form_key,
                   loss_rate, hourly_rate, price_override)
    elif menu == "📈 업체 시뮬레이션":
        page_simulation(mat_df, sel_mat_code, mat_info, form_key,
                        loss_rate, hourly_rate, price_override)
    elif menu == "🏢 업체 관리":
        page_vendors()
    elif menu == "📊 업체 비교":
        page_compare()
    elif menu == "📁 감사 이력":
        page_history()


# ══════════════════════════════════════════════════════════════
# 페이지: 역산 감사
# ══════════════════════════════════════════════════════════════
def page_audit(mat_df, sel_mat_code, mat_info, form_key,
               loss_rate, hourly_rate, price_override):
    st.title("📋 역산 감사")
    col_left, col_right = st.columns([1, 1], gap="large")

    with col_left:
        st.subheader("1. 형상 정보")
        step_file   = st.file_uploader("STEP 파일 업로드 (선택)", type=["step", "stp"])
        part_name   = ""
        part_no     = ""
        step_fname  = ""
        volume_cm3  = None
        hole_count  = None
        setup_count = None
        auto_diff   = None

        if step_file is not None:
            part_name  = Path(step_file.name).stem
            step_fname = step_file.name
            if is_step_available():
                with st.spinner("STEP 분석 중..."):
                    with tempfile.NamedTemporaryFile(suffix=".step", delete=False) as tmp:
                        tmp.write(step_file.read())
                        tmp_path = tmp.name
                    mac      = float(mat_info.get("machinability", 1.0))
                    res_step = analyze_step_file(tmp_path, mac)
                    os.unlink(tmp_path)
                if res_step["error"]:
                    st.error(f"분석 오류: {res_step['error']}")
                else:
                    bb          = res_step["bounding_box"]
                    volume_cm3  = bb.get("volume_cm3", 0)
                    holes       = res_step["holes"]
                    hole_count  = len(holes)
                    setup_count = res_step["setups"]
                    auto_diff   = res_step["difficulty"]
                    st.success("STEP 분석 완료")
                    c1, c2, c3 = st.columns(3)
                    c1.metric("X", f"{bb.get('x_mm', 0):.1f} mm")
                    c2.metric("Y", f"{bb.get('y_mm', 0):.1f} mm")
                    c3.metric("Z", f"{bb.get('z_mm', 0):.1f} mm")
            else:
                st.info("pythonOCC 미설치 — 수동 입력 모드")

        part_name  = st.text_input("부품명", value=part_name)
        part_no    = st.text_input("형번", value=part_no, help="예: A001, BRKT-001")
        step_fname = st.text_input("도면 파일명", value=step_fname)

        if volume_cm3 is None:
            volume_cm3 = st.number_input("바운딩박스 부피 (cm³)",
                                         min_value=0.1, value=100.0, step=10.0)
        else:
            st.number_input("바운딩박스 부피 (cm³) — STEP 자동",
                            value=volume_cm3, disabled=True)
        if hole_count is None:
            hole_count  = st.number_input("홀 개수", min_value=0, value=0, step=1)
        if setup_count is None:
            setup_count = st.number_input("셋업 횟수", min_value=1, value=1, step=1)

        st.divider()
        st.subheader("2. 난이도")
        if auto_diff:
            st.info(f"자동 추정: **{auto_diff['level']}등급** · {auto_diff['reason']}")
            default_diff_idx = auto_diff["level"] - 1
        else:
            default_diff_idx = 0

        diff_options     = get_all_options()
        selected_diff    = st.selectbox("난이도 등급 (최종 확정)",
                                        diff_options, index=default_diff_idx)
        difficulty_level = parse_level_from_option(selected_diff)
        difficulty_coeff = get_coefficient(difficulty_level)
        st.caption(f"계수: **{difficulty_coeff}** — {DIFFICULTY_TABLE[difficulty_level]['desc']}")

    with col_right:
        st.subheader("3. 협력사 단가")

        vendors_df = load_vendors()
        if vendors_df.empty:
            st.warning("등록된 업체가 없습니다. [업체 관리] 메뉴에서 먼저 등록해주세요.")
            vendor_id   = None
            vendor_name = st.text_input("업체명 (직접 입력)")
        else:
            vendor_options = ["직접 입력"] + vendors_df["name"].tolist()
            sel_vendor = st.selectbox("협력사 선택", vendor_options)
            if sel_vendor == "직접 입력":
                vendor_id   = None
                vendor_name = st.text_input("업체명")
            else:
                vendor_row  = vendors_df[vendors_df["name"] == sel_vendor].iloc[0]
                vendor_id   = int(vendor_row["id"])
                vendor_name = sel_vendor
                st.caption(
                    f"담당: {vendor_row.get('contact', '—') or '—'} · "
                    f"연락처: {vendor_row.get('phone', '—') or '—'}"
                )

        vendor_price     = st.number_input("협력사 단가 (원)", min_value=0,
                                           value=100000, step=1000)
        postprocess_cost = st.number_input("후처리비 (원)", min_value=0,
                                           value=0, step=1000,
                                           help="도금·열처리 등 단가에 포함된 금액")
        estimated_hours  = st.number_input("예상 가공시간 (h)",
                                           min_value=0.1, value=1.0, step=0.1)
        note             = st.text_input("메모 (선택)")

        st.divider()
        try:
            mat_result = calc_material_cost(
                volume_cm3=volume_cm3, material_code=sel_mat_code,
                form=form_key, loss_rate_override=loss_rate,
                price_override=price_override,
            )
            st.metric(
                "예상 재료비", f"{mat_result['material_cost']:,.0f} 원",
                help=f"{mat_result['weight_kg']:.3f}kg × "
                     f"로스 {mat_result['loss_rate'] * 100:.0f}% × "
                     f"{mat_result['price_used']:,.0f}원/kg"
            )
        except Exception as e:
            st.warning(f"재료비 미리보기 오류: {e}")
            mat_result = None

        run_audit = st.button("🔍 감사 실행", type="primary", use_container_width=True)

    if run_audit:
        if not part_name:
            st.warning("부품명을 입력해주세요.")
        elif mat_result is None:
            st.error("재료비 계산 오류.")
        else:
            try:
                result = audit(
                    vendor_price=vendor_price,
                    material_cost=mat_result["material_cost"],
                    hourly_rate=hourly_rate,
                    estimated_hours=estimated_hours,
                    difficulty_coeff=difficulty_coeff,
                    postprocess_cost=postprocess_cost,
                )
                st.divider()
                st.subheader("📊 감사 결과")
                render_verdict(result)
                render_cost_cards(result)

                with st.expander("상세 수치"):
                    st.dataframe(pd.DataFrame({
                        "항목": ["협력사 단가", "재료비", "후처리비", "실질 가공비",
                                 "적정 가공비", "오차율", "임률", "가공시간", "난이도 계수"],
                        "값":   [f"{result.vendor_price:,.0f} 원",
                                 f"{result.material_cost:,.0f} 원",
                                 f"{result.postprocess_cost:,.0f} 원",
                                 f"{result.actual_machining:,.0f} 원",
                                 f"{result.standard_machining:,.0f} 원",
                                 f"{result.variance_pct:.1f} %",
                                 f"{result.hourly_rate:,.0f} 원/h",
                                 f"{result.estimated_hours:.1f} h",
                                 f"{result.difficulty_coeff}"]
                    }), hide_index=True)

                st.subheader("📈 재료비 시나리오")
                adj = st.slider("재료비 조정률 (%)", -30, 30, 0, 1)
                if adj != 0:
                    sc = audit_scenario(result, adj)
                    s1, s2, s3 = st.columns(3)
                    s1.metric("조정 재료비",     f"{sc.material_cost:,.0f} 원",
                              delta=f"{adj:+d}%")
                    s2.metric("조정 실질 가공비", f"{sc.actual_machining:,.0f} 원")
                    s3.metric("조정 오차율",      f"{sc.variance_pct:.1f}%",
                              delta=f"{sc.variance_pct - result.variance_pct:+.1f}%",
                              delta_color="inverse")
                    render_verdict(sc)

                # 이력 저장
                save_to_db(
                    part_name=part_name, part_no=part_no, step_file=step_fname,
                    material_code=sel_mat_code,
                    vendor_id=vendor_id, vendor_name=vendor_name,
                    result=result, difficulty_level=difficulty_level,
                    estimated_hours=estimated_hours,
                    volume_cm3=volume_cm3, hole_count=hole_count,
                    setup_count=setup_count,
                    price_snapshot=mat_result["price_used"],
                    note=note,
                )
                st.success("이력 저장 완료")

                # 방금 저장된 ID
                try:
                    conn = sqlite3.connect(DB_PATH)
                    last_id = conn.execute(
                        "SELECT MAX(id) FROM history_log"
                    ).fetchone()[0]
                    conn.close()
                except Exception:
                    last_id = None

                # 유사 부품 매칭
                similar_df = render_similar_parts(
                    volume_cm3=float(volume_cm3),
                    hole_count=int(hole_count),
                    setup_count=int(setup_count),
                    current_vendor_price=vendor_price,
                    saved_id=last_id,
                )

                # ── 협상 근거 엑셀 다운로드 ──────────────────
                st.divider()
                st.subheader("📥 협상 근거 리포트")

                # 시뮬레이션도 함께 포함
                sim_df = calc_vendor_simulation(
                    volume_cm3=float(volume_cm3),
                    material_code=sel_mat_code,
                    form_key=form_key,
                    loss_rate=loss_rate,
                    hourly_rate=hourly_rate,
                    estimated_hours=estimated_hours,
                    difficulty_coeff=difficulty_coeff,
                    postprocess_cost=postprocess_cost,
                    price_override=price_override,
                )

                excel_bytes = build_report_excel(
                    part_name=part_name,
                    part_no=part_no or "",
                    vendor_name=vendor_name or "",
                    material_code=sel_mat_code,
                    result=result,
                    difficulty_level=difficulty_level,
                    estimated_hours=estimated_hours,
                    similar_df=similar_df if isinstance(similar_df, pd.DataFrame) else pd.DataFrame(),
                    sim_df=sim_df if not sim_df.empty else None,
                    audit_date=datetime.now().strftime("%Y-%m-%d %H:%M"),
                )
                fname = f"협상근거_{part_name}_{datetime.now().strftime('%Y%m%d')}.xlsx"
                st.download_button(
                    "📊 협상 근거 엑셀 다운로드",
                    data=excel_bytes,
                    file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                    use_container_width=True,
                )

            except Exception as e:
                st.error(f"감사 실행 오류: {e}")


# ══════════════════════════════════════════════════════════════
# 페이지: 업체별 예상 단가 시뮬레이션
# ══════════════════════════════════════════════════════════════
def page_simulation(mat_df, sel_mat_code, mat_info, form_key,
                    loss_rate, hourly_rate, price_override):
    st.title("📈 업체별 예상 단가 시뮬레이션")
    st.caption(
        "신규 부품 발주 전, 등록 업체별 과거 마진 패턴을 적용해 "
        "예상 단가를 추정합니다. 이력 3건 이상인 업체만 신뢰도가 높습니다."
    )

    col_l, col_r = st.columns([1, 1], gap="large")

    with col_l:
        st.subheader("부품 사양 입력")
        sim_part_name    = st.text_input("부품명 (참고용)")
        sim_volume       = st.number_input("바운딩박스 부피 (cm³)",
                                           min_value=0.1, value=100.0, step=10.0,
                                           key="sim_vol")
        sim_hole         = st.number_input("홀 개수", min_value=0, value=0, step=1,
                                           key="sim_hole")
        sim_setup        = st.number_input("셋업 횟수", min_value=1, value=1, step=1,
                                           key="sim_setup")
        sim_pp           = st.number_input("후처리비 (원)", min_value=0, value=0,
                                           step=1000, key="sim_pp")

        st.subheader("난이도")
        diff_opts   = get_all_options()
        sim_diff_s  = st.selectbox("난이도 등급", diff_opts, key="sim_diff")
        sim_diff_lv = parse_level_from_option(sim_diff_s)
        sim_coeff   = get_coefficient(sim_diff_lv)
        st.caption(f"계수: **{sim_coeff}**")

        sim_hours = st.number_input("예상 가공시간 (h)",
                                    min_value=0.1, value=1.0, step=0.1,
                                    key="sim_hours")

        run_sim = st.button("🔍 시뮬레이션 실행", type="primary",
                            use_container_width=True)

    with col_r:
        st.subheader("결과")
        if run_sim:
            sim_df = calc_vendor_simulation(
                volume_cm3=float(sim_volume),
                material_code=sel_mat_code,
                form_key=form_key,
                loss_rate=loss_rate,
                hourly_rate=hourly_rate,
                estimated_hours=sim_hours,
                difficulty_coeff=sim_coeff,
                postprocess_cost=sim_pp,
                price_override=price_override,
            )

            if sim_df.empty:
                st.info("등록된 업체가 없거나 이력이 부족합니다.")
            else:
                # 순위 표시
                st.markdown("##### 예상 단가 순위 (낮은 순)")
                for rank, (_, row) in enumerate(sim_df.iterrows(), 1):
                    pct   = row["예상오차율"]
                    price = row["예상단가"]
                    total = row["이력건수"]

                    if pct > 110:
                        color = "🔴"
                    elif pct < 80:
                        color = "🔵"
                    else:
                        color = "🟢"

                    reliability = "★★★" if total >= 5 else ("★★☆" if total >= 3 else "★☆☆")
                    st.metric(
                        label=f"{rank}위  {color}  {row['업체명']}  {reliability}",
                        value=f"{price:,.0f} 원",
                        delta=f"예상 오차율 {pct:.1f}%",
                        delta_color="inverse",
                    )

                st.divider()
                st.dataframe(
                    sim_df[["업체명", "예상단가", "예상가공비", "재료비",
                            "예상오차율", "이력건수", "비고"]].rename(columns={
                        "예상단가": "예상 단가(원)", "예상가공비": "예상 가공비(원)",
                        "재료비": "재료비(원)", "예상오차율": "예상 오차율(%)",
                        "이력건수": "이력 건수",
                    }),
                    hide_index=True, use_container_width=True,
                    column_config={
                        "예상 단가(원)":  st.column_config.NumberColumn(format="%,.0f"),
                        "예상 가공비(원)": st.column_config.NumberColumn(format="%,.0f"),
                        "재료비(원)":     st.column_config.NumberColumn(format="%,.0f"),
                        "예상 오차율(%)": st.column_config.NumberColumn(format="%.1f %%"),
                    }
                )
                st.caption("★★★ 이력 5건↑  ★★☆ 3~4건  ★☆☆ 1~2건 (신뢰도 낮음)")

                # 엑셀 다운로드
                excel_bytes = build_report_excel(
                    part_name=sim_part_name or "신규부품",
                    part_no="",
                    vendor_name="시뮬레이션",
                    material_code=sel_mat_code,
                    result=audit(
                        vendor_price=float(sim_df["예상단가"].iloc[0]),
                        material_cost=float(sim_df["재료비"].iloc[0]),
                        hourly_rate=hourly_rate,
                        estimated_hours=sim_hours,
                        difficulty_coeff=sim_coeff,
                        postprocess_cost=sim_pp,
                    ),
                    difficulty_level=sim_diff_lv,
                    estimated_hours=sim_hours,
                    similar_df=pd.DataFrame(),
                    sim_df=sim_df,
                    audit_date=datetime.now().strftime("%Y-%m-%d %H:%M"),
                )
                fname = f"시뮬레이션_{sim_part_name or '신규'}_{datetime.now().strftime('%Y%m%d')}.xlsx"
                st.download_button(
                    "📥 시뮬레이션 결과 엑셀 다운로드",
                    data=excel_bytes,
                    file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )


# ══════════════════════════════════════════════════════════════
# 페이지: 업체 관리
# ══════════════════════════════════════════════════════════════
def page_vendors():
    st.title("🏢 업체 관리")

    if "show_add_vendor" not in st.session_state:
        st.session_state.show_add_vendor = False

    col_title, col_btn = st.columns([4, 1])
    with col_btn:
        if st.button("➕ 업체 추가", use_container_width=True):
            st.session_state.show_add_vendor = not st.session_state.show_add_vendor

    if st.session_state.show_add_vendor:
        with st.container(border=True):
            st.subheader("신규 업체 등록")
            with st.form("add_vendor_form", clear_on_submit=True):
                c1, c2  = st.columns(2)
                name    = c1.text_input("업체명 *", placeholder="필수 입력")
                contact = c2.text_input("담당자명")
                phone   = c1.text_input("연락처", placeholder="010-0000-0000")
                email   = c2.text_input("이메일")
                note    = st.text_area("메모", height=80)
                col_s, col_c = st.columns(2)
                submitted = col_s.form_submit_button("✅ 등록", type="primary",
                                                     use_container_width=True)
                cancelled = col_c.form_submit_button("취소", use_container_width=True)

            if submitted:
                if not name:
                    st.error("업체명은 필수입니다.")
                else:
                    add_vendor(name, contact, phone, email, note)
                    st.success(f"'{name}' 등록 완료")
                    st.session_state.show_add_vendor = False
                    st.rerun()
            if cancelled:
                st.session_state.show_add_vendor = False
                st.rerun()

    st.divider()

    vendors_df = load_vendors()
    if vendors_df.empty:
        st.info("등록된 업체가 없습니다. [업체 추가] 버튼을 눌러 등록해주세요.")
        return

    stats_df = load_vendor_stats()

    for _, row in vendors_df.iterrows():
        vid  = int(row["id"])
        stat = stats_df[stats_df["id"] == vid]

        with st.expander(
            f"**{row['name']}**  ·  {row.get('contact', '') or '담당자 미등록'}"
        ):
            c1, c2, c3 = st.columns(3)
            c1.write(f"📞 {row.get('phone', '—') or '—'}")
            c2.write(f"✉️ {row.get('email', '—') or '—'}")
            c3.write(f"📝 {row.get('note', '—') or '—'}")

            if not stat.empty and int(stat.iloc[0]["total"]) > 0:
                s = stat.iloc[0]
                st.divider()
                m1, m2, m3, m4 = st.columns(4)
                m1.metric("감사 건수",   f"{int(s['total'])}건")
                m2.metric("평균 오차율", f"{s['avg_variance']}%")
                m3.metric("발주 건수",   f"{int(s['ordered_count'])}건")
                m4.metric("재료비 비중", f"{s['mat_ratio']}%")

                if int(s["total"]) >= 5:
                    avg_v = float(s["avg_variance"])
                    if avg_v > 115:
                        msg = f"⚠ 오차율 높음 (평균 {avg_v:.1f}%). 항목별 소명 요청 검토."
                    elif avg_v > 105:
                        msg = f"🔶 소폭 과다 경향 (평균 {avg_v:.1f}%). 단가 재협의 여지 있음."
                    elif avg_v >= 90:
                        msg = f"✅ 적정 수준 (평균 {avg_v:.1f}%)."
                    else:
                        msg = f"🔵 저가 경향 (평균 {avg_v:.1f}%). 품질 조건 명기 권장."
                    st.info(msg)

                hist = load_vendor_history(vid)
                if not hist.empty:
                    st.dataframe(
                        hist[["created_at", "part_name", "part_no",
                              "vendor_price", "variance_pct", "verdict"]].rename(columns={
                            "created_at": "일시", "part_name": "부품명", "part_no": "형번",
                            "vendor_price": "단가", "variance_pct": "오차율(%)", "verdict": "판정"
                        }),
                        hide_index=True, use_container_width=True,
                        column_config={
                            "오차율(%)": st.column_config.NumberColumn(format="%.1f %%"),
                            "단가":      st.column_config.NumberColumn(format="%,.0f 원"),
                        }
                    )

            st.divider()
            ec1, _, ec3 = st.columns([2, 2, 1])
            with ec1:
                with st.popover("✏️ 정보 수정"):
                    with st.form(f"edit_{vid}"):
                        nname    = st.text_input("업체명",  value=row["name"])
                        ncontact = st.text_input("담당자",  value=row.get("contact", "") or "")
                        nphone   = st.text_input("연락처",  value=row.get("phone", "") or "")
                        nemail   = st.text_input("이메일",  value=row.get("email", "") or "")
                        nnote    = st.text_area("메모",     value=row.get("note", "") or "")
                        if st.form_submit_button("저장"):
                            update_vendor(vid, nname, ncontact, nphone, nemail, nnote)
                            st.success("수정 완료")
                            st.rerun()
            with ec3:
                if st.button("🗑 비활성화", key=f"del_{vid}"):
                    if vendor_has_history(vid):
                        st.warning("이력이 있어 비활성화 처리됩니다. (이력 유지)")
                    deactivate_vendor(vid)
                    st.rerun()


# ══════════════════════════════════════════════════════════════
# 페이지: 업체 비교
# ══════════════════════════════════════════════════════════════
def page_compare():
    st.title("📊 업체 비교")

    stats_df = load_vendor_stats()
    if stats_df.empty or stats_df["total"].sum() == 0:
        st.info("감사 이력이 없습니다. 역산 감사를 먼저 실행해주세요.")
        return

    stats_df = stats_df[stats_df["total"] > 0].copy()

    cols = st.columns(max(len(stats_df), 1))
    for i, (_, row) in enumerate(stats_df.iterrows()):
        avg_v = float(row["avg_variance"]) if row["avg_variance"] else 0
        icon  = "🟢" if 90 <= avg_v <= 110 else ("🔴" if avg_v > 110 else "🔵")
        with cols[i]:
            st.metric(f"{icon} {row['name']}",
                      f"평균 {avg_v:.1f}%", f"총 {int(row['total'])}건")

    st.divider()
    st.dataframe(
        stats_df.rename(columns={
            "name": "업체명", "total": "건수",
            "avg_variance": "평균오차율(%)", "min_variance": "최저",
            "max_variance": "최고", "ordered_count": "발주건",
            "mat_ratio": "재료비비중(%)"
        })[["업체명", "건수", "평균오차율(%)", "최저", "최고", "발주건", "재료비비중(%)"]],
        hide_index=True, use_container_width=True
    )

    st.subheader("형번별 단가 비교")
    conn = sqlite3.connect(DB_PATH)
    try:
        cross_df = pd.read_sql("""
            SELECT h.part_no, h.part_name, v.name AS vendor_name,
                   h.vendor_price, h.variance_pct, h.verdict, h.created_at
            FROM history_log h
            JOIN vendors v ON h.vendor_id = v.id
            WHERE h.part_no IS NOT NULL AND h.part_no != ''
            ORDER BY h.part_no, h.created_at DESC
        """, conn)
    finally:
        conn.close()

    if cross_df.empty:
        st.info("형번이 입력된 이력이 없습니다.")
    else:
        sel_part = st.selectbox("형번 선택", cross_df["part_no"].unique().tolist())
        st.dataframe(
            cross_df[cross_df["part_no"] == sel_part][
                ["vendor_name", "vendor_price", "variance_pct", "verdict", "created_at"]
            ].rename(columns={
                "vendor_name": "업체명", "vendor_price": "단가",
                "variance_pct": "오차율(%)", "verdict": "판정", "created_at": "일시"
            }),
            hide_index=True, use_container_width=True,
            column_config={
                "오차율(%)": st.column_config.NumberColumn(format="%.1f %%"),
                "단가":      st.column_config.NumberColumn(format="%,.0f 원"),
            }
        )


# ══════════════════════════════════════════════════════════════
# 페이지: 감사 이력
# ══════════════════════════════════════════════════════════════
def page_history():
    st.title("📁 감사 이력")
    df = load_history()
    if df.empty:
        st.info("저장된 이력이 없습니다.")
        return

    with st.expander("🔎 검색 필터", expanded=True):
        r1c1, r1c2, r1c3 = st.columns(3)
        f_verdict = r1c1.multiselect("판정", df["verdict"].dropna().unique().tolist())
        f_vendor  = r1c2.multiselect("업체", df["vendor_name"].dropna().unique().tolist())
        f_text    = r1c3.text_input("통합 검색", placeholder="부품명 / 형번 / 메모")

        r2c1, r2c2 = st.columns(2)
        df["created_at"] = pd.to_datetime(df["created_at"])
        min_date   = df["created_at"].min().date()
        max_date   = df["created_at"].max().date()
        date_range = r2c1.date_input("감사 일자 범위",
                                     value=(min_date, max_date),
                                     min_value=min_date, max_value=max_date)
        v_min = float(df["variance_pct"].min()) if df["variance_pct"].notna().any() else 0.0
        v_max = float(df["variance_pct"].max()) if df["variance_pct"].notna().any() else 200.0
        variance_range = r2c2.slider("오차율 범위 (%)", 0.0, 300.0,
                                     (max(0.0, v_min), min(300.0, v_max)), 1.0)

    if f_verdict:
        df = df[df["verdict"].isin(f_verdict)]
    if f_vendor:
        df = df[df["vendor_name"].isin(f_vendor)]
    if f_text:
        mask = (
            df["part_name"].str.contains(f_text, case=False, na=False) |
            df["part_no"].str.contains(f_text, case=False, na=False) |
            df["note"].str.contains(f_text, case=False, na=False)
        )
        df = df[mask]
    if isinstance(date_range, (list, tuple)) and len(date_range) == 2:
        start_dt = pd.Timestamp(date_range[0])
        end_dt   = pd.Timestamp(date_range[1]) + pd.Timedelta(days=1)
        df = df[(df["created_at"] >= start_dt) & (df["created_at"] < end_dt)]
    df = df[
        df["variance_pct"].isna() |
        ((df["variance_pct"] >= variance_range[0]) &
         (df["variance_pct"] <= variance_range[1]))
    ]

    col_map = {
        "id": "ID", "created_at": "일시", "part_name": "부품명", "part_no": "형번",
        "vendor_name": "업체", "vendor_price": "협력사단가",
        "material_cost": "재료비", "actual_machining": "실질가공비",
        "standard_machining": "적정가공비", "variance_pct": "오차율(%)",
        "verdict": "판정", "difficulty_level": "난이도",
        "volume_cm3": "부피(cm³)", "hole_count": "홀", "setup_count": "셋업",
        "note": "메모"
    }
    show_cols = [c for c in col_map if c in df.columns]
    df_show   = df[show_cols].rename(columns=col_map)

    st.dataframe(
        df_show, hide_index=True, use_container_width=True,
        column_config={
            "오차율(%)":  st.column_config.NumberColumn(format="%.1f %%"),
            "협력사단가": st.column_config.NumberColumn(format="%,.0f 원"),
            "재료비":     st.column_config.NumberColumn(format="%,.0f 원"),
            "실질가공비": st.column_config.NumberColumn(format="%,.0f 원"),
            "적정가공비": st.column_config.NumberColumn(format="%,.0f 원"),
            "부피(cm³)":  st.column_config.NumberColumn(format="%.1f"),
        }
    )
    st.caption(f"전체 {len(load_history())}건 · 표시 {len(df_show)}건")

    st.divider()
    csv = df_show.to_csv(index=False, encoding="utf-8-sig")
    st.download_button(
        "📥 CSV 다운로드", data=csv,
        file_name=f"감사이력_{datetime.now().strftime('%Y%m%d')}.csv",
        mime="text/csv"
    )


if __name__ == "__main__":
    main()
