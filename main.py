"""
가공 견적 AI 엔진 PRO v2.1
main.py — Streamlit 메인 UI

변경사항 (v2.1 → v2.2):
- [버그] 업체 추가 탭: rerun 후 탭이 초기화되는 문제 수정
  → st.session_state로 활성 탭 위치 유지
- [강화] 감사 이력 검색: 날짜 범위 / 오차율 범위 / 통합 텍스트 검색 추가
- [신규] 유사 부품 매칭: 감사 결과 하단에 자동 표시
  → 부피 ±20%, 홀 ±2개, 셋업 동일 조건으로 전 업체 이력 검색
  → 협상 근거 문구 자동 생성
"""

import streamlit as st
import pandas as pd
import sqlite3
import tempfile
import os
from pathlib import Path
from datetime import datetime, date

from engine import (
    audit, audit_scenario, AuditResult,
    calc_material_cost, load_material_master, get_market_prices,
    get_coefficient, get_label, get_all_options, parse_level_from_option,
    DIFFICULTY_TABLE,
    analyze_step_file, is_step_available,
)

BASE_DIR = Path(__file__).parent
DB_PATH  = BASE_DIR / "data" / "history_log.db"


# ── DB 초기화 ──────────────────────────────────────────────────
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


# ── 업체 CRUD ──────────────────────────────────────────────────
def load_vendors(active_only=True) -> pd.DataFrame:
    """vendors 테이블 조회. active_only=True이면 활성 업체만 반환."""
    conn = sqlite3.connect(DB_PATH)
    try:
        q = "SELECT * FROM vendors"
        if active_only:
            q += " WHERE is_active = 1"
        return pd.read_sql(q + " ORDER BY name", conn)
    finally:
        conn.close()


def add_vendor(name: str, contact="", phone="", email="", note=""):
    """신규 업체 등록."""
    conn = sqlite3.connect(DB_PATH)
    try:
        conn.execute(
            "INSERT INTO vendors (name,contact,phone,email,note) VALUES (?,?,?,?,?)",
            (name, contact, phone, email, note)
        )
        conn.commit()
    finally:
        conn.close()


def update_vendor(vid: int, name: str, contact: str, phone: str, email: str, note: str):
    """업체 정보 수정."""
    conn = sqlite3.connect(DB_PATH)
    try:
        conn.execute(
            "UPDATE vendors SET name=?,contact=?,phone=?,email=?,note=? WHERE id=?",
            (name, contact, phone, email, note, vid)
        )
        conn.commit()
    finally:
        conn.close()


def deactivate_vendor(vid: int):
    """업체 비활성화 (이력 보존)."""
    conn = sqlite3.connect(DB_PATH)
    try:
        conn.execute("UPDATE vendors SET is_active=0 WHERE id=?", (vid,))
        conn.commit()
    finally:
        conn.close()


def vendor_has_history(vid: int) -> bool:
    """해당 업체의 감사 이력 존재 여부 확인."""
    conn = sqlite3.connect(DB_PATH)
    try:
        count = conn.execute(
            "SELECT COUNT(*) FROM history_log WHERE vendor_id=?", (vid,)
        ).fetchone()[0]
        return count > 0
    finally:
        conn.close()


# ── 이력 저장·조회 ──────────────────────────────────────────────
def save_to_db(part_name, part_no, step_file, material_code,
               vendor_id, vendor_name, result: AuditResult,
               difficulty_level, estimated_hours,
               volume_cm3=None, hole_count=None, setup_count=None,
               price_snapshot=None, note=""):
    """감사 결과 1건을 history_log에 저장."""
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
    """전체 감사 이력 조회 (최신순)."""
    if not DB_PATH.exists():
        return pd.DataFrame()
    conn = sqlite3.connect(DB_PATH)
    try:
        return pd.read_sql(
            "SELECT * FROM history_log ORDER BY created_at DESC", conn
        )
    finally:
        conn.close()


def load_vendor_history(vendor_id: int) -> pd.DataFrame:
    """특정 업체의 감사 이력 조회."""
    conn = sqlite3.connect(DB_PATH)
    try:
        return pd.read_sql(
            "SELECT * FROM history_log WHERE vendor_id=? ORDER BY created_at DESC",
            conn, params=(vendor_id,)
        )
    finally:
        conn.close()


def load_vendor_stats() -> pd.DataFrame:
    """업체별 감사 통계 집계."""
    conn = sqlite3.connect(DB_PATH)
    try:
        return pd.read_sql("""
            SELECT
                v.id, v.name,
                COUNT(h.id)                                                 AS total,
                ROUND(AVG(h.variance_pct), 1)                              AS avg_variance,
                ROUND(MIN(h.variance_pct), 1)                              AS min_variance,
                ROUND(MAX(h.variance_pct), 1)                              AS max_variance,
                SUM(COALESCE(h.ordered, 0))                                AS ordered_count,
                ROUND(AVG(h.material_cost / NULLIF(h.vendor_price,0)*100), 1) AS mat_ratio
            FROM vendors v
            LEFT JOIN history_log h ON v.id = h.vendor_id
            WHERE v.is_active = 1
            GROUP BY v.id
        """, conn)
    finally:
        conn.close()


def search_similar_parts(volume_cm3: float, hole_count: int,
                         setup_count: int, exclude_id: int = None) -> pd.DataFrame:
    """
    유사 부품 이력 검색.
    조건: 부피 ±20%, 홀 개수 ±2개, 셋업 횟수 동일.
    전 업체 이력 대상. 부피 차이 오름차순 상위 8건 반환.
    """
    if not DB_PATH.exists():
        return pd.DataFrame()
    conn = sqlite3.connect(DB_PATH)
    try:
        params = [
            volume_cm3 * 0.8, volume_cm3 * 1.2,  # 부피 범위
            hole_count - 2, hole_count + 2,        # 홀 범위
            setup_count,                            # 셋업 동일
            volume_cm3,                             # vol_diff 계산용
        ]
        exclude_clause = ""
        if exclude_id is not None:
            exclude_clause = "AND h.id != ?"
            params.append(exclude_id)

        return pd.read_sql(f"""
            SELECT
                h.id,
                h.created_at,
                h.part_name,
                h.part_no,
                h.vendor_name,
                h.vendor_price,
                h.material_cost,
                h.variance_pct,
                h.verdict,
                h.volume_cm3,
                h.hole_count,
                h.setup_count,
                h.difficulty_level,
                h.ordered,
                ABS(h.volume_cm3 - ?) AS vol_diff
            FROM history_log h
            WHERE
                h.volume_cm3  BETWEEN ? AND ?
                AND h.hole_count  BETWEEN ? AND ?
                AND h.setup_count = ?
                AND h.volume_cm3 IS NOT NULL
                {exclude_clause}
            ORDER BY vol_diff ASC
            LIMIT 8
        """, conn, params=[volume_cm3] + params[:-1] + ([exclude_id] if exclude_id else []))
    finally:
        conn.close()


# ── 판정 UI ────────────────────────────────────────────────────
def render_verdict(result: AuditResult):
    """판정 결과를 색상에 맞는 Streamlit 알림으로 표시."""
    fn = {"green": st.success, "orange": st.warning,
          "yellow": st.warning, "red": st.error}.get(result.color, st.info)
    fn(f"**{result.verdict}** ({result.variance_pct:.1f}%) — {result.action}")


def render_cost_cards(result: AuditResult):
    """재료비 / 실질 가공비 / 적정 가공비 메트릭 카드 3개 표시."""
    c1, c2, c3 = st.columns(3)
    c1.metric("재료비", f"{result.material_cost:,.0f} 원")
    c2.metric("실질 가공비", f"{result.actual_machining:,.0f} 원")
    delta = result.actual_machining - result.standard_machining
    c3.metric("적정 가공비", f"{result.standard_machining:,.0f} 원",
              delta=f"{delta:+,.0f} 원", delta_color="inverse")


def render_similar_parts(volume_cm3: float, hole_count: int,
                         setup_count: int, current_vendor_price: float,
                         saved_id: int = None):
    """
    유사 부품 이력을 검색하여 표시하고 협상 근거 문구를 자동 생성.
    volume_cm3 / hole_count / setup_count 가 모두 0이 아닌 경우에만 동작.
    """
    if not (volume_cm3 and hole_count is not None and setup_count):
        return

    similar = search_similar_parts(volume_cm3, hole_count, setup_count,
                                   exclude_id=saved_id)
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
        "volume_cm3": "부피(cm³)", "hole_count": "홀", "setup_count": "셋업",
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

    # ── 협상 근거 문구 자동 생성 ──────────────────────────────
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


# ── 메인 ──────────────────────────────────────────────────────
def main():
    st.set_page_config(
        page_title="가공 견적 AI 엔진 PRO v2.1",
        page_icon="🔧",
        layout="wide"
    )
    init_db()

    # 업체 관리 탭 위치 유지용 session_state
    if "vendor_tab" not in st.session_state:
        st.session_state.vendor_tab = 0  # 0: 목록, 1: 추가

    with st.sidebar:
        st.markdown("### 🔧 Quote AI v2.1")
        st.caption("가공 견적 역산 감사 시스템")
        st.divider()

        menu = st.radio(
            "메뉴",
            ["📋 역산 감사", "🏢 업체 관리", "📊 업체 비교", "📁 감사 이력"],
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
    elif menu == "🏢 업체 관리":
        page_vendors()
    elif menu == "📊 업체 비교":
        page_compare()
    elif menu == "📁 감사 이력":
        page_history()


# ── 페이지: 역산 감사 ──────────────────────────────────────────
def page_audit(mat_df, sel_mat_code, mat_info, form_key,
               loss_rate, hourly_rate, price_override):
    """역산 감사 메인 페이지."""
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
                    st.caption(
                        f"홀: {hole_count}개 "
                        f"(심공: {sum(1 for h in holes if h['is_deep'])}개) · "
                        f"셋업 추정: {setup_count}회"
                    )
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
            hole_count  = st.number_input("홀 개수",  min_value=0, value=0, step=1)
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

                # 방금 저장된 이력 ID 조회 (유사 부품 검색 시 자기 자신 제외용)
                try:
                    conn = sqlite3.connect(DB_PATH)
                    last_id = conn.execute(
                        "SELECT MAX(id) FROM history_log"
                    ).fetchone()[0]
                    conn.close()
                except Exception:
                    last_id = None

                # 유사 부품 매칭 표시
                render_similar_parts(
                    volume_cm3=float(volume_cm3),
                    hole_count=int(hole_count),
                    setup_count=int(setup_count),
                    current_vendor_price=vendor_price,
                    saved_id=last_id,
                )

            except Exception as e:
                st.error(f"감사 실행 오류: {e}")


# ── 페이지: 업체 관리 ──────────────────────────────────────────
def page_vendors():
    """
    업체 등록·수정·비활성화 페이지.
    탭 위치를 st.session_state.vendor_tab으로 유지하여
    신규 등록 후 rerun 시 탭이 초기화되는 버그를 수정.
    """
    st.title("🏢 업체 관리")

    # ── 신규 등록 폼 (탭 바깥, 상단 고정) ────────────────────
    # 탭 내부에서 rerun하면 항상 첫 번째 탭으로 돌아가는 Streamlit 제약을 피하기 위해
    # 등록 폼을 expander로 분리하고, session_state로 열림 여부를 제어한다.
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
                note    = st.text_area("메모", height=80,
                                       placeholder="거래 특이사항, 전문 분야 등")
                col_s, col_c = st.columns([1, 1])
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

    # ── 등록 업체 목록 ────────────────────────────────────────
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
                        msg = f"⚠ 전반적으로 오차율이 높습니다 (평균 {avg_v:.1f}%). 항목별 소명 요청을 검토하세요."
                    elif avg_v > 105:
                        msg = f"🔶 소폭 과다 경향 (평균 {avg_v:.1f}%). 단가 재협의 여지가 있습니다."
                    elif avg_v >= 90:
                        msg = f"✅ 전반적으로 적정 수준입니다 (평균 {avg_v:.1f}%)."
                    else:
                        msg = f"🔵 저가 경향 (평균 {avg_v:.1f}%). 품질 조건을 명기하고 진행하세요."
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


# ── 페이지: 업체 비교 ──────────────────────────────────────────
def page_compare():
    """업체별 오차율 통계 및 형번별 단가 비교 페이지."""
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
    conn     = sqlite3.connect(DB_PATH)
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


# ── 페이지: 감사 이력 ──────────────────────────────────────────
def page_history():
    """
    감사 이력 조회 페이지.
    검색 조건: 판정 / 업체 / 통합 텍스트 / 날짜 범위 / 오차율 범위
    """
    st.title("📁 감사 이력")
    df = load_history()
    if df.empty:
        st.info("저장된 이력이 없습니다.")
        return

    # ── 검색 필터 ─────────────────────────────────────────────
    with st.expander("🔎 검색 필터", expanded=True):
        row1_c1, row1_c2, row1_c3 = st.columns(3)
        f_verdict = row1_c1.multiselect(
            "판정",
            df["verdict"].dropna().unique().tolist()
        )
        f_vendor  = row1_c2.multiselect(
            "업체",
            df["vendor_name"].dropna().unique().tolist()
        )
        f_text    = row1_c3.text_input(
            "통합 검색",
            placeholder="부품명 / 형번 / 메모",
            help="부품명·형번·메모를 동시에 검색합니다."
        )

        row2_c1, row2_c2 = st.columns(2)

        # 날짜 범위
        df["created_at"] = pd.to_datetime(df["created_at"])
        min_date = df["created_at"].min().date()
        max_date = df["created_at"].max().date()
        date_range = row2_c1.date_input(
            "감사 일자 범위",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date,
        )

        # 오차율 범위
        v_min = float(df["variance_pct"].min()) if df["variance_pct"].notna().any() else 0.0
        v_max = float(df["variance_pct"].max()) if df["variance_pct"].notna().any() else 200.0
        variance_range = row2_c2.slider(
            "오차율 범위 (%)",
            min_value=0.0, max_value=300.0,
            value=(max(0.0, v_min), min(300.0, v_max)),
            step=1.0,
        )

    # ── 필터 적용 ─────────────────────────────────────────────
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

    # 날짜 범위 적용
    if isinstance(date_range, (list, tuple)) and len(date_range) == 2:
        start_dt = pd.Timestamp(date_range[0])
        end_dt   = pd.Timestamp(date_range[1]) + pd.Timedelta(days=1)
        df = df[(df["created_at"] >= start_dt) & (df["created_at"] < end_dt)]

    # 오차율 범위 적용
    df = df[
        df["variance_pct"].isna() |
        ((df["variance_pct"] >= variance_range[0]) &
         (df["variance_pct"] <= variance_range[1]))
    ]

    # ── 테이블 출력 ───────────────────────────────────────────
    col_map = {
        "id": "ID",
        "created_at": "일시", "part_name": "부품명", "part_no": "형번",
        "vendor_name": "업체", "vendor_price": "협력사단가",
        "material_cost": "재료비", "actual_machining": "실질가공비",
        "standard_machining": "적정가공비", "variance_pct": "오차율(%)",
        "verdict": "판정", "difficulty_level": "난이도",
        "volume_cm3": "부피(cm³)", "hole_count": "홀", "setup_count": "셋업",
        "ordered": "발주확정", "note": "메모"
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
            "발주확정":   st.column_config.CheckboxColumn(),
        }
    )

    total_cnt   = len(load_history())
    display_cnt = len(df_show)
    st.caption(f"전체 {total_cnt}건 · 표시 {display_cnt}건")

    # ── 발주 확정 처리 ────────────────────────────────────────
    st.divider()
    st.subheader("발주 확정 처리")
    st.caption("실제 발주 확정 건을 표시하면 업체 패턴 분석에 반영됩니다.")
    col_id, col_btn, _ = st.columns([1, 1, 3])
    sel_id = col_id.number_input("감사 ID", min_value=1, step=1, label_visibility="collapsed")
    if col_btn.button("✅ 발주 확정"):
        conn = sqlite3.connect(DB_PATH)
        try:
            conn.execute("UPDATE history_log SET ordered=1 WHERE id=?", (sel_id,))
            conn.commit()
        finally:
            conn.close()
        st.success(f"ID {sel_id} 발주 확정 처리")
        st.rerun()

    # ── CSV 다운로드 ──────────────────────────────────────────
    csv = df_show.to_csv(index=False, encoding="utf-8-sig")
    st.download_button(
        "📥 CSV 다운로드", data=csv,
        file_name=f"감사이력_{datetime.now().strftime('%Y%m%d')}.csv",
        mime="text/csv"
    )


if __name__ == "__main__":
    main()
