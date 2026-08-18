# -*- coding: utf-8 -*-
import io
import streamlit as st
from report_generator import load_data, build_html, build_excel, agg, cancel_rate, fmt

st.set_page_config(
    page_title="트립비토즈 공구 리포트 생성기",
    page_icon="📊",
    layout="centered",
)

st.title("📊 트립비토즈 공구 리포트 생성기")
st.caption("예약 엑셀 파일을 업로드하면 HTML · XLSX 리포트와 메일 본문을 자동 생성합니다.")
st.divider()

uploaded = st.file_uploader("예약 엑셀 파일 (.xlsx)", type=["xlsx"])

if uploaded:
    try:
        df = load_data(uploaded)
    except Exception as e:
        st.error(f"파일 로드 실패: {e}")
        st.stop()

    숙소_auto = df['숙소'].iloc[0] if '숙소' in df.columns else ''
    예약일_min = df['예약일'].min().strftime('%Y.%m.%d')
    예약일_max = df['예약일'].max().strftime('%Y.%m.%d')
    투숙_min   = df['체크인'].min().strftime('%Y.%m.%d')
    투숙_max   = df['체크아웃'].max().strftime('%Y.%m.%d')

    st.success(f"✅ {len(df):,}건 로드 완료")

    # ── 입력 폼 ──────────────────────────────────────────────
    with st.form("report_form"):
        st.subheader("리포트 정보")
        c1, c2 = st.columns(2)
        with c1:
            influencer  = st.text_input("인플루언서명", placeholder="예: 마이아, 듀")
            hotel       = st.text_input("호텔·리조트명", value=숙소_auto,
                                        placeholder="예: 소노벨 단양")
        with c2:
            sale_period = st.text_input("판매 기간",
                                        value=f"{예약일_min} ~ {예약일_max}")
            stay_period = st.text_input("투숙 기간",
                                        value=f"{투숙_min} ~ {투숙_max}")

        submitted = st.form_submit_button(
            "🚀 리포트 생성", use_container_width=True, type="primary"
        )

    # ── 생성 ─────────────────────────────────────────────────
    if submitted:
        if not influencer and not hotel:
            st.warning("인플루언서명 또는 호텔명을 입력해주세요.")
            st.stop()

        title = (
            f"{influencer} × {hotel} — 공동구매 매출 실적 요약"
            if influencer else
            f"{hotel} — 공동구매 매출 실적 요약"
        )

        with st.spinner("리포트 생성 중..."):
            html_str  = build_html(df, title, sale_period, stay_period)
            xlsx_buf  = io.BytesIO()
            build_excel(df, title, sale_period, stay_period, xlsx_buf)
            xlsx_bytes = xlsx_buf.getvalue()

        a  = agg(df)
        cr = cancel_rate(a)

        wait      = df[df['대기예약여부']]
        wait_line = ''
        if len(wait) > 0:
            wc = len(wait[wait['상태'] == '확정'])
            wr = round(wc / len(wait) * 100, 1)
            wait_line = f'\n- 대기예약 전환율: {wr}%'

        def kr_date(d):
            y, mo, day = d.replace('.', '-').split('-')
            return f"{y}년 {mo}월 {day}일"

        s_start = sale_period.split('~')[0].strip()
        s_end   = sale_period.split('~')[1].strip()

        mail = f"""\
안녕하세요.
트립비토즈 [담당자명]입니다.

{kr_date(s_start)}부터 {kr_date(s_end)}까지
진행된 {title.replace(' — 공동구매 매출 실적 요약', '')} 공동구매 판매 실적을 아래와 같이 공유드립니다.

[실적 요약]
■ 판매 기간: {sale_period}
■ 투숙 기간: {stay_period}
■ 판매 실적
- 전체 예약: {fmt(a['전체건'])}건 / {fmt(a['전체박'])}박 / {fmt(a['전체매출'])}원
- 확정 예약: {fmt(a['확정건'])}건 / {fmt(a['확정박'])}박 / {fmt(a['확정매출'])}원
- 취소 건수: {fmt(a['취소건'])}건 (취소율 {cr}%){wait_line}

세부 내역은 아래 첨부된 표를 참고 부탁드리며,
추가 문의 사항이 있으시면 언제든지 연락 주시기 바랍니다.

이번 공구에 적극적으로 협조해 주신 덕분에 좋은 성과를 거둘 수 있었습니다.
앞으로도 지속적인 협력 부탁드립니다.

감사합니다.

[담당자명] 드림
트립비토즈"""

        safe = (title
                .replace(' ', '_').replace('×', 'X')
                .replace('—', '').replace('/', '_'))[:40]

        st.session_state['report'] = {
            'html_str':   html_str,
            'xlsx_bytes': xlsx_bytes,
            'title':      title,
            'safe':       safe,
            'mail':       mail,
        }

    # ── 리포트 결과 표시 (다운로드 후에도 유지) ───────────────
    if 'report' in st.session_state:
        r = st.session_state['report']
        st.success("✅ 생성 완료!")

        c1, c2 = st.columns(2)
        with c1:
            st.download_button(
                "📄 HTML 다운로드",
                r['html_str'].encode('utf-8'),
                file_name=f"report_{r['safe']}.html",
                mime="text/html",
                use_container_width=True,
            )
        with c2:
            st.download_button(
                "📊 XLSX 다운로드",
                r['xlsx_bytes'],
                file_name=f"report_{r['safe']}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        st.subheader("📧 메일 본문")
        st.code(r['mail'], language=None)
