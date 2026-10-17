import os
import html
from datetime import datetime
import streamlit as st
import pandas as pd

import collector

# ==========================================
# 페이지 기본 설정
# ==========================================
st.set_page_config(
    page_title="Gemini AI Enlgish Conversation Tutor V1.0",
    page_icon="🎓",
    layout="wide",
)

# 사용자 DB 초기화 (기본 관리자 admin / jhlee_english01 자동 생성 보장)
collector.init_users_db()


# ==========================================
# 세션 상태(Session State) 초기화
# ==========================================
def init_session_state():
    defaults = {
        "logged_in": False,
        "user_id": "",
        "current_menu": "Log In",
        "font_size": 14,
        "topic": "",
        "save_file_path": collector.DEFAULT_EXCEL_PATH,
        "chat_history": [],  # 현재 세션 대화 표시용 리스트
        "chat_session": None,
        "gemini_client": None,
        "custom_api_key": "",
        "model_name": "gemini-2.5-pro",
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


init_session_state()


# ==========================================
# 스타일(CSS) 설정 - 상단 메뉴바, 고정 헤더 테이블, 채팅창
# ==========================================
def inject_custom_css(font_size: int):
    st.markdown(
        f"""
        <style>
        /* 메인 타이틀 스타일 */
        .main-title {{
            font-size: 28px;
            font-weight: 800;
            color: #1F4E78;
            text-align: center;
            padding: 12px 0 8px 0;
            border-bottom: 2px solid #1F4E78;
            margin-bottom: 16px;
        }}

        /* 상태 배지 바 */
        .status-bar {{
            background-color: #F4F7FA;
            border: 1px solid #D6E0EA;
            border-radius: 6px;
            padding: 8px 16px;
            margin-bottom: 14px;
            font-size: 14px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}

        /* 01_Gemini_API_Eng_Study.py 스타일 대화창 */
        .chat-display-box {{
            background-color: #FFFFFF;
            border: 1px solid #CCCCCC;
            border-radius: 6px;
            padding: 16px;
            height: 430px;
            overflow-y: auto;
            font-size: {font_size}px;
            line-height: 1.5;
            font-family: 'Segoe UI', 'Malgun Gothic', sans-serif;
            margin-bottom: 12px;
        }}
        .chat-msg-me {{
            margin-bottom: 10px;
            word-wrap: break-word;
            white-space: pre-wrap;
        }}
        .chat-msg-correction {{
            color: red;
            font-weight: bold;
            margin-bottom: 10px;
            word-wrap: break-word;
            white-space: pre-wrap;
        }}
        .chat-msg-gemini {{
            margin-bottom: 14px;
            word-wrap: break-word;
            white-space: pre-wrap;
        }}

        /* 영어공부 복습: 제목 행 고정 + 우측 상하 스크롤 테이블 */
        .review-table-container {{
            max-height: 480px;
            overflow-y: scroll;
            overflow-x: auto;
            border: 1px solid #BDC3C7;
            border-radius: 6px;
            background-color: #FFFFFF;
        }}
        table.review-table {{
            width: 100%;
            border-collapse: collapse;
            font-size: {font_size}px;
            table-layout: fixed;
        }}
        table.review-table thead th {{
            position: sticky;
            top: 0;
            background-color: #1F4E78;
            color: #FFFFFF;
            font-weight: bold;
            text-align: center;
            padding: 10px 8px;
            border: 1px solid #163A59;
            z-index: 20;
        }}
        table.review-table tbody td {{
            padding: 8px 10px;
            border: 1px solid #E0E0E0;
            vertical-align: top;
            word-wrap: break-word;
            overflow-wrap: break-word;
            white-space: pre-wrap;
        }}
        table.review-table tbody tr:nth-child(even) {{
            background-color: #F9FBFD;
        }}
        table.review-table tbody tr:hover {{
            background-color: #EEF5FF;
        }}

        /* 알림 메시지 박스 */
        .msg-modal-box {{
            padding: 16px;
            border-radius: 8px;
            border-left: 6px solid #D9534F;
            background-color: #FDF2F2;
            color: #A94442;
            font-weight: bold;
            font-size: 16px;
            margin: 12px 0;
        }}
        .msg-success-box {{
            padding: 16px;
            border-radius: 8px;
            border-left: 6px solid #28A745;
            background-color: #F2FDF4;
            color: #1E7E34;
            font-weight: bold;
            font-size: 16px;
            margin: 12px 0;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )


inject_custom_css(st.session_state.font_size)

# ==========================================
# 상단 타이틀 출력
# ==========================================
st.markdown(
    '<div class="main-title">Gemini AI Enlgish Conversation Tutor V1.0</div>',
    unsafe_allow_html=True,
)

# ==========================================
# 상단 7개 메뉴 버튼 배치
# ("신규등록", "Log In", "Gemini 영어공부", "영어공부 복습", "엑셀 저장", "모든 데이터 삭제", "Log Out")
# ==========================================
MENU_ITEMS = [
    "신규등록",
    "Log In",
    "Gemini 영어공부",
    "영어공부 복습",
    "엑셀 저장",
    "모든 데이터 삭제",
    "Log Out",
]

nav_cols = st.columns(7)
for idx, menu_name in enumerate(MENU_ITEMS):
    with nav_cols[idx]:
        is_active = st.session_state.current_menu == menu_name
        btn_type = "primary" if is_active else "secondary"
        if st.button(menu_name, key=f"nav_btn_{idx}", use_container_width=True, type=btn_type):
            if menu_name == "Log Out":
                # Log Out 처리: 로그인 전 단계로 즉시 복귀
                st.session_state.logged_in = False
                st.session_state.user_id = ""
                st.session_state.chat_history = []
                st.session_state.chat_session = None
                st.session_state.current_menu = "Log In"
                st.toast("Log Out 되었습니다. Log In 화면으로 이동합니다.")
                st.rerun()
            elif menu_name in ["Gemini 영어공부", "영어공부 복습", "엑셀 저장", "모든 데이터 삭제"]:
                if not st.session_state.logged_in:
                    st.session_state.current_menu = menu_name
                else:
                    st.session_state.current_menu = menu_name
                st.rerun()
            else:
                st.session_state.current_menu = menu_name
                st.rerun()

# 현재 로그인 상태 표시바
if st.session_state.logged_in:
    st.markdown(
        f'<div class="status-bar">'
        f'<span>🟢 <b>로그인 상태:</b> 활성화됨 &nbsp;|&nbsp; <b>접속 ID:</b> <code>{html.escape(st.session_state.user_id)}</code></span>'
        f'<span><b>현재 메뉴:</b> {html.escape(st.session_state.current_menu)} &nbsp;|&nbsp; <b>Font Size:</b> {st.session_state.font_size}px</span>'
        f'</div>',
        unsafe_allow_html=True,
    )
else:
    st.markdown(
        f'<div class="status-bar">'
        f'<span>🔴 <b>로그인 상태:</b> 로그아웃 상태 (프로그램 사용을 위해 <b>신규등록</b> 후 <b>Log In</b> 해 주세요)</span>'
        f'<span><b>현재 메뉴:</b> {html.escape(st.session_state.current_menu)}</span>'
        f'</div>',
        unsafe_allow_html=True,
    )

# 사이드바: API Key 및 모델 설정 (클라우드 배포 및 로컬 .env 모두 지원)
with st.sidebar:
    st.subheader("⚙️ 환경 설정")
    detected_key = collector.get_api_key(st.session_state.custom_api_key)
    if detected_key:
        st.success("GOOGLE_API_KEY가 정상 감지되었습니다.")
    else:
        st.warning("GOOGLE_API_KEY가 감지되지 않았습니다.")

    api_key_input = st.text_input(
        "Google Gemini API Key (선택)",
        value=st.session_state.custom_api_key,
        type="password",
        help=".env 파일에 GOOGLE_API_KEY가 설정되어 있으면 비워두셔도 됩니다.",
    )
    if api_key_input != st.session_state.custom_api_key:
        st.session_state.custom_api_key = api_key_input
        st.session_state.chat_session = None

    model_choice = st.selectbox(
        "Gemini 모델 선택",
        options=["gemini-2.5-pro", "gemini-2.5-flash", "gemini-2.0-flash"],
        index=0,
        help="01_Gemini_API_Eng_Study.py의 기본 모델은 gemini-2.5-pro 입니다.",
    )
    if model_choice != st.session_state.model_name:
        st.session_state.model_name = model_choice
        st.session_state.chat_session = None


# ==========================================
# 화면 1: "신규등록" (회원 등록 + 관리자 기반 ID/PW 분실 초기화)
# ==========================================
def render_register_view():
    st.subheader("📝 신규등록 및 회원 ID/PW 초기화")

    tab_reg, tab_reset = st.tabs(["👤 신규 회원 등록", "🔐 ID / PW 분실 초기화 (관리자)"])

    with tab_reg:
        st.markdown("신규 회원의 경우 **ID**와 **Password(2회)**를 입력하여 프로그램 사용 권한을 받으실 수 있습니다.")
        with st.form("register_form", clear_on_submit=False):
            reg_id = st.text_input("신규 회원 ID", placeholder="사용하실 ID를 입력하세요 (예: Jay)")
            reg_pw1 = st.text_input("Password 입력 (1회)", type="password", placeholder="비밀번호 입력")
            reg_pw2 = st.text_input("Password 재입력 (2회 확인)", type="password", placeholder="비밀번호 다시 입력")
            submit_reg = st.form_submit_button("신규등록 완료", type="primary", use_container_width=True)

        if submit_reg:
            success, msg = collector.register_user(reg_id, reg_pw1, reg_pw2)
            if not success:
                if msg == "회원 등록이 이미 되었습니다":
                    # 작업지도서 명세: 이미 등록된 회원의 경우 "회원 등록이 이미 되었습니다" Msg 창 생성
                    st.markdown(
                        f'<div class="msg-modal-box">⚠️ {html.escape(msg)}</div>',
                        unsafe_allow_html=True,
                    )
                    st.warning(msg, icon="⚠️")
                else:
                    st.markdown(
                        f'<div class="msg-modal-box">❌ {html.escape(msg)}</div>',
                        unsafe_allow_html=True,
                    )
                    st.error(msg)
            else:
                st.markdown(
                    f'<div class="msg-success-box">✅ {html.escape(msg)}</div>',
                    unsafe_allow_html=True,
                )
                st.success(msg)

    with tab_reset:
        st.markdown(
            "ID 또는 Password를 분실한 경우 관리자 계정(`admin`) 인증을 통해 등록된 회원 ID를 조회하거나 Password를 초기화할 수 있습니다."
        )
        with st.form("admin_reset_form", clear_on_submit=False):
            col_a1, col_a2 = st.columns(2)
            with col_a1:
                admin_id_input = st.text_input("관리자 ID", value="admin")
            with col_a2:
                admin_pw_input = st.text_input("관리자 Password", type="password", placeholder="관리자 비밀번호 입력")

            st.divider()
            action_mode = st.radio(
                "초기화 작업 선택",
                options=[
                    "비밀번호(PW) 재설정 초기화",
                    "회원 ID/계정 삭제 초기화 (신규 재등록 허용)",
                    "등록된 회원 ID 목록 조회 (ID 분실 확인)",
                ],
                horizontal=True,
            )
            target_uid = st.text_input("초기화 대상 회원 ID", placeholder="초기화할 회원의 ID 입력")
            col_p1, col_p2 = st.columns(2)
            with col_p1:
                new_pw1 = st.text_input("새 Password (1회)", type="password")
            with col_p2:
                new_pw2 = st.text_input("새 Password 확인 (2회)", type="password")

            submit_reset = st.form_submit_button("초기화 / 조회 실행", type="primary", use_container_width=True)

        if submit_reset:
            if action_mode == "등록된 회원 ID 목록 조회 (ID 분실 확인)":
                ok, user_list, msg = collector.get_registered_user_ids(admin_id_input, admin_pw_input)
                if ok:
                    st.success(msg)
                    st.info(f"현재 등록된 회원 ID 목록: {', '.join(user_list)}")
                else:
                    st.error(msg)
            elif action_mode == "회원 ID/계정 삭제 초기화 (신규 재등록 허용)":
                ok, msg = collector.reset_user_account(
                    admin_id_input,
                    admin_pw_input,
                    target_uid,
                    delete_account=True,
                )
                if ok:
                    st.success(msg)
                else:
                    st.error(msg)
            else:
                ok, msg = collector.reset_user_account(
                    admin_id_input,
                    admin_pw_input,
                    target_uid,
                    new_pw1=new_pw1,
                    new_pw2=new_pw2,
                    delete_account=False,
                )
                if ok:
                    st.success(msg)
                else:
                    st.error(msg)


# ==========================================
# 화면 2: "Log In"
# ==========================================
def render_login_view():
    st.subheader("🔑 Log In")
    st.markdown("등록된 **ID**와 **Password**를 입력하여 프로그램을 활성화하세요.")

    col_left, col_center, col_right = st.columns([1, 2, 1])
    with col_center:
        with st.form("login_form", clear_on_submit=False):
            login_id = st.text_input("ID", placeholder="등록된 ID를 입력하세요")
            login_pw = st.text_input("Password", type="password", placeholder="Password를 입력하세요")
            submit_login = st.form_submit_button("Log In", type="primary", use_container_width=True)

        if submit_login:
            ok, msg = collector.verify_login(login_id, login_pw)
            if ok:
                uid = login_id.strip()
                st.session_state.logged_in = True
                st.session_state.user_id = uid
                st.session_state.chat_session = None

                # 기존 저장된 해당 ID의 로그를 불러와 대화창에 복원
                user_df = collector.load_user_logs(uid, st.session_state.save_file_path)
                restored_history = []
                for _, row in user_df.iterrows():
                    restored_history.append(
                        {
                            "User": str(row["User"]),
                            "Topic": str(row["Topic"]),
                            "Date": str(row["Date"]),
                            "Who": str(row["Who"]),
                            "Conversation History": str(row["Conversation History"]),
                            "Error Correction": str(row["Error Correction"]),
                        }
                    )
                st.session_state.chat_history = restored_history
                st.markdown(
                    f'<div class="msg-success-box">✅ {html.escape(msg)}</div>',
                    unsafe_allow_html=True,
                )
                st.success(msg)
                st.session_state.current_menu = "Gemini 영어공부"
                st.rerun()
            else:
                # 입력된 ID나 PW가 일치하지 않으면 Error Message 창 생성
                st.markdown(
                    f'<div class="msg-modal-box">🚨 Error Message: {html.escape(msg)}</div>',
                    unsafe_allow_html=True,
                )
                st.error(msg, icon="🚨")


# ==========================================
# 화면 3: "Gemini 영어공부" (01_Gemini_API_Eng_Study.py 동일 기능 및 Tone & Manner)
# ==========================================
def render_study_view():
    st.subheader("💬 Gemini 영어공부")

    # 1) [Font Size 입력부] & [엑셀 저장 파일 설정부]
    col_font1, col_font2, col_save = st.columns([1.2, 0.8, 4.0])
    with col_font1:
        new_font_val = st.number_input(
            "Font Size (8 ~ 50):",
            min_value=8,
            max_value=50,
            value=int(st.session_state.font_size),
            step=1,
        )
    with col_font2:
        st.write("")
        st.write("")
        if st.button("적용", use_container_width=True):
            st.session_state.font_size = int(new_font_val)
            st.rerun()
    with col_save:
        st.text_input(
            "엑셀 저장 파일 (기본 경로):",
            value=st.session_state.save_file_path,
            disabled=True,
        )

    # 2) [Topic 입력부]
    topic_val = st.text_input(
        "Conversation Topic:",
        value=st.session_state.topic,
        placeholder="e.g. Movie, Travel, Hobby...",
    )
    st.session_state.topic = topic_val

    # 3) [대화창 출력부] - 01_Gemini_API_Eng_Study.py와 동일한 색상(Me: 파란색, Error Correction: 빨간색, Gemini: 초록색)
    chat_html_parts = ['<div class="chat-display-box">']
    if not st.session_state.chat_history:
        chat_html_parts.append(
            '<div style="color: #888888;"><i>Conversation Topic을 입력하고 아래 입력창에 영어 문장을 입력해 대화를 시작하세요.</i></div>'
        )
    else:
        for item in st.session_state.chat_history:
            who = item.get("Who", "")
            text = html.escape(str(item.get("Conversation History", "")))
            correction = html.escape(str(item.get("Error Correction", "")).strip())

            if who == "Me":
                chat_html_parts.append(
                    f'<div class="chat-msg-me"><b style="color: blue;">Me:</b> {text}</div>'
                )
                if correction:
                    chat_html_parts.append(
                        f'<div class="chat-msg-correction">[Error Correction]: {correction}</div>'
                    )
            else:
                chat_html_parts.append(
                    f'<div class="chat-msg-gemini"><b style="color: green;">Gemini:</b> {text}</div>'
                )
    chat_html_parts.append("</div>")
    st.markdown("".join(chat_html_parts), unsafe_allow_html=True)

    # 4) [메시지 입력부] - 엔터키 및 Send 버튼 지원
    with st.form("chat_send_form", clear_on_submit=True):
        col_msg, col_btn = st.columns([5, 1])
        with col_msg:
            user_text = st.text_input(
                "Message",
                placeholder="Type your English sentence here...",
                label_visibility="collapsed",
            )
        with col_btn:
            send_clicked = st.form_submit_button("Send", type="primary", use_container_width=True)

    if send_clicked:
        current_topic = st.session_state.topic.strip()
        clean_user_text = user_text.strip()
        if not current_topic:
            st.warning("Please enter a Topic first! (먼저 Conversation Topic을 입력해 주세요.)")
        elif not clean_user_text:
            st.warning("영어 문장을 입력해 주세요.")
        else:
            user_timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            with st.spinner("Gemini is typing..."):
                try:
                    if st.session_state.chat_session is None:
                        _, chat_sess = collector.create_gemini_chat_session(
                            api_key=st.session_state.custom_api_key,
                            model_name=st.session_state.model_name,
                        )
                        st.session_state.chat_session = chat_sess

                    result = collector.ask_gemini_tutor(
                        st.session_state.chat_session,
                        current_topic,
                        clean_user_text,
                    )
                    correction = result.get("correction", "")
                    gemini_reply = result.get("response", "")
                    gemini_timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                    # Gemini_English.xlsx 포맷 (첫번째 열 User 추가)으로 즉시 저장
                    collector.append_conversation_turn(
                        user_id=st.session_state.user_id,
                        topic=current_topic,
                        user_text=clean_user_text,
                        correction=correction,
                        gemini_response=gemini_reply,
                        user_time=user_timestamp,
                        gemini_time=gemini_timestamp,
                        excel_path=st.session_state.save_file_path,
                    )

                    # 현재 화면 대화 내역 업데이트
                    st.session_state.chat_history.append(
                        {
                            "User": st.session_state.user_id,
                            "Topic": current_topic,
                            "Date": user_timestamp,
                            "Who": "Me",
                            "Conversation History": clean_user_text,
                            "Error Correction": correction,
                        }
                    )
                    st.session_state.chat_history.append(
                        {
                            "User": st.session_state.user_id,
                            "Topic": current_topic,
                            "Date": gemini_timestamp,
                            "Who": "Gemini",
                            "Conversation History": gemini_reply,
                            "Error Correction": "",
                        }
                    )
                    st.rerun()
                except Exception as e:
                    st.error(f"API Error - Failed to get response: {str(e)}")

    # 5) [종료 (Save & Exit) 버튼] - 원본 프로그램과 동일하게 제공
    if st.button("종료 (Save & Exit)", type="primary", use_container_width=True):
        st.success(
            f"'{st.session_state.user_id}'님의 모든 영어회화 기록이 '{st.session_state.save_file_path}'에 안전하게 저장되었습니다."
        )
        st.session_state.current_menu = "영어공부 복습"
        st.rerun()


# ==========================================
# 화면 4: "영어공부 복습" (로그인 ID 필터링 + 고정 헤더 + 상하 스크롤)
# ==========================================
def render_review_view():
    st.subheader(f"📖 영어공부 복습 (User: {st.session_state.user_id})")

    # 로그인한 ID에 해당하는 데이터만 선별 (타 사용자 로그 출력 금지)
    user_df = collector.load_user_logs(st.session_state.user_id, st.session_state.save_file_path)

    if user_df.empty:
        st.info(f"현재 '{st.session_state.user_id}'님의 영어회화 저장 기록이 없습니다. 'Gemini 영어공부'에서 대화를 시작해 보세요!")
        return

    st.caption(
        f"총 {len(user_df)}건의 대화 기록이 조회되었습니다. 우측 스크롤바를 사용하여 상하로 스크롤하더라도 상단 제목 행(User | Topic | Date | Who | Conversation History | Error Correction)은 고정됩니다."
    )

    # 제목 행 고정(Sticky Header) 및 우측 상하 스크롤바가 적용된 테이블 렌더링
    rows_html = []
    for _, row in user_df.iterrows():
        u_val = html.escape(str(row.get("User", "")))
        t_val = html.escape(str(row.get("Topic", "")))
        d_val = html.escape(str(row.get("Date", "")))
        w_val = html.escape(str(row.get("Who", "")))
        c_val = html.escape(str(row.get("Conversation History", "")))
        e_val = html.escape(str(row.get("Error Correction", "")))

        who_style = "color: blue; font-weight: bold;" if w_val == "Me" else "color: green; font-weight: bold;"
        err_style = "color: #D9534F; font-weight: 600;" if e_val else ""

        rows_html.append(
            f"<tr>"
            f'<td style="text-align:center;">{u_val}</td>'
            f'<td style="text-align:center;">{t_val}</td>'
            f'<td style="text-align:center;">{d_val}</td>'
            f'<td style="text-align:center; {who_style}">{w_val}</td>'
            f"<td>{c_val}</td>"
            f'<td style="{err_style}">{e_val}</td>'
            f"</tr>"
        )

    table_html = (
        '<div class="review-table-container">'
        '<table class="review-table">'
        "<colgroup>"
        '<col style="width: 9%;">'
        '<col style="width: 11%;">'
        '<col style="width: 15%;">'
        '<col style="width: 8%;">'
        '<col style="width: 32%;">'
        '<col style="width: 25%;">'
        "</colgroup>"
        "<thead>"
        "<tr>"
        "<th>User</th>"
        "<th>Topic</th>"
        "<th>Date</th>"
        "<th>Who</th>"
        "<th>Conversation History</th>"
        "<th>Error Correction</th>"
        "</tr>"
        "</thead>"
        "<tbody>"
        + "".join(rows_html)
        + "</tbody></table></div>"
    )
    st.markdown(table_html, unsafe_allow_html=True)


# ==========================================
# 화면 5: "엑셀 저장" (로그인 ID 전용 추출 + 셀 크기 내 자동 줄바꿈 + 원하는 폴더/파일명 저장)
# ==========================================
def render_excel_save_view():
    st.subheader(f"💾 엑셀 저장 (User: {st.session_state.user_id})")

    user_df = collector.load_user_logs(st.session_state.user_id, st.session_state.save_file_path)

    if user_df.empty:
        st.warning(f"'{st.session_state.user_id}'님의 저장할 영어회화 로그 데이터가 없습니다.")
        return

    st.markdown(
        f"**`{st.session_state.user_id}`** 님의 영어회화 로그 데이터(총 **{len(user_df)}** 행)를 "
        "`User | Topic | Date | Who | Conversation History | Error Correction` 포맷 및 **셀 내부 자동 줄바꿈(Shift+Enter 효과)** 서식을 적용하여 저장합니다."
    )

    col_folder, col_fname = st.columns([2, 1])
    default_folder = os.path.dirname(st.session_state.save_file_path) or os.getcwd()
    with col_folder:
        target_folder = st.text_input("저장할 폴더 경로 (Local PC / Server 경로):", value=default_folder)
    with col_fname:
        target_filename = st.text_input("저장할 엑셀 파일명:", value="Gemini_English.xlsx")

    if not target_filename.lower().endswith(".xlsx"):
        target_filename_xlsx = target_filename + ".xlsx"
    else:
        target_filename_xlsx = target_filename

    full_target_path = os.path.join(target_folder.strip(), target_filename_xlsx.strip())

    col_btn_local, col_btn_download = st.columns(2)

    with col_btn_local:
        if st.button("📁 지정한 폴더와 파일명으로 즉시 저장", type="primary", use_container_width=True):
            ok, msg = collector.save_user_logs_to_path(
                user_id=st.session_state.user_id,
                target_file_path=full_target_path,
                excel_path=st.session_state.save_file_path,
            )
            if ok:
                st.success(msg)
            else:
                st.error(msg)

    with col_btn_download:
        excel_bytes = collector.export_user_logs_to_bytes(
            user_id=st.session_state.user_id,
            excel_path=st.session_state.save_file_path,
        )
        st.download_button(
            label=f"⬇️ 브라우저로 '{target_filename_xlsx}' 다운로드 (원하는 폴더 선택)",
            data=excel_bytes,
            file_name=target_filename_xlsx,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    st.divider()
    st.markdown("##### 📋 저장될 데이터 미리보기 (`User` 열 포함)")
    st.dataframe(user_df, use_container_width=True, hide_index=True)


# ==========================================
# 화면 6: "모든 데이터 삭제" (삭제 확인 창 + 로그인 ID 데이터만 삭제)
# ==========================================
def render_delete_view():
    st.subheader(f"🗑️ 모든 데이터 삭제 (User: {st.session_state.user_id})")

    user_df = collector.load_user_logs(st.session_state.user_id, st.session_state.save_file_path)
    record_count = len(user_df)

    st.markdown(
        f"""
        <div class="msg-modal-box">
            ⚠️ <b>데이터 삭제 확인</b><br><br>
            현재까지 저장된 <b>'{html.escape(st.session_state.user_id)}'</b> 님의 영어회화 로그(총 <b>{record_count}</b>건)를 모두 삭제하시겠습니까?<br>
            삭제를 누르면 해당 ID에 해당하는 모든 영어회화 로그 데이터가 영구 삭제됩니다. (다른 회원의 데이터는 유지됩니다.)
        </div>
        """,
        unsafe_allow_html=True,
    )

    col_del, col_cancel, _ = st.columns([1, 1, 2])
    with col_del:
        if st.button("🚨 삭제", type="primary", use_container_width=True):
            ok, deleted_cnt, msg = collector.delete_user_logs(
                user_id=st.session_state.user_id,
                excel_path=st.session_state.save_file_path,
            )
            st.session_state.chat_history = []
            if ok:
                st.success(msg)
            else:
                st.error(msg)
    with col_cancel:
        if st.button("취소 (돌아가기)", use_container_width=True):
            st.session_state.current_menu = "Gemini 영어공부"
            st.rerun()


# ==========================================
# 메뉴 라우팅 실행
# ==========================================
current = st.session_state.current_menu

if current == "신규등록":
    render_register_view()
elif current == "Log In":
    render_login_view()
else:
    # 로그인 필수 메뉴 보호
    if not st.session_state.logged_in:
        st.markdown(
            '<div class="msg-modal-box">🔒 이 기능을 사용하려면 먼저 <b>Log In</b> 해 주세요. (계정이 없으신 경우 <b>신규등록</b>을 먼저 진행해 주세요.)</div>',
            unsafe_allow_html=True,
        )
        render_login_view()
    else:
        if current == "Gemini 영어공부":
            render_study_view()
        elif current == "영어공부 복습":
            render_review_view()
        elif current == "엑셀 저장":
            render_excel_save_view()
        elif current == "모든 데이터 삭제":
            render_delete_view()
