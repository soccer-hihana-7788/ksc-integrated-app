import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta
import time
import os
import io
import base64
from PIL import Image
from streamlit_drawable_canvas import st_canvas

# --- ページ設定 (アプリ全体で1回のみ) ---
st.set_page_config(page_title="KSC 総合管理ポータル", layout="wide")

# --- 1. 共通認証の設定 (セッション管理) ---
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    st.title("🛡️ 総合ログインシステム")
    user_id = st.text_input("ログインID")
    password = st.text_input("パスワード", type="password")
    
    if st.button("ログイン"):
        try:
            correct_user = st.secrets["auth"]["user_id"]
            correct_password = st.secrets["auth"]["password"]
        except Exception:
            correct_user = "KSC"
            correct_password = "kuma2019"

        if (user_id == correct_user and password == correct_password) or (user_id == "KSC" and password == "kuma2019"):
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("IDまたはパスワードが正しくありません。")
    st.stop()

# --- 2. Google Sheets 共通クライアント設定 ---
@st.cache_resource
def get_gspread_client():
    scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    if "gcp_service_account" in st.secrets:
        creds_dict = dict(st.secrets["gcp_service_account"])
        if "private_key" in creds_dict:
            # \n を実際の改行に置き換えてPEMパディングエラーを防ぐ
            creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        credentials = Credentials.from_service_account_info(creds_dict, scopes=scope)
    else:
        key_file = "ksc-cash-app-7b96a6f1774a.json" if os.path.exists("ksc-cash-app-7b96a6f1774a.json") else "credentials.json"
        credentials = Credentials.from_service_account_file(key_file, scopes=scope)
    return gspread.authorize(credentials)

try:
    gc = get_gspread_client()
except Exception as e:
    st.error(f"Googleスプレッドシート認証エラー: {e}")
    st.stop()

# --- 3. サイドバー：共通メニュー・ログアウト ---
st.sidebar.title("📌 メニュー")
app_mode = st.sidebar.selectbox(
    "利用するアプリを選択", 
    ["ホーム（アプリ選択）", "⚽ サッカーチーム背番号管理ツール", "💰 KSC経費申請管理ツール"]
)

st.sidebar.markdown("---")
if st.sidebar.button("ログアウト"):
    st.session_state["authenticated"] = False
    st.session_state.pop("current_user", None)
    st.rerun()


# ==========================================
# ホーム画面（アプリ選択）
# ==========================================
if app_mode == "ホーム（アプリ選択）":
    st.title("⚽ KSC 総合管理ポータルへようこそ")
    st.markdown("利用したいツールを下のボタン、またはサイドバーから選択してください。")
    
    col1, col2 = st.columns(2)
    with col1:
        st.info("### ⚽ サッカーチーム背番号管理")
        st.write("選手の背番号、発注状況、納品、受け渡し状況を一元管理します。")
        if st.button("背番号管理ツールを開く"):
            st.session_state["app_select"] = "⚽ サッカーチーム背番号管理ツール"
            st.rerun()
            
    with col2:
        st.info("### 💰 KSC経費申請管理")
        st.write("交通費清算書や日当清算書（受領書）の申請、PDF印刷を行います。")
        if st.button("経費申請管理ツールを開く"):
            st.session_state["app_select"] = "💰 KSC経費申請管理ツール"
            st.rerun()

    if "app_select" in st.session_state:
        app_mode = st.session_state["app_select"]


# ==========================================
# アプリ①：サッカーチーム背番号管理ツール
# ==========================================
if app_mode == "⚽ サッカーチーム背番号管理ツール":
    try:
        sh1 = gc.open_by_key("1yOFFntvApgvi-h6Z2bYFZbQevKNqAT82cfEitjD9Q14")
    except Exception as e:
        st.error(f"背番号管理用スプレッドシート接続エラー: {e}")
        st.stop()

    st.title("⚽ サッカーチーム 背番号管理システム")
    st.markdown("---")

    def init_worksheet_if_empty(ws, start_num, end_num):
        data = ws.get_all_records()
        if not data:
            rows = []
            for num in range(start_num, end_num + 1):
                rows.append([num, "", "", "", ""])
            ws.append_rows(rows)

    def load_sheet_data(sheet_name, start_num, end_num):
        try:
            ws = sh1.worksheet(sheet_name)
        except Exception:
            ws = sh1.add_worksheet(title=sheet_name, rows=200, cols=10)
            ws.append_row(["背番号", "名前", "発注日時", "納品", "受け渡し日時"])

        init_worksheet_if_empty(ws, start_num, end_num)
        records = ws.get_all_records()
        df = pd.DataFrame(records)
        
        if df.empty:
            return ws, df

        for col in ["発注日時", "受け渡し日時"]:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col].replace(r'^\s*$', None, regex=True), errors='coerce').dt.date

        if "納品" in df.columns:
            df["納品"] = df["納品"].astype(str).replace({"✕": None, "None": None, "": None, "nan": None})

        df['is_handed'] = df['受け渡し日時'].apply(lambda x: True if pd.notna(x) and str(x).strip() != "" else False)
        return ws, df

    tab_active, tab1, tab2 = st.tabs(["🔥 利用中", "👕 1 〜 100番 管理", "👕 101 〜 200番 管理"])

    with tab_active:
        st.subheader("🔥 利用中 背番号一覧")
        _, df1 = load_sheet_data("1_100", 1, 100)
        _, df2 = load_sheet_data("101_200", 101, 200)
        all_df = pd.concat([df1, df2], ignore_index=True)
        
        if not all_df.empty:
            active_df = all_df[all_df["名前"].astype(str).str.strip() != ""].copy()
            search_active = st.text_input("🔍 利用中の背番号または名前で検索", key="search_active", placeholder="例: 10 や 山田 など")
            if search_active.strip():
                term = search_active.strip().lower()
                active_df = active_df[
                    active_df['背番号'].astype(str).str.contains(term, case=False, na=False) |
                    active_df['名前'].astype(str).str.contains(term, case=False, na=False)
                ]
            
            if active_df.empty:
                st.info("現在利用中の背番号はありません。")
            else:
                st.data_editor(
                    active_df[["背番号", "名前", "発注日時", "納品", "受け渡し日時", "is_handed"]],
                    column_config={
                        "背番号": st.column_config.NumberColumn("背番号", disabled=True),
                        "名前": st.column_config.TextColumn("名前 (選手名)", disabled=True),
                        "発注日時": st.column_config.DateColumn("発注日時 (YYYY-MM-DD)", format="YYYY-MM-DD", disabled=True),
                        "納品": st.column_config.SelectboxColumn("納品", options=[None, "〇", "✕"], disabled=True),
                        "受け渡し日時": st.column_config.DateColumn("受け渡し日時 (YYYY-MM-DD)", format="YYYY-MM-DD", disabled=True),
                        "is_handed": st.column_config.CheckboxColumn("受け渡し済ロック", disabled=True)
                    },
                    disabled=True, hide_index=True, use_container_width=True, key="editor_active_view"
                )

    def render_uniform_manager(sheet_name, start_num, end_num):
        st.subheader(f"背番号 {start_num} 〜 {end_num} 一覧管理")
        ws, df = load_sheet_data(sheet_name, start_num, end_num)
        if df.empty:
            st.warning("データが取得できませんでした。")
            return

        search_term = st.text_input(f"🔍 背番号または名前で検索 ({sheet_name})", key=f"search_{sheet_name}", placeholder="例: 10 や 山田 など")
        display_df = df.copy()
        if search_term.strip():
            term = search_term.strip().lower()
            display_df = display_df[
                display_df['背番号'].astype(str).str.contains(term, case=False, na=False) |
                display_df['名前'].astype(str).str.contains(term, case=False, na=False)
            ]

        edited_df = st.data_editor(
            display_df[["背番号", "名前", "発注日時", "納品", "受け渡し日時", "is_handed"]],
            column_config={
                "背番号": st.column_config.NumberColumn("背番号", disabled=True),
                "名前": st.column_config.TextColumn("名前 (選手名)"),
                "発注日時": st.column_config.DateColumn("発注日時 (YYYY-MM-DD)", format="YYYY-MM-DD"),
                "納品": st.column_config.SelectboxColumn("納品", options=[None, "〇", "✕"], required=False),
                "受け渡し日時": st.column_config.DateColumn("受け渡し日時 (YYYY-MM-DD)", format="YYYY-MM-DD"),
                "is_handed": st.column_config.CheckboxColumn("受け渡し済ロック", disabled=True)
            },
            disabled=["背番号", "is_handed"], hide_index=True, use_container_width=True, key=f"editor_{sheet_name}"
        )

        if st.button("💾 変更をスプレッドシートに保存", key=f"save_{sheet_name}"):
            try:
                updated_full_df = df.set_index("背番号")
                edited_set = edited_df.set_index("背番号")
                for b_num, row in edited_set.iterrows():
                    if b_num in updated_full_df.index:
                        updated_full_df.loc[b_num, "名前"] = row["名前"]
                        updated_full_df.loc[b_num, "発注日時"] = row["発注日時"]
                        updated_full_df.loc[b_num, "納品"] = row["納品"]
                        updated_full_df.loc[b_num, "受け渡し日時"] = row["受け渡し日時"]
                updated_full_df = updated_full_df.reset_index()

                update_rows = []
                for _, row in updated_full_df.iterrows():
                    b_num = row["背番号"]
                    name = str(row["名前"]) if pd.notna(row["名前"]) else ""
                    ord_date = row["発注日時"]
                    ord_str = ord_date.strftime("%Y-%m-%d") if pd.notna(ord_date) and str(ord_date).strip() != "" else ""
                    delivery = str(row["納品"]) if pd.notna(row["納品"]) and row["納品"] is not None and str(row["納品"]) != "None" else ""
                    hand_date = row["受け渡し日時"]
                    hand_str = hand_date.strftime("%Y-%m-%d") if pd.notna(hand_date) and str(hand_date).strip() != "" else ""
                    update_rows.append([int(b_num), name, ord_str, delivery, hand_str])

                ws.update(f"A2:E{len(update_rows)+1}", update_rows)
                st.cache_data.clear()
                st.success("スプレッドシートのデータを正常に更新しました！")
                time.sleep(1)
                st.rerun()
            except Exception as e:
                st.error(f"保存に失敗しました: {e}")

    with tab1:
        render_uniform_manager("1_100", 1, 100)
    with tab2:
        render_uniform_manager("101_200", 101, 200)


# ==========================================
# アプリ②：KSC経費申請管理ツール
# ==========================================
elif app_mode == "💰 KSC経費申請管理ツール":
    try:
        sh2 = gc.open_by_key("1yVYajQm6KeaoppB3KMaHismS-95NWxGGfie9DhDEEgk")
    except Exception as e:
        st.error(f"経費申請用スプレッドシート接続エラー: {e}")
        st.stop()

    if "current_user" not in st.session_state or not st.session_state["current_user"]:
        st.title("👤 氏名選択")
        name_list = ["加藤", "瀬野", "平", "稲垣", "富井", "新美", "細川", "島田"]
        selected_user = st.selectbox("あなたの氏名を選択してください", name_list)
        if st.button("決定"):
            st.session_state["current_user"] = selected_user
            st.rerun()
        st.stop()

    st.sidebar.write(f"ログイン中の氏名: **{st.session_state['current_user']}**")
    if st.sidebar.button("ユーザー変更"):
        st.session_state["current_user"] = None
        st.rerun()

    st.title("KSC経費申請管理ツール")
    form_type = st.radio("申請書種別を選択してください", ["KSC 交通費清算書", "KSC 日当清算書 兼 受領書"], horizontal=True)
    st.markdown("---")

    if form_type == "KSC 交通費清算書":
        st.header("🚗 KSC 交通費清算書")
        with st.form("transport_form", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1:
                date = st.date_input("日付", datetime.now())
                dest = st.text_input("行先")
            with col2:
                purp = st.text_input("目的")
                amt = st.number_input("金額 (円)", min_value=0, step=1)
            rem = st.text_area("備考")
            if st.form_submit_button("スプレッドシートに保存"):
                try:
                    ws = sh2.worksheet("transport_log")
                    ws.append_row([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), st.session_state["current_user"], str(date), dest, purp, int(amt), rem])
                    st.cache_data.clear()
                    st.success("データを保存しました。")
                    time.sleep(1)
                    st.rerun()
                except Exception as e: 
                    st.error(f"保存失敗: {e}")
    else:
        st.header("📋 KSC 日当清算書 兼 受領書")
        if "form_key_suffix" not in st.session_state:
            st.session_state["form_key_suffix"] = time.time()

        with st.form("allowance_form", clear_on_submit=True):
            dt = st.date_input("日時", datetime.now())
            cont = st.text_area("臨時コーチ依頼内容")
            amt = st.number_input("金額 (円)", min_value=0, step=1)
            c_c = st.checkbox("確認 (コーチ)")
            c_t = st.checkbox("確認 (臨時コーチ)")
            
            st.write("🖋️ **臨時コーチ署名**")
            sig_name = st.text_input("確認(臨時コーチ氏名)を入力してください")
            canvas_number = st.session_state['form_key_suffix']
            canvas_result = st_canvas(
                fill_color="rgba(255, 255, 255, 1)", stroke_width=2, stroke_color="#000",
                background_color="#fff", height=120, width=300, drawing_mode="freedraw", 
                key=f"canvas_{canvas_number}"
            )
            
            if st.form_submit_button("スプレッドシートに保存"):
                try:
                    sig_b64 = ""
                    if c_t and canvas_result.image_data is not None:
                        img = Image.fromarray(canvas_result.image_data.astype('uint8'), 'RGBA')
                        buffered = io.BytesIO()
                        img.save(buffered, format="PNG")
                        sig_b64 = base64.b64encode(buffered.getvalue()).decode()

                    ws = sh2.worksheet("allowance_log")
                    ws.append_row([
                        datetime.now().strftime("%Y-%m-%d %H:%M:%S"), st.session_state["current_user"], 
                        str(dt), cont, int(amt), "済" if c_c else "未", sig_name, sig_b64
                    ])
                    st.cache_data.clear()
                    st.session_state["form_key_suffix"] = time.time()
                    st.success("データを保存しました。")
                    time.sleep(1)
                    st.rerun()
                except Exception as e: 
                    st.error(f"保存失敗: {e}")

    @st.cache_data(ttl=60)
    def fetch_records(sheet_name):
        ws = sh2.worksheet(sheet_name)
        return ws.get_all_records()

    st.markdown("---")
    st.subheader(f"📊 {st.session_state['current_user']} 様の申請済み一覧")

    try:
        current_ws_name = "transport_log" if form_type == "KSC 交通費清算書" else "allowance_log"
        all_records = fetch_records(current_ws_name)
        
        if all_records:
            df_all = pd.DataFrame(all_records)
            df_all['row_idx'] = range(2, len(df_all) + 2)
            df_user = df_all[df_all["氏名"] == st.session_state["current_user"]].copy()
            
            if not df_user.empty:
                date_col = '日付' if form_type == "KSC 交通費清算書" else '日時'
                df_user[date_col] = pd.to_datetime(df_user[date_col])
                
                col_btn1, col_btn2, _ = st.columns([1.5, 1.5, 7])
                if "show_list" not in st.session_state:
                    st.session_state["show_list"] = True
                if col_btn1.button("一覧を表示"):
                    st.session_state["show_list"] = True
                if col_btn2.button("一覧を非表示"):
                    st.session_state["show_list"] = False

                col_d1, col_d2 = st.columns(2)
                with col_d1: start_date = st.date_input("開始日", df_user[date_col].min().date())
                with col_d2: end_date = st.date_input("終了日", df_user[date_col].max().date())
                
                mask = (df_user[date_col].dt.date >= start_date) & (df_user[date_col].dt.date <= end_date)
                df_filtered = df_user.loc[mask].sort_values(by=date_col, ascending=False).reset_index(drop=True)
                
                if st.session_state["show_list"]:
                    df_with_select = df_filtered.copy()
                    df_with_select.insert(0, "選択", False)
                    cols_to_show = df_with_select.columns.tolist()
                    if 'row_idx' in cols_to_show: cols_to_show.remove('row_idx')
                    if '確認(臨時コーチ署名)' in cols_to_show: cols_to_show.remove('確認(臨時コーチ署名)')
                    df_with_select[date_col] = df_with_select[date_col].dt.strftime('%Y-%m-%d')
                    
                    edited_df = st.data_editor(
                        df_with_select[cols_to_show],
                        hide_index=True, use_container_width=True,
                        column_config={"選択": st.column_config.CheckboxColumn(required=True)},
                        disabled=[c for c in cols_to_show if c != "選択"],
                        key=f"editor_{current_ws_name}"
                    )
                    selected_indices = edited_df.index[edited_df["選択"]].tolist()
                    df_to_edit = df_filtered.iloc[selected_indices]
                else:
                    st.info("一覧は非表示になっています。")
                    df_to_edit = pd.DataFrame()
            else:
                st.info("データがありません。")
    except Exception as e:
        st.error(f"エラー: {e}")
