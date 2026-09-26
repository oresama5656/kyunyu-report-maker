import streamlit as st
import datetime
from docxtpl import DocxTemplate
import io
import os
import json
import random

# ページ基本設定
st.set_page_config(
    page_title="吸入指導トレーシングレポート",
    page_icon="📄",
    layout="centered"
)

# 極限まで無駄を削ぎ落とした洗練されたモダンUIスタイル
st.markdown("""
<style>
    /* グローバル設定 */
    .stApp {
        background-color: #F8FAFC;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", "Hiragino Sans", "Meiryo", sans-serif;
        color: #0F172A;
    }

    /* メインコンテナの極限圧縮 */
    .main .block-container {
        max-width: 580px;
        padding-top: 0.6rem !important;
        padding-bottom: 1rem !important;
    }

    /* ヘッダーエリアのコンパクト化 */
    .header-box {
        margin-bottom: 0.6rem;
        display: flex;
        align-items: baseline;
        justify-content: space-between;
        border-bottom: 1px solid #E2E8F0;
        padding-bottom: 0.4rem;
    }
    .header-title {
        font-size: 1.15rem;
        font-weight: 700;
        letter-spacing: -0.01em;
        color: #0F172A;
    }
    .header-desc {
        font-size: 0.75rem;
        color: #64748B;
    }

    /* ラベルの小型化と余白削減 */
    .stTextInput > label, .stTextArea > label, .stRadio > label {
        font-size: 0.78rem !important;
        font-weight: 600 !important;
        color: #475569 !important;
        margin-bottom: 0.1rem !important;
    }

    /* 入力ウィジェットの垂直マージン圧縮 */
    div[data-testid="stTextInput"], div[data-testid="stTextArea"], div[data-testid="stRadio"] {
        margin-bottom: -0.35rem !important;
    }

    /* 入力ボックスのスリム化 */
    input[type="text"] {
        height: 34px !important;
        padding: 4px 10px !important;
        font-size: 0.85rem !important;
        border-radius: 6px !important;
        background-color: #FFFFFF !important;
        border: 1px solid #CBD5E1 !important;
    }
    input[type="text"]:focus, textarea:focus {
        border-color: #0F172A !important;
        box-shadow: 0 0 0 2px rgba(15, 23, 42, 0.08) !important;
    }

    /* テキストエリアのスリム化 */
    textarea {
        padding: 6px 10px !important;
        font-size: 0.82rem !important;
        border-radius: 6px !important;
        background-color: #FFFFFF !important;
        border: 1px solid #CBD5E1 !important;
        line-height: 1.4 !important;
    }

    /* ラジオボタンの圧縮 */
    .stRadio > div {
        gap: 1.2rem !important;
        padding-top: 0rem !important;
    }
    .stRadio div[role="radiogroup"] label {
        font-size: 0.8rem !important;
    }

    /* 2列2行ボタングリッドの小型化 */
    div.stButton > button {
        width: 100% !important;
        height: 30px !important;
        padding: 2px 8px !important;
        font-size: 0.76rem !important;
        font-weight: 500 !important;
        border-radius: 6px !important;
        background-color: #FFFFFF !important;
        border: 1px solid #E2E8F0 !important;
        color: #334155 !important;
        box-shadow: 0 1px 2px rgba(0, 0, 0, 0.02) !important;
        transition: all 0.15s ease !important;
    }
    div.stButton > button:hover {
        background-color: #F1F5F9 !important;
        border-color: #CBD5E1 !important;
        color: #0F172A !important;
    }

    /* ダウンロードボタンのスマート化 */
    .stDownloadButton > button {
        width: 100% !important;
        height: 40px !important;
        background-color: #0F172A !important;
        color: #FFFFFF !important;
        font-weight: 600 !important;
        font-size: 0.88rem !important;
        border-radius: 6px !important;
        border: none !important;
        box-shadow: 0 2px 8px rgba(15, 23, 42, 0.12) !important;
        margin-top: 0.6rem !important;
        transition: all 0.15s ease !important;
    }
    .stDownloadButton > button:hover {
        background-color: #1E293B !important;
        transform: translateY(-1px) !important;
    }
</style>
""", unsafe_allow_html=True)

TEMPLATE_PATH = "template_kyunyu_report.docx"
JSON_PATH = "templates.json"

if not os.path.exists(TEMPLATE_PATH):
    from create_template import generate_template
    generate_template()

@st.cache_data
def load_templates():
    if os.path.exists(JSON_PATH):
        with open(JSON_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

template_data = load_templates()

# --- 薬局情報の永続化（URLパラメータ方式：セキュリティ制限を受けず100%確実に動作） ---
params = st.query_params

default_pharmacy_name = params.get("p_name", "アビー薬局")
default_pharmacy_address = params.get("p_addr", "茨城県水戸市緑町1丁目2番3号")
default_pharmacy_tel = params.get("p_tel", "029-200-1234")
default_pharmacy_fax = params.get("p_fax", "029-200-1235")
default_pharmacist_name = params.get("p_phm", "薬師寺　花子")

def get_today_wareki():
    today = datetime.date.today()
    if today.year >= 2019:
        reiwa_year = today.year - 2018
        year_str = "令和元年" if reiwa_year == 1 else f"令和{reiwa_year}年"
    else:
        year_str = f"{today.year}年"
    return f"{year_str}{today.month}月{today.day}日"

DEFAULT_GUIDANCE = (
    "説明書および練習用吸入器を用いて吸入手技の確認・指導を実施しました。\n"
    "吸入手順（器具のセット、十分な息吐き、吸入速度、吸入後の息止め・うがい）を確認したところ、一連の動作を適切に実施できていました。\n"
    "実演にて手技を再確認し、正しく吸入できていることと良好な理解を確認いたしました。\n"
    "現時点で特記すべき問題は認めず、今後も経過を確認してまいります。"
)

if "input_guidance_detail" not in st.session_state:
    st.session_state["input_guidance_detail"] = DEFAULT_GUIDANCE

def set_random_template(category_key):
    items = template_data.get(category_key, [])
    if items:
        chosen_text = random.choice(items)
        st.session_state["input_guidance_detail"] = chosen_text

# ヘッダー
st.markdown("""
<div class="header-box">
    <div class="header-title">吸入指導トレーシングレポート作成</div>
    <div class="header-desc">Word形式（.docx）で出力</div>
</div>
""", unsafe_allow_html=True)

# サイドバー: 薬局設定
with st.sidebar:
    st.markdown("### 薬局情報設定")
    st.caption("ここで設定した情報がレポートに反映されます")
    
    pharmacy_name = st.text_input("薬局名称", value=default_pharmacy_name)
    pharmacy_address = st.text_input("薬局住所", value=default_pharmacy_address)
    pharmacy_tel = st.text_input("TEL", value=default_pharmacy_tel)
    pharmacy_fax = st.text_input("FAX", value=default_pharmacy_fax)
    pharmacist_name = st.text_input("担当薬剤師名", value=default_pharmacist_name)

    # 薬局情報をURLに保存・反映するボタン
    if st.button("💾 この薬局情報を反映してURLを発行", use_container_width=True):
        st.query_params["p_name"] = pharmacy_name
        st.query_params["p_addr"] = pharmacy_address
        st.query_params["p_tel"] = pharmacy_tel
        st.query_params["p_fax"] = pharmacy_fax
        st.query_params["p_phm"] = pharmacist_name
        st.rerun()

    # パラメータがセットされている場合はブックマーク案内を表示
    if "p_name" in params:
        st.success("✅ 自店情報がURLに反映されました！")
        st.info("💡 **重要：** 現在のアドレスバーのURLをブラウザの「お気に入り（ブックマーク）」に登録してください。次回からそのブックマークを開くだけで、常にこの薬局情報がセットされた状態で起動します！")

# --- メインフォーム ---

report_date = st.text_input("報告日", value=get_today_wareki())
doctor_name = st.text_input("処方医名（先生御机下）", value="相澤　一郎")
patient_name = st.text_input("患者名（様）", value="山田　太郎")
patient_dob = st.text_input("生年月日", value="平成29年3月14日")
order_no = st.text_input("オーダー番号（任意）", value="")

consent_option = st.radio(
    "患者からの同意",
    ["得た", "得ていない", "拒否（治療上重要なため報告）"],
    index=0,
    horizontal=True
)

target_drug = st.text_input("対象薬剤", value="イナビル吸入粉末剤２０ｍｇ")

# 2列2行ボタングリッド
st.markdown("<p style='font-size: 0.72rem; font-weight: 600; color: #64748B; margin-top: 0.5rem; margin-bottom: 0.25rem;'>文面生成（クリックで自動入力）</p>", unsafe_allow_html=True)

b1, b2 = st.columns(2)
with b1:
    st.button("本人吸入（手技良好）", use_container_width=True, on_click=set_random_template, args=("adult_self",))
with b2:
    st.button("小児・保護者指導", use_container_width=True, on_click=set_random_template, args=("pediatric_parent",))

b3, b4 = st.columns(2)
with b3:
    st.button("継続・リピート", use_container_width=True, on_click=set_random_template, args=("repeat_continue",))
with b4:
    st.button("手技修正・改善", use_container_width=True, on_click=set_random_template, args=("fix_improved",))

# 指導内容テキストエリア
guidance_detail = st.text_area(
    "指導内容および状況",
    height=100,
    key="input_guidance_detail"
)

# --- 出力処理 ---
consent_got_str = "☒" if consent_option == "得た" else "☐"
consent_not_got_str = "☒" if consent_option == "得ていない" else "☐"
consent_refused_str = "☒" if consent_option == "拒否（治療上重要なため報告）" else "☐"

def generate_docx():
    doc = DocxTemplate(TEMPLATE_PATH)
    context = {
        "doctor_name": doctor_name,
        "report_date": report_date,
        "patient_name": patient_name,
        "patient_dob": patient_dob,
        "order_no": order_no,
        "pharmacy_name": pharmacy_name,
        "pharmacy_address": pharmacy_address,
        "pharmacy_tel": pharmacy_tel,
        "pharmacy_fax": pharmacy_fax,
        "pharmacist_name": pharmacist_name,
        "consent_got": consent_got_str,
        "consent_not_got": consent_not_got_str,
        "consent_refused": consent_refused_str,
        "target_drug": target_drug,
        "guidance_detail": guidance_detail,
    }
    doc.render(context)
    
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

docx_bytes = generate_docx()
file_name = f"トレーシングレポート_吸入_{patient_name}.docx"

st.download_button(
    label="Wordレポートを出力する（.docx）",
    data=docx_bytes,
    file_name=file_name,
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    use_container_width=True
)
