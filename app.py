import streamlit as st
import pandas as pd
import io
import zipfile
import json
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from datetime import datetime
import os

# ドキュメント生成用ライブラリ
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import mm
from docx import Document

# ---------------------------------------------------------
# 1. ユーティリティ関数（フォント設定など）
# ---------------------------------------------------------
FONT_FILE = "ipaexg.ttf"
FONT_NAME = "IPAexGothic"

def register_font():
    """PDF生成用の日本語フォントを登録する"""
    if os.path.exists(FONT_FILE):
        pdfmetrics.registerFont(TTFont(FONT_NAME, FONT_FILE))
        return True
    else:
        return False

# ---------------------------------------------------------
# 2. メール送信機能
# ---------------------------------------------------------
def send_email_with_attachment(zip_buffer, zip_filename, contest_name):
    """
    作成したZIPファイルを添付して、指定されたアドレス（スタッフ）にメールを送信する
    StreamlitのSecretsから設定を読み込む
    """
    try:
        # Secretsから設定を取得
        smtp_server = st.secrets["email"]["smtp_server"]
        smtp_port = st.secrets["email"]["smtp_port"]
        sender_email = st.secrets["email"]["sender_email"]
        sender_password = st.secrets["email"]["sender_password"]
        receiver_email = "info@beethoven-asia.com" # 送信先（スタッフ共有用）

        # メールの作成
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = receiver_email
        msg['Subject'] = f"【自動送信】資料出力: {contest_name}"

        body = f"""
        お疲れ様です。
        
        コンクール運営アプリより、以下の資料が出力されました。
        ZIPファイルを添付します。
        
        ・コンクール名: {contest_name}
        ・出力日時: {datetime.now().strftime('%Y/%m/%d %H:%M')}
        
        ※このメールは自動送信されています。
        """
        msg.attach(MIMEText(body, 'plain'))

        # 添付ファイルの設定
        part = MIMEApplication(zip_buffer.getvalue(), Name=zip_filename)
        part['Content-Disposition'] = f'attachment; filename="{zip_filename}"'
        msg.attach(part)

        # SMTPサーバーへの接続と送信
        if smtp_port == 465:
            with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
                server.login(sender_email, sender_password)
                server.send_message(msg)
        else:
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(sender_email, sender_password)
                server.send_message(msg)
        
        return True, "メール送信成功"
    
    except Exception as e:
        return False, f"メール送信エラー: {str(e)}"

# ---------------------------------------------------------
# 3. ドキュメント生成関数群
# ---------------------------------------------------------

def create_schedule_pdf(data, output_buffer, title):
    """受付表（スケジュール表）のPDFを作成"""
    c = canvas.Canvas(output_buffer, pagesize=A4)
    width, height = A4
    
    # フォント登録確認
    if register_font():
        c.setFont(FONT_NAME, 10)
    
    y = height - 20*mm
    c.setFont(FONT_NAME, 16) if register_font() else None
    c.drawString(20*mm, y, f"受付表: {title}")
    y -= 15*mm
    
    c.setFont(FONT_NAME, 10) if register_font() else None
    # ヘッダー
    c.drawString(20*mm, y, "番号")
    c.drawString(40*mm, y, "氏名")
    c.drawString(90*mm, y, "部門")
    c.drawString(130*mm, y, "演奏曲目")
    y -= 5*mm
    c.line(20*mm, y, 190*mm, y)
    y -= 5*mm
    
    for item in data:
        if y < 20*mm: # 改ページ
            c.showPage()
            y = height - 20*mm
            c.setFont(FONT_NAME, 10) if register_font() else None
        
        c.drawString(20*mm, y, str(item.get('no', '')))
        c.drawString(40*mm, y, str(item.get('name', '')))
        c.drawString(90*mm, y, str(item.get('category', '')))
        # 曲目は長いので省略などの処理が必要だが簡易的に表示
        song = str(item.get('song', ''))[:20]
        c.drawString(130*mm, y, song)
        y -= 8*mm
        
    c.save()

def create_score_sheet_pdf(data, output_buffer, judge_name, title):
    """採点表のPDFを作成（審査員ごとに発行）"""
    c = canvas.Canvas(output_buffer, pagesize=A4)
    width, height = A4
    if register_font():
        c.setFont(FONT_NAME, 10)
        
    y = height - 20*mm
    
    # タイトルと審査員名
    c.setFont(FONT_NAME, 14) if register_font() else None
    c.drawString(20*mm, y, f"採点表: {title}")
    c.setFont(FONT_NAME, 12) if register_font() else None
    c.drawRightString(190*mm, y, f"審査員: {judge_name} 先生")
    y -= 15*mm
    
    # 表ヘッダー
    c.setFont(FONT_NAME, 10) if register_font() else None
    c.drawString(20*mm, y, "番号")
    c.drawString(35*mm, y, "氏名")
    c.drawString(80*mm, y, "曲目")
    c.drawString(140*mm, y, "点数・講評")
    y -= 5*mm
    c.line(20*mm, y, 190*mm, y)
    y -= 10*mm
    
    for item in data:
        if y < 40*mm:
            c.showPage()
            y = height - 20*mm
            c.setFont(FONT_NAME, 12) if register_font() else None
            c.drawRightString(190*mm, y, f"審査員: {judge_name} 先生")
            y -= 15*mm
            c.setFont(FONT_NAME, 10) if register_font() else None
            
        c.drawString(20*mm, y, str(item.get('no', '')))
        c.drawString(35*mm, y, str(item.get('name', '')))
        song = str(item.get('song', ''))[:15]
        c.drawString(80*mm, y, song)
        
        # 記入欄枠
        c.rect(140*mm, y - 15*mm, 50*mm, 20*mm)
        
        y -= 25*mm
        
    c.save()

def create_word_doc(data, title, doc_type="list"):
    """Wordファイルを作成（受付表、採点表、WP用など汎用）"""
    doc = Document()
    doc.add_heading(title, 0)
    
    if doc_type == "wp_schedule":
        doc.add_paragraph("WordPress用スケジュールデータ")
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = '時間'
        hdr_cells[1].text = '番号'
        hdr_cells[2].text = '氏名'
        hdr_cells[3].text = '曲目'
        
        for item in data:
            row_cells = table.add_row().cells
            row_cells[0].text = str(item.get('time_slot', ''))
            row_cells[1].text = str(item.get('no', ''))
            row_cells[2].text = str(item.get('name', ''))
            row_cells[3].text = str(item.get('song', ''))
            
    else:
        # 汎用リスト（受付表など）
        table = doc.add_table(rows=1, cols=3)
        table.style = 'Table Grid'
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = '番号'
        hdr_cells[1].text = '氏名'
        hdr_cells[2].text = '部門'

        for item in data:
            row_cells = table.add_row().cells
            row_cells[0].text = str(item.get('no', ''))
            row_cells[1].text = str(item.get('name', ''))
            row_cells[2].text = str(item.get('category', ''))

    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer

def create_summary_pdf(data, judge_list, output_buffer, title):
    """集計表PDF（審査員全員の列を作る）"""
    c = canvas.Canvas(output_buffer, pagesize=A4)
    width, height = A4
    if register_font():
        c.setFont(FONT_NAME, 8)
        
    y = height - 20*mm
    c.setFont(FONT_NAME, 14) if register_font() else None
    c.drawString(20*mm, y, f"集計表: {title}")
    y -= 15*mm
    
    # ヘッダー
    c.setFont(FONT_NAME, 8) if register_font() else None
    c.drawString(10*mm, y, "番号")
    c.drawString(20*mm, y, "氏名")
    
    # 審査員列
    x = 60*mm
    col_width = 20*mm
    for j_name in judge_list:
        c.drawString(x, y, j_name[:4]) # 長いと重なるのでカット
        x += col_width
    c.drawString(x, y, "合計")
    
    y -= 5*mm
    c.line(10*mm, y, width - 10*mm, y)
    y -= 5*mm
    
    for item in data:
        if y < 15*mm:
            c.showPage()
            y = height - 20*mm
            c.setFont(FONT_NAME, 8) if register_font() else None
            
        c.drawString(10*mm, y, str(item.get('no', '')))
        c.drawString(20*mm, y, str(item.get('name', '')))
        
        # 枠線だけ描画（点数書き込み用）
        cur_x = 60*mm
        for _ in judge_list:
            c.rect(cur_x-2*mm, y-2*mm, 15*mm, 6*mm, fill=0)
            cur_x += col_width
        
        y -= 8*mm
        
    c.save()


# ---------------------------------------------------------
# 4. メインアプリケーションUI
# ---------------------------------------------------------
def main():
    st.title("🎹 コンクール運営資料作成 & スケジュール管理")
    
    # --- サイドバー: 設定読み込み/保存 ---
    with st.sidebar:
        st.header("⚙️ 設定管理")
        uploaded_config = st.file_uploader("設定ファイル(JSON)を読み込む", type=['json'])
        if uploaded_config:
            config_data = json.load(uploaded_config)
            st.session_state.update(config_data)
            st.success("設定を復元しました")

    # --- 1. Excelアップロードとシート選択 ---
    st.header("1. 名簿データのアップロード")
    uploaded_file = st.file_uploader("ExcelまたはCSVファイル", type=['xlsx', 'xls', 'csv'])
    
    if uploaded_file:
        try:
            # CSVかExcelかで処理を分ける
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                # ExcelFileとして読み込んでシート名を取得
                xls = pd.ExcelFile(uploaded_file)
                sheet_names = xls.sheet_names
                
                # シート選択ボックス
                selected_sheet = st.selectbox("読み込むシートを選択してください", sheet_names)
                
                # 選択されたシートをDataFrameとして読み込む
                df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)

            st.write("データプレビュー:", df.head(3))
            
            # --- 2. 列の割り当て ---
            st.header("2. 列の割り当て")
            cols = df.columns.tolist()
            
            col1, col2, col3 = st.columns(3)
            with col1:
                col_no = st.selectbox("出場番号の列", cols, index=cols.index("出場番号") if "出場番号" in cols else 0)
                col_name = st.selectbox("氏名の列", cols, index=cols.index("氏名") if "氏名" in cols else 0)
            with col2:
                col_cat = st.selectbox("部門の列", cols, index=cols.index("出場部門") if "出場部門" in cols else 0)
                col_song = st.selectbox("演奏曲目の列", cols, index=cols.index("演奏曲目") if "演奏曲目" in cols else 0)
            with col3:
                col_time = st.selectbox("演奏時間の列", cols, index=cols.index("演奏時間") if "演奏時間" in cols else 0)

            # データを統一フォーマットに変換
            processed_data = []
            for index, row in df.iterrows():
                processed_data.append({
                    'no': row[col_no],
                    'name': row[col_name],
                    'category': row[col_cat],
                    'song': row[col_song],
                    'time_str': row[col_time]
                })
            
            # --- 3. スケジュール・グループ設定 ---
            st.header("3. 進行スケジュール設定")
            
            if 'groups' not in st.session_state:
                st.session_state['groups'] = [{'start_no': '', 'end_no': '', 'start_time': '10:00', 'end_time': '11:00'}]
            
            # グループ追加ボタン
            if st.button("＋ グループを追加"):
                st.session_state['groups'].append({'start_no': '', 'end_no': '', 'start_time': '', 'end_time': ''})
            
            groups_config = []
            for i, grp in enumerate(st.session_state['groups']):
                with st.expander(f"グループ {i+1}", expanded=True):
                    c1, c2, c3, c4 = st.columns(4)
                    grp['start_no'] = c1.text_input(f"開始番号 (G{i+1})", grp['start_no'], key=f"s_no_{i}")
                    grp['end_no'] = c2.text_input(f"終了番号 (G{i+1})", grp['end_no'], key=f"e_no_{i}")
                    grp['start_time'] = c3.text_input(f"開始時刻 (G{i+1})", grp['start_time'], key=f"s_time_{i}")
                    grp['end_time'] = c4.text_input(f"終了時刻 (G{i+1})", grp['end_time'], key=f"e_time_{i}")
                    groups_config.append(grp)
            
            # --- 4. 大会情報入力 ---
            st.header("4. 大会情報入力 (WP用・ファイル名用)")
            contest_name = st.text_input("コンクール名 (ファイル名に使用)", "第10回BIPCA 東京予選④")
            open_time = st.text_input("開場時刻", "09:30")
            reception_time = st.text_input("受付時刻", "09:30")
            result_announce = st.text_input("審査結果発表日時", "当日 Webにて")

            # --- 5. 審査員設定 ---
            st.header("5. 審査員登録")
            if 'judges' not in st.session_state:
                st.session_state['judges'] = ["審査員A"]
            
            if st.button("＋ 審査員を追加"):
                st.session_state['judges'].append("")
            
            updated_judges = []
            for i, judge in enumerate(st.session_state['judges']):
                val = st.text_input(f"審査員 {i+1} 氏名", judge, key=f"judge_{i}")
                updated_judges.append(val)
            st.session_state['judges'] = updated_judges

            # --- 6. 出力・プレビュー ---
            st.header("6. ファイル出力とメール送信")
            
            # 設定保存用データの作成
            config_export = {
                'contest_name': contest_name,
                'groups': groups_config,
                'judges': updated_judges
            }
            config_json = json.dumps(config_export, ensure_ascii=False, indent=2)

            if st.button("全ファイル生成 & メール送信"):
                # ZIPファイルの作成
                zip_buffer = io.BytesIO()
                
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                    
                    # 1. 受付表 PDF
                    pdf_buf = io.BytesIO()
                    create_schedule_pdf(processed_data, pdf_buf, contest_name)
                    zip_file.writestr("受付表.pdf", pdf_buf.getvalue())
                    
                    # 2. 受付表 Word
                    word_buf = create_word_doc(processed_data, f"受付表: {contest_name}")
                    zip_file.writestr("受付表.docx", word_buf.getvalue())
                    
                    # 3. 採点表 (審査員分)
                    for j_name in updated_judges:
                        if j_name: # 空欄でなければ
                            score_buf = io.BytesIO()
                            create_score_sheet_pdf(processed_data, score_buf, j_name, contest_name)
                            zip_file.writestr(f"採点表_{j_name}.pdf", score_buf.getvalue())
                    
                    # 4. WP用 Word
                    wp_data = processed_data # 実際はスケジュールでフィルタリングしたデータを使う
                    wp_buf = create_word_doc(wp_data, contest_name, doc_type="wp_schedule")
                    zip_file.writestr("HP公開用スケジュール.docx", wp_buf.getvalue())
                    
                    # 5. 集計表 PDF
                    summary_buf = io.BytesIO()
                    create_summary_pdf(processed_data, updated_judges, summary_buf, contest_name)
                    zip_file.writestr("集計表.pdf", summary_buf.getvalue())
                    
                    # 6. 設定ファイル
                    zip_file.writestr("設定データ.json", config_json)

                # メール送信処理
                is_sent, mail_msg = send_email_with_attachment(zip_buffer, f"{contest_name}.zip", contest_name)
                
                if is_sent:
                    st.success(f"メール送信完了: {mail_msg}")
                else:
                    st.error(mail_msg)
                
                # ダウンロードボタンの表示
                st.download_button(
                    label="ZIPファイルをダウンロード",
                    data=zip_buffer.getvalue(),
                    file_name=f"{contest_name}.zip",
                    mime="application/zip"
                )

        except Exception as e:
            st.error(f"エラーが発生しました: {e}")

if __name__ == "__main__":
    main()
