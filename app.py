import streamlit as st
import pandas as pd
import io
import zipfile
import json
import re
import os
import copy
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from email.header import Header
from datetime import datetime, timedelta
from docx import Document
from docx.text.paragraph import Paragraph
from docx.shared import Pt

# ---------------------------------------------------------
# 1. ユーティリティ（時間変換・Word操作・データ解析）
# ---------------------------------------------------------

def parse_jp_time_to_seconds(time_str):
    if not time_str:
        return 0
    s = str(time_str)
    minutes = re.search(r'(\d+)\s*[分m]', s)
    seconds = re.search(r'(\d+)\s*[秒s]', s)
    
    total_sec = 0
    if minutes:
        total_sec += int(minutes.group(1)) * 60
    if seconds:
        total_sec += int(seconds.group(1))
    return total_sec

def format_seconds_to_jp_label(total_seconds):
    if total_seconds <= 0:
        return "0分"
    
    minutes = total_seconds // 60
    remainder_seconds = total_seconds % 60
    
    if remainder_seconds >= 30:
        minutes += 1
        
    h = minutes // 60
    m = minutes % 60
    
    if h > 0:
        return f"{h}時間{m}分"
    else:
        return f"{m}分"

def format_time_label(text):
    if not text:
        return ""
    matches = re.findall(r'(\d{1,2})[:：](\d{2})', str(text))
    if len(matches) >= 2:
        start_time = f"{matches[0][0]}時{matches[0][1]}分"
        end_time = f"{matches[1][0]}時{matches[1][1]}分"
        return f"{start_time}～{end_time}"
    else:
        return text

def format_single_time_label(text):
    if not text:
        return ""
    match = re.search(r'(\d{1,2})[:：](\d{2})', str(text))
    if match:
        return f"{match.group(1)}時{match.group(2)}分"
    return text

def calculate_next_day_morning(date_str):
    if not date_str:
        return ""
    match = re.search(r'(\d{4})[^\d](\d{1,2})[^\d](\d{1,2})', str(date_str))
    if match:
        try:
            year, month, day = map(int, match.groups())
            dt = datetime(year, month, day)
            next_day = dt + timedelta(days=1)
            return next_day.strftime(f"%Y年%m月%d日10時00分")
        except:
            return ""
    return ""

def resolve_participants_from_string(input_str, all_data_list):
    if not input_str:
        return []

    id_map = {str(item['no']): i for i, item in enumerate(all_data_list)}
    resolved_members = []
    
    parts = [p.strip() for p in input_str.replace('、', ',').split(',')]
    
    for part in parts:
        if not part:
            continue
        if '-' in part:
            range_parts = part.split('-')
            if len(range_parts) == 2:
                start_id = range_parts[0].strip()
                end_id = range_parts[1].strip()
                if start_id in id_map and end_id in id_map:
                    s_idx = id_map[start_id]
                    e_idx = id_map[end_id]
                    if s_idx > e_idx:
                        s_idx, e_idx = e_idx, s_idx
                    for i in range(s_idx, e_idx + 1):
                        resolved_members.append(all_data_list[i])
        else:
            if part in id_map:
                idx = id_map[part]
                resolved_members.append(all_data_list[idx])
    return resolved_members

# --- Word操作系 ---

def replace_text_smart(paragraph, replacements):
    """
    強力な置換関数。
    1. まずRunごとの単純置換を試みる（スタイル維持）。
    2. それで置換しきれない（タグが分割されている）場合、
       段落内のテキストを強制的に結合して置換する。
    """
    full_text = paragraph.text
    if not any(key in full_text for key in replacements):
        return

    # 1. 単純置換
    if paragraph.runs:
        for run in paragraph.runs:
            for key, val in replacements.items():
                if key in run.text:
                    run.text = run.text.replace(key, str(val))

    # 2. 残存チェックと強制置換
    full_text_new = paragraph.text
    remaining_keys = [k for k in replacements if k in full_text_new]

    if remaining_keys:
        current_text = full_text_new
        for k in remaining_keys:
            current_text = current_text.replace(k, str(replacements[k]))
        
        for run in paragraph.runs:
            run.text = ""
        
        if paragraph.runs:
            paragraph.runs[0].text = current_text
        else:
            paragraph.add_run(current_text)

def fill_row_data(row, data_dict):
    """行内の全セルの段落に対して置換を実行"""
    for cell in row.cells:
        for paragraph in cell.paragraphs:
            replace_text_smart(paragraph, data_dict)

def replace_text_in_document_full(doc, replacements):
    """
    ドキュメント全体（本文、表、ヘッダー、フッター）を対象に置換を行う。
    """
    # 1. 本文段落
    for paragraph in doc.paragraphs:
        replace_text_smart(paragraph, replacements)
    
    # 2. 本文の表
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    replace_text_smart(paragraph, replacements)
                    
    # 3. ヘッダー・フッター（全セクション）
    for section in doc.sections:
        # ヘッダー (通常, 1ページ目, 偶数ページ)
        for header in [section.header, section.first_page_header, section.even_page_header]:
            if header:
                for paragraph in header.paragraphs:
                    replace_text_smart(paragraph, replacements)
                for table in header.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for paragraph in cell.paragraphs:
                                replace_text_smart(paragraph, replacements)
        
        # フッター
        for footer in [section.footer, section.first_page_footer, section.even_page_footer]:
            if footer:
                for paragraph in footer.paragraphs:
                    replace_text_smart(paragraph, replacements)
                for table in footer.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for paragraph in cell.paragraphs:
                                replace_text_smart(paragraph, replacements)

# ---------------------------------------------------------
# 2. メール送信機能（SSL対応版・添付ファイル名修正・使用者情報挿入）
# ---------------------------------------------------------

def send_email_callback():
    """ZIPファイルダウンロード時にメールを送信するコールバック関数"""
    if 'zip_buffer' not in st.session_state or not st.session_state['zip_buffer']:
        return

    # Streamlit Secrets から設定を取得
    try:
        smtp_server = st.secrets["email"]["smtp_server"]
        smtp_port = st.secrets["email"]["smtp_port"]
        sender_email = st.secrets["email"]["sender_email"]
        password = st.secrets["email"]["sender_password"]
    except Exception:
        # シークレットキー名が異なる場合のフォールバック（smtp or email）
        try:
            smtp_server = st.secrets["smtp"]["server"]
            smtp_port = st.secrets["smtp"]["port"]
            sender_email = st.secrets["smtp"]["sender_email"]
            password = st.secrets["smtp"]["password"]
        except:
            return

    contest_name = st.session_state.get('contest_name', '無題')
    user_email = st.session_state.get('user_email', '不明なユーザー')
    
    # ZIP内のファイルリストを取得して本文を作成
    file_list_str = ""
    try:
        # 現在のバッファ位置を保存し、先頭に戻して読み込む
        current_pos = st.session_state['zip_buffer'].tell()
        st.session_state['zip_buffer'].seek(0)
        
        with zipfile.ZipFile(st.session_state['zip_buffer'], 'r') as zf_read:
            for name in zf_read.namelist():
                file_list_str += f"・{name}\n"
        
        # バッファ位置を戻す
        st.session_state['zip_buffer'].seek(current_pos)
    except Exception as e:
        file_list_str = f"（ファイル一覧取得エラー: {e}）"

    # 生成日時（日本時間 UTC+9）
    jst_now = datetime.utcnow() + timedelta(hours=9)
    timestamp = jst_now.strftime("%Y年%m月%d日%H時%M分")

    # 件名と本文の構築
    subject = f"採点表等を作成しました：{contest_name}"
    body = f"""{user_email}が以下のファイルを生成しました。

{file_list_str}
生成日時：{timestamp}"""
    
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = sender_email  # 自分自身に送信
    msg['Subject'] = Header(subject, 'utf-8') # 件名の文字化け防止
    msg.attach(MIMEText(body, 'plain'))

    # ZIP添付
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(st.session_state['zip_buffer'].getvalue())
    encoders.encode_base64(part)
    
    # ファイル名のエンコード処理 (noname回避)
    filename = f"{contest_name}.zip"
    encoded_filename = Header(filename, 'utf-8').encode()
    part.add_header('Content-Disposition', 'attachment', filename=encoded_filename)
    
    msg.attach(part)

    try:
        # ロリポップ等はポート465でSMTP_SSLを使用する
        server = smtplib.SMTP_SSL(smtp_server, smtp_port)
        server.login(sender_email, password)
        server.send_message(msg)
        server.quit()
        print("Backup email sent successfully.")
    except Exception as e:
        print(f"Failed to send email: {e}")

# ---------------------------------------------------------
# 3. ドキュメント生成ロジック
# ---------------------------------------------------------

def generate_word_from_template(template_path_or_file, groups, all_data, global_context):
    """
    採点表・受付表用 (従来のスマート置換を使用)
    """
    doc = Document(template_path_or_file)
    
    global_replacements = {}
    for k, v in global_context.items():
        global_replacements[f"{{{{ {k} }}}}"] = v
    replace_text_in_document_full(doc, global_replacements)

    # データを挿入する表を探す
    target_table = None
    time_row_template = None
    data_row_template = None
    
    for table in doc.tables:
        t_row = None
        d_row = None
        for row in table.rows:
            row_text = "".join([c.text for c in row.cells])
            if "{{ time }}" in row_text:
                t_row = row
            if "{{ s.no }}" in row_text:
                d_row = row
        
        if t_row and d_row:
            target_table = table
            time_row_template = t_row
            data_row_template = d_row
            break
    
    if target_table:
        tbl = target_table._tbl
        time_tr = time_row_template._tr
        data_tr = data_row_template._tr
        
        tbl.remove(time_tr)
        tbl.remove(data_tr)
        
        for group in groups:
            # 1. 時間行
            new_tr_time = copy.deepcopy(time_tr)
            tbl.append(new_tr_time)
            new_time_row = target_table.rows[-1]
            
            raw_time = group['time_str']
            formatted_time = format_time_label(raw_time)
            fill_row_data(new_time_row, {'{{ time }}': formatted_time})

            # 2. メンバー行
            target_members = resolve_participants_from_string(group['member_input'], all_data)
            
            for member in target_members:
                new_tr_data = copy.deepcopy(data_tr)
                tbl.append(new_tr_data)
                new_data_row = target_table.rows[-1]
                
                replacements = {
                    '{{ s.no }}': member['no'],
                    '{{ s.name }}': member['name'],
                    '{{ s.kana }}': member.get('kana', ''),
                    '{{ s.age }}': member.get('age', ''),
                    '{{ s.tel }}': member.get('tel', ''),
                    '{{ s.song }}': member['song'],
                }
                fill_row_data(new_data_row, replacements)

    output_buffer = io.BytesIO()
    doc.save(output_buffer)
    return output_buffer


def generate_web_program_doc(template_path_or_file, groups, all_data, global_context):
    """
    WEBプログラム用（セル単位スキャン＋書式強制ロジック）
    """
    doc = Document(template_path_or_file)
    
    global_replacements = {}
    for k, v in global_context.items():
        global_replacements[f"{{{{ {k} }}}}"] = v
    
    # --- Step 1: グローバル変数の置換と太字強制 ---
    # ヘッダー・フッター含む全置換
    replace_text_in_document_full(doc, global_replacements)
    
    # 特定タグの太字化（置換後の値を検索して太字にする）
    # ※ contest_open等は対象外なので、ここでは太字にしない
    bold_target_values = [
        global_context.get('contest_name', ''),
        global_context.get('contest_date', ''),
        global_context.get('contest_hall', '')
    ]
    
    def apply_bold_to_targets(doc_obj, target_values):
        def _process_para(para):
            for run in para.runs:
                for val in target_values:
                    if val and val in run.text:
                        run.font.bold = True

        for p in doc_obj.paragraphs: _process_para(p)
        for t in doc_obj.tables:
            for r in t.rows:
                for c in r.cells:
                    for p in c.paragraphs: _process_para(p)
    
    apply_bold_to_targets(doc, bold_target_values)

    # --- Step 2: テンプレート行の特定とループ処理 ---
    template_time_para = None
    template_data_table = None
    
    for p in doc.paragraphs:
        if "{{ time }}" in p.text:
            template_time_para = p
            break
            
    if template_time_para:
        for table in doc.tables:
            txt = ""
            for r in table.rows:
                for c in r.cells:
                    txt += c.text
            if "{{ s.no }}" in txt:
                template_data_table = table
                break
        
        if template_data_table:
            # 要素のコピー
            template_p_xml = copy.deepcopy(template_time_para._p)
            template_tbl_xml = copy.deepcopy(template_data_table._tbl)
            
            # 元の削除
            parent_body = template_time_para._element.getparent()
            if parent_body is not None: parent_body.remove(template_time_para._p)
            
            parent_tbl = template_data_table._tbl.getparent()
            if parent_tbl is not None: parent_tbl.remove(template_data_table._tbl)
            
            # 行テンプレート抽出
            data_tr_list = []
            header_tr_list = []
            temp_rows = list(template_tbl_xml.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}tr'))
            start_index = -1
            rows_per_entry = 2 

            for i, tr in enumerate(temp_rows):
                text_content = "".join([t.text for t in tr.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')])
                if "{{ s.no }}" in text_content:
                    start_index = i
                    break
                else:
                    header_tr_list.append(tr)

            if start_index != -1:
                end_index = min(start_index + rows_per_entry, len(temp_rows))
                data_tr_list = temp_rows[start_index : end_index]
            
            for tr in temp_rows: template_tbl_xml.remove(tr)
            
            doc_body = doc._body._element
            
            for group in groups:
                # 1. 時間
                new_p_xml = copy.deepcopy(template_p_xml)
                doc_body.append(new_p_xml)
                new_para = Paragraph(new_p_xml, doc._body)
                raw_time = group['time_str']
                formatted_time = format_time_label(raw_time)
                replace_text_smart(new_para, {'{{ time }}': formatted_time})
                
                # 2. テーブル
                new_tbl_xml = copy.deepcopy(template_tbl_xml)
                doc_body.append(new_tbl_xml)
                for h_tr in header_tr_list: new_tbl_xml.append(copy.deepcopy(h_tr))
                
                target_members = resolve_participants_from_string(group['member_input'], all_data)
                
                for member in target_members:
                    for tr_template in data_tr_list:
                        new_tr = copy.deepcopy(tr_template)
                        new_tbl_xml.append(new_tr)
                        
                        # 直前に追加された行を取得するためにテーブルを再取得
                        # (XML操作だけではセルの中身を編集できないため)
                        current_table = doc.tables[-1] 
                        current_row = current_table.rows[-1]
                        
                        # --- 重要: セル単位スキャン & 書き込み ---
                        # 行内の全セルをチェックし、特定のタグがある場所にだけ
                        # 指定された書式で書き込む（他のセルのレイアウトは壊さない）
                        
                        for cell in current_row.cells:
                            # タグが含まれているかチェックするためにテキスト取得
                            # ※セル結合されている場合、同じセルオブジェクトが複数回回ってくる可能性があるが、
                            # 内容を書き換えるとタグが消えるため、2回目以降はヒットしないので安全。
                            cell_text = cell.text
                            
                            if "{{ s.no }}" in cell_text:
                                cell.text = "" # クリア
                                p = cell.paragraphs[0]
                                run = p.add_run(f"{member['no']}")
                                run.font.bold = True # 太字
                                
                            if "{{ s.name }}" in cell_text:
                                cell.text = "" # クリア
                                p = cell.paragraphs[0]
                                
                                # 氏名 (太字)
                                run_name = p.add_run(f"{member['name']}")
                                run_name.font.bold = True
                                
                                # スペース・カッコ (標準)
                                run_sep1 = p.add_run(" （")
                                run_sep1.font.bold = False
                                
                                # フリガナ (標準)
                                if member.get('kana'):
                                    run_kana = p.add_run(f"{member['kana']}")
                                    run_kana.font.bold = False
                                
                                # 中黒 (標準)
                                run_sep2 = p.add_run("・")
                                run_sep2.font.bold = False
                                
                                # 年齢 (標準)
                                run_age = p.add_run(f"{member.get('age', '')}")
                                run_age.font.bold = False
                                
                                # 歳・閉じカッコ (標準)
                                run_sep3 = p.add_run("歳）")
                                run_sep3.font.bold = False
                                
                            if "{{ s.song }}" in cell_text:
                                cell.text = "" # クリア
                                p = cell.paragraphs[0]
                                run_song = p.add_run(f"{member['song']}")
                                run_song.font.bold = False # 標準

                doc_body.append(copy.deepcopy(template_p_xml))
                last_p = Paragraph(doc_body[-1], doc._body)
                last_p.text = "" 

    output_buffer = io.BytesIO()
    doc.save(output_buffer)
    return output_buffer


def generate_judges_list_doc(template_path_or_file, judges_list, global_context):
    doc = Document(template_path_or_file)
    global_replacements = {}
    for k, v in global_context.items():
        global_replacements[f"{{{{ {k} }}}}"] = v
    replace_text_in_document_full(doc, global_replacements)

    # 表パターン
    for table in doc.tables:
        target_row_idx = -1
        for i, row in enumerate(table.rows):
            row_text = "".join([c.text for c in row.cells])
            if "{{ judge_name }}" in row_text:
                target_row_idx = i
                break
        
        if target_row_idx != -1:
            template_row = table.rows[target_row_idx]
            tbl = table._tbl
            tr_xml = template_row._tr
            tbl.remove(tr_xml)
            for judge in judges_list:
                new_tr = copy.deepcopy(tr_xml)
                tbl.append(new_tr)
                new_row = table.rows[-1]
                fill_row_data(new_row, {'{{ judge_name }}': judge})
            output_buffer = io.BytesIO()
            doc.save(output_buffer)
            return output_buffer

    # 段落パターン
    target_para = None
    for para in doc.paragraphs:
        if "{{ judge_name }}" in para.text:
            target_para = para
            break
            
    if target_para:
        p_element = target_para._p
        parent = target_para._parent
        template_p_xml = copy.deepcopy(p_element)
        
        if hasattr(parent, '_element'):
             try: parent._element.remove(p_element)
             except: pass
        else:
             try: doc._body._body.remove(p_element)
             except: pass
        
        for judge in judges_list:
            new_p_xml = copy.deepcopy(template_p_xml)
            doc._body._body.append(new_p_xml)
            new_para = Paragraph(new_p_xml, parent)
            replace_text_smart(new_para, {'{{ judge_name }}': judge})

    output_buffer = io.BytesIO()
    doc.save(output_buffer)
    return output_buffer

# ---------------------------------------------------------
# 4. メインアプリケーションUI
# ---------------------------------------------------------
def main():
    st.set_page_config(layout="wide", page_title="コンクール資料作成")
    
    # --- 0. メールアドレス確認 (Gateway) ---
    if 'user_email' not in st.session_state:
        st.session_state['user_email'] = None

    if not st.session_state['user_email']:
        st.title("コンクール運営資料ジェネレーター")
        st.info("メールアドレスの入力をお願いします。")
        
        with st.form("email_login_form"):
            input_email = st.text_input("ご担当者様 メールアドレス", placeholder="example@example.com")
            submit_login = st.form_submit_button("利用を開始する")
            
            if submit_login:
                if input_email and "@" in input_email:
                    st.session_state['user_email'] = input_email
                    st.rerun()
                else:
                    st.error("有効なメールアドレスを入力してください。")
        
        # メールアドレス未入力時はここで処理を止める
        st.stop()

    # --- 以下、メインコンテンツ ---
    st.title("🎹 コンクール運営資料ジェネレーター (Word版)")
    st.markdown(f"**ログイン中:** {st.session_state['user_email']}")
    
    # --- サイドバー: 設定読み込み ---
    with st.sidebar:
        st.header("⚙️ 設定管理")
        uploaded_config = st.file_uploader("設定ファイル(JSON)を読み込む", type=['json'])
        if uploaded_config:
            # 修正: ファイルポインタを先頭に戻す処理を追加
            uploaded_config.seek(0)
            config_data = json.load(uploaded_config)
            st.session_state.update(config_data)
            st.success("設定を復元しました")

    # --- 1. Excelアップロード ---
    st.header("1. 名簿データ (Excel)")
    uploaded_excel = st.file_uploader("名簿Excelファイルをアップロード", type=['xlsx', 'xls', 'csv'])
    
    all_data = []
    
    if uploaded_excel:
        try:
            if uploaded_excel.name.endswith('.csv'):
                df = pd.read_csv(uploaded_excel)
            else:
                xls = pd.ExcelFile(uploaded_excel)
                sheet = st.selectbox("シートを選択", xls.sheet_names)
                df = pd.read_excel(uploaded_excel, sheet_name=sheet)

            # 列の割り当て
            cols = df.columns.tolist()
            c1, c2, c3, c4 = st.columns(4)
            col_no = c1.selectbox("出場番号", cols, index=cols.index("出場番号") if "出場番号" in cols else 0)
            col_name = c2.selectbox("氏名", cols, index=cols.index("氏名") if "氏名" in cols else 0)
            
            default_kana = cols.index("フリガナ") if "フリガナ" in cols else 0
            col_kana = c3.selectbox("フリガナ (任意)", ["(なし)"] + cols, index=default_kana + 1)
            
            col_song = c4.selectbox("演奏曲目", cols, index=cols.index("演奏曲目") if "演奏曲目" in cols else 0)
            
            c5, c6, c7 = st.columns(3)
            default_age = cols.index("年齢") if "年齢" in cols else -1
            col_age = c5.selectbox("年齢列 (任意)", ["(なし)"] + cols, index=default_age + 1)

            default_tel = cols.index("電話番号") if "電話番号" in cols else -1
            col_tel = c6.selectbox("電話番号列 (受付表用)", ["(なし)"] + cols, index=default_tel + 1)

            default_dur = cols.index("演奏時間") if "演奏時間" in cols else -1
            col_duration = c7.selectbox("演奏時間列 (自動計算用)", ["(なし)"] + cols, index=default_dur + 1)

            st.markdown("---")

            for _, row in df.iterrows():
                kana_val = str(row[col_kana]) if col_kana != "(なし)" else ""
                age_val = str(row[col_age]) if col_age != "(なし)" else ""
                tel_val = str(row[col_tel]) if col_tel != "(なし)" else ""
                
                dur_seconds = 0
                if col_duration != "(なし)":
                    raw_dur = str(row[col_duration])
                    dur_seconds = parse_jp_time_to_seconds(raw_dur)

                all_data.append({
                    'no': str(row[col_no]), 
                    'name': str(row[col_name]),
                    'kana': kana_val,
                    'song': str(row[col_song]),
                    'age': age_val,
                    'tel': tel_val,
                    'duration_sec': dur_seconds
                })
            
            st.write(f"読み込み完了: {len(all_data)} 件のデータ")

            # --- 2. テンプレート選択 ---
            st.header("2. Wordテンプレート選択")
            
            TEMPLATE_DIR = "templates"
            template_files = []
            if os.path.exists(TEMPLATE_DIR):
                template_files = [f for f in os.listdir(TEMPLATE_DIR) if f.endswith(".docx") and not f.startswith("~$")]
            
            score_template_path = None
            reception_template_path = None
            web_template_path = None
            judges_list_template_path = None
            
            use_manual_upload = False

            if template_files:
                idx_score = 0
                idx_reception = 0
                idx_web = 0
                idx_judges = 0
                for i, f in enumerate(template_files):
                    if "採点表" in f: idx_score = i
                    if "受付表" in f: idx_reception = i
                    if "WEB" in f or "プログラム" in f: idx_web = i
                    if "審査員" in f and "リスト" not in f: idx_judges = i
                
                col_t1, col_t2 = st.columns(2)
                col_t3, col_t4 = st.columns(2)
                
                with col_t1:
                    selected_score_file = st.selectbox("採点表テンプレート", template_files, index=idx_score)
                    score_template_path = os.path.join(TEMPLATE_DIR, selected_score_file)
                
                with col_t2:
                    selected_reception_file = st.selectbox("受付表テンプレート", template_files, index=idx_reception)
                    reception_template_path = os.path.join(TEMPLATE_DIR, selected_reception_file)
                
                with col_t3:
                    selected_web_file = st.selectbox("WEBプログラムテンプレート", template_files, index=idx_web)
                    web_template_path = os.path.join(TEMPLATE_DIR, selected_web_file)

                with col_t4:
                    selected_judges_file = st.selectbox("審査員リストテンプレート", template_files, index=idx_judges)
                    judges_list_template_path = os.path.join(TEMPLATE_DIR, selected_judges_file)
                
                if st.checkbox("テンプレートを手動でアップロードする"):
                    use_manual_upload = True
            else:
                st.warning("templatesフォルダが見つからないか、docxファイルがありません。手動アップロードモードに切り替えます。")
                use_manual_upload = True

            if use_manual_upload:
                c_up1, c_up2 = st.columns(2)
                c_up3, c_up4 = st.columns(2)
                uploaded_score_template = c_up1.file_uploader("採点表テンプレート (.docx)", type=['docx'])
                uploaded_reception_template = c_up2.file_uploader("受付表テンプレート (.docx)", type=['docx'])
                uploaded_web_template = c_up3.file_uploader("WEBプログラムテンプレート (.docx)", type=['docx'])
                uploaded_judges_template = c_up4.file_uploader("審査員リストテンプレート (.docx)", type=['docx'])
                
                if uploaded_score_template: score_template_path = uploaded_score_template
                if uploaded_reception_template: reception_template_path = uploaded_reception_template
                if uploaded_web_template: web_template_path = uploaded_web_template
                if uploaded_judges_template: judges_list_template_path = uploaded_judges_template

            # --- 3. グループ・スケジュール設定 ---
            st.header("3. グループ・スケジュール設定")
            
            if 'groups' not in st.session_state:
                st.session_state['groups'] = [{'member_input': '', 'time_str': '13:00-14:10'}]
            
            def add_group():
                st.session_state['groups'].append({'member_input': '', 'time_str': ''})
            
            def move_group_up(idx):
                if idx > 0:
                    st.session_state['groups'][idx], st.session_state['groups'][idx-1] = st.session_state['groups'][idx-1], st.session_state['groups'][idx]

            def move_group_down(idx):
                if idx < len(st.session_state['groups']) - 1:
                    st.session_state['groups'][idx], st.session_state['groups'][idx+1] = st.session_state['groups'][idx+1], st.session_state['groups'][idx]
            
            def remove_group(idx):
                st.session_state['groups'].pop(idx)

            st.button("＋ グループ追加", on_click=add_group)

            for i, grp in enumerate(st.session_state['groups']):
                c_sort, c_input, c_total, c_time, c_del = st.columns([0.8, 3, 1.2, 2, 0.5])
                
                with c_sort:
                    if st.button("▲", key=f"up_{i}"):
                        move_group_up(i)
                        st.rerun()
                    if st.button("▼", key=f"down_{i}"):
                        move_group_down(i)
                        st.rerun()

                input_val = c_input.text_input(
                    f"グループ {i+1} 対象番号",
                    value=grp['member_input'],
                    key=f"g_in_{i}",
                    placeholder="例: A01-A05, C01"
                )
                st.session_state['groups'][i]['member_input'] = input_val

                current_members = resolve_participants_from_string(input_val, all_data)
                total_sec = sum(m['duration_sec'] for m in current_members)
                time_display = format_seconds_to_jp_label(total_sec)
                
                with c_total:
                    st.markdown(f"""
                    <div style="margin-bottom: 0px;">
                        <label style="font-size: 14px; color: rgb(49, 51, 63); margin-bottom: 0.5rem; display: block;">
                            合計演奏時間
                        </label>
                        <div style="
                            background-color: rgba(28, 131, 225, 0.1); 
                            border: 1px solid rgba(28, 131, 225, 0.1);
                            border-radius: 0.5rem;
                            padding: 0px 10px;
                            min-height: 2.5rem;
                            height: auto;
                            display: flex;
                            align-items: center;
                            color: rgb(0, 66, 128);
                            font-size: 1rem;
                            line-height: 1.5;
                        ">
                            計: {time_display}
                        </div>
                    </div>
                    """, unsafe_allow_html=True)

                time_val = c_time.text_input(
                    "時間",
                    value=grp['time_str'],
                    key=f"g_time_{i}",
                    placeholder="例: 13:00-14:00"
                )
                st.session_state['groups'][i]['time_str'] = time_val

                with c_del:
                    st.markdown("<div style='margin-top: 1.8rem;'></div>", unsafe_allow_html=True)
                    if st.button("×", key=f"del_{i}"):
                        remove_group(i)
                        st.rerun()

            # --- 4. 審査員設定 ---
            st.header("4. 審査員設定")
            if 'judges' not in st.session_state:
                st.session_state['judges'] = ["審査員A"]
            
            if st.button("＋ 審査員追加"):
                st.session_state['judges'].append("")
                st.rerun()

            for i in range(len(st.session_state['judges'])):
                val = st.text_input(f"審査員 {i+1}", value=st.session_state['judges'][i], key=f"judge_input_{i}")
                st.session_state['judges'][i] = val

            contest_name = st.text_input("コンクール名 (ファイル名等に使用)", "第10回BIPCA 東京予選④")
            st.session_state['contest_name'] = contest_name # セッションに保存(メール件名用)

            # --- 5. 審査会詳細 ---
            st.header("5. 審査会詳細")
            st.info("※ここで入力した内容はWord出力時に自動的に形式変換されて挿入されます。")
            
            if 'contest_details' not in st.session_state:
                st.session_state['contest_details'] = {
                    'date': '', 'hall': '', 'open': '10:00', 'reception': '10:45-15:30',
                    'start': '11:00', 'end': '14:00', 'result': '', 'method': '公式サイト上で掲載'
                }
            
            det = st.session_state['contest_details']

            def on_date_change():
                current_date = st.session_state['detail_date']
                calculated = calculate_next_day_morning(current_date)
                if calculated:
                    st.session_state['contest_details']['result'] = calculated

            col_d1, col_d2 = st.columns(2)
            det['date'] = col_d1.text_input("開催日時 (例: 2025年12月21日)", value=det['date'], key="detail_date", on_change=on_date_change)
            det['hall'] = col_d2.text_input("会場", value=det['hall'])
            
            col_d3, col_d4, col_d5, col_d6 = st.columns(4)
            det['open'] = col_d3.text_input("開場時刻 (例: 10:00)", value=det['open'])
            det['start'] = col_d4.text_input("審査開始 (例: 11:00)", value=det['start'])
            det['end'] = col_d5.text_input("審査終了 (例: 14:00)", value=det['end'])
            det['reception'] = col_d6.text_input("受付時間 (例: 10:45-15:30)", value=det['reception'])

            col_d7, col_d8 = st.columns(2)
            det['result'] = col_d7.text_input("結果発表日時 (自動計算)", value=det['result'])
            
            det['method'] = col_d8.selectbox("結果発表方式", [
                "公式サイト上で掲載",
                "会場ロビーもしくはホワイエで掲載",
                "表彰式にて発表",
                "その他"
            ], index=["公式サイト上で掲載", "会場ロビーもしくはホワイエで掲載", "表彰式にて発表", "その他"].index(det['method']) if det['method'] in ["公式サイト上で掲載", "会場ロビーもしくはホワイエで掲載", "表彰式にて発表", "その他"] else 0)

            # --- 6. ファイル出力 ---
            st.header("6. ファイル出力")
            if st.button("ファイル生成を実行", type="primary"):
                # テンプレートチェック
                if not score_template_path:
                    st.error("採点表テンプレートが選択されていません。")
                    return
                if not web_template_path:
                    st.warning("WEBプログラムテンプレートが選択されていません。WEBプログラムは生成されません。")
                if not judges_list_template_path:
                    st.warning("審査員リストテンプレートが選択されていません。")

                valid_judges = [j for j in st.session_state['judges'] if j.strip()]
                
                details_formatted = {
                    'contest_date': det['date'],
                    'contest_hall': det['hall'],
                    'contest_open': format_single_time_label(det['open']),
                    'contest_reception': format_time_label(det['reception']),
                    'contest_start': format_single_time_label(det['start']),
                    'contest_end': format_single_time_label(det['end']),
                    'contest_result': det['result'],
                    'contest_method': det['method']
                }

                config_json = json.dumps({
                    'groups': st.session_state['groups'],
                    'judges': valid_judges,
                    'contest_name': contest_name,
                    'contest_details': det
                }, ensure_ascii=False, indent=2)

                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                    
                    base_context = {
                        'contest_name': contest_name,
                        **details_formatted
                    }

                    # 1. 採点表生成
                    for judge in valid_judges:
                        try:
                            if hasattr(score_template_path, 'seek'): score_template_path.seek(0)
                            context = base_context.copy()
                            context['judge_name'] = judge
                            doc_io = generate_word_from_template(score_template_path, st.session_state['groups'], all_data, context)
                            zf.writestr(f"採点表_{judge}.docx", doc_io.getvalue())
                        except Exception as e:
                            st.error(f"採点表生成エラー ({judge}): {e}")

                    # 2. 受付表生成
                    if reception_template_path:
                        try:
                            if hasattr(reception_template_path, 'seek'): reception_template_path.seek(0)
                            context = base_context.copy()
                            context['judge_name'] = '受付用'
                            doc_io = generate_word_from_template(reception_template_path, st.session_state['groups'], all_data, context)
                            zf.writestr("受付表.docx", doc_io.getvalue())
                        except Exception as e:
                            st.error(f"受付表生成エラー: {e}")

                    # 3. WEBプログラム生成（修正版）
                    if web_template_path:
                        try:
                            if hasattr(web_template_path, 'seek'): web_template_path.seek(0)
                            context = base_context.copy()
                            context['judge_name'] = ''
                            doc_io = generate_web_program_doc(web_template_path, st.session_state['groups'], all_data, context)
                            zf.writestr("WEBプログラム.docx", doc_io.getvalue())
                        except Exception as e:
                            st.error(f"WEBプログラム生成エラー: {e}")
                            
                    # 4. 審査員リスト生成
                    if judges_list_template_path:
                         try:
                            if hasattr(judges_list_template_path, 'seek'): judges_list_template_path.seek(0)
                            context = base_context.copy()
                            doc_io = generate_judges_list_doc(judges_list_template_path, valid_judges, context)
                            zf.writestr("本日の審査員.docx", doc_io.getvalue())
                         except Exception as e:
                            st.error(f"審査員リスト生成エラー: {e}")

                    # 5. PDFファイルの同梱
                    if os.path.exists(TEMPLATE_DIR):
                        pdf_files = [f for f in os.listdir(TEMPLATE_DIR) if f.endswith(".pdf")]
                        for pdf_file in pdf_files:
                            pdf_path = os.path.join(TEMPLATE_DIR, pdf_file)
                            zf.write(pdf_path, arcname=pdf_file)

                    # 設定ファイル
                    zf.writestr("設定データ.json", config_json)
                
                # ZIPバッファをセッションステートに保存
                st.session_state['zip_buffer'] = zip_buffer
                st.success("生成完了！下のボタンからダウンロードしてください。")
            
            # ダウンロードボタン表示（生成後のみ）
            if 'zip_buffer' in st.session_state and st.session_state['zip_buffer']:
                st.download_button(
                    label="ZIPファイルをダウンロード",
                    data=st.session_state['zip_buffer'].getvalue(),
                    file_name=f"{contest_name}.zip",
                    mime="application/zip",
                    on_click=send_email_callback  # ダウンロード時にメール送信実行
                )

        except Exception as e:
            st.error(f"エラーが発生しました: {e}")

if __name__ == "__main__":
    main()

