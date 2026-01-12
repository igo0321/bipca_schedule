import streamlit as st
import pandas as pd
import io
import zipfile
import json
import smtplib
import re  # 正規表現用に追加
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from datetime import datetime
import copy
from docx import Document

# ---------------------------------------------------------
# 1. ユーティリティ（時間変換・Word操作）
# ---------------------------------------------------------

def format_time_label(text):
    """
    入力された時間文字列（例: 13:00-14:10, 13:00〜14:10）から
    数字を抽出して「13時00分～14時10分」の形式に変換する。
    マッチしない場合は元のテキストを返す。
    """
    if not text:
        return ""
    # 数字1〜2桁 + コロン + 数字2桁 を探す (全角コロンも対応)
    matches = re.findall(r'(\d{1,2})[:：](\d{2})', str(text))
    
    # 開始と終了の2つが見つかった場合のみ変換
    if len(matches) >= 2:
        start_time = f"{matches[0][0]}時{matches[0][1]}分"
        end_time = f"{matches[1][0]}時{matches[1][1]}分"
        return f"{start_time}～{end_time}"
    else:
        return text

def copy_table_row(table, row):
    """表の行を複製して末尾に追加"""
    tbl = table._tbl
    new_tr = copy.deepcopy(row._tr)
    tbl.append(new_tr)
    return table.rows[-1]

def replace_text_in_paragraph(paragraph, replacements):
    """段落内のテキストを置換"""
    # 完全に一致するRunがあれば置換（書式維持のため）
    for key, value in replacements.items():
        if key in paragraph.text:
            replaced = False
            for run in paragraph.runs:
                if key in run.text:
                    run.text = run.text.replace(key, str(value))
                    replaced = True
            
            # Run単位で置換できなかった場合、段落全体を書き換え
            # (注意: 途中で書式が変わっていると書式がリセットされる場合があります)
            if not replaced:
                full_text = paragraph.text
                new_text = full_text.replace(key, str(value))
                if paragraph.runs:
                    paragraph.runs[0].text = new_text
                    for r in paragraph.runs[1:]:
                        r.text = ""

def fill_row_data(row, data_dict):
    """行内の全セルのテキストを置換"""
    for cell in row.cells:
        for paragraph in cell.paragraphs:
            replace_text_in_paragraph(paragraph, data_dict)

def delete_row(table, row_idx):
    """指定行を削除"""
    tbl = table._tbl
    tr = table.rows[row_idx]._tr
    tbl.remove(tr)

def replace_text_in_document_body(doc, replacements):
    """
    表以外の本文やヘッダー内のテキストを置換する
    """
    # 1. 本文の段落
    for paragraph in doc.paragraphs:
        replace_text_in_paragraph(paragraph, replacements)
    
    # 2. ヘッダー/フッター（セクションごと）
    for section in doc.sections:
        # ヘッダー
        for paragraph in section.header.paragraphs:
            replace_text_in_paragraph(paragraph, replacements)
        # フッター
        for paragraph in section.footer.paragraphs:
            replace_text_in_paragraph(paragraph, replacements)

# ---------------------------------------------------------
# 2. ドキュメント生成メインロジック
# ---------------------------------------------------------

def generate_word_from_template(template_file, groups, all_data, global_context):
    """
    template_file: Wordテンプレート
    groups: グループ設定リスト
    all_data: Excelから読み込んだ参加者データリスト
    global_context: { 'contest_name': '...', 'judge_name': '...' } などの共通情報
    """
    doc = Document(template_file)
    
    # --- A. 全体情報の置換（ヘッダーやタイトルなど） ---
    # 置換用タグの作成 (例: {{ contest_name }})
    global_replacements = {}
    for k, v in global_context.items():
        global_replacements[f"{{{{ {k} }}}}"] = v  # {{ key }} 形式に変換
    
    replace_text_in_document_body(doc, global_replacements)

    # --- B. 表データの生成 ---
    if not doc.tables:
        # 表がない場合でもエラーにせず、そのまま返す（表紙だけの場合など考慮）
        output_buffer = io.BytesIO()
        doc.save(output_buffer)
        return output_buffer
    
    table = doc.tables[0] # 最初の表を対象とする
    
    # テンプレート構造の前提:
    # 0行目: ヘッダー
    # 1行目: 時間区切り用の行（ひな形）
    # 2行目: データ表示用の行（ひな形）
    
    if len(table.rows) < 3:
        raise Exception("テンプレートの表は少なくとも3行（ヘッダー、時間行、データ行）必要です。")

    # ひな形の行を参照・コピーしておく
    time_row_template = table.rows[1]
    data_row_template = table.rows[2]
    
    # ひな形行をテーブルから削除
    delete_row(table, 2)
    delete_row(table, 1)
    
    # グループごとに処理
    for group in groups:
        # 1. 時間行を追加
        new_time_row = copy_table_row(table, time_row_template)
        
        # 時間文字列の変換処理
        raw_time = group['time_str']
        formatted_time = format_time_label(raw_time)
        
        # 時間行の中にある {{ time }} タグを置換
        fill_row_data(new_time_row, {'{{ time }}': formatted_time})

        # 2. そのグループに該当するデータを抽出
        target_members = []
        s_no = group['start_no']
        e_no = group['end_no']
        in_range = False
        
        for item in all_data:
            current_no = str(item['no'])
            if s_no and current_no == s_no:
                in_range = True
            if in_range:
                target_members.append(item)
            if e_no and current_no == e_no:
                in_range = False
        
        # 3. メンバーごとにデータ行を追加
        for member in target_members:
            new_data_row = copy_table_row(table, data_row_template)
            
            # データ行の置換辞書
            replacements = {
                '{{ s.no }}': member['no'],
                '{{ s.name }}': member['name'],
                '{{ s.kana }}': member.get('kana', ''), # フリガナ追加
                '{{ s.age }}': member.get('age', ''),
                '{{ s.song }}': member['song'],
            }
            fill_row_data(new_data_row, replacements)

    output_buffer = io.BytesIO()
    doc.save(output_buffer)
    return output_buffer

# ---------------------------------------------------------
# 3. メール送信機能 (変更なし)
# ---------------------------------------------------------
def send_email_with_attachment(zip_buffer, zip_filename, contest_name):
    # (既存のコードと同じため省略しませんが、スペース節約のため中身は変更なし)
    try:
        if "email" not in st.secrets:
             return False, "Secretsにメール設定がありません。"
        smtp_server = st.secrets["email"]["smtp_server"]
        smtp_port = st.secrets["email"]["smtp_port"]
        sender_email = st.secrets["email"]["sender_email"]
        sender_password = st.secrets["email"]["sender_password"]
        receiver_email = "info@beethoven-asia.com"

        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = receiver_email
        msg['Subject'] = f"【自動送信】資料出力: {contest_name}"
        body = f"コンクール名: {contest_name}\n出力日時: {datetime.now().strftime('%Y/%m/%d %H:%M')}\n\n資料を添付します。"
        msg.attach(MIMEText(body, 'plain'))

        part = MIMEApplication(zip_buffer.getvalue(), Name=zip_filename)
        part['Content-Disposition'] = f'attachment; filename="{zip_filename}"'
        msg.attach(part)

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
# 4. メインアプリケーションUI
# ---------------------------------------------------------
def main():
    st.title("🎹 コンクール運営資料ジェネレーター (Word版)")
    
    # --- サイドバー: 設定読み込み ---
    with st.sidebar:
        st.header("⚙️ 設定管理")
        uploaded_config = st.file_uploader("設定ファイル(JSON)を読み込む", type=['json'])
        if uploaded_config:
            config_data = json.load(uploaded_config)
            st.session_state.update(config_data)
            st.success("設定を復元しました")

    # --- 1. Excelアップロード ---
    st.header("1. 名簿データ (Excel)")
    uploaded_excel = st.file_uploader("名簿Excelファイルをアップロード", type=['xlsx', 'xls', 'csv'])
    
    if uploaded_excel:
        try:
            if uploaded_excel.name.endswith('.csv'):
                df = pd.read_csv(uploaded_excel)
            else:
                xls = pd.ExcelFile(uploaded_excel)
                sheet = st.selectbox("シートを選択", xls.sheet_names)
                df = pd.read_excel(uploaded_excel, sheet_name=sheet)

            st.write("データプレビュー:", df.head(3))
            
            # 列の割り当て
            cols = df.columns.tolist()
            c1, c2, c3, c4 = st.columns(4)
            col_no = c1.selectbox("出場番号", cols, index=cols.index("出場番号") if "出場番号" in cols else 0)
            col_name = c2.selectbox("氏名", cols, index=cols.index("氏名") if "氏名" in cols else 0)
            
            # フリガナ列 (任意)
            default_kana_idx = cols.index("フリガナ") if "フリガナ" in cols else 0
            col_kana = c3.selectbox("フリガナ (任意)", ["(なし)"] + cols, index=default_kana_idx if "フリガナ" in cols else 0)
            
            col_song = c4.selectbox("演奏曲目", cols, index=cols.index("演奏曲目") if "演奏曲目" in cols else 0)
            
            # 年齢列 (任意)
            default_age_idx = cols.index("年齢") + 1 if "年齢" in cols else 0
            col_age = st.selectbox("年齢列 (任意)", ["(なし)"] + cols, index=default_age_idx)
            
            # データ変換
            all_data = []
            for _, row in df.iterrows():
                # 任意の列の取得
                age_val = str(row[col_age]) if col_age != "(なし)" else ""
                kana_val = str(row[col_kana]) if col_kana != "(なし)" else ""
                
                all_data.append({
                    'no': str(row[col_no]), 
                    'name': str(row[col_name]),
                    'kana': kana_val,
                    'song': str(row[col_song]),
                    'age': age_val
                })

            # --- 2. テンプレートアップロード ---
            st.header("2. Wordテンプレート")
            st.info("""
            以下のタグが使用可能です：
            - 文書全体: {{ contest_name }}, {{ judge_name }}
            - 表の時間行: {{ time }}
            - 表のデータ行: {{ s.no }}, {{ s.name }}, {{ s.kana }}, {{ s.age }}, {{ s.song }}
            """)
            uploaded_template = st.file_uploader("Wordテンプレート (.docx)", type=['docx'])

            # --- 3. スケジュール設定 ---
            st.header("3. グループ・スケジュール設定")
            if 'groups' not in st.session_state:
                st.session_state['groups'] = [{'start_no': '', 'end_no': '', 'time_str': '13:00-14:10'}]
            
            if st.button("＋ グループ追加"):
                st.session_state['groups'].append({'start_no': '', 'end_no': '', 'time_str': ''})
            
            groups_config = []
            for i, grp in enumerate(st.session_state['groups']):
                with st.expander(f"グループ {i+1}", expanded=True):
                    c1, c2, c3 = st.columns([1, 1, 2])
                    grp['start_no'] = c1.text_input(f"開始番号", grp['start_no'], key=f"s_{i}")
                    grp['end_no'] = c2.text_input(f"終了番号", grp['end_no'], key=f"e_{i}")
                    grp['time_str'] = c3.text_input(f"時間 (例: 13:00-14:10)", grp['time_str'], key=f"t_{i}")
                    groups_config.append(grp)

            # --- 4. 審査員設定 ---
            st.header("4. 審査員設定")
            if 'judges' not in st.session_state:
                st.session_state['judges'] = ["審査員A"]
            
            if st.button("＋ 審査員追加"):
                st.session_state['judges'].append("")
            
            judges_list = []
            for i, j in enumerate(st.session_state['judges']):
                judges_list.append(st.text_input(f"審査員 {i+1}", j, key=f"j_{i}"))
            st.session_state['judges'] = [j for j in judges_list if j] # 空白除去
            
            # コンクール名
            contest_name = st.text_input("コンクール名 (ファイル名・置換用)", "第10回BIPCA 東京予選④")

            # --- 5. 出力 ---
            if st.button("ファイル生成を実行"):
                if not uploaded_template:
                    st.error("Wordテンプレートをアップロードしてください。")
                else:
                    # 設定保存データの作成
                    config_json = json.dumps({
                        'groups': groups_config,
                        'judges': judges_list,
                        'contest_name': contest_name
                    }, ensure_ascii=False, indent=2)

                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                        
                        # 1. 採点表 (審査員ごと)
                        for judge in st.session_state['judges']:
                            uploaded_template.seek(0)
                            try:
                                # コンテキスト（共通情報）の作成
                                context = {
                                    'contest_name': contest_name,
                                    'judge_name': judge
                                }
                                
                                doc_io = generate_word_from_template(uploaded_template, groups_config, all_data, context)
                                zf.writestr(f"採点表_{judge}.docx", doc_io.getvalue())
                            except Exception as e:
                                st.error(f"採点表生成エラー ({judge}): {e}")

                        # 設定ファイル
                        zf.writestr("設定データ.json", config_json)
                    
                    st.success("生成完了！")
                    
                    # メール送信
                    sent, msg = send_email_with_attachment(zip_buffer, f"{contest_name}.zip", contest_name)
                    if sent:
                        st.info(f"メール送信完了: {msg}")
                    else:
                        st.warning(msg)
                    
                    # ダウンロード
                    st.download_button(
                        "ZIPファイルをダウンロード",
                        zip_buffer.getvalue(),
                        f"{contest_name}.zip",
                        "application/zip"
                    )

        except Exception as e:
            st.error(f"予期せぬエラー: {e}")

if __name__ == "__main__":
    main()
