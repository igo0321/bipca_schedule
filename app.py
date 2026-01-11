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
import copy
from docx import Document

# ---------------------------------------------------------
# 1. Word操作用ユーティリティ（行の複製・置換など）
# ---------------------------------------------------------

def copy_table_row(table, row):
    """
    表の指定された行（row）を、XMLレベルで複製して表の末尾に追加する。
    スタイル（罫線、高さ、フォントなど）を維持する。
    """
    tbl = table._tbl
    new_tr = copy.deepcopy(row._tr)
    tbl.append(new_tr)
    return table.rows[-1]

def replace_text_in_paragraph(paragraph, replacements):
    """
    段落内のテキストを指定された辞書に基づいて置換する。
    """
    for key, value in replacements.items():
        if key in paragraph.text:
            replaced = False
            for run in paragraph.runs:
                if key in run.text:
                    run.text = run.text.replace(key, str(value))
                    replaced = True
            
            if not replaced:
                full_text = paragraph.text
                new_text = full_text.replace(key, str(value))
                if paragraph.runs:
                    paragraph.runs[0].text = new_text
                    for r in paragraph.runs[1:]:
                        r.text = ""

def fill_row_data(row, data_dict):
    """行内の全セルのテキストを置換データに基づいて更新する"""
    for cell in row.cells:
        for paragraph in cell.paragraphs:
            replace_text_in_paragraph(paragraph, data_dict)

def delete_row(table, row_idx):
    """指定されたインデックスの行を削除する"""
    tbl = table._tbl
    tr = table.rows[row_idx]._tr
    tbl.remove(tr)

# ---------------------------------------------------------
# 2. ドキュメント生成メインロジック
# ---------------------------------------------------------

def generate_word_from_template(template_file, groups, all_data):
    """
    Wordテンプレートを読み込み、グループ設定とデータに基づいて行を増殖させる。
    """
    doc = Document(template_file)
    
    if not doc.tables:
        raise Exception("テンプレート内に表が見つかりません。")
    
    table = doc.tables[0]
    
    # テンプレート構造の前提:
    # 0行目: ヘッダー
    # 1行目: 時間区切り用の行（ひな形）
    # 2行目: データ表示用の行（ひな形）
    
    if len(table.rows) < 3:
        raise Exception("テンプレートの表は少なくとも3行（ヘッダー、時間行、データ行）必要です。")

    # ひな形の行を取得（参照を保持）
    time_row_template = table.rows[1]
    data_row_template = table.rows[2]
    
    # ひな形行をテーブルから一旦削除する
    delete_row(table, 2) # データ行を削除
    delete_row(table, 1) # 時間行を削除
    
    # グループごとに処理
    for group in groups:
        # 1. 時間行を追加
        new_time_row = copy_table_row(table, time_row_template)
        # 時間のテキストを置換
        if group['time_str']:
            # セルの最初の段落を書き換える
            if new_time_row.cells[0].paragraphs:
                 new_time_row.cells[0].paragraphs[0].text = group['time_str']
            else:
                 new_time_row.cells[0].add_paragraph(group['time_str'])

        # 2. そのグループに該当するデータを抽出
        target_members = []
        
        s_no = group['start_no']
        e_no = group['end_no']
        
        in_range = False
        
        # 全データを走査して範囲内のメンバーを抽出
        for item in all_data:
            current_no = str(item['no'])
            
            # 開始番号と一致したら範囲内フラグON
            if s_no and current_no == s_no:
                in_range = True
            
            if in_range:
                target_members.append(item)
            
            # 終了番号と一致したら、この人を含めて終了（次回からOFF）
            if e_no and current_no == e_no:
                in_range = False
        
        # 3. メンバーごとにデータ行を追加
        for member in target_members:
            new_data_row = copy_table_row(table, data_row_template)
            
            # 置換用辞書の作成
            replacements = {
                '{{ s.no }}': member['no'],
                '{{ s.name }}': member['name'],
                '{{ s.age }}': member.get('age', ''),
                '{{ s.song }}': member['song'],
            }
            fill_row_data(new_data_row, replacements)

    output_buffer = io.BytesIO()
    doc.save(output_buffer)
    return output_buffer

# ---------------------------------------------------------
# 3. メール送信機能
# ---------------------------------------------------------
def send_email_with_attachment(zip_buffer, zip_filename, contest_name):
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
            c1, c2, c3 = st.columns(3)
            col_no = c1.selectbox("出場番号列", cols, index=cols.index("出場番号") if "出場番号" in cols else 0)
            col_name = c2.selectbox("氏名列", cols, index=cols.index("氏名") if "氏名" in cols else 0)
            col_song = c3.selectbox("曲目列", cols, index=cols.index("演奏曲目") if "演奏曲目" in cols else 0)
            
            # 年齢列の選択（インデックス計算を修正済み）
            default_age_idx = cols.index("年齢") + 1 if "年齢" in cols else 0
            col_age = st.selectbox("年齢列 (任意)", ["(なし)"] + cols, index=default_age_idx)
            
            # データ変換
            all_data = []
            for _, row in df.iterrows():
                # 年齢データの取得処理を修正
                age_val = ""
                if col_age != "(なし)":
                    age_val = str(row[col_age])
                
                all_data.append({
                    'no': str(row[col_no]), 
                    'name': str(row[col_name]),
                    'song': str(row[col_song]),
                    'age': age_val
                })

            # --- 2. テンプレートアップロード ---
            st.header("2. Wordテンプレート")
            st.info("2行目に「時間行」、3行目に「データ行({{ s.name }}等)」があるWordファイルをアップロードしてください。")
            uploaded_template = st.file_uploader("Wordテンプレート (.docx)", type=['docx'])

            # --- 3. スケジュール設定 ---
            st.header("3. グループ・スケジュール設定")
            if 'groups' not in st.session_state:
                st.session_state['groups'] = [{'start_no': '', 'end_no': '', 'time_str': '13:00〜14:10'}]
            
            if st.button("＋ グループ追加"):
                st.session_state['groups'].append({'start_no': '', 'end_no': '', 'time_str': ''})
            
            groups_config = []
            for i, grp in enumerate(st.session_state['groups']):
                with st.expander(f"グループ {i+1}", expanded=True):
                    c1, c2, c3 = st.columns([1, 1, 2])
                    grp['start_no'] = c1.text_input(f"開始番号", grp['start_no'], key=f"s_{i}")
                    grp['end_no'] = c2.text_input(f"終了番号", grp['end_no'], key=f"e_{i}")
                    grp['time_str'] = c3.text_input(f"表示時間 (例: 13:00〜14:10)", grp['time_str'], key=f"t_{i}")
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
            st.session_state['judges'] = judges_list
            
            # コンクール名
            contest_name = st.text_input("コンクール名 (ファイル名用)", "第10回BIPCA 東京予選④")

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
                        for judge in judges_list:
                            if not judge: continue
                            uploaded_template.seek(0)
                            try:
                                doc_io = generate_word_from_template(uploaded_template, groups_config, all_data)
                                zf.writestr(f"採点表_{judge}.docx", doc_io.getvalue())
                            except Exception as e:
                                st.error(f"採点表生成エラー ({judge}): {e}")

                        # 2. 受付表
                        uploaded_template.seek(0)
                        try:
                            doc_io = generate_word_from_template(uploaded_template, groups_config, all_data)
                            zf.writestr("受付表.docx", doc_io.getvalue())
                        except Exception as e:
                            pass

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
