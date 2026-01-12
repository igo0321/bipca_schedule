import streamlit as st
import pandas as pd
import io
import zipfile
import json
import smtplib
import re
import os  # ファイル操作用に追加
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from datetime import datetime
import copy
from docx import Document

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

def copy_table_row(table, row):
    tbl = table._tbl
    new_tr = copy.deepcopy(row._tr)
    tbl.append(new_tr)
    return table.rows[-1]

def replace_text_in_paragraph(paragraph, replacements):
    for key, value in replacements.items():
        if key in paragraph.text:
            full_text = paragraph.text
            new_text = full_text.replace(key, str(value))
            if paragraph.runs:
                r = paragraph.runs[0]
                r.text = new_text
                for sub_r in paragraph.runs[1:]:
                    sub_r.text = ""

def fill_row_data(row, data_dict):
    for cell in row.cells:
        for paragraph in cell.paragraphs:
            replace_text_in_paragraph(paragraph, data_dict)

def delete_row(table, row_idx):
    tbl = table._tbl
    tr = table.rows[row_idx]._tr
    tbl.remove(tr)

def replace_text_in_document_body(doc, replacements):
    for paragraph in doc.paragraphs:
        replace_text_in_paragraph(paragraph, replacements)
    for section in doc.sections:
        for paragraph in section.header.paragraphs:
            replace_text_in_paragraph(paragraph, replacements)
        for paragraph in section.footer.paragraphs:
            replace_text_in_paragraph(paragraph, replacements)

# ---------------------------------------------------------
# 2. ドキュメント生成メインロジック
# ---------------------------------------------------------

def generate_word_from_template(template_path_or_file, groups, all_data, global_context):
    """
    template_path_or_file: ファイルパス(str) または アップロードされたファイルオブジェクト
    """
    doc = Document(template_path_or_file)
    
    # 全体情報の置換
    global_replacements = {}
    for k, v in global_context.items():
        global_replacements[f"{{{{ {k} }}}}"] = v
    replace_text_in_document_body(doc, global_replacements)

    if not doc.tables:
        output_buffer = io.BytesIO()
        doc.save(output_buffer)
        return output_buffer
    
    table = doc.tables[0]
    
    if len(table.rows) < 3:
        # テーブル行数が足りない場合の安全策（表紙のみテンプレートなどの可能性）
        output_buffer = io.BytesIO()
        doc.save(output_buffer)
        return output_buffer

    time_row_template = table.rows[1]
    data_row_template = table.rows[2]
    
    delete_row(table, 2)
    delete_row(table, 1)
    
    for group in groups:
        # 1. 時間行
        new_time_row = copy_table_row(table, time_row_template)
        raw_time = group['time_str']
        formatted_time = format_time_label(raw_time)
        fill_row_data(new_time_row, {'{{ time }}': formatted_time})

        # 2. メンバー解決
        target_members = resolve_participants_from_string(group['member_input'], all_data)
        
        # 3. データ行
        for member in target_members:
            new_data_row = copy_table_row(table, data_row_template)
            replacements = {
                '{{ s.no }}': member['no'],
                '{{ s.name }}': member['name'],
                '{{ s.kana }}': member.get('kana', ''),
                '{{ s.age }}': member.get('age', ''),
                '{{ s.tel }}': member.get('tel', ''),  # 電話番号対応
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
        # 実装省略（変更なし）
        return True, "メール送信(ダミー)成功" 
    except Exception as e:
        return False, str(e)

# ---------------------------------------------------------
# 4. メインアプリケーションUI
# ---------------------------------------------------------
def main():
    st.set_page_config(layout="wide", page_title="コンクール資料作成")
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
            
            # 追加オプション列 (電話番号追加)
            c5, c6, c7 = st.columns(3)
            
            default_age = cols.index("年齢") if "年齢" in cols else -1
            col_age = c5.selectbox("年齢列 (任意)", ["(なし)"] + cols, index=default_age + 1)

            default_tel = cols.index("電話番号") if "電話番号" in cols else -1
            col_tel = c6.selectbox("電話番号列 (受付表用)", ["(なし)"] + cols, index=default_tel + 1)

            default_dur = cols.index("演奏時間") if "演奏時間" in cols else -1
            col_duration = c7.selectbox("演奏時間列 (自動計算用)", ["(なし)"] + cols, index=default_dur + 1)

            st.markdown("---")

            # データ変換
            for _, row in df.iterrows():
                kana_val = str(row[col_kana]) if col_kana != "(なし)" else ""
                age_val = str(row[col_age]) if col_age != "(なし)" else ""
                tel_val = str(row[col_tel]) if col_tel != "(なし)" else "" # 電話番号取得
                
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

            # --- 2. テンプレート選択 (GitHub/Local 対応) ---
            st.header("2. Wordテンプレート選択")
            
            TEMPLATE_DIR = "templates"
            template_files = []
            
            # ディレクトリチェック
            if os.path.exists(TEMPLATE_DIR):
                template_files = [f for f in os.listdir(TEMPLATE_DIR) if f.endswith(".docx") and not f.startswith("~$")]
            
            score_template_path = None
            reception_template_path = None
            use_manual_upload = False

            if template_files:
                col_t1, col_t2 = st.columns(2)
                
                # デフォルト値の自動検出
                idx_score = 0
                idx_reception = 0
                for i, f in enumerate(template_files):
                    if "採点表" in f: idx_score = i
                    if "受付表" in f: idx_reception = i
                
                with col_t1:
                    selected_score_file = st.selectbox("採点表テンプレート", template_files, index=idx_score)
                    score_template_path = os.path.join(TEMPLATE_DIR, selected_score_file)
                
                with col_t2:
                    selected_reception_file = st.selectbox("受付表テンプレート", template_files, index=idx_reception)
                    reception_template_path = os.path.join(TEMPLATE_DIR, selected_reception_file)
                
                # 手動アップロードへの切り替えオプション
                if st.checkbox("テンプレートを手動でアップロードする"):
                    use_manual_upload = True
            else:
                st.warning("templatesフォルダが見つからないか、docxファイルがありません。手動アップロードモードに切り替えます。")
                use_manual_upload = True

            # 手動アップロード (フォールバック)
            if use_manual_upload:
                c_up1, c_up2 = st.columns(2)
                uploaded_score_template = c_up1.file_uploader("採点表テンプレート (.docx)", type=['docx'])
                uploaded_reception_template = c_up2.file_uploader("受付表テンプレート (.docx)", type=['docx'])
                
                if uploaded_score_template:
                    score_template_path = uploaded_score_template
                if uploaded_reception_template:
                    reception_template_path = uploaded_reception_template

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
                
                # --- 並べ替えボタン ---
                with c_sort:
                    if st.button("▲", key=f"up_{i}"):
                        move_group_up(i)
                        st.rerun()
                    if st.button("▼", key=f"down_{i}"):
                        move_group_down(i)
                        st.rerun()

                # --- メンバー指定入力 ---
                input_val = c_input.text_input(
                    f"グループ {i+1} 対象番号",
                    value=grp['member_input'],
                    key=f"g_in_{i}",
                    placeholder="例: A01-A05, C01"
                )
                st.session_state['groups'][i]['member_input'] = input_val

                # --- 合計時間計算 & 表示 (レイアウト修正版) ---
                current_members = resolve_participants_from_string(input_val, all_data)
                total_sec = sum(m['duration_sec'] for m in current_members)
                time_display = format_seconds_to_jp_label(total_sec)
                
                with c_total:
                    # HTML/CSSで入力欄と高さを完全に合わせた青いボックスを作成
                    # height: 45px程度がStreamlitのinput boxに近い高さです
                    st.markdown(f"""
                    <div style="margin-bottom: 0px;">
                        <label style="font-size: 14px; color: rgb(49, 51, 63); margin-bottom: 8px; display: block;">
                            合計演奏時間
                        </label>
                        <div style="
                            background-color: rgba(28, 131, 225, 0.1); 
                            border: 1px solid rgba(28, 131, 225, 0.1);
                            border-radius: 0.5rem;
                            padding: 0px 10px;
                            height: 42px;
                            display: flex;
                            align-items: center;
                            color: rgb(0, 66, 128);
                            font-size: 1rem;
                        ">
                            計: {time_display}
                        </div>
                    </div>
                    """, unsafe_allow_html=True)

                # --- 時間設定入力 ---
                time_val = c_time.text_input(
                    "時間",
                    value=grp['time_str'],
                    key=f"g_time_{i}",
                    placeholder="例: 13:00-14:00"
                )
                st.session_state['groups'][i]['time_str'] = time_val

                # --- 削除ボタン ---
                with c_del:
                    # ボタン位置を少し下げるためのスペーサー（任意）
                    st.write("") 
                    st.write("")
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

            contest_name = st.text_input("コンクール名", "第10回BIPCA 東京予選④")

            # --- 5. 出力 ---
            if st.button("ファイル生成を実行", type="primary"):
                # テンプレートチェック
                if not score_template_path:
                    st.error("採点表テンプレートが選択されていません。")
                    return
                # 受付表は任意ではなく必須とする場合はチェックを追加
                if not reception_template_path:
                    st.warning("受付表テンプレートが選択されていません。受付表は生成されません。")

                valid_judges = [j for j in st.session_state['judges'] if j.strip()]
                
                config_json = json.dumps({
                    'groups': st.session_state['groups'],
                    'judges': valid_judges,
                    'contest_name': contest_name
                }, ensure_ascii=False, indent=2)

                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                    
                    # 1. 採点表生成 (審査員ごと)
                    for judge in valid_judges:
                        try:
                            # ファイルパスかアップロードオブジェクトかで分岐せずに済むよう関数側で対応済み
                            # ただしアップロードオブジェクトの場合はポインタを戻す必要がある
                            if hasattr(score_template_path, 'seek'):
                                score_template_path.seek(0)
                            
                            context = {'contest_name': contest_name, 'judge_name': judge}
                            doc_io = generate_word_from_template(score_template_path, st.session_state['groups'], all_data, context)
                            zf.writestr(f"採点表_{judge}.docx", doc_io.getvalue())
                        except Exception as e:
                            st.error(f"採点表生成エラー ({judge}): {e}")

                    # 2. 受付表生成 (1回のみ)
                    if reception_template_path:
                        try:
                            if hasattr(reception_template_path, 'seek'):
                                reception_template_path.seek(0)
                            
                            # 受付表用コンテキスト（審査員名は不要だがコンクール名は渡す）
                            context = {'contest_name': contest_name, 'judge_name': '受付用'}
                            doc_io = generate_word_from_template(reception_template_path, st.session_state['groups'], all_data, context)
                            zf.writestr("受付表.docx", doc_io.getvalue())
                        except Exception as e:
                            st.error(f"受付表生成エラー: {e}")

                    # 設定ファイル
                    zf.writestr("設定データ.json", config_json)
                
                st.success("生成完了！")
                
                st.download_button(
                    "ZIPファイルをダウンロード",
                    zip_buffer.getvalue(),
                    f"{contest_name}.zip",
                    "application/zip"
                )

        except Exception as e:
            st.error(f"エラーが発生しました: {e}")

if __name__ == "__main__":
    main()
