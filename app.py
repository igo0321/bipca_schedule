import streamlit as st
import pandas as pd
import io
import zipfile
import json
import smtplib
import re
import os
import copy
from datetime import datetime, timedelta
from docx import Document
from docx.text.paragraph import Paragraph
# XML操作用
from docx.oxml import OxmlElement

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

def copy_table_row(table, row):
    tbl = table._tbl
    new_tr = copy.deepcopy(row._tr)
    tbl.append(new_tr)
    return table.rows[-1]

def replace_text_in_paragraph_merged(paragraph, replacements):
    """
    【修正版】タグが分割されていても、Runを結合して正しく置換するロジック。
    文字位置計算のズレを防ぐため、単純化して処理する。
    """
    # まずテキスト全体にキーが含まれているか確認（高速化）
    full_text = paragraph.text
    if not any(k in full_text for k in replacements):
        return

    # キーが含まれている場合、Runの構造を整理する
    # シンプルな戦略: 全テキストを取得し、置換を行い、それを「最初のRun」に入れて、残りのRunをクリアする
    # ★重要: これだと書式が「最初のRun」のものに統一されてしまうが、
    # タグ（{{ ... }}）の途中で書式が変わることは稀であると仮定する。
    # むしろ変に計算して壊れるより安全。
    
    # ただし、タグ以外の場所（例: "開場: " の太字部分）まで巻き込まないように注意が必要。
    # よって、「タグ部分だけ」を特定して、その範囲のRunをマージする処理が必要だが、
    # WordのXML構造上、インデックス計算はリスクが高い。
    
    # 折衷案: 
    # 段落内のテキストをスキャンし、タグが見つかったら、そのタグを構成しているRun群を特定して書き換える。
    
    # 簡易実装（フェールセーフ）:
    # もしタグがそのまま1つのRunに入っていれば単純置換（これが理想）
    for key, value in replacements.items():
        for run in paragraph.runs:
            if key in run.text:
                run.text = run.text.replace(key, str(value))
                
    # 分割されている場合の処理
    # XML操作を行わず、python-docxのレベルで解決を試みる
    # テキスト全体を再取得
    full_text = paragraph.text
    for key, value in replacements.items():
        if key in full_text:
            # まだ置換されていない（＝分割されている）
            
            # 戦略: 
            # 1. 全Runのテキストをリスト化
            # 2. 結合文字列上で置換を実行
            # 3. 置換後の文字列を、最初のRunに書き戻し、以降のRunをクリア...
            #    これだと段落全体の書式が最初のもので統一されてしまう。
            #    → "開場: {{ time }}" の場合、"開場: "の書式が適用されるならOKだが、
            #      もし "開場: " がRun1(Bold), "{{ time }}" がRun2(Normal) だとすると、Run1に統合されるとBoldになる。
            
            # 今回の不具合（{{ cont...）は、インデックス計算のズレが原因。
            # 安全策として、「段落内の全Runを統合して1つにする」処理を行う。
            # 書式の細かい混在（1行の中で赤と青が混ざるなど）は犠牲になる可能性があるが、
            # 文字化けやタグ破損よりはマシである。
            
            # ただし、極力既存のテキスト（タグ以外）を守るため、
            # 「タグを含むRunの範囲」だけを統合したい。
            
            runs = paragraph.runs
            if not runs: continue
            
            # 全結合して置換
            new_text = full_text.replace(key, str(value))
            
            # 全Runをクリア
            for run in runs:
                run.text = ""
                
            # 最初のRunに新しいテキストを設定
            # (※これにより、段落全体の書式は「最初のRun」のものになる)
            runs[0].text = new_text

def fill_row_data(row, data_dict):
    for cell in row.cells:
        for paragraph in cell.paragraphs:
            replace_text_in_paragraph_merged(paragraph, data_dict)

def delete_row(table, row):
    tbl = table._tbl
    tr = row._tr
    tbl.remove(tr)

def replace_text_in_document_body(doc, replacements):
    for paragraph in doc.paragraphs:
        replace_text_in_paragraph_merged(paragraph, replacements)
    for section in doc.sections:
        for paragraph in section.header.paragraphs:
            replace_text_in_paragraph_merged(paragraph, replacements)
        for paragraph in section.footer.paragraphs:
            replace_text_in_paragraph_merged(paragraph, replacements)

# ---------------------------------------------------------
# 2. ドキュメント生成メインロジック
# ---------------------------------------------------------

def generate_word_from_template(template_path_or_file, groups, all_data, global_context):
    """
    採点表・受付表用（1つの表の中で完結するタイプ）
    """
    doc = Document(template_path_or_file)
    
    global_replacements = {}
    for k, v in global_context.items():
        global_replacements[f"{{{{ {k} }}}}"] = v
    replace_text_in_document_body(doc, global_replacements)

    # 表処理
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
            if "{{ s.no }}" in row_text or "{{ s.name }}" in row_text:
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
    WEBプログラム用（段落(時間)＋表(データ) のセットを繰り返すタイプ）
    """
    doc = Document(template_path_or_file)
    
    global_replacements = {}
    for k, v in global_context.items():
        global_replacements[f"{{{{ {k} }}}}"] = v
    replace_text_in_document_body(doc, global_replacements)
    
    # 1. テンプレートとなる「時間段落」と「データ表」を探す
    template_time_para = None
    template_data_table = None
    
    # 段落を走査
    para_index = -1
    for i, p in enumerate(doc.paragraphs):
        if "{{ time }}" in p.text:
            template_time_para = p
            para_index = i
            break
            
    # その段落より後にある最初の表を探す
    if template_time_para:
        # python-docxでは段落と表が混在する順序を正確に追うのが難しい場合があるが、
        # document.element.body の子要素を順に見ていくのが確実。
        
        body_elements = doc._body._element.getchildren() # 全要素（段落、表など）
        
        found_time = False
        target_p_xml = template_time_para._p
        target_tbl_xml = None
        
        # XML要素レベルで検索
        for elem in body_elements:
            if elem == target_p_xml:
                found_time = True
                continue
            
            if found_time and elem.tag.endswith('tbl'):
                # 時間段落の後に最初に見つかった表
                # 中身にタグがあるか確認（念のため）
                if "{{ s.name }}" in elem.xml or "{{ s.no }}" in elem.xml: # 簡易チェック
                    target_tbl_xml = elem
                    break
        
        if target_tbl_xml:
            # テンプレート要素を確保（ディープコピー）
            template_p_copy = copy.deepcopy(target_p_xml)
            template_tbl_copy = copy.deepcopy(target_tbl_xml)
            
            # 元の要素をドキュメントから削除（XML操作）
            doc._body._element.remove(target_p_xml)
            doc._body._element.remove(target_tbl_xml)
            
            # ループ生成
            for group in groups:
                # 1. 時間段落の追加
                new_p_xml = copy.deepcopy(template_p_copy)
                doc._body._element.append(new_p_xml)
                new_para = Paragraph(new_p_xml, doc._body)
                
                raw_time = group['time_str']
                formatted_time = format_time_label(raw_time)
                replace_text_in_paragraph_merged(new_para, {'{{ time }}': formatted_time})
                
                # 2. データ表の追加（メンバー分行を増やす処理含む）
                # まず表の枠を追加
                new_tbl_xml = copy.deepcopy(template_tbl_copy)
                doc._body._element.append(new_tbl_xml)
                
                # 追加された表オブジェクトを取得（再構築）
                # doc.tables はキャッシュされている可能性があるが、末尾に追加したので最後の表を取得
                new_table = doc.tables[-1] 
                
                # この表の中のデータ行テンプレートを探す
                data_row_template = None
                for row in new_table.rows:
                    row_text = "".join([c.text for c in row.cells])
                    if "{{ s.no }}" in row_text or "{{ s.name }}" in row_text:
                        data_row_template = row
                        break
                
                if data_row_template:
                    tbl_inner = new_table._tbl
                    tr_template = data_row_template._tr
                    tbl_inner.remove(tr_template) # テンプレート行を削除
                    
                    target_members = resolve_participants_from_string(group['member_input'], all_data)
                    
                    for member in target_members:
                        new_tr = copy.deepcopy(tr_template)
                        tbl_inner.append(new_tr)
                        new_row = new_table.rows[-1]
                        
                        replacements = {
                            '{{ s.no }}': member['no'],
                            '{{ s.name }}': member['name'],
                            '{{ s.kana }}': member.get('kana', ''),
                            '{{ s.age }}': member.get('age', ''),
                            '{{ s.tel }}': member.get('tel', ''),
                            '{{ s.song }}': member['song'],
                        }
                        fill_row_data(new_row, replacements)

    output_buffer = io.BytesIO()
    doc.save(output_buffer)
    return output_buffer


def generate_judges_list_doc(template_path_or_file, judges_list, global_context):
    doc = Document(template_path_or_file)
    global_replacements = {}
    for k, v in global_context.items():
        global_replacements[f"{{{{ {k} }}}}"] = v
    replace_text_in_document_body(doc, global_replacements)

    # パターン1: 表
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

    # パターン2: 段落
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
            replace_text_in_paragraph_merged(new_para, {'{{ judge_name }}': judge})

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

                    # 3. WEBプログラム生成（専用ロジック）
                    if web_template_path:
                        try:
                            if hasattr(web_template_path, 'seek'): web_template_path.seek(0)
                            context = base_context.copy()
                            context['judge_name'] = ''
                            # ★ここで新設した関数を呼ぶ
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
