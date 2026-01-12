import streamlit as st
import pandas as pd
import io
import zipfile
import json
import smtplib
import re
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
    """
    「2分30秒」「4分」などの文字列を秒数(int)に変換する
    """
    if not time_str:
        return 0
    s = str(time_str)
    # 分と秒を抽出 (全角半角対応)
    minutes = re.search(r'(\d+)\s*[分m]', s)
    seconds = re.search(r'(\d+)\s*[秒s]', s)
    
    total_sec = 0
    if minutes:
        total_sec += int(minutes.group(1)) * 60
    if seconds:
        total_sec += int(seconds.group(1))
    
    # "分"などの単位がなく数字だけの入力などのエッジケース対応が必要ならここに追加
    # 今回はフォーマットが決まっている前提とする
    return total_sec

def format_seconds_to_jp_label(total_seconds):
    """
    秒数を「X時間Y分」または「Y分」形式に変換
    ルール: 秒数が30秒以上なら繰り上げ
    """
    if total_seconds <= 0:
        return "0分"
    
    # 繰り上げ判定
    minutes = total_seconds // 60
    remainder_seconds = total_seconds % 60
    
    if remainder_seconds >= 30:
        minutes += 1
        
    # 時間・分 表記作成
    h = minutes // 60
    m = minutes % 60
    
    if h > 0:
        return f"{h}時間{m}分"
    else:
        return f"{m}分"

def format_time_label(text):
    """
    入力された時間文字列（例: 13:00-14:10）を「13時00分～14時10分」に変換
    """
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
    """
    入力文字列（例: "A01-A03, B05, C01-C02"）を解析し、
    all_data_listの中から該当する辞書データのリストを返す。
    ※ all_data_listはExcelの並び順で格納されている前提
    """
    if not input_str:
        return []

    # 全データから {出場番号: (index, data)} のマップを作成
    id_map = {str(item['no']): i for i, item in enumerate(all_data_list)}
    
    resolved_members = []
    
    # カンマで区切る
    parts = [p.strip() for p in input_str.replace('、', ',').split(',')]
    
    for part in parts:
        if not part:
            continue
            
        # 範囲指定 (ハイフン) の場合
        if '-' in part:
            range_parts = part.split('-')
            if len(range_parts) == 2:
                start_id = range_parts[0].strip()
                end_id = range_parts[1].strip()
                
                # Excelリスト上のインデックスを取得
                if start_id in id_map and end_id in id_map:
                    s_idx = id_map[start_id]
                    e_idx = id_map[end_id]
                    
                    # 順序が逆なら入れ替え
                    if s_idx > e_idx:
                        s_idx, e_idx = e_idx, s_idx
                    
                    # リストからスライスで取得して追加
                    # インデックスベースなのでExcelの並び順で取得される
                    for i in range(s_idx, e_idx + 1):
                        resolved_members.append(all_data_list[i])
        else:
            # 単体指定の場合
            if part in id_map:
                idx = id_map[part]
                resolved_members.append(all_data_list[idx])
                
    # 重複排除はせず、指定された順（範囲指定内はExcel順）に出す仕様とする
    return resolved_members

# --- Word操作系 (変更なし・微調整) ---

def copy_table_row(table, row):
    tbl = table._tbl
    new_tr = copy.deepcopy(row._tr)
    tbl.append(new_tr)
    return table.rows[-1]

def replace_text_in_paragraph(paragraph, replacements):
    for key, value in replacements.items():
        if key in paragraph.text:
            # シンプルな置換（書式崩れ防止のためRun単位で確認推奨だが簡略化）
            # Run単位だと分割されてマッチしないことがあるため、
            # ここでは段落テキスト全体を置換してから、最初のRunに入れ直す手法をとる
            full_text = paragraph.text
            new_text = full_text.replace(key, str(value))
            
            # 既存のRunをクリアして再設定（書式は最初のRunのものを継承）
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

def generate_word_from_template(template_file, groups, all_data, global_context):
    doc = Document(template_file)
    
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
    
    # テンプレート構造チェック
    if len(table.rows) < 3:
        raise Exception("テンプレートの表は少なくとも3行（ヘッダー、時間行、データ行）必要です。")

    time_row_template = table.rows[1]
    data_row_template = table.rows[2]
    
    delete_row(table, 2)
    delete_row(table, 1)
    
    # グループごとに処理
    for group in groups:
        # 1. 時間行を追加
        new_time_row = copy_table_row(table, time_row_template)
        raw_time = group['time_str']
        formatted_time = format_time_label(raw_time)
        fill_row_data(new_time_row, {'{{ time }}': formatted_time})

        # 2. メンバー解決（入力文字列からリスト化）
        target_members = resolve_participants_from_string(group['member_input'], all_data)
        
        # 3. データ行を追加
        for member in target_members:
            new_data_row = copy_table_row(table, data_row_template)
            replacements = {
                '{{ s.no }}': member['no'],
                '{{ s.name }}': member['name'],
                '{{ s.kana }}': member.get('kana', ''),
                '{{ s.age }}': member.get('age', ''),
                '{{ s.song }}': member['song'],
            }
            fill_row_data(new_data_row, replacements)

    output_buffer = io.BytesIO()
    doc.save(output_buffer)
    return output_buffer

# ---------------------------------------------------------
# 3. メール送信機能 (既存維持)
# ---------------------------------------------------------
def send_email_with_attachment(zip_buffer, zip_filename, contest_name):
    # (省略なしで実装可能ですが、既存と同じなのでここでは枠のみ記載)
    try:
        if "email" not in st.secrets:
             return False, "Secretsにメール設定がありません。"
        # ... (既存の送信コード) ...
        # 本番環境ではここに既存コードが入ります
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
            # 古い形式(start_noなど)が含まれている場合の互換性は考慮しつつ上書き
            st.session_state.update(config_data)
            st.success("設定を復元しました")

    # --- 1. Excelアップロード ---
    st.header("1. 名簿データ (Excel)")
    uploaded_excel = st.file_uploader("名簿Excelファイルをアップロード", type=['xlsx', 'xls', 'csv'])
    
    all_data = [] # グローバルスコープ初期化
    
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
            col_kana = c3.selectbox("フリガナ (任意)", ["(なし)"] + cols, index=default_kana + 1) # (なし)でズレるため調整
            
            col_song = c4.selectbox("演奏曲目", cols, index=cols.index("演奏曲目") if "演奏曲目" in cols else 0)
            
            # 追加オプション列
            c5, c6 = st.columns(2)
            default_age = cols.index("年齢") if "年齢" in cols else -1
            col_age = c5.selectbox("年齢列 (任意)", ["(なし)"] + cols, index=default_age + 1)
            
            # ★演奏時間列の追加
            default_dur = cols.index("演奏時間") if "演奏時間" in cols else -1
            col_duration = c6.selectbox("演奏時間列 (自動計算用)", ["(なし)"] + cols, index=default_dur + 1)

            st.markdown("---")

            # データ変換
            for _, row in df.iterrows():
                # 任意の列処理
                kana_val = str(row[col_kana]) if col_kana != "(なし)" else ""
                age_val = str(row[col_age]) if col_age != "(なし)" else ""
                
                # 演奏時間（秒換算）
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
                    'duration_sec': dur_seconds  # 計算用
                })
            
            st.write(f"読み込み完了: {len(all_data)} 件のデータ")

            # --- 2. テンプレートアップロード ---
            st.header("2. Wordテンプレート")
            uploaded_template = st.file_uploader("Wordテンプレート (.docx)", type=['docx'])

            # --- 3. グループ・スケジュール設定 ---
            st.header("3. グループ・スケジュール設定")
            
            # セッションステート初期化
            if 'groups' not in st.session_state:
                # 構造変更: member_input で管理
                st.session_state['groups'] = [{'member_input': '', 'time_str': '13:00-14:10'}]
            
            # --- グループ操作用コールバック関数 ---
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

            # グループ追加ボタン
            st.button("＋ グループ追加", on_click=add_group)

            # グループリスト描画
            for i, grp in enumerate(st.session_state['groups']):
                # レイアウト: [操作ボタン(極小)] [番号入力(大)] [合計時間(表示のみ)] [時間入力(中)] [削除]
                c_sort, c_input, c_total, c_time, c_del = st.columns([0.8, 3, 1.2, 2, 0.5])
                
                # 並べ替えボタン
                with c_sort:
                    if st.button("▲", key=f"up_{i}"):
                        move_group_up(i)
                        st.rerun()
                    if st.button("▼", key=f"down_{i}"):
                        move_group_down(i)
                        st.rerun()

                # メンバー指定入力
                input_val = c_input.text_input(
                    f"グループ {i+1} 対象番号 (例: A01-A05, C01)",
                    value=grp['member_input'],
                    key=f"g_in_{i}"
                )
                st.session_state['groups'][i]['member_input'] = input_val # 値更新

                # 合計時間計算 & 表示
                # 入力された文字列を解析してメンバー特定 -> 時間合算
                current_members = resolve_participants_from_string(input_val, all_data)
                total_sec = sum(m['duration_sec'] for m in current_members)
                time_display = format_seconds_to_jp_label(total_sec)
                
                c_total.info(f"計: {time_display}") # 色付きボックスで表示

                # 時間設定入力
                time_val = c_time.text_input(
                    "時間 (例: 13:00-14:00)",
                    value=grp['time_str'],
                    key=f"g_time_{i}"
                )
                st.session_state['groups'][i]['time_str'] = time_val

                # 削除ボタン
                with c_del:
                    if st.button("×", key=f"del_{i}"):
                        remove_group(i)
                        st.rerun()

            # --- 4. 審査員設定 ---
            st.header("4. 審査員設定")
            if 'judges' not in st.session_state:
                st.session_state['judges'] = ["審査員A"]
            
            if st.button("＋ 審査員追加"):
                st.session_state['judges'].append("")
                st.rerun() # 追加を即時反映

            # 修正: 入力ループ内では削除・フィルタリングを行わない
            # enumerateで回して、インデックスを使ってsession_stateを直接更新する
            for i in range(len(st.session_state['judges'])):
                val = st.text_input(f"審査員 {i+1}", value=st.session_state['judges'][i], key=f"judge_input_{i}")
                st.session_state['judges'][i] = val # 入力値を即座にステートに反映

            # コンクール名
            contest_name = st.text_input("コンクール名 (ファイル名・置換用)", "第10回BIPCA 東京予選④")

            # --- 5. 出力 ---
            if st.button("ファイル生成を実行", type="primary"):
                if not uploaded_template:
                    st.error("Wordテンプレートをアップロードしてください。")
                else:
                    # 最終的なデータのクレンジング（空の審査員削除など）
                    valid_judges = [j for j in st.session_state['judges'] if j.strip()]
                    
                    # 設定保存データの作成
                    config_json = json.dumps({
                        'groups': st.session_state['groups'],
                        'judges': valid_judges,
                        'contest_name': contest_name
                    }, ensure_ascii=False, indent=2)

                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                        
                        # 1. 採点表 (審査員ごと)
                        for judge in valid_judges:
                            uploaded_template.seek(0)
                            try:
                                context = {
                                    'contest_name': contest_name,
                                    'judge_name': judge
                                }
                                
                                doc_io = generate_word_from_template(
                                    uploaded_template, 
                                    st.session_state['groups'], 
                                    all_data, 
                                    context
                                )
                                zf.writestr(f"採点表_{judge}.docx", doc_io.getvalue())
                            except Exception as e:
                                st.error(f"採点表生成エラー ({judge}): {e}")

                        # 設定ファイル
                        zf.writestr("設定データ.json", config_json)
                    
                    st.success("生成完了！")
                    
                    # ダウンロードボタン
                    st.download_button(
                        "ZIPファイルをダウンロード",
                        zip_buffer.getvalue(),
                        f"{contest_name}.zip",
                        "application/zip"
                    )

        except Exception as e:
            st.error(f"エラーが発生しました: {e}")
            st.write(e) # デバッグ用詳細表示

if __name__ == "__main__":
    main()
