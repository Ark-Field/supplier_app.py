import io
import datetime
import pandas as pd
import streamlit as st

# --------------------------------------------------
# 1. ページ初期設定
# --------------------------------------------------
st.set_page_config(
    page_title="納期確認リスト・手配先別グループ分けツール",
    layout="wide"
)

st.title("📄 納期確認リスト 手配先別グループ分け ＆ 依頼書作成ツール")
st.write("エクセルを読み込み、手配先別にグループ分け。日付項目（受注日・指定納期・納品予定）を `****/**/**` 形式に整えてExcelファイルを作成します。")

# --------------------------------------------------
# 2. サイドバー：出力条件の設定（タイトル・返信期限）
# --------------------------------------------------
st.sidebar.header("📝 依頼文言・期限の設定")
default_title = st.sidebar.text_input("タイトル（ヘッダー文言）", value="【ご依頼】納期確認および回答のお願い")
deadline_date = st.sidebar.date_input("返信期限", value=datetime.date.today() + datetime.timedelta(days=7))
deadline_str = f"{deadline_date.year}年{deadline_date.month}月{deadline_date.day}日"

# --------------------------------------------------
# 3. メイン画面：ファイルアップロード ＆ ヘッダー行設定
# --------------------------------------------------
uploaded_file = st.file_uploader("📂 納期確認リストのエクセルファイル (.xlsx) をアップロードしてください", type=["xlsx"])

if uploaded_file is not None:
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
        
        col_s1, col_s2 = st.columns(2)
        with col_s1:
            selected_sheet = st.selectbox("読み込むシートを選択してください", sheet_names) if len(sheet_names) > 1 else sheet_names[0]
        with col_s2:
            # 3行目をヘッダーとするためインデックス2を初期値に設定
            header_row = st.number_input("ヘッダー（列名）がある行番号（0始まり: 2 = 3行目）", value=2, min_value=0, step=1)

        df_excel = pd.read_excel(uploaded_file, sheet_name=selected_sheet, header=header_row)
        
        # 列名の重複・空欄対策
        new_columns = []
        seen_cols = {}
        for idx, col in enumerate(df_excel.columns):
            col_str = str(col).strip()
            if col_str == "" or col_str.lower() == "nan" or col_str.startswith("Unnamed"):
                col_str = f"予備_{idx}"
            if col_str in seen_cols:
                seen_cols[col_str] += 1
                col_str = f"{col_str}_{seen_cols[col_str]}"
            else:
                seen_cols[col_str] = 0
            new_columns.append(col_str)
            
        df_excel.columns = new_columns
        df_excel.columns = df_excel.columns.str.replace(r'[\r\n\s]', '', regex=True)
        
        # --- 日付列（受注日、指定納期、納品予定）のフォーマット変換 (****/**/** 形式) ---
        date_cols = ["受注日", "指定納期", "納品予定"]
        for d_col in date_cols:
            if d_col in df_excel.columns:
                # 日付型やタイムスタンプ、文字列をパースして YYYY/MM/DD に変換
                parsed_dates = pd.to_datetime(df_excel[d_col], errors='coerce')
                df_excel[d_col] = parsed_dates.dt.strftime('%Y/%m/%d')
                # パースできなかった（もともと空欄や別の文字だった）ものは元の値（または空欄）に戻す
                df_excel[d_col] = df_excel[d_col].fillna(df_excel[d_col].astype(str).replace(["NaT", "nan", "None"], ""))
        
        # --- 出力対象列を「納品予定」までに絞り込む ---
        target_columns = [
            "予備_0", "手配先", "受注№", "品番", "S番号", "部番", 
            "パネル名称", "注文番号", "図面番号", "数", "ユーザ材質", 
            "予定重量", "受注日", "指定納期", "納品予定"
        ]
        
        valid_target_cols = [col for col in target_columns if col in df_excel.columns]
        df_excel = df_excel[valid_target_cols]
        
        st.success(f"エクセルファイルの読み込み成功（日付を ****/**/** 形式に整形し、出力列を絞り込みました）！")
        
        with st.expander("👀 読み込んだエクセルデータのプレビュー（最初の5行）", expanded=True):
            st.dataframe(df_excel.head(), use_container_width=True)
            
        st.markdown("---")
        
        # グループ分け基準の列を選択
        available_cols = list(df_excel.columns)
        default_index = 0
        for name in ["手配先", "手配先名", "発注先", "発注先名", "手配先番号"]:
            if name in available_cols:
                default_index = available_cols.index(name)
                break
                
        selected_group_col = st.selectbox("🔑 グループ分けの基準にする列", available_cols, index=default_index)
        
        st.markdown("---")
        
        if st.button("🚀 手配先別にグループ分けを実行する", type="primary"):
            df_work = df_excel.copy()
            df_work[selected_group_col] = df_work[selected_group_col].astype(str).str.strip()
            df_work.loc[df_work[selected_group_col] == "", selected_group_col] = "（手配先未設定）"
            df_work.loc[df_work[selected_group_col].str.lower().isin(["nan", "none", "nat"]), selected_group_col] = "（手配先未設定）"
            
            groups = dict(list(df_work.groupby(selected_group_col)))
            st.session_state["groups"] = groups
            st.session_state["group_col"] = selected_group_col
            
            st.success(f"グループ分け完了！合計 **{len(groups)}** 個の手配先に分かれました。")

        if "groups" in st.session_state:
            groups = st.session_state["groups"]
            
            st.markdown("---")
            st.markdown("### 📊 手配先別グループ一覧・ファイルダウンロード")
            st.info(f"📌 設定中の返信期限: **{deadline_str}** まで")
            
            tab_names = [str(name) for name in groups.keys()]
            tabs = st.tabs(tab_names)
            
            for tab, (supp_name, group_df) in zip(tabs, groups.items()):
                with tab:
                    st.write(f"**手配先:** {supp_name} （対象件数: {len(group_df)}件）")
                    st.dataframe(group_df, use_container_width=True)
                    
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        group_df.to_excel(writer, sheet_name="納期確認", index=False, startrow=3)
                        
                        worksheet = writer.sheets["納期確認"]
                        worksheet['A1'] = default_title
                        worksheet['A2'] = f"【返信期限】 {deadline_str} までにご回答をお願いいたします。"
                        worksheet['A3'] = f"【手配先】 {supp_name}"
                        
                    excel_data = output.getvalue()
                    
                    st.download_button(
                        label=f"⬇️ 「{supp_name}」のExcelファイルをダウンロード",
                        data=excel_data,
                        file_name=f"納期確認依頼_{supp_name}_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_excel_{supp_name}"
                    )
                    
    except Exception as e:
        st.error(f"ファイルの処理中にエラーが発生しました: {e}")