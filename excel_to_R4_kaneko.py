import streamlit as st
import pandas as pd
from io import BytesIO

def app1():
    st.title("金子宝泉堂")

    with st.expander('🐴前提事項'):
        st.write('①消費税コードは、軽減税率に◯がある場合にのみ32(課税対応仕入)が反映される。')
        st.write('※科目コードがない場合は、R4側の科目設定に基づいてインポートする仕様になっているので把握しておいた方がいいです。')
        st.write('②インボイスは経過措置期間に応じた設計になっていない。将来的には注意が必要。')
    with st.expander('🐸変更点'):
        st.write('①Excelの種類ごとの選択項目に変更しました。')
        st.write('②共通の部門を99から4に変更しました。 ※本部経費に◯がある場合に4が設定されます。')
        st.write('③本部経費に◯がない場合に①で選択した部門コードを設定するように変更しました。')
        st.write('※BS項目にも部門コードが設定されます。不要であればR4側の設定を変更して下さい。')
        st.write('④補助科目が反映されるように変更しました。')
        st.write('⑤確認用にプレビューを追加しました。')
        st.write('⑥レイアウトを左寄りに変更しました。')

    uploaded_file = st.file_uploader("Excelファイルをアップロードしてください", type=["xlsx"])

    if uploaded_file:
        dfs = pd.read_excel(uploaded_file, sheet_name=None, header=1)
        sheet_names = list(dfs.keys())
        selected_sheet = st.selectbox("処理対象のシートを選択してください", sheet_names)

        account_options = {
            "現金-文具": '1',
            "現金-表具①（城南)": '2',
            "現金-表具②（きらぼし)": '3'
        }
        selected_default = st.selectbox("選択してください", list(account_options.keys()))
        default_value = account_options[selected_default]

        if "科目マスタ" in dfs:
            st.subheader("科目マスタ プレビュー")
            df_master = dfs["科目マスタ"]
            st.dataframe(df_master)
        else:
            st.warning("科目マスタシートが見つかりません。")

        if st.button("処理を実行する"):
            df_selectsheet = dfs[selected_sheet]
            df_selectsheet = df_selectsheet[df_selectsheet['日'].notna() & (df_selectsheet['日'] != '')]
            df_master = dfs['科目マスタ']

            st.subheader("選択したシート プレビュー")
            st.dataframe(df_selectsheet)

            output_columns = [
                "月種別", "種類", "形式", "作成方法", "付箋", "伝票日付", "伝票番号", "伝票摘要", "枝番", 
                "借方部門", "借方部門名", "借方科目", "借方科目名", "借方補助", "借方補助科目名", "借方金額", 
                "借方消費税コード", "借方消費税業種", "借方消費税税率", "借方資金区分", "借方任意項目１", 
                "借方任意項目２", "借方インボイス情報", "貸方部門", "貸方部門名", "貸方科目", "貸方科目名", 
                "貸方補助", "貸方補助科目名", "貸方金額", "貸方消費税コード", "貸方消費税業種", "貸方消費税税率", 
                "貸方資金区分", "貸方任意項目１", "貸方任意項目２", "貸方インボイス情報", "摘要", "期日", "証番号", 
                "入力マシン", "入力ユーザ", "入力アプリ", "入力会社", "入力日付", "コメント"
            ]
            output_df = pd.DataFrame(index=df_selectsheet.index, columns=output_columns)

            df_selectsheet[['年', '月', '日']] = df_selectsheet[['年', '月', '日']].astype(int)
            df_selectsheet['伝票日付'] = (
                df_selectsheet['年'].astype(str) +
                df_selectsheet['月'].apply(lambda x: f"{x:02}") +
                df_selectsheet['日'].apply(lambda x: f"{x:02}")
            )
            output_df['伝票日付'] = df_selectsheet['伝票日付']

            df_selectsheet['借方金額'] = df_selectsheet[['入金', '出金']].sum(axis=1, skipna=True)
            df_selectsheet['貸方金額'] = df_selectsheet['借方金額']
            output_df['借方金額'] = df_selectsheet['借方金額'].astype(int)
            output_df['貸方金額'] = df_selectsheet['貸方金額'].astype(int)

            output_df['摘要'] = df_selectsheet['摘要']

            sales_account_dict = pd.Series(df_master['入金科目'].values, index=df_master['入金科目一覧']).to_dict()
            df_selectsheet['貸方科目'] = df_selectsheet.apply(
                lambda row: int(sales_account_dict.get(row['入金科目'])) if pd.notna(row['入金科目']) else None,
                axis=1
            )
            output_df['貸方科目'] = df_selectsheet['貸方科目'].fillna(100).astype(int)

            sales_subaccount_dict = pd.Series(df_master['入金補助'].values, index=df_master['入金科目一覧']).to_dict()
            output_df['貸方補助'] = df_selectsheet.apply(
                lambda row: sales_subaccount_dict.get(row['入金科目']) if pd.notna(row['入金科目']) else default_value,
                axis=1
            )

            expense_account_dict = pd.Series(df_master['出金科目'].values, index=df_master['支出科目一覧']).to_dict()
            df_selectsheet['借方科目'] = df_selectsheet.apply(
                lambda row: int(expense_account_dict.get(row['出金科目'])) if pd.notna(row['出金科目']) else None,
                axis=1
            )
            output_df['借方科目'] = df_selectsheet['借方科目'].fillna(100).astype(int)

            expense_subaccount_dict = pd.Series(df_master['出金補助'].values, index=df_master['支出科目一覧']).to_dict()
            output_df['借方補助'] = df_selectsheet.apply(
                lambda row: expense_subaccount_dict.get(row['出金科目']) if pd.notna(row['出金科目']) else default_value,
                axis=1
            )

            df_selectsheet['借方消費税コード'] = df_selectsheet['軽減税率'].apply(lambda x: '32' if x == '○' else None)
            df_selectsheet['借方消費税税率'] = df_selectsheet['軽減税率'].apply(lambda x: 'K8' if x == '○' else None)
            output_df['借方消費税コード'] = df_selectsheet['借方消費税コード']
            output_df['借方消費税税率'] = df_selectsheet['借方消費税税率']

            df_selectsheet['借方インボイス情報'] = df_selectsheet['ｲﾝﾎﾞｲｽ'].apply(lambda x: 8 if x == '○' else None)
            output_df['借方インボイス情報'] = df_selectsheet['借方インボイス情報']

            df_selectsheet['借方部門'] = df_selectsheet['本部経費'].apply(lambda x: 4 if x == '○' else default_value)
            output_df['借方部門'] = df_selectsheet['借方部門']
            df_selectsheet['貸方部門'] = df_selectsheet['本部経費'].apply(lambda x: 4 if x == '○' else default_value)
            output_df['貸方部門'] = df_selectsheet['貸方部門']

            output_df['借方補助'] = output_df['借方補助'].fillna(0).astype(int)
            output_df['貸方補助'] = output_df['貸方補助'].fillna(0).astype(int)

            st.subheader('R4形式仕訳データ プレビュー')
            st.dataframe(output_df)

            csv_buffer = BytesIO()
            output_df.to_csv(csv_buffer, encoding='cp932', index=False)
            csv_buffer.seek(0)
            st.download_button(label="CSVダウンロード", data=csv_buffer, file_name="仕訳データ.csv", mime="text/csv")

