import pandas as pd
import streamlit as st
import os
import plotly.express as px
import numpy as np
import argparse
import sys

# --- Streamlit ページ設定 (必ず最初に実行) ---
# Streamlitのコードが実行されるとき用のページ設定
if 'streamlit' in sys.modules and hasattr(st, 'session_state'):
    st.set_page_config(layout="wide", page_title="工数分析ビューア")

# --- 仕様に基づく変更点 (Stream Cloud対応) ---
# 1. ローカルファイル関連機能の調整
# 2. 例外処理と表示の強化
# 3. ファイルアップロードUIの改善
# 4. CUI実行時のヘルプ表示の追加
# -----------------------------------------

# --- CUIモード用のコマンドライン引数解析 ---
def parse_arguments():
    parser = argparse.ArgumentParser(description='工数分析ビューア')
    parser.add_argument('--file', '-f', type=str, help='分析するExcelファイルのパス')
    parser.add_argument('--decimal', '-d', type=int, default=2, help='小数点以下の桁数（デフォルト: 2）')
    return parser.parse_args()

# --- 定数定義 ---
BLANK_STR = "[空白]"
BASE_COLUMN_ORDER = [
    "USER_FIELD_01", "USER_FIELD_02", "USER_FIELD_03",
    "業務内容1", "業務内容2", "業務内容3", "業務内容4", "業務内容5"
]
UNIT_COL = "UNIT"
EFFORT_COL = "作業時間(h)"
DEFAULT_LOCAL_FILE = "merged_efforts.xlsx"

# --- メイン関数 ---
def main():
    # CUI実行時の処理
    if not _is_streamlit_running():
        args = parse_arguments()
        if len(sys.argv) == 1:
            print("Usage: python effort_analysis_viewer.py --file <エクセルファイルパス> [--decimal <小数点桁数>]")
            print("または: streamlit run effort_analysis_viewer.py")
            sys.exit(0)
        
        if args.file:
            print(f"ファイル '{args.file}' を分析中...")
            try:
                analyze_file(args.file, decimal_places=args.decimal)
            except Exception as e:
                print(f"エラー: {e}")
                sys.exit(1)
        else:
            print("ファイルが指定されていません。")
            sys.exit(1)
    else:
        # Streamlit UI
        run_streamlit_app()

# --- CUIモード実行時の分析処理 ---
def analyze_file(file_path, decimal_places=2):
    """コマンドライン実行時のファイル分析"""
    try:
        excel_data = pd.read_excel(file_path, sheet_name=None)
        df = pd.concat(excel_data.values(), ignore_index=True)
        df.columns = df.columns.str.strip()
        
        # 集計処理の簡易表示
        group_cols = [col for col in BASE_COLUMN_ORDER if col in df.columns]
        if UNIT_COL in df.columns and UNIT_COL not in group_cols:
            group_cols.append(UNIT_COL)
        
        if EFFORT_COL in df.columns:
            result = df.groupby(group_cols, dropna=False)[EFFORT_COL].sum().reset_index()
            result = result.sort_values(by=EFFORT_COL, ascending=False)
            
            # 表示用のフォーマット
            for col in group_cols:
                if result[col].isnull().any():
                    result[col] = result[col].fillna(BLANK_STR)
            
            result[EFFORT_COL] = result[EFFORT_COL].apply(lambda x: f"{x:.{decimal_places}f}")
            
            # 結果出力
            print(f"\n集計結果 (上位10件):")
            print("-" * 80)
            print(result.head(10).to_string(index=False))
            print("-" * 80)
            print(f"合計作業時間: {df[EFFORT_COL].sum():.{decimal_places}f} h")
            print(f"データ行数: {len(df)}")
            print("\n分析完了。詳細な分析とグラフ表示には Streamlit インターフェースをご利用ください。")
        else:
            print(f"エラー: '{EFFORT_COL}' 列がデータ内に見つかりません。")
    except Exception as e:
        print(f"ファイル分析中にエラーが発生しました: {e}")
        raise

# --- Streamlit実行確認関数 ---
def _is_streamlit_running():
    """Streamlitフレームワークで実行中かどうかを確認"""
    return 'streamlit' in sys.modules and hasattr(st, 'session_state')

# --- データ読み込みキャッシュ関数 ---
@st.cache_data
def load_data(source):
    is_uploaded_file = hasattr(source, 'seek')
    try:
        excel_data = pd.read_excel(source, sheet_name=None)
        df = pd.concat(excel_data.values(), ignore_index=True)
        df.columns = df.columns.str.strip()
        st.success(f"{'アップロードファイル' if is_uploaded_file else 'ローカルファイル'}が正常に読み込まれました。")
        return df
    except FileNotFoundError:
        st.error(f"ファイルが見つかりません: {source}")
        return None
    except Exception as e:
        st.error(f"データ読み込み中にエラーが発生しました: {e}")
        st.error(f"詳細: {str(e)}")
        return None

# --- Streamlitアプリケーション本体 ---
def run_streamlit_app():
    # --- Streamlit ページタイトル ---
    st.title("工数分析ビューア")
    
    with st.expander("📋 使い方"):
        st.markdown("""
        ### 使い方
        1. Excel形式の工数データをアップロードしてください
        2. サイドバーのフィルター機能で分析したいデータを絞り込めます
        3. 集計結果はテーブルとグラフで表示されます
        
        ### データ要件
        - 複数のワークシートを含むExcelファイル
        - 必須列: `作業時間(h)` および `USER_FIELD_01`～`USER_FIELD_03`、`業務内容1`～`業務内容5` など
        
        ### ファイル形式
        - `.xlsx` または `.xls` 形式のExcelファイル
        """)

    # ファイルアップロード
    uploaded_file = st.file_uploader(
        "分析するExcelファイルをアップロード", 
        type=["xlsx", "xls"],
        help="複数のワークシートを含むExcelファイルをアップロードしてください。"
    )
    
    # Stream Cloud環境ではローカルファイルチェックを調整
    use_local = False
    local_file_exists = os.path.exists(DEFAULT_LOCAL_FILE)
    
    # Stream Cloud対応: セッション状態を利用した処理
    if 'show_local_option' not in st.session_state:
        st.session_state.show_local_option = local_file_exists
    
    if st.session_state.show_local_option:
        use_local = st.checkbox(
            f"ローカルファイルを使用する ({DEFAULT_LOCAL_FILE})",
            value=False,
            disabled=(uploaded_file is not None),
            help="ローカル環境時のみ有効。Stream Cloudでは機能しません。"
        )
    
    # データソース決定とロード
    df_source = None
    load_error = False
    
    if use_local and local_file_exists:
        df_source = load_data(DEFAULT_LOCAL_FILE)
        if df_source is None:
            load_error = True
    elif uploaded_file is not None:
        df_source = load_data(uploaded_file)
        if df_source is None:
            load_error = True
    
    # データがなければ処理中断
    if df_source is None:
        if not load_error:
            st.info("分析するファイルをアップロードしてください。")
            st.stop()
        else:
            st.stop()

    # --- データ情報表示 ---
    with st.expander("📊 データ概要", expanded=False):
        st.markdown(f"**行数**: {len(df_source)}")
        st.markdown(f"**列数**: {df_source.shape[1]}")
        available_cols = [col for col in df_source.columns]
        st.markdown(f"**主要列**: {', '.join([col for col in available_cols if col in BASE_COLUMN_ORDER + [UNIT_COL, EFFORT_COL]])}")

    # --- フィルター処理 (サイドバー) ---
    st.sidebar.header("フィルター条件")

    available_base_columns = [col for col in BASE_COLUMN_ORDER if col in df_source.columns]
    unit_col_exists_in_source = UNIT_COL in df_source.columns

    filtered_df = df_source.copy()
    applied_filters = {}

    # 基本列のカスケードフィルター
    for col in available_base_columns:
        if col not in filtered_df.columns:
            continue
        
        options_with_blank = filtered_df[col].fillna(BLANK_STR).unique().tolist()
        try:
            options_with_blank.sort(key=lambda x: str(x))
        except TypeError:
            pass
        
        if not options_with_blank:
            continue

        selected = st.sidebar.multiselect(
            f"{col} で絞り込み", 
            options=options_with_blank, 
            key=f"filter_{col}", 
            default=[]
        )
        
        if selected:
            applied_filters[col] = selected
            filter_values_actual = [np.nan if v == BLANK_STR else v for v in selected]
            is_nan_selected = np.nan in filter_values_actual
            non_nan_values = [v for v in filter_values_actual if v is not np.nan]
            
            if col in filtered_df.columns:
                if is_nan_selected:
                    filtered_df = filtered_df[filtered_df[col].isin(non_nan_values) | filtered_df[col].isna()]
                else:
                    filtered_df = filtered_df[filtered_df[col].isin(non_nan_values)]

    # --- 「UNIT」フィルター（常に最後）---
    unit_filter_selected = []
    if unit_col_exists_in_source:
        st.sidebar.markdown("---")
        unit_options = df_source[UNIT_COL].fillna(BLANK_STR).unique().tolist()
        try:
            unit_options.sort(key=lambda x: str(x))
        except TypeError:
            pass

        unit_filter_selected = st.sidebar.multiselect(
            f"{UNIT_COL} で絞り込み (AND条件)", 
            options=unit_options, 
            key=f"filter_{UNIT_COL}", 
            default=[]
        )
        
        if unit_filter_selected:
            unit_filter_values_actual = [np.nan if v == BLANK_STR else v for v in unit_filter_selected]
            is_nan_selected_unit = np.nan in unit_filter_values_actual
            non_nan_unit_values = [v for v in unit_filter_values_actual if v is not np.nan]
            
            if UNIT_COL in filtered_df.columns:
                if is_nan_selected_unit:
                    filtered_df = filtered_df[filtered_df[UNIT_COL].isin(non_nan_unit_values) | filtered_df[UNIT_COL].isna()]
                else:
                    filtered_df = filtered_df[filtered_df[UNIT_COL].isin(non_nan_unit_values)]

    # --- 集計と表示 ---
    st.header("集計結果")
    
    # 集計結果の概要メトリクス
    metrics_cols = st.columns(3)
    with metrics_cols[0]:
        st.metric("データ総数", f"{len(filtered_df):,} 件")
    with metrics_cols[1]:
        if EFFORT_COL in filtered_df.columns:
            total_effort = filtered_df[EFFORT_COL].sum()
            st.metric("合計作業時間", f"{total_effort:.2f} h")
    with metrics_cols[2]:
        if len(applied_filters) > 0:
            st.metric("適用フィルター数", f"{len(applied_filters)} 件")
    
    # 集計キー決定
    group_cols = []
    last_selected_base_filter_index = -1
    
    if applied_filters:
        for col in reversed(available_base_columns):
            if col in applied_filters:
                try:
                    last_selected_base_filter_index = available_base_columns.index(col)
                    break
                except ValueError:
                    pass
        
        if last_selected_base_filter_index != -1:
            group_cols_candidate_indices = range(last_selected_base_filter_index + 1, len(available_base_columns))
            group_cols_candidate = [available_base_columns[i] for i in group_cols_candidate_indices]
            group_cols = [c for c in group_cols_candidate if c in filtered_df.columns]
    else:
        group_cols = [c for c in available_base_columns if c in filtered_df.columns]

    if unit_col_exists_in_source and UNIT_COL not in group_cols:
        if UNIT_COL in filtered_df.columns:
            group_cols.append(UNIT_COL)

    if EFFORT_COL not in filtered_df.columns:
        st.error(f"データに「{EFFORT_COL}」列が見つかりません。集計できません。")
        st.stop()

    if group_cols:
        try:
            result_df = filtered_df.groupby(group_cols, dropna=False, observed=True)[EFFORT_COL].sum().reset_index()
            for col in group_cols:
                if result_df[col].isnull().any():
                    try:
                        result_df[col] = result_df[col].astype(object).fillna(BLANK_STR)
                    except Exception:
                        result_df[col] = result_df[col].fillna(BLANK_STR)

            st.subheader("集計テーブル")
            col1, col2, col3 = st.columns([2, 2, 1])
            with col1:
                sort_options = result_df.columns.tolist()
                default_sort_col = EFFORT_COL if EFFORT_COL in sort_options else (sort_options[0] if sort_options else None)
                sort_column = st.selectbox(
                    "ソート列", 
                    sort_options, 
                    index=sort_options.index(default_sort_col) if default_sort_col and default_sort_col in sort_options else 0
                )
            with col2:
                sort_ascending = st.radio("ソート順", ["降順", "昇順"], index=0, horizontal=True) == "昇順"
            with col3:
                decimal_places = st.number_input("表示小数点桁数", 0, 4, 2)

            result_df_sorted = result_df
            if sort_column and not result_df.empty:
                try:
                    result_df_sorted = result_df.sort_values(
                        by=sort_column, 
                        ascending=sort_ascending,
                        key=lambda col: col.astype(str) if col.dtype == 'object' else col
                    )
                except Exception as e:
                    st.warning(f"ソート中に問題: {e}")

            final_columns = result_df_sorted.columns.tolist()
            if EFFORT_COL in final_columns:
                final_columns.remove(EFFORT_COL)
                if unit_col_exists_in_source and UNIT_COL in final_columns:
                    final_columns.remove(UNIT_COL)
                    final_columns.append(UNIT_COL)
                final_columns.append(EFFORT_COL)

            result_df_display = result_df_sorted[final_columns].copy()
            if EFFORT_COL in result_df_display.columns:
                result_df_display[EFFORT_COL] = result_df_display[EFFORT_COL].apply(lambda x: f"{x:.{decimal_places}f}")

            st.dataframe(result_df_display, use_container_width=True, hide_index=True)

            # --- CSV出力ボタン ---
            if not result_df_display.empty:
                csv_data = result_df_display.to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    label="CSVでダウンロード",
                    data=csv_data,
                    file_name="集計結果.csv",
                    mime="text/csv",
                )

            st.header("グラフ表示")
            col1_graph, col2_graph = st.columns(2)
            with col1_graph:
                graph_type = st.selectbox("グラフの種類", ["横棒グラフ", "縦棒グラフ"])
            with col2_graph:
                num_items_options = [10, 20, 50, 100, "すべて"]
                num_items_to_show = st.selectbox("表示件数", num_items_options, index=1)

            try:
                plot_df = result_df_sorted.copy()
                item_cols = [col for col in group_cols if col in plot_df.columns]
                if item_cols:
                    plot_df["項目"] = plot_df[item_cols].astype(str).agg(" / ".join, axis=1)
                else:
                    plot_df["項目"] = "合計"

                sort_direction = "上位" if not sort_ascending else "下位"
                if sort_column != EFFORT_COL:
                    sort_direction = "ソート順"

                if isinstance(num_items_to_show, int) and num_items_to_show < len(plot_df):
                    plot_df_n = plot_df.head(num_items_to_show)
                    graph_title = f"{sort_direction}{num_items_to_show}件の作業時間 ({' / '.join(item_cols)})"
                else:
                    plot_df_n = plot_df
                    graph_title = f"すべての作業時間 ({' / '.join(item_cols)})"

                if plot_df_n.empty or EFFORT_COL not in plot_df_n.columns:
                    st.info("グラフ表示対象のデータがありません。")
                else:
                    max_item_length = max(plot_df_n["項目"].astype(str).apply(len)) if not plot_df_n.empty else 10

                    if graph_type == "横棒グラフ":
                        plot_df_n_h = plot_df_n[::-1]
                        fig = px.bar(plot_df_n_h, x=EFFORT_COL, y="項目", orientation="h", title=graph_title)
                        fig.update_layout(
                            xaxis_side='top', 
                            xaxis_title="作業時間 [h]",
                            yaxis={"tickfont": {"size": 10}},
                            margin=dict(l=min(350, max(150, max_item_length * 7)), r=30, t=110, b=20),
                            height=max(400, 20 * len(plot_df_n_h)),
                        )
                        hover_template = '%{y}: %{x:.2f} h'
                    else:  # 縦棒グラフ
                        fig = px.bar(plot_df_n, x="項目", y=EFFORT_COL, title=graph_title)
                        fig.update_layout(
                            xaxis={
                                'categoryorder': 'array', 
                                'categoryarray': plot_df_n["項目"].tolist(), 
                                "tickfont": {"size": 10}
                            },
                            xaxis_tickangle=-45, 
                            yaxis_title="作業時間 [h]",
                            margin=dict(l=70, r=30, t=80, b=min(300, max(100, max_item_length * 6))),
                            height=max(500, 350 + min(300, max(100, max_item_length * 6))),
                        )
                        hover_template = '%{x}: %{y:.2f} h'

                    fig.update_layout(
                        font=dict(size=12), 
                        plot_bgcolor='rgba(240,240,240,0.5)', 
                        bargap=0.2, 
                        title_font=dict(size=16)
                    )
                    fig.update_traces(hovertemplate=hover_template)
                    st.plotly_chart(fig, use_container_width=True)

            except Exception as e:
                st.error(f"グラフ描画中にエラー: {e}")
                st.exception(e)

        except Exception as e:
            st.error(f"集計処理中にエラー: {e}")
            st.exception(e)
    else:
        # 集計対象列がない場合の処理
        if not filtered_df.empty and EFFORT_COL in filtered_df.columns:
            total_effort = filtered_df[EFFORT_COL].sum()
            st.metric("絞り込み結果の合計作業時間", f"{total_effort:.2f} h")
            
            # 表示列を動的に決定
            display_cols_fallback = [col for col in available_base_columns if col in filtered_df.columns]
            if unit_col_exists_in_source and UNIT_COL in filtered_df.columns:
                display_cols_fallback.append(UNIT_COL)
            if EFFORT_COL in filtered_df.columns:
                display_cols_fallback.append(EFFORT_COL)
            
            st.dataframe(filtered_df[display_cols_fallback], hide_index=True)
        else:
            st.info("フィルター条件に一致するデータがありません。")

    st.caption("分析完了")

# エントリーポイント
if __name__ == "__main__":
    main()

# エントリーポイント
if __name__ == "__main__":
    main()