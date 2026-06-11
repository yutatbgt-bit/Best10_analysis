import os
import sys
import glob
import argparse
import pandas as pd
import numpy as np
from openpyxl.styles import Font

def find_latest_file(pattern):
    """
    指定されたパターンにマッチするファイルの中から、最終更新日時が最も新しいファイルを返す。
    """
    files = glob.glob(pattern)
    if not files:
        return None
    return max(files, key=os.path.getmtime)

def main():
    # 日本語出力の文字化けを防ぐ
    if hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8')
        
    # コマンドライン引数の解析
    parser = argparse.ArgumentParser(description="ベスト10実績と今月売上の突合ツール")
    parser.add_argument(
        "--store", 
        "-s", 
        default="全店", 
        help="参照する基準Excelの店舗シート名（デフォルト: 全店）"
    )
    args = parser.parse_args()
    target_store = args.store.strip()
        
    input_dir = "input_data"
    output_dir = "output_data"
    
    # フォルダの自動作成
    os.makedirs(input_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)
        
    # input_data フォルダから最新のExcelファイルを自動検出
    best10_pattern = os.path.join(input_dir, "ベスト１０実績_*.xlsx")
    sales_pattern = os.path.join(input_dir, "販売管理表(単品別売上実績)-*.xlsx")
    
    best10_file = find_latest_file(best10_pattern)
    sales_file = find_latest_file(sales_pattern)
    
    if not best10_file:
        print(f"Error: '{input_dir}' フォルダ内に 'ベスト１０実績_*.xlsx' が見つかりません。")
        sys.exit(1)
    if not sales_file:
        print(f"Error: '{input_dir}' フォルダ内に '販売管理表(単品別売上実績)-*.xlsx' が見つかりません。")
        sys.exit(1)
        
    print(f"使用するベスト10基準ファイル: {best10_file}")
    print(f"使用する最新販売管理表ファイル: {sales_file}")
    
    # 指定された店舗シートが存在するかチェック
    try:
        xls_best = pd.ExcelFile(best10_file)
        if target_store not in xls_best.sheet_names:
            print(f"\n[ERROR] 指定された店舗シート '{target_store}' はExcelファイル内に存在しません。")
            # 「今期部門売上」などのシステム用シートを除いた、店舗名シートのみを案内するためにフィルタ
            invalid_sheets = ["今期部門売上", "今期売上実績", "コードなど", "前期ベスト10に対して今期順位", 
                              "今期ベスト10に対して前期順位", "前期部門売上", "前期売上実績　データ"]
            valid_stores = [s for s in xls_best.sheet_names if s not in invalid_sheets]
            print(f"利用可能な店舗名: {', '.join(valid_stores)}")
            sys.exit(1)
    except Exception as e:
        print(f"Excelファイルの検証中にエラーが発生しました: {e}")
        sys.exit(1)
        
    print(f"参照店舗（シート名）: {target_store}")
    
    # 1. 販売管理表のロードとマッピング用辞書の作成
    print("最新販売管理表を読み込み中...")
    df_sales_raw = pd.read_excel(sales_file, sheet_name="Sheet1", skiprows=6, header=None)

    current_sales_map = {}
    for idx, row in df_sales_raw.iterrows():
        code = row.iloc[0]
        name = row.iloc[1]
        
        if pd.isna(code) or str(code).strip() in ["合計", "総計"] or "対象期間" in str(code):
            continue
            
        try:
            if isinstance(code, float) and code.is_integer():
                code_str = str(int(code))
            else:
                code_str = str(code).strip().split('.')[0]
        except:
            code_str = str(code).strip()
            
        if not code_str.isdigit():
            continue
            
        sales_amt = pd.to_numeric(row.iloc[3], errors='coerce')
        ratio_val = pd.to_numeric(row.iloc[6], errors='coerce')
        
        current_sales_map[code_str] = {
            "sales": sales_amt if pd.notna(sales_amt) else 0,
            "ratio": ratio_val if pd.notna(ratio_val) else np.nan
        }

    print(f"販売管理表から {len(current_sales_map)} 件の商品をマッピングしました。")

    # 2. ベスト10実績の店舗シートのロードとパース
    print(f"ベスト10基準データ（{target_store}シート）を読み込み中...")
    df_base = pd.read_excel(best10_file, sheet_name=target_store)

    dept_data = {}
    dept_order = []
    current_dept = None

    for idx, row in df_base.iterrows():
        val0 = row.iloc[0]
        nan_count = row.isna().sum()
        
        # 部門名の判定
        if pd.notna(val0) and nan_count >= len(row) - 3:
            val0_str = str(val0).strip()
            if val0_str not in ["順位", "今期", "合計", "ベスト10小計"] and not val0_str.isdigit():
                current_dept = val0_str
                if current_dept not in dept_data:
                    dept_data[current_dept] = []
                    dept_order.append(current_dept)
                continue
                
        if current_dept is None:
            continue
            
        try:
            rank = int(val0)
            if 1 <= rank <= 10:
                r = row.tolist()
                
                code_str = ""
                if pd.notna(r[2]):
                    try:
                        if isinstance(r[2], float) and r[2].is_integer():
                            code_str = str(int(r[2]))
                        else:
                            code_str = str(r[2]).strip().split('.')[0]
                    except:
                        code_str = str(r[2]).strip()
                
                # 今月の実績値を取得
                current_sales = np.nan
                current_ratio = np.nan
                if code_str in current_sales_map:
                    current_sales = current_sales_map[code_str]["sales"]
                    current_ratio = current_sales_map[code_str]["ratio"]

                dept_data[current_dept].append({
                    "今期順位": rank,
                    "前期順位": pd.to_numeric(r[1], errors='coerce'),
                    "商品コード": code_str,
                    "商品名": str(r[3]).strip() if pd.notna(r[3]) else "",
                    "基準売上実績": pd.to_numeric(r[4], errors='coerce'),
                    "基準前年売上": pd.to_numeric(r[6], errors='coerce'),
                    "基準比較日比(%)": pd.to_numeric(r[8], errors='coerce'),
                    "今月売上実績": current_sales,
                    "今月比較日比(%)": current_ratio,
                    "基準打数実績": pd.to_numeric(r[9], errors='coerce'),
                    "基準単価実績": pd.to_numeric(r[13], errors='coerce')
                })
        except (ValueError, TypeError):
            pass

    # 3. ExcelWriter を使用して複数シートのExcelファイルを作成
    # ファイル名に店舗名を動的に含める (例: best10_with_current_sales_王子店.xlsx)
    output_excel = os.path.join(output_dir, f"best10_with_current_sales_{target_store}.xlsx")
    print(f"結果Excelファイルを作成中: {output_excel}...")
    
    red_bold_font = Font(color="FF0000", bold=True)
    
    with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
        for dept in dept_order:
            items = dept_data[dept]
            if not items:
                continue
            df_dept = pd.DataFrame(items)
            
            # シート名の文字数制限への対応
            sheet_name = dept
            for char in [':', '\\', '/', '?', '*', '[', ']']:
                sheet_name = sheet_name.replace(char, '_')
            sheet_name = sheet_name[:31]
            
            df_dept.to_excel(writer, sheet_name=sheet_name, index=False)
            
            worksheet = writer.sheets[sheet_name]
            
            # 赤太字条件付き書式の適用
            for row_idx in range(2, worksheet.max_row + 1):
                cell_ratio = worksheet.cell(row=row_idx, column=9)
                ratio_val = cell_ratio.value
                
                if ratio_val is not None and isinstance(ratio_val, (int, float)) and ratio_val <= 100:
                    cell_name = worksheet.cell(row=row_idx, column=4)
                    cell_name.font = red_bold_font
                    cell_ratio.font = red_bold_font
            
            # 列幅の自動調整
            for col in worksheet.columns:
                max_len = 0
                col_letter = col[0].column_letter
                for cell in col:
                    if cell.value is not None:
                        val_str = str(cell.value)
                        str_len = sum(2 if ord(c) > 0x7F else 1 for c in val_str)
                        max_len = max(max_len, str_len)
                worksheet.column_dimensions[col_letter].width = max(max_len + 3, 10)
            
    print(f"突合Excelレポートの作成（店舗: {target_store}）が完了しました。")

if __name__ == "__main__":
    main()
