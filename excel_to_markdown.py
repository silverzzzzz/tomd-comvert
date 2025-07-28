import pandas as pd
import sys
import os
from pathlib import Path

def excel_to_markdown(excel_path, output_path=None):
    """
    エクセルファイルをマークダウン形式に変換する
    
    Args:
        excel_path: 入力エクセルファイルのパス
        output_path: 出力マークダウンファイルのパス（指定しない場合は同じディレクトリに.mdで保存）
    """
    try:
        # ファイルの存在確認
        if not os.path.exists(excel_path):
            print(f"エラー: ファイルが見つかりません - {excel_path}")
            return False
        
        # 出力パスの設定
        if output_path is None:
            base_path = Path(excel_path)
            output_path = base_path.parent / f"{base_path.stem}.md"
        
        # エクセルファイルを読み込み
        # .xlsファイルの場合
        if excel_path.endswith('.xls'):
            excel_file = pd.ExcelFile(excel_path, engine='xlrd')
        else:
            excel_file = pd.ExcelFile(excel_path)
        
        # マークダウンの内容を構築
        markdown_content = []
        
        # ファイル名をタイトルとして追加
        file_name = os.path.basename(excel_path)
        markdown_content.append(f"# {file_name}\n")
        
        # 各シートを処理
        for sheet_name in excel_file.sheet_names:
            # シート名を見出しとして追加
            markdown_content.append(f"\n## {sheet_name}\n")
            
            # シートのデータを読み込み
            df = pd.read_excel(excel_file, sheet_name=sheet_name)
            
            # 空のDataFrameの場合はスキップ
            if df.empty:
                markdown_content.append("*（データなし）*\n")
                continue
            
            # DataFrameをマークダウンテーブルに変換
            # ヘッダー行
            headers = "|" + "|".join(str(col) for col in df.columns) + "|"
            separator = "|" + "|".join("---" for _ in df.columns) + "|"
            
            markdown_content.append(headers)
            markdown_content.append(separator)
            
            # データ行
            for _, row in df.iterrows():
                row_data = "|" + "|".join(str(val).replace('\n', '<br>') if pd.notna(val) else "" for val in row) + "|"
                markdown_content.append(row_data)
            
            markdown_content.append("")  # 空行を追加
        
        # マークダウンファイルに書き込み
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(markdown_content))
        
        print(f"変換完了: {output_path}")
        return True
        
    except Exception as e:
        print(f"エラーが発生しました: {str(e)}")
        return False

def main():
    if len(sys.argv) < 2:
        print("使用方法: python excel_to_markdown.py <エクセルファイルパス> [出力マークダウンファイルパス]")
        sys.exit(1)
    
    excel_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else None
    
    excel_to_markdown(excel_path, output_path)

if __name__ == "__main__":
    main()