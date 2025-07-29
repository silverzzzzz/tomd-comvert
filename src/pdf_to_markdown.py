import sys
import os
from pathlib import Path
import pdfplumber
import pandas as pd
import re

def clean_text(text):
    """テキストのクリーニング"""
    if text is None:
        return ""
    # 余分な空白を削除
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def extract_tables_from_pdf(pdf_path):
    """PDFからテーブルを抽出してマークダウン形式に変換"""
    tables_markdown = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            # ページからテーブルを抽出
            tables = page.extract_tables()
            
            if tables:
                for table_num, table in enumerate(tables, 1):
                    if not table or len(table) == 0:
                        continue
                    
                    # テーブル番号を追加
                    if len(tables) > 1:
                        tables_markdown.append(f"\n### ページ{page_num} - テーブル{table_num}\n")
                    else:
                        tables_markdown.append(f"\n### ページ{page_num} - テーブル\n")
                    
                    # テーブルデータをクリーニング
                    cleaned_table = []
                    for row in table:
                        cleaned_row = [clean_text(str(cell)) if cell is not None else "" for cell in row]
                        # 空行をスキップ
                        if any(cell for cell in cleaned_row):
                            cleaned_table.append(cleaned_row)
                    
                    if not cleaned_table:
                        continue
                    
                    # マークダウンテーブルを作成
                    # ヘッダー行
                    header = "|" + "|".join(cleaned_table[0]) + "|"
                    separator = "|" + "|".join("---" for _ in cleaned_table[0]) + "|"
                    
                    tables_markdown.append(header)
                    tables_markdown.append(separator)
                    
                    # データ行
                    for row in cleaned_table[1:]:
                        row_data = "|" + "|".join(cell.replace('\n', '<br>') for cell in row) + "|"
                        tables_markdown.append(row_data)
                    
                    tables_markdown.append("")  # 空行を追加
    
    return tables_markdown

def extract_text_from_pdf(pdf_path):
    """PDFからテキストを抽出"""
    text_content = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            # ページからテキストを抽出
            text = page.extract_text()
            
            if text:
                text_content.append(f"\n## ページ{page_num}\n")
                # 段落ごとに分割
                paragraphs = text.split('\n\n')
                for para in paragraphs:
                    cleaned = clean_text(para)
                    if cleaned:
                        text_content.append(cleaned + "\n")
    
    return text_content

def pdf_to_markdown(pdf_path, output_path=None, mode="both"):
    """
    PDFファイルをマークダウン形式に変換する
    
    Args:
        pdf_path: 入力PDFファイルのパス
        output_path: 出力マークダウンファイルのパス（指定しない場合は同じディレクトリに.mdで保存）
        mode: 変換モード ("text": テキストのみ, "tables": テーブルのみ, "both": 両方)
    """
    try:
        # ファイルの存在確認
        if not os.path.exists(pdf_path):
            print(f"エラー: ファイルが見つかりません - {pdf_path}")
            return False
        
        # 出力パスの設定
        if output_path is None:
            base_path = Path(pdf_path)
            output_path = base_path.parent / f"{base_path.stem}.md"
        
        # マークダウンの内容を構築
        markdown_content = []
        
        # ファイル名をタイトルとして追加
        file_name = os.path.basename(pdf_path)
        markdown_content.append(f"# {file_name}\n")
        
        # テキストとテーブルの抽出
        if mode in ["text", "both"]:
            print("テキストを抽出中...")
            text_content = extract_text_from_pdf(pdf_path)
            if text_content and mode == "text":
                markdown_content.extend(text_content)
        
        if mode in ["tables", "both"]:
            print("テーブルを抽出中...")
            tables_content = extract_tables_from_pdf(pdf_path)
            if tables_content:
                if mode == "both" and text_content:
                    markdown_content.append("\n---\n\n# テーブルデータ\n")
                markdown_content.extend(tables_content)
        
        # マークダウンファイルに書き込み
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(markdown_content))
        
        print(f"変換完了: {output_path}")
        return True
        
    except Exception as e:
        print(f"エラーが発生しました: {str(e)}")
        print("pdfplumberがインストールされていない場合は、以下のコマンドでインストールしてください:")
        print("pip install pdfplumber")
        return False

def main():
    if len(sys.argv) < 2:
        print("使用方法: python pdf_to_markdown.py <PDFファイルパス> [出力マークダウンファイルパス] [モード]")
        print("モード: text (テキストのみ), tables (テーブルのみ), both (両方, デフォルト)")
        sys.exit(1)
    
    pdf_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else None
    mode = sys.argv[3] if len(sys.argv) > 3 else "both"
    
    if mode not in ["text", "tables", "both"]:
        print("エラー: モードは 'text', 'tables', 'both' のいずれかを指定してください")
        sys.exit(1)
    
    pdf_to_markdown(pdf_path, output_path, mode)

if __name__ == "__main__":
    main()