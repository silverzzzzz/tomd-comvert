import sys
import os
from pathlib import Path
import re
from typing import List, Optional
import zipfile
from xml.etree import ElementTree as ET

# python-docx for .docx files
try:
    from docx import Document
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

# odfpy for .odt files
try:
    from odf import text, table
    from odf.opendocument import load
    from odf.element import Element
    ODT_AVAILABLE = True
except ImportError:
    ODT_AVAILABLE = False

def clean_text(text):
    """テキストのクリーニング"""
    if text is None:
        return ""
    # 余分な空白を削除
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def docx_to_markdown(docx_path):
    """Wordドキュメント(.docx)をマークダウンに変換"""
    if not DOCX_AVAILABLE:
        raise ImportError("python-docxがインストールされていません。pip install python-docx を実行してください。")
    
    doc = Document(docx_path)
    markdown_content = []
    
    for element in doc.element.body:
        # 段落の処理
        if element.tag.endswith('p'):
            para = Paragraph(element, doc)
            text = para.text.strip()
            
            if text:
                # 見出しレベルの判定
                style_name = para.style.name if para.style else ""
                
                if "Heading 1" in style_name or "見出し 1" in style_name:
                    markdown_content.append(f"# {text}\n")
                elif "Heading 2" in style_name or "見出し 2" in style_name:
                    markdown_content.append(f"## {text}\n")
                elif "Heading 3" in style_name or "見出し 3" in style_name:
                    markdown_content.append(f"### {text}\n")
                else:
                    # リストアイテムの処理
                    if para.style and ("List" in para.style.name or "箇条書き" in para.style.name):
                        markdown_content.append(f"- {text}")
                    else:
                        # 通常の段落
                        # 太字・斜体の処理
                        formatted_text = text
                        for run in para.runs:
                            if run.bold and run.italic:
                                formatted_text = formatted_text.replace(run.text, f"***{run.text}***")
                            elif run.bold:
                                formatted_text = formatted_text.replace(run.text, f"**{run.text}**")
                            elif run.italic:
                                formatted_text = formatted_text.replace(run.text, f"*{run.text}*")
                        
                        markdown_content.append(formatted_text)
                markdown_content.append("")
        
        # テーブルの処理
        elif element.tag.endswith('tbl'):
            tbl = Table(element, doc)
            if len(tbl.rows) > 0:
                markdown_content.append("")
                
                # ヘッダー行
                header_cells = []
                for cell in tbl.rows[0].cells:
                    header_cells.append(clean_text(cell.text))
                
                header = "|" + "|".join(header_cells) + "|"
                separator = "|" + "|".join("---" for _ in header_cells) + "|"
                
                markdown_content.append(header)
                markdown_content.append(separator)
                
                # データ行
                for row in tbl.rows[1:]:
                    row_cells = []
                    for cell in row.cells:
                        row_cells.append(clean_text(cell.text).replace('\n', '<br>'))
                    row_data = "|" + "|".join(row_cells) + "|"
                    markdown_content.append(row_data)
                
                markdown_content.append("")
    
    return markdown_content

def extract_text_from_odt_element(element):
    """ODT要素からテキストを再帰的に抽出"""
    text_content = []
    
    if hasattr(element, 'data'):
        text_content.append(element.data)
    
    if hasattr(element, 'childNodes'):
        for child in element.childNodes:
            text_content.extend(extract_text_from_odt_element(child))
    
    return text_content

def odt_to_markdown(odt_path):
    """OpenDocumentテキスト(.odt)をマークダウンに変換"""
    if not ODT_AVAILABLE:
        raise ImportError("odfpyがインストールされていません。pip install odfpy を実行してください。")
    
    doc = load(odt_path)
    markdown_content = []
    
    # 全てのテキスト要素を取得
    for element in doc.getElementsByType(text.P):
        text_content = ''.join(extract_text_from_odt_element(element))
        text_content = clean_text(text_content)
        
        if text_content:
            # スタイル名を取得
            style_name = element.getAttribute('stylename') or ""
            
            # 見出しの判定
            if "Heading" in style_name or "heading" in style_name:
                level = 1  # デフォルトレベル
                # レベル番号を抽出
                match = re.search(r'\d+', style_name)
                if match:
                    level = min(int(match.group()), 6)
                
                prefix = "#" * level
                markdown_content.append(f"{prefix} {text_content}\n")
            else:
                markdown_content.append(text_content)
                markdown_content.append("")
    
    # テーブルの処理
    for tbl in doc.getElementsByType(table.Table):
        markdown_content.append("")
        
        rows = tbl.getElementsByType(table.TableRow)
        if rows:
            # 各行を処理
            table_data = []
            for row in rows:
                cells = row.getElementsByType(table.TableCell)
                row_data = []
                for cell in cells:
                    cell_text = ''.join(extract_text_from_odt_element(cell))
                    row_data.append(clean_text(cell_text))
                table_data.append(row_data)
            
            if table_data:
                # ヘッダー行
                header = "|" + "|".join(table_data[0]) + "|"
                separator = "|" + "|".join("---" for _ in table_data[0]) + "|"
                
                markdown_content.append(header)
                markdown_content.append(separator)
                
                # データ行
                for row in table_data[1:]:
                    row_text = "|" + "|".join(cell.replace('\n', '<br>') for cell in row) + "|"
                    markdown_content.append(row_text)
                
                markdown_content.append("")
    
    return markdown_content

def doc_to_markdown(doc_path, output_path=None):
    """
    ドキュメントファイルをマークダウン形式に変換する
    
    Args:
        doc_path: 入力ドキュメントファイルのパス（.docx, .odt）
        output_path: 出力マークダウンファイルのパス（指定しない場合は同じディレクトリに.mdで保存）
    """
    try:
        # ファイルの存在確認
        if not os.path.exists(doc_path):
            print(f"エラー: ファイルが見つかりません - {doc_path}")
            return False
        
        # 出力パスの設定
        if output_path is None:
            base_path = Path(doc_path)
            output_path = base_path.parent / f"{base_path.stem}.md"
        
        # ファイル拡張子による処理の分岐
        file_ext = Path(doc_path).suffix.lower()
        
        # マークダウンの内容を構築
        markdown_content = []
        
        # ファイル名をタイトルとして追加
        file_name = os.path.basename(doc_path)
        markdown_content.append(f"# {file_name}\n")
        
        if file_ext == '.docx':
            print("Word文書(.docx)を処理中...")
            content = docx_to_markdown(doc_path)
            markdown_content.extend(content)
            
        elif file_ext in ['.odt', '.ods']:
            print("OpenDocument形式を処理中...")
            content = odt_to_markdown(doc_path)
            markdown_content.extend(content)
            
        else:
            print(f"エラー: サポートされていないファイル形式です - {file_ext}")
            print("サポート形式: .docx, .odt")
            return False
        
        # マークダウンファイルに書き込み
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(markdown_content))
        
        print(f"変換完了: {output_path}")
        return True
        
    except ImportError as e:
        print(f"ライブラリエラー: {str(e)}")
        return False
    except Exception as e:
        print(f"エラーが発生しました: {str(e)}")
        print("\n必要なライブラリ:")
        print("- Word文書(.docx): pip install python-docx")
        print("- OpenDocument(.odt): pip install odfpy")
        return False

def main():
    if len(sys.argv) < 2:
        print("使用方法: python doc_to_markdown.py <ドキュメントファイルパス> [出力マークダウンファイルパス]")
        print("\nサポート形式:")
        print("- Word文書: .docx")
        print("- OpenDocument: .odt")
        sys.exit(1)
    
    doc_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else None
    
    doc_to_markdown(doc_path, output_path)

if __name__ == "__main__":
    main()