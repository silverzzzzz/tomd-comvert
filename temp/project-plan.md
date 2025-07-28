# マークダウンコンバーター プロジェクト計画

## 概要
様々なファイル形式をマークダウンに変換するTypeScriptライブラリ

## 対象ファイル形式（候補）
- **テキスト系**: .txt, .csv, .json, .yaml, .xml
- **ドキュメント系**: .docx, .pdf, .html
- **コード系**: .js, .ts, .py, .java, .cpp
- **その他**: .xlsx, .pptx, jpg, .png

## 必要なライブラリ候補

### ファイル処理
- `fs-extra` - ファイル操作の拡張
- `glob` - ファイルパターンマッチング
- `mime-types` - ファイル形式判定

### 特定フォーマット処理
- `pdf-parse` - PDF解析
- `mammoth` - DOCX → HTML変換
- `xlsx` - Excel処理
- `csv-parser` - CSV処理
- `xml2js` - XML解析
- `turndown` - HTML → Markdown変換

### 開発環境
- `typescript` - TypeScript
- `ts-node` - 開発時実行
- `jest` - テスト
- `eslint` + `prettier` - コード品質
- `commander` - CLI

## アーキテクチャ
```
src/
├── converters/        # 各形式のコンバーター
├── core/             # 核となる変換ロジック
├── utils/            # ユーティリティ
├── cli/              # コマンドラインインターフェース
└── index.ts          # エントリーポイント
```