# Excel-PDF自動照合＆マークアップツール (Python)

## 概要

このツールは、Azure Document IntelligenceとPython/LLMを組み合わせ、Excelファイルと出典PDFを読み込み、記載内容の自動検証、PDFへのマークアップ、そして照合レポートの生成までを行うワークフローを提案します。表記の差異や単位の変換を考慮した柔軟な照合が可能です。

## 主な機能

- Excelファイルからの入力値セルと計算式セルの抽出
- 出典PDFからのテキストと座標情報の抽出 (Azure Document Intelligence)
- LLMによる入力値とPDFテキストの同値性判定（単位変換や表記揺れを考慮）
- PDFへの照合結果（✅/❌）のマークアップとセル位置情報の追記
- 照合結果のExcel/CSVレポート出力

## インストール
```bash
uv sync

# For dev
# uv pip install -e .
```

## 利用ライブラリ

- **出典 PDF の読み込み・座標抽出**: [`azure-ai-documentintelligence`](https://pypi.org/project/azure-ai-documentintelligence/) (Azure AI Document Intelligence)
- **Excel ファイル読み込み**: [`openpyxl`](https://pypi.org/project/openpyxl/), [`pandas`](https://pypi.org/project/pandas/)
- **PDF 書き込み／マークアップ**: [`PyMuPDF`](https://pypi.org/project/PyMuPDF/) (fitz)
- **LLM連携**: [`langchain`](https://pypi.org/project/langchain/)

## 全体ワークフロー

1.  **Excel ファイル読み込み**:
    - `openpyxl` を使用してワークブックを開き、「入力値」と「計算結果」を抽出します。
    - 照合対象となるセルの位置（行・列）と値のリストを構築します。
2.  **出典 PDF 読み込み**:
    - `azure-ai-formrecognizer` の `prebuilt-layout` モデルを利用してPDFを解析し、テキスト情報と各単語のバウンディングボックス座標を取得します。
    - ページごとに単語と座標を構造化された形式で保存します。
3.  **マッチングロジック (LLM 使用)**:
    - Excelから抽出した各入力値と、PDFから抽出したテキスト候補をLLMに送信し、それらの同値性を判定します。
    - LLMは、`1km` ↔️ `1000m` のような単位変換や、「株式会社」↔️ 「(株)」のような表記の揺れを文脈に基づいて解釈し、マッチするかどうかを判定します。
    - LLMの出力は以下のJSON形式を想定しています。
        ```json
        [
          {"cell": "A1", "match": true, "reason": "...", "matched_text": "...", "source_path": "..."},
          ...
        ]
        ```
4.  **PDF マークアップ**:
    - `PyMuPDF` を使用して出典PDFを開き、LLMからの `match` フラグに基づいて、該当する座標周辺にマークアップを挿入します。
    - 必要に応じて、マークにシート名とセル座標を追記することで、トレーサビリティを向上させます。
5.  **レポート生成**:
    - LLMからのJSON出力を `pandas` の DataFrame に変換し、ExcelまたはCSV形式でエクスポートします。
    - `pandas` の機能を利用して、HTML形式や追加の集計グラフを含むダッシュボード形式でのレポート出力も可能です。

## コードスニペット例

```python:main.py
from langchain_google_genai import ChatGoogleGenerativeAI
# from langchain_openai import ChatOpenAI
from dotenv import load_dotenv

from pathlib import Path
import logging

import excel_sheet_matching_agent as esma

load_dotenv()
logging.basicConfig(level=logging.INFO)

excel_path = Path("examples/計算シートサンプル.xlsx")
excel_pdf_path = Path("examples/計算シートサンプル_シート1.pdf")
sheet_name = "シート1"
source_pdfs = [Path("examples/出典サンプル.pdf")]

# prebuild-layout
analyzed_markdown_paths = []
analyzed_json_paths = []
for source_pdf in source_pdfs:
    md, js = esma.analyze_local_pdf(source_pdf)
    analyzed_markdown_paths.append(md)
    analyzed_json_paths.append(js)

# load Inputs/Outputs from Excel sheet
inputs = esma.extract_data(excel_path, sheet_name).input

# matching
llm = ChatGoogleGenerativeAI(model="gemini-2.0-flash-lite", temperature=0)
# llm = ChatOpenAI(model="gpt-4o-mini", temperature=0)
matching_results = esma.match(llm, inputs, analyzed_markdown_paths)
for (inp, result) in zip(inputs, matching_results):
    print("-"*10)
    print(inp)
    print(result)

# markup
esma.markup(excel_path, excel_pdf_path, inputs, matching_results, source_pdfs, analyzed_json_paths)
```