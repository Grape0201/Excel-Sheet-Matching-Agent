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
