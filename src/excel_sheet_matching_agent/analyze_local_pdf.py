'''
This script will:
- save results in `data/document_intelligence`
  * markdown files: `data/document_intelligence/markdown`
  * json files: `data/document_intelligence/json`

# 追加
多言語対応OCRでは、どうしても中国語の簡体字や繁体字が混ざってしまう
紛れ込んでしまう中国語の漢字を日本語の漢字に変換する
'''
from azure.core.credentials import AzureKeyCredential
from azure.ai.documentintelligence import DocumentIntelligenceClient
from dotenv import load_dotenv
from opencc import OpenCC

import json
from pathlib import Path
import os
import logging


load_dotenv()
logger = logging.getLogger(__name__)

# DocumentIntelligenceClient の作成
client = DocumentIntelligenceClient(
    endpoint=os.environ["AZURE_DOCUMENT_INTELLIGENCE_ENDPOINT"],
    credential=AzureKeyCredential(os.environ["AZURE_DOCUMENT_INTELLIGENCE_API_KEY"])
)

# OpenCCのセットアップ
cc_s2t = OpenCC('s2t')    # 簡体字 → 繁体字
cc_t2jp = OpenCC('t2jp')  # 繁体字 → 日本語の漢字

def cc(text: str) -> str:
    # 1. 簡体字 → 繁体字
    converted = cc_s2t.convert(text)
    # 2. 繁体字 → 日本語漢字
    converted = cc_t2jp.convert(converted)
    return converted

def analyze_local_pdf(image_path: Path, output_dir: Path = Path("./data/document_intelligence"), dryrun: bool = False) -> tuple[Path, Path]:
    logger.info(f"calling prebuilt-layout API: {image_path}")

    # path
    output_markdown_path = output_dir / "markdown" / f"{image_path.stem}.md"
    output_json_path = output_dir / "json" / f"{image_path.stem}.json"
    if output_markdown_path.exists() and output_json_path.exists():
        logger.info(f"analyzed result already exists, skipping: {output_markdown_path}")
        return output_markdown_path, output_json_path

    output_markdown_path.parent.mkdir(parents=True, exist_ok=True)
    output_json_path.parent.mkdir(parents=True, exist_ok=True)

    if dryrun:
        logger.info(f"will be saved in {output_markdown_path}")
        return output_markdown_path, output_json_path
    
    # 分析実行：output_content_format を Markdown に指定
    with open(image_path, "rb") as f:
        poller = client.begin_analyze_document(
            "prebuilt-layout",  # レイアウト解析モデル
            f,
            output_content_format="markdown"
        )
    
    # 長時間実行される場合は poller.result() で待機
    result = poller.result()

    # Markdown 形式の文字列（人が読めるフォーマット）が result.content に入っています
    markdown_output = result.content
    with output_markdown_path.open("w", encoding="utf-8") as mf:
        mf.write(cc(markdown_output))
    
    # 解析結果の全体データを辞書形式に変換して JSON ファイルとして保存
    with output_json_path.open("w", encoding="utf-8") as jf:
        text = json.dumps(result.as_dict(), ensure_ascii=False)
        jf.write(cc(text))  # 中国語漢字変換
    
    return output_markdown_path, output_json_path

if __name__ == "__main__":
    import sys
    analyze_local_pdf(Path(sys.argv[1]), dryrun=False)