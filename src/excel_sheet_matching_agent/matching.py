from langchain_core.language_models.chat_models import BaseChatModel

from typing import List
from pathlib import Path

from .prompts import matching_prompt
from .models import ExcelCellInputData, MatchingBatchResult, MatchingResult


def batch_verify_inputs_with_llm(verify_chain, inputs: List[ExcelCellInputData], document_text: str) -> list[MatchingResult]:
    # Excelセル情報をプレーンテキスト化して渡す
    input_descriptions = []
    for item in inputs:
        hint = ", ".join([v for v in item.metadata.values() if v])
        input_descriptions.append(f"- Cell: {item.cell}\n  Value: {item.value}\n  Hint: {hint}")
    
    result = verify_chain.invoke({
        "document_text": document_text,
        "excel_inputs": "\n".join(input_descriptions)
    })
    return result.results # type: ignore

def match(llm: BaseChatModel, inputs: List[ExcelCellInputData], analyzed_markdown_paths: list[Path]) -> list[MatchingResult]:
    document_text = ""
    for i, markdown_path in enumerate(analyzed_markdown_paths):
        with markdown_path.open() as f:
            document_text = f"### Doucment #{i+1}\n"
            document_text = f"source_path: {markdown_path}\n"
            document_text += f.read()
            document_text += "\n\n"
    verify_chain = matching_prompt | llm.with_structured_output(MatchingBatchResult)
    return batch_verify_inputs_with_llm(verify_chain, inputs, document_text)
