from pydantic import BaseModel
from typing import Optional, List

class ExcelCellInputData(BaseModel):
    sheet: str
    cell: str
    value: int | float
    metadata: dict

class ExcelCellFormulaData(BaseModel):
    sheet: str
    cell: str
    value: str
    metadata: dict

class ExcelSheetData(BaseModel):
    input: list[ExcelCellInputData]
    formula: list[ExcelCellFormulaData]

class MatchingResult(BaseModel):
    cell: str
    match: bool
    reason: str
    matched_text: Optional[str] = None
    source_path: Optional[str] = None

class MatchingBatchResult(BaseModel):
    results: List[MatchingResult]

