from langchain_core.prompts import ChatPromptTemplate

matching_prompt = ChatPromptTemplate([
    (
        "system", 
        """You are a document verification assistant.
Your job is to determine whether each numeric input from an Excel sheet is justified by the provided source document.
Use semantic understanding, and consider units, value conversions, and contextual meaning.
"""
    ),
    (
        "human",
        """Here are some examples:

EXAMPLE 1:
- Value: 1000
- Hint: 単位: m
- PDF Text: この道路の長さは1kmである。
Expected Match: true
Reason: 1km = 1000m, which matches the value.
Matched Text: "1km"

EXAMPLE 2:
- Value: 500
- Hint: 金額
- PDF Text: 報酬は月額500円である。
Expected Match: true
Reason: The PDF mentions 500円 which matches the input.
Matched Text: "500円"

EXAMPLE 3:
- Value: 200
- Hint: 重さ, kg
- PDF Text: 180kgと記載されている。
Expected Match: false
Reason: The input value (200kg) differs from the source (180kg).

EXAMPLE 4:
- Value: 300
- Hint: 単位: ml
- PDF Text: この液体の体積は300cm3である。
Expected Match: true
Reason: 300ml = 300cm3, which matches the value.
Matched Text: "300cm3"

Now verify the following inputs against the document.
"""
    ),
    (
        "human",
        """
### Source Document
```
{document_text}
```

### Excel Inputs
```
{excel_inputs}
```
"""
    )
])


extract_inputs_prompt = ChatPromptTemplate([
    ("human", """You are an expert data extraction tool capable of analyzing HTML representations of spreadsheets and extracting specific information based on class attributes.

Here is an HTML table representing an Excel sheet:
```html
{html_table}
```

Your task is to find all `<td>` elements that have the class `input`.
For each such cell, extract the following information and format it as a JSON object conforming to the `ExcelCellInputData` schema provided below.

**Instructions for fields:**
- `cell`: Determine the Excel cell reference (e.g., "A1", "B2") by combining the `data-col` attribute (which is the Excel column letter like A, B, C) and the `data-row` attribute (which is the Excel row number like 1, 2, 3). For example, if `data-col="B"` and `data-row="2"`, the cell reference is "B2".
- `value`: Extract the numeric value directly from the content of the `<td>` tag. Ensure it is parsed as an integer or float. If the cell is empty, the value should be null.
- `metadata`: Analyze the cells surrounding this `input` cell (especially in the same row or column, like potential headers or labels) and infer a list of strings that describe or provide context for this value. **Prioritize extracting units if they are present near the cell.** If no clear metadata is apparent from the HTML context, return an empty list `[]`.

Provide the output as a JSON array of these objects.
""")])


if __name__ == "__main__":
    print(matching_prompt.invoke({"document_text": "abcdefg", "excel_inputs": "1234567890", "results": ""}))