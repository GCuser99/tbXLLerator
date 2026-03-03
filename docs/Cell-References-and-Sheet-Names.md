# Cell References and Sheet Names

Some UDFs need to know *where* a cell is, not just its value. For example, a function that returns a cell's address, or one that extracts the sheet name. These patterns require working with raw cell references — which is only possible with type-`U` registration.

---

## Getting a Cell's Row and Column

Use `btSingleCellRef` to validate that the input is a single cell and extract its coordinates:

```vba
[DllExport]
Public Function TBXLL_CellAddress(ByRef pRef As XLOPER12) As LongPtr
    Dim xTemp As XLOPER12
    Dim v As Variant

    ' Validate single cell — returns #VALUE! for multi-cell ranges
    If Not BindU(pRef, btSingleCellRef, v, xTemp) Then GoTo ReturnResult

    ' v(0) = 1-based row, v(1) = 1-based column
    Dim args(1) As XLOPER12
    args(0) = GetXLVariant12(v(0))
    args(1) = GetXLVariant12(v(1))

    ' Use Excel's ADDRESS function to format it
    If Excel12v(xlfAddress, xTemp, 2, args) <> 0 Then
        SetErrorResult xTemp
    End If

ReturnResult:
    Return AllocResultToCaller(xTemp)
End Function
```

In the worksheet: `=TBXLL_CellAddress(A1)` returns `$A$1`. Passing a multi-cell range like `A1:B2` returns `#VALUE!`.

---

## Getting a Cell's Value

Use `btValue` to extract the scalar value from a single cell, preserving its original type:

```vba
Dim v As Variant
If Not BindU(pRef, btValue, v, xTemp) Then GoTo ReturnResult

If IsEmpty(v) Then
    ' Cell was blank
ElseIf VarType(v) = vbDouble Then
    ' Cell contained a number
ElseIf VarType(v) = vbString Then
    ' Cell contained text
End If
```

You can combine `btSingleCellRef` and `btValue` — first validate that it's a single cell, then extract its value:

```vba
If Not BindU(pRef, btSingleCellRef, v, xTemp) Then GoTo ReturnResult
If Not BindU(pRef, btValue, v, xTemp) Then GoTo ReturnResult
```

---

## Getting the Sheet Name

The `xlSheetNm` callback extracts the sheet name from a cell reference:

```vba
[DllExport]
Public Function TBXLL_SheetName(ByRef pRef As XLOPER12) As LongPtr
    Dim xTemp As XLOPER12
    Dim xArgs(0) As XLOPER12
    Dim xOut(0) As XLOPER12
    Dim xFree(0) As XLOPER12

    xArgs(0) = pRef

    If Excel12v(xlSheetNm, xOut(0), 1, xArgs) <> 0 Then
        SetErrorResult xTemp
        GoTo ReturnResult
    End If

    xTemp = GetXLString12(Xloper12StrValue(xOut(0)))

    xFree(0) = xOut(0)
    Excel12v xlFree, ByVal vbNullPtr, 1, xFree

ReturnResult:
    Return AllocResultToCaller(xTemp)
End Function
```

Returns a string like `[Book1.xlsx]Sheet1`.

### MacroEquivalent registration

`xlSheetNm` is a macro-only callback. Functions that use it should be registered with `MacroEquivalent = True`, which is mutually exclusive with `ThreadSafe = True`:

```vba
Set udf = New UDF
With udf
    .ProcName = "TBXLL_SheetName"
    .FuncHelp = "Returns the sheet name from a cell reference"
    .ThreadSafe = False
    .MacroEquivalent = True
    .AddArgument Name:="range", Help:="A cell reference"
    .Register
End With
```

In practice, `xlSheetNm` often works without `MacroEquivalent = True` in current Excel versions, but formally it requires it.

---

## Why This Only Works with Type U

Type-`U` registration passes the raw XLOPER12 from Excel, which can be `xltypeSRef` (same-sheet reference) or `xltypeRef` (cross-sheet reference). These contain the row, column, and sheet identity.

Type-`Q` registration tells Excel to coerce references to values *before* calling your function. By the time you receive the argument, it's already been resolved to `xltypeNum`, `xltypeStr`, etc. — the reference information is gone.

This is why `btSingleCellRef` always fails with `BindQ` — there's no reference to inspect.

---

## Detecting Empty Cells

Use `btValue` to distinguish between blank cells and cells containing zero or empty strings:

```vba
[DllExport]
Public Function TBXLL_IsEmptyCell(ByRef cell As XLOPER12) As LongPtr
    Dim xTemp As XLOPER12
    Dim v As Variant

    If Not BindU(cell, btValue, v, xTemp) Then GoTo ReturnResult

    xTemp = GetXLBool12(IsEmpty(v))
ReturnResult:
    Return AllocResultToCaller(xTemp)
End Function
```

A blank cell returns `Empty` (VarType = `vbEmpty`), while a cell containing `0` returns `0.0` (VarType = `vbDouble`) and a cell containing `""` returns `""` (VarType = `vbString`).

---

## Next Steps

- [[U vs Q Registration]] — why reference access requires type U
- [[Argument Binding]] — the full btSingleCellRef and btValue documentation
- [[Delegating to Excel Built‑ins]] — using `xlSheetNm` and `xlfAddress`
