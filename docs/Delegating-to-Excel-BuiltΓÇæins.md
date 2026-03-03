# Delegating to Excel Built‑ins

One of the most powerful features of XLL development is the ability to call Excel's own built-in functions from within your UDF. This lets you leverage SUM, COUNTIF, TRANSPOSE, ROUND, and hundreds of other functions as building blocks, rather than reimplementing their logic.

---

## The Basic Pattern

All calls to Excel built-ins go through `Excel12v`:

```vba
Dim xResult As XLOPER12
Dim args(0) As XLOPER12

args(0) = pRange   ' pass the input argument through

If Excel12v(xlfSum, xResult, 1, args) <> 0 Then
    SetErrorResult xTemp
    GoTo ReturnResult
End If
```

- **First argument** — the function code (e.g., `xlfSum`, `xlfCountif`, `xlfTranspose`)
- **Second argument** — the output XLOPER12 that receives the result
- **Third argument** — the number of input arguments
- **Fourth argument** — the array of input XLOPER12s

`Excel12v` returns `0` on success and a nonzero error code on failure. Always check the return value.

---

## Direct Pass-Through

The simplest pattern passes a UDF argument directly to a built-in. This works because type-`U` arguments arrive as raw XLOPER12s that Excel can interpret natively — including cell references.

```vba
' =TBXLL_SumArray(A1:A10) → equivalent to =SUM(A1:A10)
[DllExport]
Public Function TBXLL_SumArray(ByRef pArr As XLOPER12) As LongPtr
    Dim xTemp As XLOPER12
    Dim args(0) As XLOPER12
    args(0) = pArr
    If Excel12v(xlfSum, xTemp, 1, args) <> 0 Then
        SetErrorResult xTemp
    End If
    Return AllocResultToCaller(xTemp)
End Function
```

When the built-in succeeds, `xTemp` already contains the result — you can return it directly. No intermediate `xlFree` is needed because `AllocResultToCaller` copies the value to the heap and the original `xTemp` is on the stack.

### Important nuance: when to xlFree

If you read the result of a built-in call and then build your own return value (rather than returning the built-in's result directly), you must `xlFree` the built-in's result:

```vba
' Call SUM, read the value, then build our own result
If Excel12v(xlfSum, xSum(0), 1, args) <> 0 Then
    SetErrorResult xTemp
    GoTo ReturnResult
End If

Dim total As Double
total = Xloper12NumValue(xSum(0))

' Free Excel's result — we've extracted what we need
xFree(0) = xSum(0)
Excel12v xlFree, ByVal vbNullPtr, 1, xFree

' Build our own result
xTemp = GetXLNum12(total * 2)
```

---

## Multiple Arguments

For built-ins that take multiple arguments, size the args array accordingly:

```vba
' =TBXLL_CountIf(A1:A10, ">5") → equivalent to =COUNTIF(A1:A10, ">5")
Dim args(1) As XLOPER12
args(0) = pRange
args(1) = pCriteria

If Excel12v(xlfCountif, xTemp, 2, args) <> 0 Then
    SetErrorResult xTemp
End If
```

---

## Zero-Argument Built-ins

Some functions take no arguments, like `NOW()`:

```vba
If Excel12v(xlfNow, xTemp, 0) <> 0 Then
    SetErrorResult xTemp
End If
```

Pass `0` for the count. The `args` array is not needed.

---

## Chaining Built-ins

When you need to call multiple built-ins in sequence, free each intermediate result before the next call (or before returning):

```vba
' Compute AVERAGE as SUM / COUNT
Dim sumRes As XLOPER12, cntRes As XLOPER12
Dim freeArgs(0) As XLOPER12

' Step 1: SUM
If Excel12v(xlfSum, sumRes, 1, args) <> 0 Then
    SetErrorResult xTemp
    GoTo ReturnResult
End If

' Step 2: COUNT — if this fails, free sumRes
If Excel12v(xlfCount, cntRes, 1, args) <> 0 Then
    freeArgs(0) = sumRes
    Excel12v xlFree, ByVal vbNullPtr, 1, freeArgs
    SetErrorResult xTemp
    GoTo ReturnResult
End If

' Step 3: compute and free both
xTemp = GetXLNum12(Xloper12NumValue(sumRes) / Xloper12NumValue(cntRes))

freeArgs(0) = sumRes
Excel12v xlFree, ByVal vbNullPtr, 1, freeArgs
freeArgs(0) = cntRes
Excel12v xlFree, ByVal vbNullPtr, 1, freeArgs
```

The discipline here is: if a later call fails, free any results you've already obtained before jumping to the error path.

---

## Passing Optional Arguments with GetXLMissing12

Some Excel built-ins have optional parameters. To omit them, pass `GetXLMissing12()`:

```vba
' =ROUND(3.7) with digits omitted → defaults to 0 → returns 4
Dim args(1) As XLOPER12
args(0) = pNum
args(1) = GetXLMissing12()   ' omit the digits argument

If Excel12v(xlfRound, xOut(0), 2, args) <> 0 Then
    SetErrorResult xTemp
    GoTo ReturnResult
End If
```

---

## Array-Returning Built-ins

Some built-ins return arrays (e.g., TRANSPOSE). The result is an `xltypeMulti` XLOPER12 that you can return directly:

```vba
' =TBXLL_Transpose(A1:C2) → transposed array
Dim args(0) As XLOPER12
args(0) = pArr

If Excel12v(xlfTranspose, xTemp, 1, args) <> 0 Then
    SetErrorResult xTemp
End If

Return AllocResultToCaller(xTemp)
```

The user enters this as an array formula (Ctrl+Shift+Enter) or it spills in Excel 365.

---

## Common Function Codes

| Code | Excel Function | Arguments |
|------|---------------|-----------|
| `xlfSum` | SUM | 1+ ranges |
| `xlfCount` | COUNT | 1+ ranges |
| `xlfCountif` | COUNTIF | range, criteria |
| `xlfAverage` | AVERAGE | 1+ ranges |
| `xlfUpper` | UPPER | text |
| `xlfLower` | LOWER | text |
| `xlfTranspose` | TRANSPOSE | array |
| `xlfRound` | ROUND | number, digits |
| `xlfNow` | NOW | (none) |
| `xlfToday` | TODAY | (none) |
| `xlfAddress` | ADDRESS | row, column |

The full list of function codes is in the `ExcelFunctionNumbers` enum in `ExcelSDK`.

---

## Callback-Only Functions

Some functions are only available to XLLs, not to worksheet formulas:

| Code | Purpose |
|------|---------|
| `xlSheetNm` | Get the sheet name from a cell reference |
| `xlfCaller` | Get info about the calling cell |
| `xlSheetId` | Get the sheet ID |
| `xlCoerce` | Coerce an XLOPER12 to a different type |
| `xlFree` | Free an Excel-allocated XLOPER12 |

Functions like `xlSheetNm` and `xlfCaller` require `MacroEquivalent = True` registration, which is mutually exclusive with `ThreadSafe = True`.

---

## Next Steps

- [[Memory Management]] — ownership rules for `Excel12v` results
- [[Optional Arguments]] — using `GetXLMissing12` and `xltypeMissing`
- [[Cell References and Sheet Names]] — extracting location info from references
