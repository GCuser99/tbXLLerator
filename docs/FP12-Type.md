# FP12 Type

The `FP12` registration type (`rdtFP12`, type char `K%`) is a specialized alternative to XLOPER12 for functions that work exclusively with numeric arrays. It provides a more direct path to the data — no `Variant` boxing, no per-element type checking, and no `xlCoerce` overhead.

---

## What FP12 Is

When you register an argument as `rdtFP12`, Excel passes a pointer to a memory block with this layout:

```
Offset 0:  Long   — row count
Offset 4:  Long   — column count
Offset 8:  Double — first element (row 0, col 0)
Offset 16: Double — second element (row 0, col 1)
...                  (row-major order)
```

The data is a contiguous array of `Double` values. Non-numeric cells are converted to `0.0` by Excel before the call.

---

## Reading FP12 Data

The `ReadFP12` helper extracts the dimensions and copies the data into a 2D `Double` array:

```vba
Dim rows As Long, cols As Long
Dim arr() As Double

If Not ReadFP12(lpFP12, rows, cols, arr) Then
    SetErrorResult xTemp
    GoTo ReturnResult
End If
```

After a successful read:
- `arr(0, 0)` is the top-left element
- `arr(rows - 1, cols - 1)` is the bottom-right element
- All values are `Double` — no type checking needed

---

## Complete Example

```vba
[DllExport]
Public Function TBXLL_SumFP12(ByVal lpFP12 As LongPtr) As LongPtr
    Dim xTemp As XLOPER12
    Dim rows As Long, cols As Long
    Dim arr() As Double

    If Not ReadFP12(lpFP12, rows, cols, arr) Then
        SetErrorResult xTemp
        GoTo ReturnResult
    End If

    Dim total As Double
    Dim r As Long, c As Long
    For r = 0 To rows - 1
        For c = 0 To cols - 1
            total = total + arr(r, c)
        Next c
    Next r

    xTemp = GetXLNum12(total)
ReturnResult:
    Return AllocResultToCaller(xTemp)
End Function
```

Registration:

```vba
Set udf = New UDF
With udf
    .ProcName = "TBXLL_SumFP12"
    .FuncHelp = "Sums a numeric range using FP12 input"
    .ThreadSafe = True
    .AddArgument Name:="range", Help:="Numeric range", Type:=rdtFP12
    .Register
End With
```

---

## Function Signature

Note the signature difference from XLOPER12-based functions:

```vba
' XLOPER12 argument (U or Q)
Public Function MyFunc(ByRef pArg As XLOPER12) As LongPtr

' FP12 argument (K%)
Public Function MyFunc(ByVal lpFP12 As LongPtr) As LongPtr
```

FP12 arguments are received as a `LongPtr` (pointer) passed `ByVal`. The return type is still `LongPtr` — you still return an XLOPER12 via `AllocResultToCaller`.

---

## Tradeoffs

| | FP12 (`K%`) | XLOPER12 + btArray (`U`) |
|---|---|---|
| **Element type** | Always `Double` | `Variant` — can be Double, String, Error, Empty |
| **Error handling** | Non-numeric cells become `0.0` — errors are invisible | Errors preserved as `vbError` — can propagate or handle |
| **Blank cells** | Become `0.0` | Become `vbEmpty` — distinguishable from zero |
| **String cells** | Become `0.0` | Preserved as `vbString` |
| **Memory layout** | Contiguous doubles — cache-friendly | Variant array — larger per-element footprint |
| **Overhead** | Minimal — direct copy | `xlCoerce` + per-element Variant construction |

---

## When to Use FP12

FP12 is a good fit when:

- Your function only processes numeric data
- You don't need to distinguish between blank cells, errors, and zero
- Performance on large numeric arrays is critical
- The input range is known to contain only numbers

FP12 is **not** a good fit when:

- You need to detect or propagate errors
- The range may contain text or blanks that should be handled differently from zero
- You need to return a specific error for non-numeric input

---

## Next Steps

- [[Working with Arrays]] — the general-purpose array pattern using btArray
- [[Performance Tuning]] — comparing FP12 vs XLOPER12 array performance
- [[RegDataTypes Reference]] — all registration type options
