# Working with Arrays

Many real-world UDFs need to accept ranges as input, process arrays of data, or return multi-cell results. This page covers the patterns for array input via `btArray`, element-by-element processing, and array output via `GetXLMulti12`.

---

## Array Input with btArray

`btArray` coerces any input — a cell reference, a literal array, or a single value — into a 2D `Variant()` array, 0-based in both dimensions:

```vba
Dim arr() As Variant
If Not BindU(pRange, btArray, arr, xTemp) Then GoTo ReturnResult
```

After a successful bind:
- `arr(0, 0)` is the top-left element
- `UBound(arr, 1)` is the last row index
- `UBound(arr, 2)` is the last column index
- A single cell becomes a 1×1 array: `arr(0, 0)`

### Element types

Each element preserves its original Excel type as a VBA `Variant`:

| Excel cell content | VarType | Example value |
|-------------------|---------|---------------|
| Number | `vbDouble` | `3.14` |
| Integer | `vbLong` | `42` |
| Text | `vbString` | `"hello"` |
| Boolean | `vbBoolean` | `True` |
| Error (#N/A, etc.) | `vbError` | `CVErr(xlerrNA)` |
| Blank cell | `vbEmpty` | `Empty` |

Always use `VarType()` to check element types before processing. Don't assume all elements are numeric — ranges can contain mixed data.

---

## Iterating an Array

The standard pattern is a nested loop with `VarType` discrimination:

```vba
Dim total As Double
Dim r As Long, c As Long

For r = 0 To UBound(arr, 1)
    For c = 0 To UBound(arr, 2)
        Select Case VarType(arr(r, c))
            Case vbDouble
                total = total + arr(r, c)
            Case vbLong
                total = total + CDbl(arr(r, c))
            Case vbError
                ' Skip errors, or propagate them
            Case vbEmpty
                ' Skip blank cells
            Case Else
                ' Handle unexpected types
        End Select
    Next c
Next r
```

This mirrors how Excel's own SUM function works — it silently skips text, blanks, and booleans, sums numbers, and propagates errors.

---

## Array Output with GetXLMulti12

To return a multi-cell result (an array formula), build a 2D `Variant()` array and convert it with `GetXLMulti12`:

```vba
Dim rows As Long = UBound(arr, 1) + 1
Dim cols As Long = UBound(arr, 2) + 1
Dim arrOut() As Variant
ReDim arrOut(rows - 1, cols - 1)

' Fill it with your computed values
Dim r As Long, c As Long
For r = 0 To rows - 1
    For c = 0 To cols - 1
        arrOut(r, c) = arr(r, c) * 2
    Next c
Next r

' Convert to XLOPER12
Dim xMulti As XLOPER12
xMulti = GetXLMulti12(arrOut)

If xMulti.xltype <> xltypeMulti Then
    SetErrorResult xTemp
    GoTo ReturnResult
End If

' Do NOT call FreeXLMulti12 here — xlAutoFree12 will handle it
xTemp = xMulti
```

The output array elements can be any type that `GetXLVariant12` supports: `Double`, `Long`, `Integer`, `Boolean`, `String`, or error variants via `CVErr()`.

### Important: do not free array results yourself

When you assign `xMulti` to `xTemp` and return it via `AllocResultToCaller`, ownership of the element array transfers to Excel. The `xlAutoFree12` callback frees the element array and any string buffers inside it. Calling `FreeXLMulti12` before returning would cause a double-free crash.

---

## String Array Output

`GetXLMulti12` handles string elements through `GetXLVariant12`, which calls `GetXLString12` for each string. The string buffers are allocated with `GlobalAlloc`, and `xlAutoFree12` frees them when Excel is done.

```vba
arrOut(r, c) = prefix & "_" & CStr(arr(r, c))
```

No special handling is needed — just assign strings to the output array like any other type.

---

## FP12 — An Alternative for Numeric-Only Arrays

If your function works exclusively with `Double` values (no strings, no errors, no blanks), the `FP12` type offers a more direct path. Register the argument with `Type:=rdtFP12`, and Excel passes a pointer to a flat `Double` array:

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

FP12 avoids the overhead of `Variant` arrays and `xlCoerce`, but can't represent errors, strings, or blanks. See [[FP12 Type]] for details.

---

## Next Steps

- [[Delegating to Excel Built‑ins]] — passing ranges to SUM, COUNTIF, etc.
- [[Performance Tuning]] — array vs scalar UDF performance considerations
- [[FP12 Type]] — the numeric-only array alternative
