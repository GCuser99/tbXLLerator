# Returning Errors

Excel has a well-defined set of error values that users recognize. Returning the appropriate error from your UDF makes it behave like a native Excel function — errors propagate correctly through dependent formulas, and users understand what went wrong.

---

## Error Codes

| Constant | Excel Display | When to use |
|----------|--------------|-------------|
| `xlerrDiv0` | `#DIV/0!` | Division by zero |
| `xlerrValue` | `#VALUE!` | Wrong argument type or invalid input |
| `xlerrNum` | `#NUM!` | Numeric domain error (e.g., sqrt of negative) |
| `xlerrNA` | `#N/A` | Value not found / not applicable |
| `xlerrRef` | `#REF!` | Invalid cell reference |
| `xlerrName` | `#NAME?` | Unrecognized name (rarely returned by UDFs) |
| `xlerrNull` | `#NULL!` | Invalid range intersection (rarely returned by UDFs) |
| `xlerrGettingData` | `#GETTING_DATA` | Async operation in progress |

---

## Returning an Error

Use `GetXLErr12` with the appropriate error constant:

```vba
' Return #DIV/0!
xTemp = GetXLErr12(xlerrDiv0)

' Return #NUM!
xTemp = GetXLErr12(xlerrNum)

' Return #N/A
xTemp = GetXLErr12(xlerrNA)
```

### The SetErrorResult shortcut

`SetErrorResult` is a convenience that sets the XLOPER12 to `#VALUE!`:

```vba
SetErrorResult xTemp   ' equivalent to: xTemp = GetXLErr12(xlerrValue)
```

This is the default error for "something went wrong" and is what `BindU` uses when coercion fails.

---

## Choosing the Right Error

Match your error to Excel's conventions so users get familiar feedback:

```vba
' Division by zero — use #DIV/0!
If b = 0 Then
    xTemp = GetXLErr12(xlerrDiv0)
    GoTo ReturnResult
End If
xTemp = GetXLNum12(a / b)
```

```vba
' Negative number for square root — use #NUM!
If n < 0 Then
    xTemp = GetXLErr12(xlerrNum)
    GoTo ReturnResult
End If
xTemp = GetXLNum12(Sqr(n))
```

```vba
' Lookup didn't find a match — use #N/A
xTemp = GetXLErr12(xlerrNA)
```

```vba
' Wrong input type — use #VALUE!
If (pText.xltype And xltypeStr) = 0 Then
    SetErrorResult xTemp
    GoTo ReturnResult
End If
```

---

## Propagating Errors from Input Arrays

When processing arrays, check for error elements and propagate them to the output:

```vba
For r = 0 To UBound(arr, 1)
    For c = 0 To UBound(arr, 2)
        If VarType(arr(r, c)) = vbError Then
            arrOut(r, c) = arr(r, c)   ' pass error through
        ElseIf VarType(arr(r, c)) = vbDouble Then
            arrOut(r, c) = arr(r, c) * 2
        Else
            arrOut(r, c) = CVErr(xlerrValue)
        End If
    Next c
Next r
```

This follows Excel's convention: if an input cell contains an error, the corresponding output cell shows the same error.

---

## Error-to-String Conversion

The `XlErrToString` helper converts error codes to their display strings:

```vba
Dim errText As String
errText = XlErrToString(xlerrDiv0)   ' returns "#DIV/0!"
```

This is useful in functions like `TBXLL_Join` that need to represent errors as text.

---

## Next Steps

- [[Argument Binding]] — how `BindU` generates errors on coercion failure
- [[Working with Arrays]] — propagating errors in array processing
