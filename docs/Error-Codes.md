# Error Codes

Excel defines a fixed set of error values. XLL UDFs return these via `GetXLErr12()` or propagate them from input arrays via `CVErr()`.

---

## XloperErrorCodes Enum

```vba
Public Enum XloperErrorCodes
    xlerrNull = 0
    xlerrDiv0 = 7
    xlerrValue = 15
    xlerrRef = 23
    xlerrName = 29
    xlerrNum = 36
    xlerrNA = 42
    xlerrGettingData = 43
End Enum
```

---

## Error Reference

| Constant | Value | Excel Display | Meaning | Typical UDF Use |
|----------|-------|--------------|---------|-----------------|
| `xlerrNull` | 0 | `#NULL!` | Invalid range intersection | Rarely returned by UDFs |
| `xlerrDiv0` | 7 | `#DIV/0!` | Division by zero | Denominator is zero |
| `xlerrValue` | 15 | `#VALUE!` | Wrong value type | Invalid argument type, coercion failure |
| `xlerrRef` | 23 | `#REF!` | Invalid cell reference | Deleted or invalid reference |
| `xlerrName` | 29 | `#NAME?` | Unrecognized name | Rarely returned by UDFs |
| `xlerrNum` | 36 | `#NUM!` | Invalid numeric value | Domain error (negative sqrt, overflow) |
| `xlerrNA` | 42 | `#N/A` | Value not available | Lookup not found, no applicable result |
| `xlerrGettingData` | 43 | `#GETTING_DATA` | Data pending | Async operations in progress |

---

## Returning Errors

```vba
' Return a specific error
xTemp = GetXLErr12(xlerrDiv0)

' Return #VALUE! via the convenience helper
SetErrorResult xTemp
```

## Propagating Errors from Arrays

When iterating input arrays, check for `vbError` and pass errors through to the output:

```vba
If VarType(arr(r, c)) = vbError Then
    arrOut(r, c) = arr(r, c)        ' preserve the original error
End If
```

To create a new error in an output array:

```vba
arrOut(r, c) = CVErr(xlerrValue)    ' insert #VALUE! into the output
```

## Converting Errors to Text

The `XlErrToString` helper returns the display string for an error code:

```vba
XlErrToString(xlerrNull)        ' → "#NULL!"
XlErrToString(xlerrDiv0)        ' → "#DIV/0!"
XlErrToString(xlerrValue)       ' → "#VALUE!"
XlErrToString(xlerrRef)         ' → "#REF!"
XlErrToString(xlerrName)        ' → "#NAME?"
XlErrToString(xlerrNum)         ' → "#NUM!"
XlErrToString(xlerrNA)          ' → "#N/A"
XlErrToString(xlerrGettingData) ' → "#GETTING_DATA"
```

---

## Next Steps

- [[Returning Errors]] — patterns for error handling in UDFs
- [[Working with Arrays]] — propagating errors through array processing
