# Optional Arguments

Excel functions commonly have optional parameters — think of `ROUND(number, [digits])` where digits defaults to 0. XLL UDFs support the same pattern by detecting `xltypeMissing`.

---

## How It Works

When a user omits an argument in a formula, Excel passes an XLOPER12 with `xltype = xltypeMissing`. You check for this before calling `BindU`:

```vba
[DllExport]
Public Function TBXLL_AddOptional( _
    ByRef pA As XLOPER12, _
    ByRef pB As XLOPER12, _
    ByRef pC As XLOPER12) As LongPtr

    Dim xTemp As XLOPER12
    Dim a As Double, b As Double, c As Double

    If Not BindU(pA, btNumber, a, xTemp) Then GoTo ReturnResult
    If Not BindU(pB, btNumber, b, xTemp) Then GoTo ReturnResult

    ' Optional argument: default = 0
    If pC.xltype = xltypeMissing Then
        c = 0
    ElseIf Not BindU(pC, btNumber, c, xTemp) Then
        GoTo ReturnResult
    End If

    xTemp = GetXLNum12(a + b + c)
ReturnResult:
    Return AllocResultToCaller(xTemp)
End Function
```

The user can call this as either `=TBXLL_AddOptional(1, 2)` (returns 3) or `=TBXLL_AddOptional(1, 2, 10)` (returns 13).

---

## The Pattern

The check always follows this structure:

```vba
If pArg.xltype = xltypeMissing Then
    ' Use your default value
    myVar = defaultValue
ElseIf Not BindU(pArg, btSomeType, myVar, xTemp) Then
    GoTo ReturnResult
End If
```

Check for `xltypeMissing` **first**, before calling `BindU`. If you call `BindU` on a missing argument, it will fail and set `xTemp` to `#VALUE!`.

---

## Optional Boolean Example

Optional flags with a boolean default:

```vba
Dim transpose As Boolean

If pTranspose.xltype = xltypeMissing Then
    transpose = False
ElseIf Not BindU(pTranspose, btBool, transpose, xTemp) Then
    GoTo ReturnResult
End If
```

---

## Optional Arguments in Built-in Delegation

When delegating to an Excel built-in that has optional parameters, use `GetXLMissing12()` to pass the omitted argument:

```vba
' ROUND(number, [digits]) — omit digits to default to 0
Dim args(1) As XLOPER12
args(0) = pNum
args(1) = GetXLMissing12()

If Excel12v(xlfRound, xOut(0), 2, args) <> 0 Then
    SetErrorResult xTemp
    GoTo ReturnResult
End If
```

---

## Registration

Optional arguments are registered the same way as required arguments — Excel doesn't distinguish them at the registration level. The optional behavior is entirely in your function's logic. However, the help text should indicate that the argument is optional:

```vba
.AddArgument Name:="c", Help:="optional - defaults to 0 if omitted"
```

---

## Next Steps

- [[Argument Binding]] — the full set of bind types
- [[Delegating to Excel Built‑ins]] — passing `GetXLMissing12()` to built-ins
