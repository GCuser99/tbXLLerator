# Writing Your First UDF

This guide walks through building a UDF from scratch — not by copying a template, but by understanding each decision. We'll build a function that takes a number and returns its square root, with proper error handling for negative inputs.

---

## Step 1 — The Function Signature

Every XLL UDF has the same shape:

```vba
[DllExport]
Public Function TBXLL_MySqrt(ByRef pNum As XLOPER12) As LongPtr
```

- **`[DllExport]`** — makes the function visible to Excel
- **`ByRef pNum As XLOPER12`** — Excel passes a pointer to an XLOPER12
- **`As LongPtr`** — you return a pointer to the result XLOPER12

The `TBXLL_` prefix is a convention, not a requirement. It helps avoid name collisions with Excel's built-in functions.

---

## Step 2 — Declare Locals

```vba
    Dim xTemp As XLOPER12
    Dim n As Double
```

`xTemp` holds the result. It's a local variable — never `Static` — so the function is safe for multithreaded recalculation. `n` will hold the coerced input value.

---

## Step 3 — Bind the Input

```vba
    If Not BindU(pNum, btNumber, n, xTemp) Then GoTo ReturnResult
```

This converts the raw XLOPER12 to a `Double`. If the user passed something that can't be converted to a number (e.g., a text string like "hello"), `BindU` returns `False` and sets `xTemp` to `#VALUE!`. The `GoTo ReturnResult` skips to the return statement.

---

## Step 4 — Compute and Handle Errors

```vba
    If n < 0 Then
        xTemp = GetXLErr12(xlerrNum)
    Else
        xTemp = GetXLNum12(Sqr(n))
    End If
```

For negative inputs, we return `#NUM!` (the standard Excel error for invalid numeric operations). For valid inputs, we compute the square root and package it as an XLOPER12 number.

---

## Step 5 — Return

```vba
ReturnResult:
    Return AllocResultToCaller(xTemp)
End Function
```

`AllocResultToCaller` copies `xTemp` to the heap and returns the pointer. Excel will call `xlAutoFree12` to free it later.

---

## The Complete Function

```vba
[DllExport]
Public Function TBXLL_MySqrt(ByRef pNum As XLOPER12) As LongPtr
    Dim xTemp As XLOPER12
    Dim n As Double

    If Not BindU(pNum, btNumber, n, xTemp) Then GoTo ReturnResult

    If n < 0 Then
        xTemp = GetXLErr12(xlerrNum)
    Else
        xTemp = GetXLNum12(Sqr(n))
    End If

ReturnResult:
    Return AllocResultToCaller(xTemp)
End Function
```

---

## Step 6 — Register It

In `xlAutoOpen`, add:

```vba
Set udf = New UDF
With udf
    .ProcName = "TBXLL_MySqrt"
    .FuncText = "TBXLL_MySqrt"
    .Category = "My Add-In"
    .FuncHelp = "Returns the square root of a number, or #NUM! if negative"
    .Volatile = False
    .ThreadSafe = True
    .AddArgument Name:="number", Help:="The number to take the square root of"
    .Register
End With
udfs.Add udf
```

---

## Step 7 — Test

Build the XLL, load it in Excel, and try:

| Cell formula | Expected result |
|-------------|----------------|
| `=TBXLL_MySqrt(16)` | `4` |
| `=TBXLL_MySqrt(2)` | `1.41421356...` |
| `=TBXLL_MySqrt(-1)` | `#NUM!` |
| `=TBXLL_MySqrt("hello")` | `#VALUE!` |
| `=TBXLL_MySqrt(A1)` | Square root of whatever is in A1 |

The `#VALUE!` for "hello" comes from `BindU` failing to coerce a non-numeric string. The `#NUM!` for negative numbers comes from your explicit check. Both are standard Excel error semantics that users expect.

---

## The Decision Framework

When writing any UDF, ask yourself these questions:

1. **What types do my inputs need to be?** This determines your `BindType` (`btNumber`, `btString`, `btArray`, etc.)
2. **Are any arguments optional?** Check for `xltypeMissing` before binding. See [[Optional Arguments]].
3. **Can this function run on multiple threads?** If yes, register as `ThreadSafe = True` and use only local variables. If it needs shared state, use `ThreadSafe = False`.
4. **Does it need to be volatile?** Only if the result depends on something other than its arguments (like the current time). Most functions should be non-volatile.
5. **What errors should it return?** Map domain errors to the appropriate Excel error code (`xlerrNum`, `xlerrDiv0`, `xlerrNA`, `xlerrValue`).

---

## Next Steps

- [[Working with Arrays]] — when your function needs range inputs or array outputs
- [[Delegating to Excel Built‑ins]] — calling SUM, COUNTIF, TRANSPOSE from your UDF
- [[Returning Errors]] — the full set of Excel error codes and when to use each
