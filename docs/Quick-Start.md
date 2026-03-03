# Quick Start

This page walks through a minimal working XLL — a single UDF that multiplies two numbers. By the end you'll understand the three pieces every XLL UDF requires: the function itself, its registration, and its cleanup.

---

## The UDF Function

Every XLL UDF follows the same shape:

```vba
[DllExport]
Public Function TBXLL_Multiply( _
    ByRef pA As XLOPER12, _
    ByRef pB As XLOPER12) As LongPtr

    Dim xTemp As XLOPER12
    Dim a As Double, b As Double

    If Not BindU(pA, btNumber, a, xTemp) Then GoTo ReturnResult
    If Not BindU(pB, btNumber, b, xTemp) Then GoTo ReturnResult

    xTemp = GetXLNum12(a * b)
ReturnResult:
    Return AllocResultToCaller(xTemp)
End Function
```

Here's what's happening:

1. **`[DllExport]`** tells twinBASIC to export this function from the DLL so Excel can find it.
2. **`ByRef pA As XLOPER12`** — Excel passes arguments as pointers to `XLOPER12` structures. These are Excel's universal data containers that can hold numbers, strings, booleans, errors, arrays, or cell references.
3. **`BindU(pA, btNumber, a, xTemp)`** converts the raw `XLOPER12` into a `Double`. If the input can't be converted (e.g., the user passed a text string that isn't numeric), `BindU` returns `False` and sets `xTemp` to `#VALUE!` automatically.
4. **`GetXLNum12(a * b)`** packages the result back into an `XLOPER12` for return to Excel.
5. **`AllocResultToCaller(xTemp)`** allocates the result on the heap and sets the `xlbitDLLFree` flag so Excel knows to call `xlAutoFree12` when it's done with the value. This is what makes thread-safe UDFs possible.

---

## Registration

Excel doesn't discover UDFs automatically. You must register each one in `xlAutoOpen`, which Excel calls when the XLL loads:

```vba
Private udfs As New Collection

[DllExport]
Public Function xlAutoOpen() As Long
    Dim udf As UDF
    Set udf = New UDF
    With udf
        .ProcName = "TBXLL_Multiply"
        .FuncText = "TBXLL_Multiply"
        .Category = "My Add-In"
        .FuncHelp = "Multiplies two numbers"
        .Volatile = False
        .ThreadSafe = True
        .AddArgument Name:="Num1", Help:="First number"
        .AddArgument Name:="Num2", Help:="Second number"
        .Register
    End With
    udfs.Add udf

    xlAutoOpen = 1
End Function
```

The `UDF` class builds the registration type-text string and calls `xlfRegister` for you. The `udfs` collection keeps the registration objects alive for the lifetime of the add-in, which is required because Excel retains pointers into the registration data.

---

## Cleanup

When the add-in unloads, `xlAutoClose` must unregister every UDF:

```vba
[DllExport]
Public Function xlAutoClose() As Long
    Dim udf As UDF
    For Each udf In udfs
        udf.UnRegister
    Next udf
    xlAutoClose = 1
End Function
```

---

## Testing It

After building and loading the XLL in Excel:

1. In any cell, type `=TBXLL_Multiply(3, 4)` and press Enter
2. The cell should display `12`
3. Open the Function Wizard (**fx** button) — you should see `TBXLL_Multiply` under the "My Add-In" category with the help text you provided

---

## The Pattern

Every UDF you write will follow this same three-part pattern:

1. **Function body** — `[DllExport]`, receive `XLOPER12` args, bind them, compute, return via `AllocResultToCaller`
2. **Registration** — create a `UDF` object in `xlAutoOpen`, set properties, call `.Register`, store in collection
3. **Cleanup** — iterate the collection in `xlAutoClose` and call `.UnRegister`

The rest of this wiki explores each part in depth. Start with [[XLL Fundamentals]] for the conceptual foundation, or jump to a specific How-To Guide from the sidebar.
