# twinBASIC XLL Framework

A framework for building Excel XLL add-ins in [twinBASIC](https://twinbasic.com), a modern replacement for VBA/VB6 that supports native DLL exports. This framework provides a complete foundation for developing high-performance, thread-safe Excel UDFs without requiring C/C++.

---

## Overview

Excel XLL add-ins are native DLLs that integrate directly with Excel's calculation engine. They offer significant performance advantages over VBA and COM add-ins, including support for multithreaded calculation. Traditionally, XLL development requires C or C++. This framework enables XLL development entirely in twinBASIC.

The framework handles the low-level XLOPER12 memory layout, argument binding, type coercion, memory management, and Excel callback mechanics, allowing UDF authors to focus on business logic.

---

## Key Features

- **Native twinBASIC** — no C/C++ required, no external build tools
- **Thread-safe UDF support** — dynamic allocation pattern with `xlbitDLLFree` and `xlAutoFree12` enables concurrent recalculation across CPU cores
- **Unified argument binding** — `Bind()` dispatcher handles `btNumber`, `btString`, `btBool`, `btDate`, `btArray`, `btSingleCellRef`, `btValue`
- **Array input and output** — `CoerceToArray` and `GetXLMulti12` handle full round-trip array processing
- **Excel built-in delegation** — pass arguments directly to `xlfSum`, `xlfTranspose`, `xlfRound`, etc.
- **Structured memory management** — two well-defined patterns (Static and Dynamic) with `AllocXLOPER12Result`, `FreeXLMulti12`, and `xlAutoFree12`
- **UDF registration class** — `ThreadSafe`, `Volatile`, `MacroEquivalent`, `Visible` properties with automatic type string construction
- **Comprehensive demo module** — 30+ UDFs demonstrating every major pattern

---

## Requirements

- [twinBASIC](https://twinbasic.com) (64-bit)
- Microsoft Excel (64-bit)
- Jon Johnson's [ExcelSDK](https://github.com/fafalone/TBXLLUDF)

---

## Quick Example

This shows a typical UDF callback:
```vba
[DllExport]
' Converts a number to its Roman Numeral representation - thread-safe
Public Function TBXLL_RomanNumeral(pIn As XLOPER12) As LongPtr
    Dim num As Long
    Dim xTemp As XLOPER12
    ' Convert the input XLOPER12 to a number
    If Bind(pIn, btNumber, num, xTemp) Then
        ' Do the calculations and convert string to XLOPER12 for return to worksheet
        xTemp = GetXLString12(num_getroman(num)) 'num_getroman does all the work (written by Jon Johnson)
    End If
    Return AllocResultToCaller(xTemp)
End Function
```
Here is the corresponding registration pattern:
```vba
Private udfs As New Collection

[DllExport]
Public Function xlAutoOpen() As Long
    ' Required, handles registration
    Dim udf As UDF
    Set udf = New UDF
    With udf
        .ProcName = "TBXLL_RomanNumeral"
        .FuncText = "TBXLL_RomanNumeral"
        .Category = "tB XLL UDF Add-In"
        .FuncHelp = "Converts a number to its Roman Numeral representation"
        .Visible = True
        .Volatile = False
        .ThreadSafe = True '<- this is needed to support fast multi-threaded calculation
        .AddArgument Name:="range", Help:="range"
        .Register
    End With
    udfs.Add udf
    ' ... repeat pattern above for each UDF
    xlAutoOpen = 1
End Function

[DllExport]
Public Function xlAutoClose() As Long
    Dim udf As UDF
    For Each udf In udfs
        udf.UnRegister
    Next udf
    xlAutoClose = 1
End Function
```

## Architecture

### Modules

| Module | Purpose |
|--------|---------|
| `ExcelSDK` | XLOPER12 struct, constants, enums, Excel12v declaration (written by Jon Johnson)|
| `Helpers` | Bind framework, coercion helpers, GetXL* helpers, memory management |
| `Auto_Callbacks` | xlAutoOpen, xlAutoClose, xlAutoRemove, xlAutoFree12, xlAddInManagerInfo12 |
| `UDF` | Convenience wrapper class for UDF registration |
| `Demos` | Demo UDFs illustrating every supported pattern |

### XLOPER12 Layout (twinBASIC 64-bit)

The twinBASIC XLOPER12 struct differs from the C SDK layout:
```vba
Public Type XLOPER12
    val(2) As LongLong   ' 24 bytes — union storage at offsets 0, 8, 16
    xltype As XloperTypes' 4 bytes  — at offset 24
End Type
```

In the C SDK, `xltype` precedes `val`. In twinBASIC, `val` precedes `xltype`. All memory operations in this framework account for this difference explicitly.

---

## Memory Management

Two patterns are supported. Pattern 2 is preferred for all new UDFs.

Use when the UDF maintains persistent state, calls non-thread-safe Excel APIs, or is volatile. Register with `ThreadSafe = False`.

### Pattern 1: Dynamic / xlbitDLLFree (Thread-safe, preferred)
```vba
[DllExport]
Public Function TBXLL_Example(ByRef pN As XLOPER12) As LongPtr
    Dim xTemp As XLOPER12
    Dim n As Double
    If Bind(pN, btNumber, n, xTemp) Then
        xTemp = GetXLNum12(n * 2)
    End If
    Return AllocResultToCaller(xTemp)
End Function
```
Each call allocates an independent heap XLOPER12. Excel calls `xlAutoFree12` when done. Register with `ThreadSafe = True` to enable concurrent recalculation.

### Pattern 2: Static (Non-thread-safe)
```vba
[DllExport]
Public Function TBXLL_Example(ByRef pN As XLOPER12) As LongPtr
    Static xResult As XLOPER12
    Dim n As Double
    If Bind(pN, btNumber, n, xResult) Then
        xResult = GetXLNum12(n * 2)
    End If
    Return VarPtr(xResult)
End Function
```

---

## Argument Binding

`Bind()` is the unified entry point for all argument type coercion:
```vba
Public Function Bind( _
    ByRef pIn As XLOPER12, _
    ByVal target As BindType, _
    ByRef outValue As Variant, _
    ByRef xResult As XLOPER12) As Boolean
```

On failure, `Bind` sets `xResult` to `#VALUE!` automatically. Supported bind types:

| BindType | Output | Notes |
|----------|--------|-------|
| `btNumber` | `Double` | Coerces from num, int, bool, str, ref |
| `btString` | `String` | Coerces from str, num, int, bool, ref |
| `btBool` | `Boolean` | Coerces from bool, num, int, ref |
| `btDate` | `Double` | Excel serial date via numeric coercion |
| `btArray` | `Variant()` | 2D array via xlCoerce to xltypeMulti |
| `btSingleCellRef` | `Variant` | Validates single cell, rejects ranges |
| `btValue` | `Variant` | Extracts single cell value |

---

## Examples

### Scalar numeric UDF
```vba
' Example: =TBXLL_Multiply(3, 4) -> 12
[DllExport]
Public Function TBXLL_Multiply( _
    ByRef pA As XLOPER12, _
    ByRef pB As XLOPER12) As LongPtr
    Dim xTemp As XLOPER12
    Dim a As Double, b As Double
    If Not Bind(pA, btNumber, a, xTemp) Then GoTo ReturnResult
    If Not Bind(pB, btNumber, b, xTemp) Then GoTo ReturnResult
    xTemp = GetXLNum12(a * b)
ReturnResult:
    Return AllocResultToCaller(xTemp)
End Function
```

### Optional argument
```vba
' Example: =TBXLL_AddOptional(1, 2) -> 3  |  =TBXLL_AddOptional(1, 2, 3) -> 6
[DllExport]
Public Function TBXLL_AddOptional( _
    ByRef pA As XLOPER12, _
    ByRef pB As XLOPER12, _
    ByRef pC As XLOPER12) As LongPtr
    Dim xTemp As XLOPER12
    Dim a As Double, b As Double, c As Double
    If Not Bind(pA, btNumber, a, xTemp) Then GoTo ReturnResult
    If Not Bind(pB, btNumber, b, xTemp) Then GoTo ReturnResult
    If pC.xltype = xltypeMissing Then
        c = 0
    ElseIf Not Bind(pC, btNumber, c, xTemp) Then
        GoTo ReturnResult
    End If
    xTemp = GetXLNum12(a + b + c)
ReturnResult:
    Return AllocResultToCaller(xTemp)
End Function
```

### Array input and output
```vba
' Example: =TBXLL_MultiplyArrays({1,2;3,4}, {2,2;2,2}) -> {2,4;6,8}  [Ctrl-Shift-Enter]
[DllExport]
Public Function TBXLL_MultiplyArrays( _
    ByRef pArr1 As XLOPER12, _
    ByRef pArr2 As XLOPER12) As LongPtr
    Dim xTemp As XLOPER12
    Dim arr1() As Variant, arr2() As Variant
    If Not Bind(pArr1, btArray, arr1, xTemp) Then GoTo ReturnResult
    If Not Bind(pArr2, btArray, arr2, xTemp) Then GoTo ReturnResult
    If UBound(arr1, 1) <> UBound(arr2, 1) Or _
       UBound(arr1, 2) <> UBound(arr2, 2) Then
        SetErrorResult xTemp
        GoTo ReturnResult
    End If
    Dim rows As Long = UBound(arr1, 1) + 1
    Dim cols As Long = UBound(arr1, 2) + 1
    Dim arrOut() As Variant
    ReDim arrOut(rows - 1, cols - 1)
    Dim r As Long, c As Long
    For r = 0 To rows - 1
        For c = 0 To cols - 1
            If VarType(arr1(r, c)) = vbDouble And _
               VarType(arr2(r, c)) = vbDouble Then
                arrOut(r, c) = CDbl(arr1(r, c)) * CDbl(arr2(r, c))
            Else
                arrOut(r, c) = CVErr(xlerrValue)
            End If
        Next c
    Next r
    Dim xMulti As XLOPER12
    xMulti = GetXLMulti12(arrOut)
    If xMulti.xltype <> xltypeMulti Then
        SetErrorResult xTemp
        GoTo ReturnResult
    End If
    xTemp = xMulti  ' ownership transfers to xlAutoFree12
ReturnResult:
    Return AllocResultToCaller(xTemp)
End Function
```

### Delegating to an Excel built-in
```vba
' Example: =TBXLL_SumArray(A1:A10) -> SUM(A1:A10)
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

### Computationally intensive thread-safe UDF
```vba
' Example: =TBXLL_SlowCalcSafe(2) -> result, runs concurrently across cells
[DllExport]
Public Function TBXLL_SlowCalcSafe(ByRef pN As XLOPER12) As LongPtr
    Dim xTemp As XLOPER12
    Dim n As Double
    If Not Bind(pN, btNumber, n, xTemp) Then GoTo ReturnResult
    Dim i As Long, total As Double
    For i = 1 To 1000000
        total = total + Sqr(i) * n
    Next i
    xTemp = GetXLNum12(total)
ReturnResult:
    Return AllocResultToCaller(xTemp)
End Function
```

---

## UDF Registration

UDFs are registered in `xlAutoOpen` using the `UDF` class:
```vba
Dim u As New UDF
With u
    .ProcName    = "TBXLL_Multiply"
    .FuncText    = "TBXLL_Multiply"
    .Category    = "My Add-In"
    .FuncHelp    = "Demo: btNumber binding, scalar return"
    .Visible     = True
    .Volatile    = False
    .ThreadSafe  = True
    .AddArgument Name:="a", Help:="First number"
    .AddArgument Name:="b", Help:="Second number"
    .Register
End With
```

### Registration properties

| Property | Type | Notes |
|----------|------|-------|
| `ProcName` | String | Exported function name |
| `FuncText` | String | Name shown in Function Wizard |
| `Category` | String | Function Wizard category |
| `FuncHelp` | String | Function description |
| `Visible` | Boolean | Show in Function Wizard |
| `Volatile` | Boolean | Adds `!` to type string |
| `ThreadSafe` | Boolean | Adds `$` to type string, mutually exclusive with MacroEquivalent |
| `MacroEquivalent` | Boolean | Adds `#` to type string, enables macro-only API calls |

---

## xlAutoFree12

`xlAutoFree12` is called by Excel when it finishes with any XLOPER12 result that has `xlbitDLLFree` set. It handles all return types:
```vba
[DllExport]
Public Sub xlAutoFree12(ByVal pResult As LongPtr)
    If pResult = 0 Then Exit Sub
    Dim xltype As Long
    CopyMemory xltype, ByVal (pResult + 24), 4  ' xltype at offset 24 in twinBASIC layout
    xltype = xltype And &H0FFF
    Select Case xltype
        Case xltypeStr
            Dim lpStr As LongPtr
            CopyMemory lpStr, ByVal pResult, LenB(Of LongPtr)
            If lpStr <> 0 Then GlobalFree lpStr
        Case xltypeMulti
            ' frees element array and any embedded string buffers
            ...
    End Select
    GlobalFree pResult
End Sub
```

Note: `xlAutoFree12` is declared `ByVal LongPtr` rather than `ByRef XLOPER12` due to a twinBASIC pointer semantics difference from the C SDK. This is verified ABI-equivalent from Excel's perspective.

---

## Thread Safety

Multithreaded recalculation performance was verified by comparing `TBXLL_SlowCalcUnsafe` (Pattern 1, `ThreadSafe = False`) against `TBXLL_SlowCalcSafe` (Pattern 2, `ThreadSafe = True`) across 100 cells. On a multi-core machine the thread-safe version recalculates dramatically faster due to concurrent execution across all available cores.

Memory correctness of `xlAutoFree12` was verified by monitoring Excel process memory in Task Manager across repeated F9 recalculations with and without `xlAutoFree12` exported — stable memory with, steady growth without.

---

## Known Limitations

- `xltypeRef` (multi-sheet or external references) is not supported in `CoerceToSingleCellRef`
- `btSingleCellRef` returns a dummy value due to XLREF12 round-trip issues through twinBASIC Variant; use `btValue` to extract the cell's value
- `xlAutoFree12` for `xltypeMulti` frees embedded `xltypeStr` elements but does not recurse into nested `xltypeMulti` elements (not a practical limitation as Excel does not produce nested multis)
- `MacroEquivalent` (`#`) and `ThreadSafe` (`$`) are mutually exclusive in Excel's registration model

---

## License

MIT

