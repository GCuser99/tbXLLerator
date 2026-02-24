# tbXLLerator - A twinBASIC XLL Framework

A framework for building Excel XLL add-ins in [twinBASIC](https://twinbasic.com), a modern replacement for VBA/VB6 that supports native DLL exports. This framework provides a complete foundation for developing high-performance, thread-safe Excel UDFs without requiring C/C++.

---

## Overview

Excel XLL add-ins are native DLLs that integrate directly with Excel's calculation engine. They offer significant advantages over VBA and COM add-ins, including support for high-performance multithreaded calculation, and Excel's Function Wizard support. Traditionally, XLL development requires C or C++. This framework enables XLL development entirely in twinBASIC.

The framework handles the low-level XLOPER12 memory layout, argument binding, type coercion, memory management, and Excel callback mechanics, allowing UDF authors to focus on modeling logic.

---

## Why?

I am a scientist, not a developer. I wrote this as my first [completed] AI-assisted coding project, which I probably could not have done with this quality and speed on my own. And certainly not without the non-artificial intelligence of Jon Johnson who wrote the [ExcelSDK](https://github.com/fafalone/TBXLLUDF) that this framework wraps!

My goal in writing this wrapper was to easily design and use high-performance UDF's for large and complex spreadsheet models in a language that I am familiar with - [twinBASIC](https://twinbasic.com). Specifically, I'm using UDFs in spreadsheet models along with the (as yet unstarred! :-) [SolverWrapper](https://github.com/GCuser99/SolverWrapper) for model parameter optimization, which requires very fast worksheet execution.

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
- Microsoft Excel (64-bit) in MS Windows
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
        ' num_getroman does all the work (written by Jon Johnson)
        xTemp = GetXLString12(num_getroman(num))
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
| `ExcelSDK` | XLOPER12 struct, constants, enums, Excel12v declaration (**written by Jon Johnson**)|
| `Helpers` | Bind framework, coercion helpers, GetXL* helpers, memory management |
| `Auto_Callbacks` | xlAutoOpen, xlAutoClose, xlAutoRemove, xlAutoFree12, xlAddInManagerInfo12 |
| `UDF` | Convenience wrapper class for UDF registration |
| `Demos` | Demo UDFs illustrating every supported pattern |

---

## Memory Management

Two patterns are supported. Pattern 1 is preferred for modern UDFs, with a few exceptions.

### Pattern 1: Dynamic / xlbitDLLFree (Thread-safe, preferred)
Each call allocates an independent heap XLOPER12. Excel calls `xlAutoFree12` when done. Register with `ThreadSafe = True` to enable concurrent recalculation.
```vba
[DllExport]
Public Function TBXLL_Example(ByRef pN As XLOPER12) As LongPtr
    Dim xTemp As XLOPER12
    Dim n As Double
    If Bind(pN, btNumber, n, xTemp) Then
        xTemp = GetXLNum12(n * 2)
    End If
    Return AllocResultToCaller(xTemp) '<-- required
End Function
```

### Pattern 2: Static (Non-thread-safe)
Use only when the UDF requires shared state across calls (e.g. a recalculation counter).
All other UDFs should use Pattern 2 regardless of how they are registered.
```vba
' Demonstrates: Static variable persistence across recalculations with volatile registration (udf.Volatile)
' Note: Volatile functions cannot be registered with .SafeThread = True; Use .Volatile = True
[DllExport]
Public Function TBXLL_RecalcCounter() As LongPtr
    Static xResult As XLOPER12 '<-- required
    Static counter As Long

    counter = counter + 1
    xResult = GetXLInt12(counter)

    Return VarPtr(xResult) '<-- required
End Function
```

---

## Argument Binding

`Bind()` is the unified entry point for all UDF argument type coercion:
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

## More Examples

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
' Demonstrates: btArray binding for two ranges, dimension validation, GetXLMulti12 array return
' Example: =TBXLL_MultiplyArrays({1,2;3,4}, {2,2;2,2}) -> {2,4;6,8}  [Ctrl-Shift-Enter]
[DllExport]
Public Function TBXLL_MultiplyArrays( _
    ByRef pArr1 As XLOPER12, _
    ByRef pArr2 As XLOPER12) As LongPtr

    Dim xTemp As XLOPER12
    Dim arr1() As Variant
    Dim arr2() As Variant

    ' Bind both input arrays
    If Not Bind(pArr1, btArray, arr1, xTemp) Then GoTo ReturnResult
    If Not Bind(pArr2, btArray, arr2, xTemp) Then GoTo ReturnResult

    ' Validate dimensions match
    If UBound(arr1, 1) <> UBound(arr2, 1) Or _
       UBound(arr1, 2) <> UBound(arr2, 2) Then
        SetErrorResult xTemp
        GoTo ReturnResult
    End If

    ' Build result variant array
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

    ' Convert result array to XLOPER12
    Dim xMulti As XLOPER12
    xMulti = GetXLMulti12(arrOut)

    If xMulti.xltype <> xltypeMulti Then
        SetErrorResult xTemp
        GoTo ReturnResult
    End If

    ' Do NOT call FreeXLMulti12 here - xlAutoFree12 will free the element array
    xTemp = xMulti
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

---

## UDF Registration

UDFs are registered in `xlAutoOpen` using the `UDF` class:
```vba
Dim udf As New UDF
With udf
    .ProcName    = "TBXLL_Multiply"
    .FuncText    = "TBXLL_Multiply"
    .Category    = "My Add-In"
    .FuncHelp    = "Demo: btNumber binding, scalar return"
    .Volatile    = False
    .ThreadSafe  = True
    .AddArgument Name:="a", Help:="First number"
    .AddArgument Name:="b", Help:="Second number"
    .Register
End With
```

### Registration properties and methods

| Property | Type | Notes |
|----------|------|-------|
| `ProcName` | String | Exported function name |
| `FuncText` | String | Name shown in Function Wizard |
| `Category` | String | Function Wizard category |
| `FuncHelp` | String | Function description |
| `Visible` | Boolean | Show in Function Wizard (default = True) |
| `Volatile` | Boolean | Adds `!` to type string |
| `ThreadSafe` | Boolean | Adds `$` to type string, mutually exclusive with MacroEquivalent |
| `MacroEquivalent` | Boolean | Adds `#` to type string, enables macro-only API calls |

| Methods | Arguments | Notes |
|----------|------|-------|
| `AddArgument` | name, help | Adds a new argument definition to the UDF |
| `Register` | N/A | Registers the UDF for use as a worksheet function |
| `Unregister` | N/A | Unregisters the UDF for use as a worksheet function |

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

---

## Installation

1. Install [twinBASIC](https://twinbasic.com)
2. Clone or download this repository
3. Clone or download Jon Johnson's [ExcelSDK](https://github.com/fafalone/TBXLLUDF) and add the SDK source files to your twinBASIC project
4. Open the `.twinproj` file in twinBASIC
5. Set the bitness to 64-bit
6. Build the project — twinBASIC will produce a `.xll` file in the output folder
7. In Excel, go to **File → Options → Add-ins → Manage: Excel Add-ins → Go**
8. Click **Browse** and select the `.xll` file
9. The add-in will load and UDFs will be available in the Function Wizard under the category defined in `xlAutoOpen`

### To use as a starting point for your own XLL

- Copy the `Helpers` module and `UDF` class into your own twinBASIC project
- Add your UDF functions following the patterns in the `Demos` module
- Register each UDF in `xlAutoOpen` using the `UDF` class
- Unregister in `xlAutoClose` by iterating the `udfs` collection

### Notes

- The `.xll` must match Excel's bitness — this framework currently targets 64-bit only
- Excel must be fully closed before replacing or updating the `.xll` file
- If Excel shows `#NAME?` after reloading, close Excel fully and reopen
---

## Limitations

This has not yet been tested on a 32-bit version of Excel.

---

## License

[MIT License](https://github.com/GCuser99/tbXLLerator?tab=MIT-1-ov-file)
