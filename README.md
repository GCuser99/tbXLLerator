# tbXLLerator - A twinBASIC XLL Framework

A framework for building Excel XLL add-ins using [twinBASIC](https://twinbasic.com), a modern replacement for VBA/VB6 that supports native DLL exports. This framework provides a complete foundation for developing high-performance, thread-safe **Excel-User-Defined-Functions** (UDFs) without requiring C/C++.

---

## Overview

Excel XLL add-ins are native DLLs that integrate directly with Excel's calculation engine. They offer significant advantages over VBA and COM add-ins, including support for high-performance multithreaded calculation, and Excel's Function Wizard support. Traditionally, XLL development requires C or C++. This framework enables XLL development entirely in [twinBASIC](https://twinbasic.com).

This framework wraps the very excellent [ExcelSDK](https://github.com/fafalone/TBXLLUDF) written by Jon Johnson. It handles argument binding, type coercion, memory management, and Excel callback mechanics, allowing UDF authors to focus on modeling logic.

---

## Why?

My goal in writing this wrapper was to be able to easily design and use high-performance UDF's for large and complex spreadsheet models in a language that I am familiar with - [twinBASIC](https://twinbasic.com). More specifically, I'm using UDFs in spreadsheet models along with the [SolverWrapper](https://github.com/GCuser99/SolverWrapper) for model parameter optimization, which requires very fast worksheet execution.

---

## Key Features

- **Native twinBASIC** — no C/C++ required, no external build tools
- **Thread-safe UDF support** — dynamic allocation pattern with `xlbitDLLFree` and `xlAutoFree12` enables concurrent recalculation across CPU cores
- **Unified argument binding** — `BindU()` and `BindQ()` dispatchers handle `btNumber`, `btString`, `btBool`, `btDate`, `btArray`, `btSingleCellRef`, `btValue`
- **Array input and output** — `CoerceToArray` and `GetXLMulti12` handle full round-trip array processing
- **Excel built-in delegation** — pass arguments directly to `xlfSum`, `xlfTranspose`, `xlfRound`, etc.
- **Structured memory management** — a well-defined pattern with `AllocXLOPER12Result` and `xlAutoFree12`
- **UDF registration class** — High-level mechanism including `ThreadSafe` and `Volatile` properties and automatic type-text construction
- **Comprehensive demo module** — 30+ UDFs demonstrating every major pattern
- **Complete Wiki Documentation** — See the [tBXLLerator Wiki](https://github.com/GCuser99/tbXLLerator/wiki) for help on writing your first XLL UDF.

---

## Requirements

- Wayne Phillip's [twinBASIC](https://twinbasic.com)
- Microsoft Excel 2010 or greater (both 32-bit and 64-bit) in MS Windows
- Jon Johnson's [ExcelSDK.twin](https://github.com/fafalone/TBXLLUDF)

---

## Quick Example

This shows a typical UDF callback:
```vba
' Converts a number to its Roman Numeral representation - thread-safe
' In Excel: =TBXLL_RomanNumeral(9) --> "IX"
[DllExport]
Public Function TBXLL_RomanNumeral(pIn As XLOPER12) As LongPtr
    Dim num As Long
    Dim xTemp As XLOPER12
    ' Convert the input XLOPER12 to a number
    If BindU(pIn, btNumber, num, xTemp) Then
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
        .Category = "tB XLL UDF Add-In"
        .FuncHelp = "Converts a number to its Roman Numeral representation"
        .Volatile = False
        .ThreadSafe = True '<-- supports fast multi-threaded calculation
        .AddArgument Name:="number", Help:="Number to convert"
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
---

## Architecture

### Modules

| Module | Purpose |
|--------|---------|
| `ExcelSDK` | XLOPER12 struct, constants, enums, Excel12v declaration (**written by Jon Johnson**)|
| `Helpers` | Bind framework, coercion helpers, higher-level memory management |
| `AutoCallbacks` | xlAutoOpen, xlAutoClose, xlAutoRemove, xlAutoFree12, xlAddInManagerInfo12 |
| `UDFReg` | Convenience wrapper classes for UDF registration |
| `Demos` | Demo UDFs illustrating every supported pattern |

---

## More Information?

See the comprehensive [tbXLLerator Wiki](https://github.com/GCuser99/tbXLLerator/wiki) for the details.

---

## License

This project is licensed under the [MIT License](https://github.com/GCuser99/tbXLLerator?tab=MIT-1-ov-file).
