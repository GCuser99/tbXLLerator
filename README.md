# \# twinBASIC XLL Framework

# 

# A framework for building Excel XLL add-ins in \[twinBASIC](https://twinbasic.com), a modern replacement for VBA/VB6 that supports native DLL exports. This framework provides a complete foundation for developing high-performance, thread-safe Excel UDFs without requiring C/C++.

# 

# ---

# 

# \## Overview

# 

# Excel XLL add-ins are native DLLs that integrate directly with Excel's calculation engine. They offer significant performance advantages over VBA and COM add-ins, including support for multithreaded recalculation. Traditionally, XLL development requires C or C++. This framework enables XLL development entirely in twinBASIC.

# 

# The framework handles the low-level XLOPER12 memory layout, argument binding, type coercion, memory management, and Excel callback mechanics, allowing UDF authors to focus on business logic.

# 

# ---

# 

# \## Key Features

# 

# \- \*\*Native twinBASIC\*\* — no C/C++ required, no external build tools

# \- \*\*Thread-safe UDF support\*\* — dynamic allocation pattern with `xlbitDLLFree` and `xlAutoFree12` enables concurrent recalculation across CPU cores

# \- \*\*Unified argument binding\*\* — `Bind()` dispatcher handles `btNumber`, `btString`, `btBool`, `btDate`, `btArray`, `btSingleCellRef`, `btValue`

# \- \*\*Array input and output\*\* — `CoerceToArray` and `GetXLMulti12` handle full round-trip array processing

# \- \*\*Excel built-in delegation\*\* — pass arguments directly to `xlfSum`, `xlfTranspose`, `xlfRound`, etc.

# \- \*\*Structured memory management\*\* — two well-defined patterns (Static and Dynamic) with `AllocXLOPER12Result`, `FreeXLMulti12`, and `xlAutoFree12`

# \- \*\*UDF registration class\*\* — `ThreadSafe`, `Volatile`, `MacroEquivalent`, `Visible` properties with automatic type string construction

# \- \*\*Comprehensive demo module\*\* — 30+ UDFs demonstrating every major pattern

# 

# ---

# 

# \## Requirements

# 

# \- \[twinBASIC](https://twinbasic.com) (64-bit)

# \- Microsoft Excel (64-bit)

# 

# ---

# 

# \## Architecture

# 

# \### Modules

# 

# | Module | Purpose |

# |--------|---------|

# | `ExcelSDK` | XLOPER12 struct, constants, enums, Excel12v declaration |

# | `Helpers` | Bind framework, coercion helpers, GetXL\* helpers, memory management |

# | `MainModule` | xlAutoOpen, xlAutoRemove, xlAutoFree12, xlAddInManagerInfo12 |

# | `Usage` | Demo UDFs illustrating every supported pattern |

# 

# \### XLOPER12 Layout (twinBASIC 64-bit)

# 

# The twinBASIC XLOPER12 struct differs from the C SDK layout:

# ```vb

# Public Type XLOPER12

# &nbsp;   val(2) As LongLong   ' 24 bytes — union storage at offsets 0, 8, 16

# &nbsp;   xltype As XloperTypes' 4 bytes  — at offset 24

# End Type

# ```

# 

# In the C SDK, `xltype` precedes `val`. In twinBASIC, `val` precedes `xltype`. All memory operations in this framework account for this difference explicitly.

# 

# ---

# 

# \## Memory Management

# 

# Two patterns are supported. Pattern 2 is preferred for all new UDFs.

# 

# \### Pattern 1: Static (Non-thread-safe)

# ```vb

# \[DllExport]

# Public Function TBXLL\_Example(ByRef pN As XLOPER12) As LongPtr

# &nbsp;   Static xResult As XLOPER12

# &nbsp;   Dim n As Double

# &nbsp;   If Not Bind(pN, btNumber, n, xResult) Then Return VarPtr(xResult)

# &nbsp;   xResult = GetXLNum12(n \* 2)

# &nbsp;   Return VarPtr(xResult)

# End Function

# ```

# 

# Use when the UDF maintains persistent state, calls non-thread-safe Excel APIs, or is volatile. Register with `ThreadSafe = False`.

# 

# \### Pattern 2: Dynamic / xlbitDLLFree (Thread-safe, preferred)

# ```vb

# \[DllExport]

# Public Function TBXLL\_Example(ByRef pN As XLOPER12) As LongPtr

# &nbsp;   Dim xTemp As XLOPER12

# &nbsp;   Dim n As Double

# &nbsp;   If Not Bind(pN, btNumber, n, xTemp) Then GoTo ReturnResult

# &nbsp;   xTemp = GetXLNum12(n \* 2)

# ReturnResult:

# &nbsp;   Return AllocXLOPER12Result(xTemp)

# End Function

# ```

# 

# Each call allocates an independent heap XLOPER12. Excel calls `xlAutoFree12` when done. Register with `ThreadSafe = True` to enable concurrent recalculation.

# 

# ---

# 

# \## Argument Binding

# 

# `Bind()` is the unified entry point for all argument type coercion:

# ```vb

# Public Function Bind( \_

# &nbsp;   ByRef pIn As XLOPER12, \_

# &nbsp;   ByVal target As BindType, \_

# &nbsp;   ByRef outValue As Variant, \_

# &nbsp;   ByRef xResult As XLOPER12) As Boolean

# ```

# 

# On failure, `Bind` sets `xResult` to `#VALUE!` automatically. Supported bind types:

# 

# | BindType | Output | Notes |

# |----------|--------|-------|

# | `btNumber` | `Double` | Coerces from num, int, bool, str, ref |

# | `btString` | `String` | Coerces from str, num, int, bool, ref |

# | `btBool` | `Boolean` | Coerces from bool, num, int, ref |

# | `btDate` | `Double` | Excel serial date via numeric coercion |

# | `btArray` | `Variant()` | 2D array via xlCoerce to xltypeMulti |

# | `btSingleCellRef` | `Variant` | Validates single cell, rejects ranges |

# | `btValue` | `Variant` | Extracts single cell value |

# 

# ---

# 

# \## Examples

# 

# \### Scalar numeric UDF

# ```vb

# ' Example: =TBXLL\_Multiply(3, 4) -> 12

# \[DllExport]

# Public Function TBXLL\_Multiply( \_

# &nbsp;   ByRef pA As XLOPER12, \_

# &nbsp;   ByRef pB As XLOPER12) As LongPtr

# &nbsp;   Dim xTemp As XLOPER12

# &nbsp;   Dim a As Double, b As Double

# &nbsp;   If Not Bind(pA, btNumber, a, xTemp) Then GoTo ReturnResult

# &nbsp;   If Not Bind(pB, btNumber, b, xTemp) Then GoTo ReturnResult

# &nbsp;   xTemp = GetXLNum12(a \* b)

# ReturnResult:

# &nbsp;   Return AllocXLOPER12Result(xTemp)

# End Function

# ```

# 

# \### Optional argument

# ```vb

# ' Example: =TBXLL\_AddOptional(1, 2) -> 3  |  =TBXLL\_AddOptional(1, 2, 3) -> 6

# \[DllExport]

# Public Function TBXLL\_AddOptional( \_

# &nbsp;   ByRef pA As XLOPER12, \_

# &nbsp;   ByRef pB As XLOPER12, \_

# &nbsp;   ByRef pC As XLOPER12) As LongPtr

# &nbsp;   Dim xTemp As XLOPER12

# &nbsp;   Dim a As Double, b As Double, c As Double

# &nbsp;   If Not Bind(pA, btNumber, a, xTemp) Then GoTo ReturnResult

# &nbsp;   If Not Bind(pB, btNumber, b, xTemp) Then GoTo ReturnResult

# &nbsp;   If pC.xltype = xltypeMissing Then

# &nbsp;       c = 0

# &nbsp;   ElseIf Not Bind(pC, btNumber, c, xTemp) Then

# &nbsp;       GoTo ReturnResult

# &nbsp;   End If

# &nbsp;   xTemp = GetXLNum12(a + b + c)

# ReturnResult:

# &nbsp;   Return AllocXLOPER12Result(xTemp)

# End Function

# ```

# 

# \### Array input and output

# ```vb

# ' Example: =TBXLL\_MultiplyArrays({1,2;3,4}, {2,2;2,2}) -> {2,4;6,8}  \[Ctrl-Shift-Enter]

# \[DllExport]

# Public Function TBXLL\_MultiplyArrays( \_

# &nbsp;   ByRef pArr1 As XLOPER12, \_

# &nbsp;   ByRef pArr2 As XLOPER12) As LongPtr

# &nbsp;   Dim xTemp As XLOPER12

# &nbsp;   Dim arr1() As Variant, arr2() As Variant

# &nbsp;   If Not Bind(pArr1, btArray, arr1, xTemp) Then GoTo ReturnResult

# &nbsp;   If Not Bind(pArr2, btArray, arr2, xTemp) Then GoTo ReturnResult

# &nbsp;   If UBound(arr1, 1) <> UBound(arr2, 1) Or \_

# &nbsp;      UBound(arr1, 2) <> UBound(arr2, 2) Then

# &nbsp;       SetErrorResult xTemp

# &nbsp;       GoTo ReturnResult

# &nbsp;   End If

# &nbsp;   Dim rows As Long = UBound(arr1, 1) + 1

# &nbsp;   Dim cols As Long = UBound(arr1, 2) + 1

# &nbsp;   Dim arrOut() As Variant

# &nbsp;   ReDim arrOut(rows - 1, cols - 1)

# &nbsp;   Dim r As Long, c As Long

# &nbsp;   For r = 0 To rows - 1

# &nbsp;       For c = 0 To cols - 1

# &nbsp;           If VarType(arr1(r, c)) = vbDouble And \_

# &nbsp;              VarType(arr2(r, c)) = vbDouble Then

# &nbsp;               arrOut(r, c) = CDbl(arr1(r, c)) \* CDbl(arr2(r, c))

# &nbsp;           Else

# &nbsp;               arrOut(r, c) = CVErr(xlerrValue)

# &nbsp;           End If

# &nbsp;       Next c

# &nbsp;   Next r

# &nbsp;   Dim xMulti As XLOPER12

# &nbsp;   xMulti = GetXLMulti12(arrOut)

# &nbsp;   If xMulti.xltype <> xltypeMulti Then

# &nbsp;       SetErrorResult xTemp

# &nbsp;       GoTo ReturnResult

# &nbsp;   End If

# &nbsp;   xTemp = xMulti  ' ownership transfers to xlAutoFree12

# ReturnResult:

# &nbsp;   Return AllocXLOPER12Result(xTemp)

# End Function

# ```

# 

# \### Delegating to an Excel built-in

# ```vb

# ' Example: =TBXLL\_SumArray(A1:A10) -> SUM(A1:A10)

# \[DllExport]

# Public Function TBXLL\_SumArray(ByRef pArr As XLOPER12) As LongPtr

# &nbsp;   Dim xTemp As XLOPER12

# &nbsp;   Dim args(0) As XLOPER12

# &nbsp;   args(0) = pArr

# &nbsp;   If Excel12v(xlfSum, xTemp, 1, args) <> 0 Then

# &nbsp;       SetErrorResult xTemp

# &nbsp;   End If

# &nbsp;   Return AllocXLOPER12Result(xTemp)

# End Function

# ```

# 

# \### Computationally intensive thread-safe UDF

# ```vb

# ' Example: =TBXLL\_SlowCalcSafe(2) -> result, runs concurrently across cells

# \[DllExport]

# Public Function TBXLL\_SlowCalcSafe(ByRef pN As XLOPER12) As LongPtr

# &nbsp;   Dim xTemp As XLOPER12

# &nbsp;   Dim n As Double

# &nbsp;   If Not Bind(pN, btNumber, n, xTemp) Then GoTo ReturnResult

# &nbsp;   Dim i As Long, total As Double

# &nbsp;   For i = 1 To 1000000

# &nbsp;       total = total + Sqr(i) \* n

# &nbsp;   Next i

# &nbsp;   xTemp = GetXLNum12(total)

# ReturnResult:

# &nbsp;   Return AllocXLOPER12Result(xTemp)

# End Function

# ```

# 

# ---

# 

# \## UDF Registration

# 

# UDFs are registered in `xlAutoOpen` using the `UDF` class:

# ```vb

# Dim u As New UDF

# With u

# &nbsp;   .ProcName    = "TBXLL\_Multiply"

# &nbsp;   .FuncText    = "TBXLL\_Multiply"

# &nbsp;   .Category    = "My Add-In"

# &nbsp;   .FuncHelp    = "Demo: btNumber binding, scalar return"

# &nbsp;   .Visible     = True

# &nbsp;   .Volatile    = False

# &nbsp;   .ThreadSafe  = True

# &nbsp;   .AddArgument Name:="a", Help:="First number"

# &nbsp;   .AddArgument Name:="b", Help:="Second number"

# &nbsp;   .Register

# End With

# ```

# 

# \### Registration properties

# 

# | Property | Type | Notes |

# |----------|------|-------|

# | `ProcName` | String | Exported function name |

# | `FuncText` | String | Name shown in Function Wizard |

# | `Category` | String | Function Wizard category |

# | `FuncHelp` | String | Function description |

# | `Visible` | Boolean | Show in Function Wizard |

# | `Volatile` | Boolean | Adds `!` to type string |

# | `ThreadSafe` | Boolean | Adds `$` to type string, mutually exclusive with MacroEquivalent |

# | `MacroEquivalent` | Boolean | Adds `#` to type string, enables macro-only API calls |

# 

# ---

# 

# \## xlAutoFree12

# 

# `xlAutoFree12` is called by Excel when it finishes with any XLOPER12 result that has `xlbitDLLFree` set. It handles all return types:

# ```vb

# \[DllExport]

# Public Sub xlAutoFree12(ByVal pResult As LongPtr)

# &nbsp;   If pResult = 0 Then Exit Sub

# &nbsp;   Dim xltype As Long

# &nbsp;   CopyMemory xltype, ByVal (pResult + 24), 4  ' xltype at offset 24 in twinBASIC layout

# &nbsp;   xltype = xltype And \&H0FFF

# &nbsp;   Select Case xltype

# &nbsp;       Case xltypeStr

# &nbsp;           Dim lpStr As LongPtr

# &nbsp;           CopyMemory lpStr, ByVal pResult, LenB(Of LongPtr)

# &nbsp;           If lpStr <> 0 Then GlobalFree lpStr

# &nbsp;       Case xltypeMulti

# &nbsp;           ' frees element array and any embedded string buffers

# &nbsp;           ...

# &nbsp;   End Select

# &nbsp;   GlobalFree pResult

# End Sub

# ```

# 

# Note: `xlAutoFree12` is declared `ByVal LongPtr` rather than `ByRef XLOPER12` due to a twinBASIC pointer semantics difference from the C SDK. This is verified ABI-equivalent from Excel's perspective.

# 

# ---

# 

# \## Thread Safety

# 

# Multithreaded recalculation performance was verified by comparing `TBXLL\_SlowCalcUnsafe` (Pattern 1, `ThreadSafe = False`) against `TBXLL\_SlowCalcSafe` (Pattern 2, `ThreadSafe = True`) across 100 cells. On a multi-core machine the thread-safe version recalculates dramatically faster due to concurrent execution across all available cores.

# 

# Memory correctness of `xlAutoFree12` was verified by monitoring Excel process memory in Task Manager across repeated F9 recalculations with and without `xlAutoFree12` exported — stable memory with, steady growth without.

# 

# ---

# 

# \## Known Limitations

# 

# \- `xltypeRef` (multi-sheet or external references) is not supported in `CoerceToSingleCellRef`

# \- `btSingleCellRef` returns a dummy value due to XLREF12 round-trip issues through twinBASIC Variant; use `btValue` to extract the cell's value

# \- `xlAutoFree12` for `xltypeMulti` frees embedded `xltypeStr` elements but does not recurse into nested `xltypeMulti` elements (not a practical limitation as Excel does not produce nested multis)

# \- `MacroEquivalent` (`#`) and `ThreadSafe` (`$`) are mutually exclusive in Excel's registration model

# 

# ---

# 

# \## License

# 

# MIT

