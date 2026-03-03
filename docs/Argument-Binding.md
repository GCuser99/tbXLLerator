# Argument Binding

When Excel calls your UDF, it passes raw `XLOPER12` structures. These can contain any type — a number, a string, a boolean, an error, a cell reference, or even a 2D array. Before you can use the data, you need to coerce it into a typed value. That's what `BindU()` and `BindQ()` do.

---

## BindU — The Primary Dispatcher

`BindU` is the unified entry point for all type coercion when your arguments are registered with type `U` (raw `XLOPER12`):

```vba
Public Function BindU( _
    ByRef pIn As XLOPER12, _
    ByVal target As BindType, _
    ByRef outValue As Variant, _
    ByRef xResult As XLOPER12) As Boolean
```

- **`pIn`** — the raw `XLOPER12` from Excel
- **`target`** — what type you want (e.g., `btNumber`, `btString`, `btArray`)
- **`outValue`** — receives the coerced value on success
- **`xResult`** — set to `#VALUE!` automatically on failure

Returns `True` on success, `False` on failure. On failure, `xResult` is already set to an error — you can go straight to `ReturnResult`.

### Basic usage

```vba
Dim n As Double
If Not BindU(pIn, btNumber, n, xTemp) Then GoTo ReturnResult
' n is now a Double — use it
```

```vba
Dim s As String
If Not BindU(pIn, btString, s, xTemp) Then GoTo ReturnResult
' s is now a String — use it
```

```vba
Dim arr() As Variant
If Not BindU(pRange, btArray, arr, xTemp) Then GoTo ReturnResult
' arr is now a 2D Variant array — iterate it
```

---

## Bind Types

| BindType | Output Type | What it accepts | Coercion behavior |
|----------|-------------|----------------|-------------------|
| `btNumber` | `Double` | num, int, bool, str, ref, multi | Uses `xlCoerce` for refs/strings |
| `btString` | `String` | str, num, int, bool, ref, multi | Converts numbers to text |
| `btBool` | `Boolean` | bool, num, int, ref, multi | Non-zero = True |
| `btDate` | `Double` | Same as btNumber | Validates range 0–2958465 |
| `btArray` | `Variant()` | ref, sref, multi, scalars | Always returns 2D (0-based) |
| `btSingleCellRef` | `Variant` | sref, ref | Rejects multi-cell ranges |
| `btValue` | `Variant` | Any scalar or single-cell | Returns typed Variant |

### btNumber

Coerces the input to an IEEE 754 `Double`. Accepts numeric types directly, coerces cell references and strings through Excel's `xlCoerce` mechanism. This mirrors how Excel's own built-in functions accept numeric arguments.

```vba
Dim a As Double
If Not BindU(pA, btNumber, a, xTemp) Then GoTo ReturnResult
```

### btString

Coerces the input to a `String`. Numbers are converted via `CStr()`, booleans become `"TRUE"` or `"FALSE"`, and cell references are resolved through `xlCoerce`.

```vba
Dim label As String
If Not BindU(pLabel, btString, label, xTemp) Then GoTo ReturnResult
```

### btBool

Coerces to `Boolean`. Numbers coerce to `True` if non-zero, `False` if zero. Strings and references go through `xlCoerce` to numeric first, then to boolean.

```vba
Dim flag As Boolean
If Not BindU(pFlag, btBool, flag, xTemp) Then GoTo ReturnResult
```

### btDate

Same as `btNumber` but validates that the result falls within Excel's date range (0 to 2958465, which is December 31, 9999). Excel stores dates as serial numbers — the integer part is the day count from January 0, 1900, and the fractional part is the time of day.

```vba
Dim dt As Double
If Not BindU(pDate, btDate, dt, xTemp) Then GoTo ReturnResult
' dt is an Excel serial date — add days directly: dt + 30
```

### btArray

Coerces the input to a 2D `Variant()` array (0-based in both dimensions). Cell references are resolved, and scalar inputs become 1×1 arrays. Each element preserves its type: `vbDouble` for numbers, `vbString` for text, `vbError` for errors, `vbEmpty` for blank cells.

```vba
Dim arr() As Variant
If Not BindU(pRange, btArray, arr, xTemp) Then GoTo ReturnResult
' arr(0, 0) is the top-left cell
' UBound(arr, 1) is the last row index
' UBound(arr, 2) is the last column index
```

See [[Working with Arrays]] for detailed patterns.

### btSingleCellRef

Validates that the input is a reference to exactly one cell. If the user passes a multi-cell range (e.g., `A1:B2`), the bind fails with `#VALUE!`. On success, returns a 2-element `Variant` array containing the 1-based row and column: `outValue(0)` is the row, `outValue(1)` is the column.

```vba
Dim rc As Variant
If Not BindU(pRef, btSingleCellRef, rc, xTemp) Then GoTo ReturnResult
' rc(0) = worksheet row (1-based)
' rc(1) = worksheet column (1-based)
```

This is useful for functions that need to know *where* a cell is, not just its value.

### btValue

Extracts the scalar value from a single cell or literal. For references, it coerces to a 1×1 multi-array and returns element (0,0). The returned `Variant` preserves the original type — `Double` for numbers, `String` for text, `Boolean` for booleans, `Empty` for blank cells, or an error `Variant` for errors.

```vba
Dim v As Variant
If Not BindU(pCell, btValue, v, xTemp) Then GoTo ReturnResult
If IsEmpty(v) Then
    ' Cell was blank
End If
```

---

## BindQ — For Pre-Coerced Arguments

`BindQ` is the counterpart to `BindU` for arguments registered with type `Q`. When you register an argument as type `Q`, Excel coerces references to values *before* calling your function — you receive the resolved data, not the raw reference.

`BindQ` has the same signature and bind types as `BindU`, but since references are already resolved, `btSingleCellRef` always fails (there's no reference information left to inspect).

See [[U vs Q Registration]] for guidance on when to use each.

---

## How Coercion Works Under the Hood

For simple types (`xltypeNum`, `xltypeStr`, `xltypeBool`, `xltypeInt`), `BindU` reads the value directly from the `XLOPER12` structure using the appropriate `Xloper12*Value` function.

For references (`xltypeSRef`, `xltypeRef`) and multi-cell ranges, `BindU` calls Excel's `xlCoerce` callback to resolve the reference to a value. This is the same mechanism Excel uses internally when a built-in function needs to read a cell value.

The coercion helpers (`CoerceToNumber`, `CoerceToString`, `CoerceToBool`, `CoerceToArray`) handle the `xlCoerce` call, read the result, free the intermediate `XLOPER12` via `xlFree`, and return the typed value.

---

## Error Handling in BindU

When `BindU` fails, it sets `xResult` to `xlerrValue` (`#VALUE!`) automatically via `SetErrorResult`. You don't need to set the error yourself — just `GoTo ReturnResult`:

```vba
If Not BindU(pA, btNumber, a, xTemp) Then GoTo ReturnResult
' xTemp is already set to #VALUE! — AllocResultToCaller will return it
```

If you need to detect specific input types before binding (e.g., reject non-string input), you can inspect `pIn.xltype` directly:

```vba
If (pText.xltype And xltypeStr) = 0 Then
    SetErrorResult xTemp
    GoTo ReturnResult
End If
```

---

## Next Steps

- [[U vs Q Registration]] — when to use `BindU` vs `BindQ`
- [[Working with Arrays]] — detailed patterns for `btArray` input and `GetXLMulti12` output
- [[Optional Arguments]] — handling `xltypeMissing` for optional parameters
