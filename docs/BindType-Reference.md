# BindType Reference

The `BindType` enum specifies the target type for `BindU()` and `BindQ()` argument coercion.

---

## Enum Values

```vba
Public Enum BindType
    btNumber = 0
    btString = 1
    btBool = 2
    btDate = 3
    btArray = 4
    btSingleCellRef = 5
    btValue = 6
End Enum
```

---

## Detailed Reference

### btNumber

| | |
|---|---|
| **Output type** | `Double` |
| **Accepts** | `xltypeNum`, `xltypeInt`, `xltypeBool`, `xltypeStr`, `xltypeSRef`, `xltypeRef`, `xltypeMulti` |
| **Coercion** | Direct read for numeric types. `xlCoerce` to `xltypeNum` for refs, strings, and arrays. Booleans convert to 1.0 or 0.0. |
| **Failure** | Returns `False`, sets `xResult` to `#VALUE!` |
| **BindQ support** | Yes — reads directly from coerced XLOPER12 |

### btString

| | |
|---|---|
| **Output type** | `String` |
| **Accepts** | `xltypeStr`, `xltypeNum`, `xltypeInt`, `xltypeBool`, `xltypeSRef`, `xltypeRef`, `xltypeMulti` |
| **Coercion** | Direct read for strings. `CStr()` for numbers and integers. `"TRUE"`/`"FALSE"` for booleans. `xlCoerce` to `xltypeStr` for refs and arrays. |
| **Failure** | Returns `False`, sets `xResult` to `#VALUE!` |
| **BindQ support** | Yes |

### btBool

| | |
|---|---|
| **Output type** | `Boolean` |
| **Accepts** | `xltypeBool`, `xltypeNum`, `xltypeInt`, `xltypeStr`, `xltypeSRef`, `xltypeRef`, `xltypeMulti` |
| **Coercion** | Direct read for booleans. Non-zero numbers → `True`, zero → `False`. Refs and strings coerced via `xlCoerce` to `xltypeNum`, then tested for non-zero. |
| **Failure** | Returns `False`, sets `xResult` to `#VALUE!` |
| **BindQ support** | Yes |

### btDate

| | |
|---|---|
| **Output type** | `Double` (Excel serial date) |
| **Accepts** | Same as `btNumber` |
| **Coercion** | Same as `btNumber`, plus validation that the result falls in the range 0 to 2958465 (January 0, 1900 through December 31, 9999). |
| **Failure** | Returns `False`, sets `xResult` to `#VALUE!` |
| **BindQ support** | Yes — delegates to btNumber internally, then validates range |

### btArray

| | |
|---|---|
| **Output type** | `Variant()` — 2D, 0-based in both dimensions |
| **Accepts** | `xltypeSRef`, `xltypeRef`, `xltypeMulti`, and scalar types (wrapped to 1×1) |
| **Coercion** | Uses `xlCoerce` to `xltypeMulti`, then reads each element preserving its type (`vbDouble`, `vbString`, `vbError`, `vbEmpty`, etc.) |
| **Failure** | Returns `False`, sets `xResult` to `#VALUE!` |
| **BindQ support** | Yes — `Q` delivers `xltypeMulti` directly for ranges, scalars wrapped to 1×1 |

### btSingleCellRef

| | |
|---|---|
| **Output type** | `Variant` — 2-element array: `(row, column)`, 1-based |
| **Accepts** | `xltypeSRef`, `xltypeRef` (single cell only) |
| **Coercion** | Validates that the reference covers exactly one cell (`rwFirst = rwLast` and `colFirst = colLast`). Returns 1-based row and column. |
| **Failure** | Returns `False`, sets `xResult` to `#VALUE!`. Fails for multi-cell ranges, strings, numbers, etc. |
| **BindQ support** | **No** — always fails because `Q` resolves references before the function is called |

### btValue

| | |
|---|---|
| **Output type** | `Variant` — typed value preserving the original Excel type |
| **Accepts** | Any scalar type, single-cell refs, multi-cell refs (returns top-left only) |
| **Coercion** | Coerces to 1×1 `xltypeMulti` via `xlCoerce`, then extracts element (0,0). Returns `Double` for numbers, `String` for text, `Boolean` for booleans, `Empty` for blanks, error `Variant` for errors. |
| **Failure** | Returns `False`, sets `xResult` to `#VALUE!` |
| **BindQ support** | Yes — reads directly from the coerced XLOPER12 |

---

## Next Steps

- [[Argument Binding]] — usage patterns and examples for each bind type
- [[U vs Q Registration]] — how registration type affects binding behavior
