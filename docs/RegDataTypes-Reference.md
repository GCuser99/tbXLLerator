# RegDataTypes Reference

The `RegDataTypes` enum controls how arguments and return values are passed between Excel and your UDF at the registration level. The choice affects what data your function receives, how much work Excel does before calling you, and what type-text character is emitted in the registration string.

---

## Enum Values

```vba
Public Enum RegDataTypes
    rdtXLOPER12U = 0   ' Raw XLOPER12 (type char: U)
    rdtXLOPER12X = 1   ' Async XLOPER12 (type char: X) — return type only
    rdtXLOPER12Q = 2   ' Coerced XLOPER12 (type char: Q)
    rdtFP12 = 3         ' Numeric array (type char: K%)
    rdtBoolean = 4      ' Boolean (type char: A)
    rdtDouble = 5        ' Double (type char: B)
    rdtLong = 6          ' Long integer (type char: J)
    rdtString1 = 7       ' Null-terminated Unicode (type char: C%)
    rdtString2 = 8       ' Length-counted Unicode (type char: D%)
End Enum
```

---

## Detailed Reference

### rdtXLOPER12U — Raw XLOPER12 (`U`)

The default and most flexible type. Excel passes the raw XLOPER12 without modification — your function receives whatever the user provided (number, string, cell reference, array, error, or missing).

Use with `BindU()` to coerce to the desired type. This is the only type that preserves cell reference information for `btSingleCellRef`.

**Function signature:** `ByRef pArg As XLOPER12`

### rdtXLOPER12Q — Coerced XLOPER12 (`Q`)

Excel resolves cell references to their values before calling your function. A reference to a cell containing `42` arrives as `xltypeNum` with value `42`, not as `xltypeSRef`. A reference to a range arrives as `xltypeMulti`.

Use with `BindQ()`. Slightly less overhead than `U` for functions that process every element, since it avoids the `xlCoerce` callback in your code. Cannot extract cell coordinates or sheet names.

**Function signature:** `ByRef pArg As XLOPER12`

### rdtXLOPER12X — Async XLOPER12 (`X`)

For asynchronous UDF return types. Excel provides a handle that you use later to deliver the result via `xlAsyncReturn`. Return type only — not used for arguments.

**Note:** Not currently tested in tbXLLerator.

### rdtFP12 — Numeric Array (`K%`)

Excel passes a pointer to an `FP12` structure: a header with row and column counts followed by a flat array of `Double` values. All non-numeric cells are converted to `0.0` or cause an error.

Use with `ReadFP12()` helper. Best for high-performance numeric array processing where you don't need to distinguish errors, blanks, or strings. See [[FP12 Type]].

**Function signature:** `ByVal lpFP12 As LongPtr`

### rdtDouble — Double (`B`)

Excel extracts the numeric value directly and passes it as a native `Double`. No XLOPER12 overhead, no binding step, no heap allocation for the return value.

The fastest option for simple numeric functions, but you lose the ability to return errors, handle non-numeric input gracefully, or detect missing arguments.

**Function signature (argument):** `ByVal arg As Double`
**Function signature (return):** `As Double`

### rdtLong — Long Integer (`J`)

Same as `rdtDouble` but for 32-bit integers.

**Function signature:** `ByVal arg As Long`

### rdtBoolean — Boolean (`A`)

Native boolean pass-through.

**Function signature:** `ByVal arg As Boolean`

### rdtString1 — Null-Terminated Unicode (`C%`)

Excel passes a pointer to a null-terminated wide string. Use `ReadNullTerminatedString()` to convert to a VBA `String`.

Generally, `rdtXLOPER12U` with `btString` is preferred because it handles coercion from numbers and booleans automatically.

**Function signature:** `ByVal lpStr As LongPtr`

### rdtString2 — Length-Counted Unicode (`D%`)

Excel passes a pointer to a length-counted wide string (first 2 bytes are the character count). Use `ReadCountedString()` to convert.

Same recommendation as `rdtString1` — prefer `rdtXLOPER12U` with `btString` for most cases.

**Function signature:** `ByVal lpStr As LongPtr`

---

## Type-Text Characters

The registration type-text string is built by concatenating one character per argument (plus the return type first). The `UDF` class does this automatically.

| RegDataType | Type char | Notes |
|-------------|-----------|-------|
| `rdtXLOPER12U` | `U` | Default for args and return |
| `rdtXLOPER12Q` | `Q` | Pre-coerced XLOPER12 |
| `rdtXLOPER12X` | `X` | Async return only |
| `rdtFP12` | `K%` | Two characters |
| `rdtBoolean` | `A` | |
| `rdtDouble` | `B` | |
| `rdtLong` | `J` | |
| `rdtString1` | `C%` | Two characters |
| `rdtString2` | `D%` | Two characters |

Behavioral flags are appended after all argument types:

| Flag | Type char | Set by |
|------|-----------|--------|
| Volatile | `!` | `.Volatile = True` |
| Thread-safe | `$` | `.ThreadSafe = True` |
| Macro-equivalent | `#` | `.MacroEquivalent = True` |

**Example:** A thread-safe function returning `XLOPER12`, taking two `Double` arguments: type-text = `UBB$`

---

## Next Steps

- [[U vs Q Registration]] — choosing between U and Q for XLOPER12 arguments
- [[FP12 Type]] — using the numeric array type
- [[Performance Tuning]] — how registration type affects speed
