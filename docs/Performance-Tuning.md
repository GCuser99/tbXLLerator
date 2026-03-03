# Performance Tuning

XLL UDFs are already faster than VBA and COM add-ins by virtue of being native compiled code with direct access to Excel's calculation engine. But there are choices within the XLL framework that can further affect performance by an order of magnitude or more.

---

## ThreadSafe = True

The single biggest performance lever for workbooks with many cells calling the same UDF. On an 8-core machine, thread-safe UDFs can recalculate up to 8x faster than single-threaded ones.

The demo functions `TBXLL_SpeedSafe` and `TBXLL_SpeedUnsafe` are identical in logic but differ in registration. Fill a column with 10,000+ calls and compare F9 recalculation times.

See [[Thread Safety]] for the requirements.

---

## Argument Registration: U vs Q vs B

The registration type for your arguments affects how much work happens before your function starts:

| Type | What Excel does | Best for |
|------|----------------|----------|
| `U` (default) | Passes raw XLOPER12 — your code resolves refs | Delegating to built-ins, cell references, large ranges |
| `Q` | Excel coerces refs to values first | Scalar inputs, processing every element |
| `B` (Double) | Excel extracts the double directly | Pure numeric functions — zero XLOPER12 overhead |

### Type B — fastest for numeric scalars

If your function takes and returns only `Double` values, register with `rdtDouble`:

```vba
.ReturnType = rdtDouble
.AddArgument Name:="num1", Type:=rdtDouble
.AddArgument Name:="num2", Type:=rdtDouble
```

```vba
[DllExport]
Public Function TBXLL_MultiplyB(ByVal a As Double, ByVal b As Double) As Double
    Return a * b
End Function
```

This eliminates all XLOPER12 creation, binding, heap allocation, and `xlAutoFree12` cleanup. The tradeoff is losing the ability to return errors, handle non-numeric inputs, or accept optional arguments.

### Type Q — skip the coercion callback

With type `U`, `BindU` calls `xlCoerce` to resolve cell references — an Excel callback with some overhead. With type `Q`, Excel does this coercion before calling you, so `BindQ` reads the value directly.

However, for functions that pass ranges to built-ins like `xlfSum`, type `U` is faster because you pass a lightweight reference instead of a materialized array.

See [[U vs Q Registration]] for the full comparison.

---

## Array Functions vs Per-Cell Functions

Consider a workbook where 10,000 cells each need a computation. Two design choices:

**Per-cell UDF** — each cell calls `=MY_FUNC(A1)` independently. With `ThreadSafe = True`, Excel dispatches across cores. Good for simple computations with low per-call overhead.

**Array UDF** — one cell calls `=MY_FUNC_ALL(A1:A10000)` and returns a 10,000-element array. Processes everything in a single call. Better when there's setup cost or when you want to avoid 10,000 separate heap allocations.

The demo functions `TBXLL_SpeedSafe` (per-cell) and `TBXLL_SpeedSafeArray` (array) illustrate this tradeoff.

---

## FP12 for Numeric Arrays

If your array function only processes `Double` values, `rdtFP12` avoids the overhead of `Variant` arrays and per-element type checking:

```vba
.AddArgument Name:="range", Type:=rdtFP12
```

Excel passes a contiguous block of doubles — no `Variant` boxing, no `xlCoerce`. See [[FP12 Type]].

---

## Avoid Unnecessary Volatility

Volatile functions (`Volatile = True`) recalculate on every F9 press, regardless of whether their inputs changed. In a workbook with thousands of volatile UDF calls, this means thousands of unnecessary recalculations every time the user presses F9 or any cell changes.

Only use `Volatile = True` when the function's result genuinely depends on something outside its arguments (like the current time).

---

## Minimize Excel Callbacks

Each `Excel12v` call is a callback into Excel's engine — fast, but not free. If you can compute a result directly in twinBASIC without calling Excel, that's typically faster for small inputs.

For example, summing a 10-element array in a twinBASIC loop avoids the `xlfSum` callback overhead. But for very large ranges with type-`U` registration, `xlfSum` may be faster because it operates on the raw reference without materializing the array.

---

## Measuring Performance

The `TBXLL_Timestamp` demo function returns a high-resolution timestamp using `QueryPerformanceCounter`. Use it to measure recalculation time in the worksheet:

1. Put `=TBXLL_Timestamp()` in a cell (stamps when recalc reaches it)
2. Put your UDF calls in a range
3. Put `=TBXLL_Timestamp(SomeCell)` in a cell that depends on the last UDF result
4. The difference in seconds is the recalculation time for the UDFs between the two timestamps

Because `TBXLL_Timestamp` accepts an optional input argument, you can create dependency chains that control the order Excel evaluates the timestamps relative to your UDFs.

---

## Next Steps

- [[Thread Safety]] — ensuring your UDFs qualify for concurrent execution
- [[U vs Q Registration]] — choosing the right argument type for performance
- [[FP12 Type]] — the numeric-only array alternative
