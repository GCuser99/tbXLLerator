# U vs Q Registration

When registering UDF arguments as XLOPER12, you choose between type `U` (raw) and type `Q` (pre-coerced). This choice affects what data your function receives, what operations are available, and performance characteristics.

---

## The Difference

| | Type U (raw) | Type Q (coerced) |
|---|---|---|
| **What you receive** | Exactly what Excel has — could be a number, string, cell reference, array, error, or missing | Excel resolves references first — you get values, not refs |
| **Cell references** | Preserved as `xltypeSRef` or `xltypeRef` | Resolved to `xltypeNum`, `xltypeStr`, `xltypeMulti`, etc. |
| **Binding function** | `BindU()` | `BindQ()` |
| **btSingleCellRef** | Works — can extract row/column | Always fails — reference info is gone |
| **Pass-through to built-ins** | Efficient — pass the lightweight reference directly | Inefficient — pass a materialized array |
| **Large range overhead** | Minimal — reference is a small struct | High — Excel materializes the entire range before calling you |

---

## When to Use U (Default)

Use type `U` when your function:

- **Delegates to Excel built-ins** like `xlfSum`, `xlfCountif`, `xlfTranspose`. With `U`, you pass the raw cell reference, and the built-in operates on it efficiently. With `Q`, Excel would materialize the entire range into an array first, which is slower and uses more memory.

- **Needs cell coordinates.** `btSingleCellRef` only works with `U` because it needs the raw `xltypeSRef`/`xltypeRef` structure to extract row and column.

- **Needs sheet identity.** Functions using `xlSheetNm` require the original reference.

- **Works with very large ranges.** A reference to `A1:A1000000` is a tiny `xltypeSRef` struct with `U`. With `Q`, Excel would allocate a million-element array before your function starts.

- **Only uses a subset of the range.** If your function only reads the first element of a range, `U` avoids materializing elements you'll never use.

```vba
' Registration with U (default)
.AddArgument Name:="range", Help:="Input range"
' or explicitly:
.AddArgument Name:="range", Help:="Input range", Type:=rdtXLOPER12U
```

---

## When to Use Q

Use type `Q` when your function:

- **Processes every element** of the input. If you're going to iterate the entire array anyway, having Excel pre-coerce it saves you the `xlCoerce` callback inside `BindU`.

- **Only needs scalar values.** For functions that take individual numbers, strings, or booleans, `Q` avoids the overhead of checking for references and calling `xlCoerce` in your binding code.

- **Doesn't need cell references or sheet names.** If you never use `btSingleCellRef` or `xlSheetNm`, there's no cost to losing the reference information.

```vba
' Registration with Q
.AddArgument Name:="num", Help:="A number", Type:=rdtXLOPER12Q
```

---

## Performance Comparison

Consider `=MY_FUNC(A1:A100000)`:

| Scenario | Type U | Type Q |
|----------|--------|--------|
| Pass to `xlfSum` | Fast — passes tiny reference | Slow — passes 100K-element array |
| Iterate all elements | Slightly slower — `xlCoerce` callback | Slightly faster — already materialized |
| Read only first element | Fast — coerce just one cell | Wasteful — materialized entire range |

For scalar arguments like `=MY_FUNC(42)` or `=MY_FUNC(A1)` where A1 contains a single value, the difference is negligible.

---

## Mixing U and Q

You can mix types within a single function. Each argument is registered independently:

```vba
' First arg is a range we'll pass to xlfSum (use U)
' Second arg is a scalar multiplier (use Q)
.AddArgument Name:="range", Help:="Input range", Type:=rdtXLOPER12U
.AddArgument Name:="factor", Help:="Multiplier", Type:=rdtXLOPER12Q
```

In the function body, use `BindU` for U-type arguments and `BindQ` for Q-type arguments.

---

## Summary

| Need | Use |
|------|-----|
| Pass range to Excel built-in | U |
| Extract cell row/column | U |
| Get sheet name | U |
| Large range, only read subset | U |
| Process every element of array | Q |
| Simple scalar input | Q (or B for doubles) |
| Default / unsure | U |

---

## Next Steps

- [[RegDataTypes Reference]] — all registration types including B, J, K%, C%, D%
- [[Performance Tuning]] — measuring the impact of registration choices
- [[Argument Binding]] — how BindU and BindQ differ in behavior
