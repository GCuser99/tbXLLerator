# Thread Safety

One of the biggest advantages of XLL UDFs over VBA is multithreaded recalculation. When you register a UDF with `ThreadSafe = True`, Excel can call it concurrently across multiple CPU cores, dramatically improving recalculation speed for workbooks with many cells calling the same function.

This page explains how thread safety works, what it requires from your code, and when to opt out.

---

## How Excel Multithreaded Recalculation Works

Excel's calculation engine divides the dependency tree into chains that can be evaluated independently. When multiple cells call the same thread-safe UDF and there's no dependency between them, Excel dispatches those calls to a thread pool — typically one thread per CPU core.

For a workbook with 10,000 cells calling `TBXLL_Multiply` on an 8-core machine, Excel can process roughly 8 cells simultaneously instead of 1. The speedup is nearly linear for functions that don't contend on shared resources.

---

## What Makes a UDF Thread-Safe

A function is thread-safe when multiple threads can execute it concurrently without interfering with each other. In practice, this means:

**No shared mutable state.** The function must not read or write any `Static` variables, module-level variables, or global data that could be modified by another thread at the same time.

**No non-reentrant API calls.** The function must not call Windows APIs or library functions that maintain internal state across calls (unless those APIs are documented as thread-safe).

**Independent allocations.** Each call must allocate its own result memory. This is exactly what `AllocResultToCaller` provides — every call gets a fresh heap allocation.

### Safe patterns

```vba
' All local variables — fully thread-safe
[DllExport]
Public Function TBXLL_Multiply(ByRef pA As XLOPER12, ByRef pB As XLOPER12) As LongPtr
    Dim xTemp As XLOPER12          ' local
    Dim a As Double, b As Double   ' local
    If Not BindU(pA, btNumber, a, xTemp) Then GoTo ReturnResult
    If Not BindU(pB, btNumber, b, xTemp) Then GoTo ReturnResult
    xTemp = GetXLNum12(a * b)
ReturnResult:
    Return AllocResultToCaller(xTemp)
End Function
```

### Unsafe patterns

```vba
' Static mutable state — NOT thread-safe
Static counter As Long
counter = counter + 1   ' race condition if two threads hit this simultaneously
```

```vba
' Module-level mutable state — NOT thread-safe
Private cachedResult As Double
cachedResult = n * 2    ' another thread could overwrite this
```

---

## When to Use ThreadSafe = False

Some UDFs genuinely need shared state or can't avoid non-reentrant calls. Register these as `ThreadSafe = False`:

- **Functions using `Static` variables** for persistence across recalculations (e.g., `TBXLL_RecalcCounter`)
- **Functions calling macro-only Excel callbacks** like `xlSheetNm` or `xlfCaller` (these require `MacroEquivalent = True`, which is mutually exclusive with `ThreadSafe`)
- **Functions accessing shared external resources** (files, databases, COM objects) without synchronization

When `ThreadSafe = False`, Excel serializes all calls to that function onto a single thread. Other thread-safe functions in your XLL still run concurrently — the restriction only applies to the specific function registered as unsafe.

---

## The AllocResultToCaller Pattern

The key mechanism enabling thread safety is dynamic allocation. Compare the two approaches:

### Static (NOT thread-safe)
```
Thread 1: writes result to Static xResult
Thread 2: writes result to Static xResult  ← overwrites Thread 1's result
Thread 1: returns pointer to xResult       ← returns Thread 2's value!
```

### Dynamic (thread-safe)
```
Thread 1: allocates pResult1, writes result
Thread 2: allocates pResult2, writes result
Thread 1: returns pResult1  ← correct
Thread 2: returns pResult2  ← correct
```

`AllocResultToCaller` implements the dynamic approach. Each call gets its own heap block, so there's no contention.

---

## Static State — When It's Actually OK

There is one case where `Static` variables are acceptable in thread-safe UDFs: **immutable-after-initialization** data. If a `Static` variable is set once and never modified again, concurrent reads are safe because all threads see the same value.

The `TBXLL_Timestamp` demo illustrates this:

```vba
Static freq As LongLong
Static initialized As Boolean

If Not initialized Then
    QueryPerformanceFrequency freq   ' set once
    initialized = True               ' set once
End If

' freq is now read-only — safe to access from any thread
QueryPerformanceCounter counter
xResult = GetXLNum12(counter / freq)
```

The race condition on `initialized` is benign — if two threads both initialize `freq`, they write the same value. This is a standard pattern in concurrent programming called "benign races on initialization."

---

## Performance Impact

The performance difference between `ThreadSafe = True` and `False` depends on your workbook:

- **Many cells, lightweight functions** — thread-safe gives the biggest speedup. Excel can dispatch thousands of cells across all cores.
- **Few cells, heavyweight functions** — less impact, since there's less opportunity for parallelism.
- **Single cell or array formula** — no benefit. A single call runs on one thread regardless of the flag.

The demo functions `TBXLL_SpeedSafe` and `TBXLL_SpeedUnsafe` are identical in logic but registered differently, allowing you to compare recalculation times in a worksheet with many cells.

---

## Checklist

Before registering a UDF as `ThreadSafe = True`, verify:

- [ ] No `Static` mutable variables (immutable-after-init is OK)
- [ ] No module-level mutable variables accessed by the function
- [ ] No shared collections, dictionaries, or arrays modified during execution
- [ ] No non-reentrant Windows API calls
- [ ] Result returned via `AllocResultToCaller` (not a `Static` XLOPER12)
- [ ] No calls to `MacroEquivalent`-only Excel functions (`xlSheetNm`, `xlfCaller`)

---

## Next Steps

- [[Performance Tuning]] — measuring and optimizing UDF performance
- [[Memory Management]] — how `AllocResultToCaller` and `xlAutoFree12` work together
- [[U vs Q Registration]] — how argument type affects threading behavior
