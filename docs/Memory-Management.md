# Memory Management

Memory management is the most critical aspect of XLL development. Get it wrong and you'll see crashes, memory leaks, or corrupt results. tbXLLerator provides a well-defined pattern that handles the complexity for you — but understanding *why* the pattern exists will help you use it correctly and debug issues when they arise.

---

## The Core Problem

Excel calls your UDF and expects a pointer to an `XLOPER12` result. But that result has to live somewhere in memory — and it can't be a local variable on the stack, because the stack frame is destroyed when your function returns. The result must persist until Excel is done reading it.

There are two classic approaches:

1. **Static storage** — use a `Static` variable in the function. Simple, but only one result can exist at a time, which breaks multithreaded recalculation.
2. **Dynamic allocation** — allocate each result on the heap. Each call gets its own independent memory, enabling full concurrency.

tbXLLerator uses approach 2, which is why thread-safe UDFs work.

---

## The Allocate-Return-Free Pattern

Every UDF follows the same memory lifecycle:

```
Your UDF                          Excel
────────                          ─────
1. Compute result
2. AllocResultToCaller(xTemp)
   → GlobalAlloc a new XLOPER12
   → Copy xTemp into it
   → Set xlbitDLLFree flag
   → Return pointer to Excel
                                  3. Excel reads the value
                                  4. Excel calls xlAutoFree12(ptr)
                                     → Framework frees the memory
```

In code, it always looks like this:

```vba
[DllExport]
Public Function TBXLL_Example(ByRef pN As XLOPER12) As LongPtr
    Dim xTemp As XLOPER12          ' local, per-call — never Static
    Dim n As Double

    If Not BindU(pN, btNumber, n, xTemp) Then GoTo ReturnResult

    xTemp = GetXLNum12(n * 2)

ReturnResult:
    Return AllocResultToCaller(xTemp)   ' <-- every path ends here
End Function
```

The key rules:

- **`xTemp` is always a local variable**, never `Static`. This is what makes concurrent execution safe.
- **Every code path must set `xTemp`** before reaching `AllocResultToCaller`. On the success path you set it to your result. On error paths, `BindU` sets it to `#VALUE!` automatically, or you call `SetErrorResult xTemp` explicitly.
- **Every function returns through `AllocResultToCaller`**, which handles the heap allocation and `xlbitDLLFree` flag.

---

## What `AllocResultToCaller` Does

```vba
Public Function AllocResultToCaller(ByRef x As XLOPER12) As LongPtr
    Dim pResult As LongPtr
    pResult = GlobalAlloc(GPTR, LenB(Of XLOPER12))
    If pResult = 0 Then Return 0
    x.xltype = x.xltype Or xlbitDLLFree
    CopyMemory ByVal pResult, x, LenB(Of XLOPER12)
    Return pResult
End Function
```

It allocates a fixed-size block on the Windows heap via `GlobalAlloc`, sets the `xlbitDLLFree` flag on the xltype (which tells Excel "call `xlAutoFree12` when you're done"), copies the `XLOPER12` into the heap block, and returns the pointer.

---

## What `xlAutoFree12` Does

When Excel is finished with your result, it calls `xlAutoFree12` with the pointer that your UDF returned. The framework's implementation handles these cases:

- **`xltypeStr`** — frees the string buffer (which was allocated by `GetXLString12` via `GlobalAlloc`), then frees the outer `XLOPER12` struct.
- **`xltypeMulti`** — iterates all elements in the array, frees any string buffers found inside, frees the element array itself, then frees the outer struct.
- **`xltypeRef`** — frees the `XLMREF12` allocation, then frees the outer struct.
- **All other types** (`xltypeNum`, `xltypeBool`, `xltypeInt`, `xltypeErr`, `xltypeNil`) — no inner allocations exist, so only the outer struct is freed.

---

## Who Frees What — Ownership Rules

Memory ownership in XLL development follows strict rules. Getting this wrong causes either leaks or double-frees (crashes):

| Allocation | Owner | How it's freed |
|------------|-------|----------------|
| `AllocResultToCaller` output | Excel | Excel calls `xlAutoFree12` |
| `GetXLString12` buffer (inside a result) | Excel | Freed inside `xlAutoFree12` |
| `GetXLMulti12` element array (inside a result) | Excel | Freed inside `xlAutoFree12` |
| `Excel12v` output (e.g., `xlfSum` result) | Excel | You must call `xlFree` |
| `xlGetName` output | Excel | Do NOT free — Excel owns it |
| Your local `XLOPER12` variables | Stack | Freed automatically on return |

The two most important rules to remember:

1. **Results from `Excel12v` must be freed with `xlFree`** when you're done with them. If you call `xlfSum` and read the result, free it before returning.
2. **Results you return via `AllocResultToCaller` must NOT be freed by you** — Excel will free them through `xlAutoFree12`.

---

## Common Patterns

### Calling an Excel built-in and freeing the intermediate result

```vba
Dim xSum(0) As XLOPER12
Dim xFree(0) As XLOPER12
Dim args(0) As XLOPER12

args(0) = pRange

' Excel allocates the result — we must free it
If Excel12v(xlfSum, xSum(0), 1, args) <> 0 Then
    SetErrorResult xTemp
    GoTo ReturnResult
End If

' Extract the value we need
Dim total As Double
total = Xloper12NumValue(xSum(0))

' Free Excel's allocation
xFree(0) = xSum(0)
Excel12v xlFree, ByVal vbNullPtr, 1, xFree

' Now use 'total' in our own result
xTemp = GetXLNum12(total * 2)
```

### Chaining two Excel built-ins

When chaining multiple built-in calls, you must free each intermediate result before returning. If an intermediate call fails, free any results you've already obtained:

```vba
' Get sum
If Excel12v(xlfSum, sumRes, 1, args) <> 0 Then
    SetErrorResult xTemp
    GoTo ReturnResult
End If

' Get count — if this fails, must free sumRes first
If Excel12v(xlfCount, cntRes, 1, args) <> 0 Then
    freeArgs(0) = sumRes
    Excel12v xlFree, ByVal vbNullPtr, 1, freeArgs
    SetErrorResult xTemp
    GoTo ReturnResult
End If

' Use both results
xTemp = GetXLNum12(Xloper12NumValue(sumRes) / Xloper12NumValue(cntRes))

' Free both
freeArgs(0) = sumRes
Excel12v xlFree, ByVal vbNullPtr, 1, freeArgs
freeArgs(0) = cntRes
Excel12v xlFree, ByVal vbNullPtr, 1, freeArgs
```

---

## Error Path Discipline

The `GoTo ReturnResult` pattern ensures every code path reaches `AllocResultToCaller`. There are two categories of error-setters:

- **`BindU` and `BindQ`** set `xTemp` to `#VALUE!` automatically on failure. You just need `GoTo ReturnResult`.
- **`Excel12v` and `GetXLMulti12`** do NOT set `xTemp` on failure. You must call `SetErrorResult xTemp` explicitly before `GoTo ReturnResult`.

```vba
' BindU handles the error automatically
If Not BindU(pA, btNumber, a, xTemp) Then GoTo ReturnResult

' Excel12v does NOT — you must set the error yourself
If Excel12v(xlfSum, xTemp, 1, args) <> 0 Then
    SetErrorResult xTemp    ' <-- required
    GoTo ReturnResult
End If
```

---

## Next Steps

- [[Thread Safety]] — how the allocation pattern enables concurrent recalculation
- [[Working with Arrays]] — memory considerations for `GetXLMulti12` array returns
- [[Argument Binding]] — how `BindU()` handles type coercion
