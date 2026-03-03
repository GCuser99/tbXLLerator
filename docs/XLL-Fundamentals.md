# XLL Fundamentals

This page explains the core concepts behind Excel XLL development. If you're coming from VBA, some of this will feel unfamiliar — XLLs operate at a lower level than VBA, closer to how Excel's own built-in functions work.

---

## What is an XLL?

An XLL is a native DLL (Dynamic Link Library) with a `.xll` file extension. When Excel loads it, Excel calls a series of well-known entry points (callbacks) to discover what functions the DLL provides, register them, and later clean up when the add-in unloads.

The key difference from VBA: your code runs **inside Excel's process** as native compiled code, not interpreted. This means direct memory access, no COM marshaling overhead, and the ability to participate in Excel's multithreaded recalculation engine.

---

## The XLOPER12 — Excel's Universal Data Type

Everything flowing between Excel and your XLL goes through `XLOPER12` structures. Think of an `XLOPER12` as a tagged union — it can hold any of these types:

| xltype | Meaning | Typical use |
|--------|---------|-------------|
| `xltypeNum` | IEEE 754 double | Numbers |
| `xltypeStr` | Length-counted Unicode string | Text |
| `xltypeBool` | Boolean (TRUE/FALSE) | Logical values |
| `xltypeErr` | Error code (#VALUE!, #N/A, etc.) | Error returns |
| `xltypeInt` | 32-bit integer | Small integers |
| `xltypeMulti` | 2D array of XLOPER12s | Ranges, array formulas |
| `xltypeSRef` | Single-sheet cell reference | Cell/range references |
| `xltypeRef` | Multi-sheet cell reference | External references |
| `xltypeMissing` | Omitted argument | Optional parameters |
| `xltypeNil` | Empty/null | Blank cells |

The `xltype` field tells you which member of the union is active. The framework's `BindU()` and `BindQ()` functions handle the complexity of reading these structures so you don't have to work with raw memory offsets.

---

## The Callback Lifecycle

When Excel loads your XLL, it calls these entry points in order:

### On Load
1. **`xlAutoOpen`** — Called first. This is where you register all your UDFs with Excel. Return `1` for success.
2. **`xlAutoAdd`** — Called only when the user explicitly adds the XLL via the Add-in Manager (not on automatic startup load). Typically a no-op.
3. **`xlAddInManagerInfo12`** — Excel calls this to get the display name shown in the Add-in Manager dialog.

### During Use
4. **Your UDF functions** — Called by Excel whenever a cell containing your formula recalculates.
5. **`xlAutoFree12`** — Called by Excel when it's done with a result your UDF returned. This is where allocated memory is freed.

### On Unload
6. **`xlAutoRemove`** — Called when the user explicitly removes the XLL via the Add-in Manager.
7. **`xlAutoClose`** — Called after `xlAutoRemove`, and also when Excel closes. This is where you unregister UDFs.

The framework's `AutoCallbacks` module implements all of these for you. You primarily interact with `xlAutoOpen` (to register UDFs) and `xlAutoClose` (to unregister them).

---

## How a UDF Call Works

When a cell containing `=TBXLL_Multiply(3, 4)` recalculates, here's the sequence:

1. Excel resolves the cell values `3` and `4` into `XLOPER12` structures (both `xltypeNum`)
2. Excel calls `TBXLL_Multiply`, passing pointers to the two `XLOPER12`s
3. Your function uses `BindU()` to extract the doubles, computes the result, and packages it into a new `XLOPER12` via `GetXLNum12()`
4. `AllocResultToCaller()` copies the result to heap memory, sets the `xlbitDLLFree` flag, and returns the pointer to Excel
5. Excel reads the result and displays it in the cell
6. When Excel no longer needs the result, it calls `xlAutoFree12` with the pointer, and the framework frees the heap memory

This allocate-return-free cycle is central to how XLL UDFs work, and it's what enables thread safety — each call gets its own independent allocation, so there's no shared state to contend over.

---

## VBA vs XLL — Key Differences

If you're coming from VBA UDFs, these are the conceptual shifts:

**No Application object.** XLL UDFs don't have access to `Application.WorksheetFunction` or the Excel object model during recalculation. Instead, you call Excel built-in functions via `Excel12v` with function codes like `xlfSum` and `xlfCountif`.

**No automatic type conversion.** VBA silently converts between types. In an XLL, you receive raw `XLOPER12` data and must explicitly coerce it. The `BindU()` framework handles this, but you choose the target type.

**Explicit memory management.** VBA handles memory automatically. In an XLL, you're responsible for allocating and freeing memory correctly. The framework's `AllocResultToCaller` / `xlAutoFree12` pattern handles this, but you need to follow it consistently.

**Thread safety is opt-in.** VBA UDFs always run single-threaded. XLL UDFs can be registered as thread-safe, enabling Excel to call them concurrently. This requires that your function has no shared mutable state — see [[Thread Safety]].

**Return by pointer, not by value.** VBA functions return values directly. XLL UDFs return a `LongPtr` pointing to an `XLOPER12` on the heap. The framework's `AllocResultToCaller` handles this.

---

## Next Steps

- [[Memory Management]] — understand the allocation and cleanup model in detail
- [[Argument Binding]] — learn how `BindU()` and `BindQ()` convert XLOPER12s to usable types
- [[Writing Your First UDF]] — hands-on walkthrough of building a UDF from scratch
