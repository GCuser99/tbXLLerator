# Architecture

This page describes how the tbXLLerator framework is organized and how the pieces fit together.

---

## Module Overview

| Module | Purpose |
|--------|---------|
| `ExcelSDK` | Low-level foundation: XLOPER12 struct, constants, enums, `Excel12v` declaration, and XLOPER12 builder/reader functions (`GetXLNum12`, `GetXLString12`, `Xloper12NumValue`, etc.). Written by Jon Johnson. |
| `Helpers` | The framework layer: `BindU`/`BindQ` dispatchers, coercion helpers (`CoerceToNumber`, `CoerceToArray`, etc.), `AllocResultToCaller`, `SetErrorResult`, `GetXLMulti12`, `GetXLVariant12`, `ReadFP12`, and string reader utilities. |
| `AutoCallbacks` | All `xlAuto*` entry points: `xlAutoOpen` (registration), `xlAutoClose` (unregistration), `xlAutoFree12` (memory cleanup), `xlAddInManagerInfo12` (add-in display name), `xlAutoAdd`, `xlAutoRemove`, and `xlAutoRegister12`. |
| `UDF` (class) | High-level registration wrapper. Encapsulates `xlfRegister` call with properties for `ProcName`, `FuncText`, `Category`, `ThreadSafe`, `Volatile`, `HelpTopic`, etc. Builds the type-text string automatically. |
| `Argument` (class) | Simple data class holding `Name`, `Help`, and `Type` for a single UDF argument. Created via `UDF.AddArgument`. |
| `cMemPool` (class) | Small bump allocator used internally by `GetXLActiveRef12` for `XLMREF12` allocations. Not typically used directly by UDF authors. |
| `Demos` | 30+ demo UDFs illustrating every major pattern in the framework. |

---

## Data Flow

Here's how data flows through the framework during a typical UDF call:

```
Excel worksheet cell
    │
    ▼
Excel calculation engine
    │  passes XLOPER12 args
    ▼
Your [DllExport] function
    │
    ├── BindU() ──► CoerceToNumber / CoerceToArray / etc.
    │                   │
    │                   └── Excel12v(xlCoerce, ...) if needed
    │
    ├── Your computation logic
    │
    ├── GetXLNum12 / GetXLString12 / GetXLMulti12
    │
    └── AllocResultToCaller(xTemp)
            │  heap-allocates result with xlbitDLLFree
            ▼
        Return LongPtr to Excel
            │
            ▼
        Excel reads result, displays in cell
            │
            ▼
        Excel calls xlAutoFree12(ptr)
            │  frees heap allocation + inner buffers
            ▼
        Memory released
```

---

## Registration Flow

```
Excel loads .xll
    │
    ▼
xlAutoOpen()
    │
    ├── For each UDF:
    │       UDF.Register
    │           │
    │           ├── xlGetName (get DLL path)
    │           ├── Build type-text string (U/Q/B + flags)
    │           └── Excel12v(xlfRegister, ...) with def() array
    │
    └── Store UDF objects in udfs collection
```

```
Excel unloads .xll
    │
    ▼
xlAutoClose()
    │
    └── For each UDF in udfs:
            UDF.UnRegister
                │
                ├── Excel12v(xlfRegisterId, ...) to get registration ID
                ├── Excel12v(xlfUnregister, ...) to remove
                └── Re-register as hidden + unregister again (Excel workaround)
```

---

## Key Design Decisions

**Dynamic allocation over static storage.** Every UDF result is heap-allocated via `GlobalAlloc` in `AllocResultToCaller`. This enables `ThreadSafe = True` registration because each concurrent call gets its own independent memory. The cost is one `GlobalAlloc` + one `GlobalFree` per call, which is negligible compared to Excel's own overhead.

**Unified binding via Variant.** `BindU` and `BindQ` use `Variant` as the output type for all bind targets. This provides a single function signature that handles numbers, strings, booleans, arrays, and references. The tradeoff is an implicit boxing step for scalar types, which is insignificant for most use cases.

**Registration data persists for the add-in lifetime.** The `def()` array in the `UDF` class is a class-level field, not a local variable. Excel retains pointers into the registration strings, so they must remain valid until `UnRegister` is called. The `udfs` collection in `AutoCallbacks` keeps the `UDF` objects alive.

**xlAutoFree12 handles nested allocations.** String buffers inside `xltypeStr` results and element arrays inside `xltypeMulti` results are all allocated with `GlobalAlloc`, so `xlAutoFree12` can free them uniformly with `GlobalFree`. This avoids mixed allocation strategies.

---

## Next Steps

- [[Memory Management]] — detailed explanation of the allocation and cleanup model
- [[UDF Registration]] — how the `UDF` class works
- [[XLL Fundamentals]] — the callback lifecycle
