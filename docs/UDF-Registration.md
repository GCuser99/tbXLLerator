# UDF Registration

Excel doesn't discover XLL functions automatically. Every UDF must be explicitly registered in `xlAutoOpen` so that Excel knows the function name, its argument types, help text, and behavioral flags. The `UDF` class encapsulates this process.

---

## How Registration Works

Under the hood, registration calls `xlfRegister` with an array of `XLOPER12` values that describe the function. The most critical piece is the **type-text string** — a compact encoding that tells Excel the return type, argument types, and behavioral flags (volatile, thread-safe, macro-equivalent).

The `UDF` class builds this type-text string automatically from the properties you set. You don't need to construct it manually.

---

## Basic Registration Pattern

```vba
Private udfs As New Collection

[DllExport]
Public Function xlAutoOpen() As Long
    Dim udf As UDF
    Set udf = New UDF
    With udf
        .ProcName   = "TBXLL_Multiply"
        .FuncText   = "TBXLL_Multiply"
        .Category   = "My Add-In"
        .FuncHelp   = "Multiplies two numbers"
        .Volatile   = False
        .ThreadSafe = True
        .AddArgument Name:="Num1", Help:="First number"
        .AddArgument Name:="Num2", Help:="Second number"
        .Register
    End With
    udfs.Add udf

    xlAutoOpen = 1
End Function
```

The `udfs` collection serves two purposes: it keeps the `UDF` objects alive (Excel retains pointers into the registration data), and it provides the list of UDFs to unregister in `xlAutoClose`.

---

## Properties

### Required

| Property | Type | Description |
|----------|------|-------------|
| `ProcName` | `String` | The exported function name in the DLL. Must match the `[DllExport]` function name exactly. |
| `FuncText` | `String` | The name shown to users in worksheet formulas and the Function Wizard. Often the same as `ProcName`. |

At least one of `ProcName` or `FuncText` must be supplied. If only one is provided, the other defaults to the same value.

### Function Wizard

| Property | Type | Default | Description |
|----------|------|---------|-------------|
| `Category` | `String` | `""` | Custom category name in the Function Wizard. If set, takes precedence over `CategoryExcel`. |
| `CategoryExcel` | `ExcelCategories` | `ecUserDefined` | Places the UDF under a built-in Function Wizard category (Financial, Math & Trig, etc.). |
| `FuncHelp` | `String` | `""` | Description shown when the function is selected in the Function Wizard. |
| `HelpTopic` | `String` | `""` | URL or CHM path opened when the user clicks "Help on this function" in the Wizard. |
| `MacroType` | `MacroTypes` | `mtVisibleUDF` | Controls visibility. `mtVisibleUDF` (1) shows in the Wizard, `mtHiddenUDF` (0) hides it. |

### Behavioral Flags

| Property | Type | Default | Description |
|----------|------|---------|-------------|
| `Volatile` | `Boolean` | `False` | When `True`, the UDF recalculates on every F9 press, even if its arguments haven't changed. Use sparingly — volatile functions slow down large workbooks. |
| `ThreadSafe` | `Boolean` | `True` | When `True`, Excel may call the UDF concurrently on multiple threads during recalculation. The function must have no shared mutable state. See [[Thread Safety]]. |
| `MacroEquivalent` | `Boolean` | `False` | When `True`, grants macro-sheet equivalent privileges, enabling calls to macro-only functions like `xlSheetNm` and `xlfCaller`. Mutually exclusive with `ThreadSafe`. |
| `ReturnType` | `RegDataTypes` | `rdtXLOPER12U` | The return type encoding. Most UDFs return `XLOPER12` (`U`). See [[RegDataTypes Reference]] for alternatives. |

### Arguments

Arguments are added via the `AddArgument` method:

```vba
.AddArgument Name:="range", Help:="The input range", Type:=rdtXLOPER12U
```

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `Name` | `String` | `""` | Argument name shown in the Function Wizard and formula tooltips |
| `Help` | `String` | `""` | Help text shown for this argument in the Wizard |
| `Type` | `RegDataTypes` | `rdtXLOPER12U` | Argument type encoding. See [[RegDataTypes Reference]]. |

---

## Unregistration

Every UDF registered in `xlAutoOpen` must be unregistered in `xlAutoClose`:

```vba
[DllExport]
Public Function xlAutoClose() As Long
    Dim udf As UDF
    For Each udf In udfs
        udf.UnRegister
    Next udf
    xlAutoClose = 1
End Function
```

The `UnRegister` method handles an Excel quirk: visible UDFs must be re-registered as hidden before they can be fully unregistered. The class handles this automatically.

---

## Help Topics

The `HelpTopic` property supports two forms:

**URL form** — opens a web page when the user clicks "Help on this function":
```vba
.HelpTopic = "https://github.com/GCuser99/tbXLLerator"
```
The class automatically appends `!0` if needed (Excel requires this suffix for URL help topics).

**CHM form** — opens a specific topic in a local help file:
```vba
.HelpTopic = "C:\path\to\help.chm!1001"
```
The number after `!` is the numeric Help Context ID.

---

## Category Options

You have two ways to categorize your UDF in the Function Wizard:

**Custom category** — creates a new category with your chosen name:
```vba
.Category = "My Custom Add-In"
```

**Built-in category** — places the UDF under one of Excel's standard categories:
```vba
.CategoryExcel = ecMathTrig    ' appears under "Math & Trig"
```

If both are set, `Category` (the custom string) takes precedence.

Available built-in categories include `ecFinancial`, `ecDateTime`, `ecMathTrig`, `ecStatistical`, `ecLookupReference`, `ecDatabase`, `ecText`, `ecLogical`, `ecInformation`, `ecUserDefined`, `ecEngineering`, `ecCompatibility`, and `ecWeb`.

---

## Type-Text String Construction

The `UDF` class builds the type-text string automatically. For reference, here's how it maps to properties:

The first character is the return type (default `U` for `XLOPER12`). Then one character per argument (default `U` each). Then optional suffix flags:

| Suffix | Property | Meaning |
|--------|----------|---------|
| `!` | `Volatile = True` | Recalculates every time |
| `$` | `ThreadSafe = True` | Safe for concurrent execution |
| `#` | `MacroEquivalent = True` | Macro-sheet privileges |

For example, a thread-safe function taking two `XLOPER12` arguments and returning an `XLOPER12` gets type-text `UUU$`.

---

## Next Steps

- [[Thread Safety]] — what `ThreadSafe = True` means for your code
- [[RegDataTypes Reference]] — all available argument and return type encodings
- [[U vs Q Registration]] — choosing the right argument type
