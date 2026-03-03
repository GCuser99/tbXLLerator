# Installation

---

## Prerequisites

- Wayne Phillips' [twinBASIC](https://twinbasic.com) compiler
- Microsoft Excel 2010 or later (32-bit or 64-bit) on Windows
- Jon Johnson's [ExcelSDK.twin](https://github.com/fafalone/TBXLLUDF) module

---

## Building the XLL

1. Install [twinBASIC](https://twinbasic.com)
2. Clone or download the [tbXLLerator](https://github.com/GCuser99/tbXLLerator) repository
3. Clone or download Jon Johnson's latest [ExcelSDK.twin](https://github.com/fafalone/TBXLLUDF) module and import it into your twinBASIC project
4. Open the `.twinproj` file in twinBASIC
5. Set the twinBASIC bitness (32-bit or 64-bit) to match your Excel version's bitness
6. Build the project — twinBASIC will produce a `.xll` file in the `Win32` or `Win64` output folder

> **Tip:** Under **Project → Project Settings → Build Output Path**, set the path to `${SourcePath}\${Architecture}\${ProjectName}.xll` so the output file gets the `.xll` extension automatically.

---

## Loading in Excel

1. In Excel, go to **File → Options → Add-ins → Manage: Excel Add-ins → Go** (you can also access this via the **Developer Tab → Excel Add-ins**)
2. Click **Browse** and select the `.xll` file you built
3. The add-in will load and UDFs will be available in the Function Wizard under the category defined in `xlAutoOpen`

---

## Using tbXLLerator as a Starting Point

To use the framework in your own project:

1. Create a new **Standard DLL** project in twinBASIC
2. Import the `ExcelSDK`, `Helpers`, `AutoCallbacks`, and `UDF` class files into your project
3. Set the build output path to produce a `.xll` file (see tip above)
4. Add your UDF functions following the patterns in the `Demos` module
5. Register each UDF in `xlAutoOpen` using the `UDF` class
6. Unregister in `xlAutoClose` by iterating the `udfs` collection

---

## Important Notes

- The `.xll` bitness **must match** Excel's bitness. A 64-bit XLL will not load in 32-bit Excel and vice versa.
- Excel must be **fully closed** before replacing or updating the `.xll` file. Excel locks the DLL while it's loaded.
- If the XLL fails to load silently, check the bitness first — it's the most common cause.
