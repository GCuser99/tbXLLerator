# tbXLLerator Wiki

Welcome to the **tbXLLerator** wiki — the companion documentation for the [tbXLLerator](https://github.com/GCuser99/tbXLLerator) framework, a twinBASIC XLL framework for building high-performance Excel User-Defined Functions.

---

## What is tbXLLerator?

tbXLLerator is a framework that lets you build Excel XLL add-ins entirely in [twinBASIC](https://twinbasic.com), without requiring C or C++. It wraps Jon Johnson's [ExcelSDK](https://github.com/fafalone/TBXLLUDF) and provides a structured layer for argument binding, type coercion, memory management, UDF registration, and Excel callback mechanics.

The result is that you can focus on your modeling logic while the framework handles the low-level plumbing that makes XLL development notoriously difficult.

---

## Why XLL?

Excel supports several add-in technologies. Here's where XLL fits in:

| Technology | Language | Thread-Safe Calc | Performance | Function Wizard |
|------------|----------|-----------------|-------------|-----------------|
| **XLL** | C/C++, twinBASIC | Yes | Fastest | Full support |
| **COM Add-in** | VB.NET, C# | No | Moderate | Limited |
| **VBA** | VBA | No | Slowest | No |
| **Office.js** | JavaScript | N/A | Varies | Partial |

XLL add-ins are native DLLs that plug directly into Excel's calculation engine. The key advantages are multithreaded recalculation (Excel can call your UDF concurrently across CPU cores) and full Function Wizard integration (descriptions, argument help, help topics).

---

## Wiki Structure

This wiki is organized into three sections:

**Getting Started** covers installation, prerequisites, and a minimal working example to get you up and running.

**Core Concepts** explains the foundational ideas behind XLL development — how memory works, how arguments flow between Excel and your code, how registration tells Excel about your functions, and how thread safety is achieved.

**How-To Guides** are task-oriented pages that walk through specific patterns: arrays, error handling, optional arguments, delegating to Excel built-ins, cell references, and performance tuning.

**Reference** provides lookup tables for enums, type codes, and architectural details.

---

## Quick Links

- **New to XLL development?** Start with [[XLL Fundamentals]] then [[Writing Your First UDF]]
- **Setting up?** See [[Installation]]
- **Looking for a specific pattern?** Check the How-To Guides in the sidebar
- **Migrating from VBA UDFs?** Read [[XLL Fundamentals]] for the conceptual shift, then [[Quick Start]] for the practical steps
