# Excel-DNA Utility Functions

This repository provides a collection of high-performance, thread-safe User Defined Functions (UDFs) for Microsoft Excel, developed using [Excel-DNA](https://excel-dna.net/). These functions are designed to enhance Excel's capabilities, offering advanced features for power users and developers.

Author: Nicolas Pepin
Date: 2025-06
Version: 3.1.1
Licensing: MIT

## Table of Contents

- [Overview](#overview)
- [Available Functions](#available-functions)
- [Integration with eSharper](#integration-with-esharper)
- [C# Version Compatibility](#c-version-compatibility)
- [Building and Deployment](#building-and-deployment)
- [License](#license)
- [Appendix A: Excel-DNA Technical Overview](#appendix-a-excel-dna-technical-overview)

## Overview

The UDFs in this collection are implemented in C# and can be integrated into Excel through the Excel-DNA framework. They are particularly useful for tasks that require:

- Advanced data manipulation.
- Enhanced control over Excel's calculation settings.
- Improved worksheet function capabilities.

## Available Functions

## Summary of Functions


1. **VEXCELDNA()**  
   - Returns the current version of the UDF collection  
   - **Usage**: `=VEXCELDNA()`  
   - **Returns**: String with the version number  

2. **SETTARGETVERSION(version)**  
   - Sets the target version for backward compatibility  
   - **Usage**: `=SETTARGETVERSION("2.0.0")`  
   - **Returns**: Confirmation string with the previous and current target version  

3. **GETTARGETVERSION()**  
   - Gets the current target version for backward compatibility  
   - **Usage**: `=GETTARGETVERSION()`  
   - **Returns**: String with the current target version  

4. **RECALCALL()**  
   - Triggers a full recalculation of the workbook  
   - **Usage**: `=RECALCALL()`  
   - **Returns**: `"TRUE"` on success  

5. **GETITERATIONSTATUS()**  
   - Returns Excel's iterative calculation settings  
   - **Usage**: `=GETITERATIONSTATUS()`  
   - **Returns**: String with status (ON/OFF), max iterations, and max change  

6. **SETITERATION(IterationOn, [maxIterations], [maxChange])**  
   - Configures Excel's iterative calculation settings  
   - **Usage**: `=SETITERATION(TRUE, 100, 0.001)`  
   - **Returns**: Confirmation string with current settings  

7. **ISVISIBLE([cachingTime])**  
   - Checks if a cell is visible (not hidden by rows/columns)  
   - **Usage**: `=ISVISIBLE(10)` (10 second cache duration)  
   - **Returns**: `"TRUE"` if visible, `"FALSE"` if hidden  

8. **DESCRIBE(cell_reference)**  
   - Returns a description of the cell's content type  
   - **Usage**: `=DESCRIBE(A1)`  
   - **Returns**: String describing the value type  

9. **INJECTVALUE(cell_reference, value)**  
   - Injects a value into a cell (stateful operation)  
   - **Usage**: `=INJECTVALUE(B2, "Test Value")`  
   - **Returns**: The injected value  

10. **FINDPOS(text, substring, instance)**  
    - Finds positions of substrings (case-insensitive)  
    - **Usage**: `=FINDPOS("Hello World", "o", 1)`  
    - **Returns**: Position number or error if not found  

11. **PUTOBJECT(name, value, [force], [debug])**  
    - Stores an object in temporary storage  
    - **Usage**: `=PUTOBJECT("temp1", A1:A10, TRUE)`  
    - **Returns**: The stored object  

12. **GETOBJECT(name, [debug])**  
    - Retrieves an object from temporary storage  
    - **Usage**: `=GETOBJECT("temp1")`  
    - **Returns**: The stored object or error  

13. **PURGEOBJECTS()**  
    - Clears all objects from temporary storage  
    - **Usage**: `=PURGEOBJECTS()`  
    - **Returns**: `"TRUE"` on success  

14. **TRUESPLIT(input_array, delimiter)**  
    - Splits strings into dynamic arrays  
    - **Usage**: `=TRUESPLIT(A1:A3, ",")`  
    - **Returns**: 2D array of split components  

15. **ISMEMBEROF(array1, array2)**  
    - Checks for common elements between arrays  
    - **Usage**: `=ISMEMBEROF(A1:A10, B1:B20)`  
    - **Returns**: `TRUE` if any match found  

16. **GETTHREADS()**  
    - Returns Excel's current thread count for calculations  
    - **Usage**: `=GETTHREADS()`  
    - **Returns**: Integer thread count  

17. **SETTHREADS(threadCount)**  
    - Configures Excel's calculation thread count  
    - **Usage**:  
      `=SETTHREADS(4)` (Use 4 threads)  
      `=SETTHREADS(0)` (Use all processors)  
    - **Returns**: Actual thread count set  

18. **HASHARRAY(input_array, [hashLength])**  
    - Computes a consistent hash value for an array of values  
    - **Usage**: `=HASHARRAY(A1:A10, 8)`  
    - **Returns**: Hash string (default length 8, range 4–32)  

19. **ISLOCALIP(ipAddress_string)**  
    - Checks if an IP address is a local IP (private or loopback)  
    - **Usage**: `=ISLOCALIP(ipAddress_string)`  
    - **Returns**: `TRUE` if local IP, `FALSE` otherwise or `#N/A` if invalid input  

20. **ARRAYSUBTRACT(arrayA, arrayB)**  
    - Subtracts values in `arrayB` from `arrayA`, preserving shape where possible  
    - **Usage**: `=ARRAYSUBTRACT(A1:A10, B1:B3)`  
    - **Returns**: Dynamic array of values from `arrayA` that are not present in `arrayB`  

21. **EXTRACTSUBSTR(inputString, startMarker, [endMarker])**  
    - Extracts a substring between start and end markers  
    - **Usage**: `=EXTRACTSUBSTR("A=[123] Z", "A=[", "]")`  
    - **Returns**: Extracted substring or `#N/A` if markers are not found  

22. **STRING_COMMON(s1, s2, minLength)**  
    - Returns maximal common substrings with a minimum length  
    - **Usage**: `=STRING_COMMON("Hello there, how are you", "Hello there how are you", 5)`  
    - **Returns**: Dynamic array of common substrings (empty if none meet `minLength`)  


24. **TRIM_RIGHT(s, x)**  
  - Trims `x` characters from the right end of string `s`  
  - **Usage**: `=TRIM_RIGHT("abcdef", 2)`  
  - **Returns**: `"abcd"` (removes last 2 characters; returns empty string if `x` >= length of `s`)  

25. **TRIM_LEFT(s, x)**  
  - Trims `x` characters from the left end of string `s`  
  - **Usage**: `=TRIM_LEFT("abcdef", 2)`  
  - **Returns**: `"cdef"` (removes first 2 characters; returns empty string if `x` >= length of `s`)  

## Integration with eSharper

To simplify the management and usage of these UDFs within Excel 365, this project leverages the [eSharper](https://vlasovstudio.com/esharper/) Excel add-in container.

## C# Version Compatibility

These UDFs use features from C# 10. Attempting to use syntax from later C# versions may cause compilation errors.

**Compatibility Notes:**
- Excel-DNA supports .NET Framework 4.5.2+ and .NET 6+/8.
- eSharper relies on the .NET version available within Excel, potentially limiting newer features.

## Building and Deployment

**Requirements:**
- Visual Studio 2022+
- .NET Framework 4.7.2 SDK or .NET 6.0 SDK
- Excel-DNA NuGet package

**Steps:**
1. Clone the repository.
2. Open in Visual Studio.
3. Build to generate `.xll`.
4. Load `.xll` in Excel via Add-ins menu.

## License

MIT License. See `LICENSE` file.

---

## Appendix A: Excel-DNA Technical Overview


#### **What is Excel-DNA?**

ExcelDNA is a powerful library that allows developers to create high-performance Excel add-ins using .NET languages (like C# or VB.NET). Here's a technical breakdown of how it works:

##### **1. Core Architecture**

ExcelDNA bridges Excel's native C API (the **Excel XLL SDK**) with the .NET runtime. It does this by:

* **Compiling .NET code into an XLL**: An XLL is a DLL specifically designed for Excel. ExcelDNA generates a thin native XLL stub that loads the .NET runtime and hosts your managed code.

* **Using Managed/Unmanaged Interop**: The XLL acts as a bridge between Excel (unmanaged C/C++ world) and .NET (managed world) using P/Invoke and COM Interop.

##### **2. Key Components**

* **ExcelDna.Integration.dll**: Provides the core API for registering functions, handling callbacks, and marshaling data between Excel and .NET.

* **ExcelDna.Loader.dll**: Manages the dynamic loading of .NET assemblies into Excel.

* **ExcelDnaPack**: A tool that bundles custom .NET assemblies and dependencies into a single `.xll` file for deployment.

##### **3. Function Registration**

When Excel loads the XLL:

* **ExcelDNA scans your .NET assembly** for methods marked with Excel-specific attributes (e.g., `[ExcelFunction]`).

* **It generates Excel-compatible exports** (via `xlAutoOpen` and `xlAddInManagerInfo` callbacks).

* **Wraps .NET methods** in native XLL-compatible functions, handling type conversion between:

  * Excel `XLOPER`/`XLOPER12` types ↔ .NET types (double, string, object\[,], etc.).

  * Excel arrays ↔ .NET `object[,]` or `double[,]`.

##### **4. Marshaling & Memory Management**

* **Arguments passed from Excel** are converted into .NET types.

* **Return values** from .NET are packed back into Excel-compatible structures.

* **ExcelDNA manages memory** to prevent leaks (e.g., freeing temporary `XLOPER`s).

##### **5. Asynchronous & Multithreading Support**

* Excel is single-threaded (STA), but ExcelDNA allows **async functions** via `[ExcelAsync]`.

* Uses **.NET Tasks** to run computations in the background and return results later.

##### **6. RTD (Real-Time Data) Support**

* Implements Excel's **RTD server** interface for push-based real-time updates.

* Managed .NET code can push data to Excel cells in real time.

##### **7. COM & Ribbon Integration**

* If needed, ExcelDNA can expose .NET classes to Excel via COM (for UDFs or macros).

* Supports customizing the Ribbon UI via **Fluent UI XML**.

##### **8. Debugging & Deployment**

* Works with **Visual Studio debugging** (attach to Excel process).

* Packaged as a single `.xll` file (no separate installer needed).

##### **9. Performance Considerations**

* Minimal overhead (\~native speed) due to direct XLL integration.

* Avoids COM where possible for better performance.

##### **10. Comparison to Other Tech (VSTO, COM Add-ins)**

* **Faster** than VSTO (no COM overhead).

* **Lighter** than VSTO (no need for a separate runtime).

* **More flexible** than VBA (full .NET ecosystem access).

##### **Example Flow (Calling a .NET Function from Excel)**

1. User enters `=MyNetFunction(A1)` in Excel.

2. Excel calls the XLL’s exported stub.

3. ExcelDNA marshals arguments to .NET.

4. Your `[ExcelFunction]` method runs in .NET.

5. Return value is marshaled back to Excel.

ExcelDNA essentially makes .NET a first-class citizen in Excel while maintaining high performance and compatibility.

#### **How does it compare with Python-based approaches?**

ExcelDNA (for .NET) and Python integration in Excel serve different purposes and have distinct technical approaches. Here’s a detailed comparison:

##### **1. Technical Implementation**

| **Aspect**            | **ExcelDNA (.NET)**                       | **Python in Excel**                                                                 |
| --------------------- | ----------------------------------------- | ----------------------------------------------------------------------------------- |
| **Integration Level** | Deep XLL integration (native Excel C API) | Officially supported by Microsoft (via PyXLL, xlwings, or built-in Python in Excel) |
| **Performance**       | Near-native (minimal overhead)            | Slower (Python interpreter + marshaling)                                            |
| **Language**          | C#, F#, VB.NET                            | Python                                                                              |
| **Deployment**        | Single `.xll` file                        | Requires Python runtime, dependencies                                               |
| **Concurrency**       | Supports async via `[ExcelAsync]`         | Limited (Python's GIL can bottleneck multithreading)                                |
| **Real-Time Data**    | RTD support (push updates)                | Possible with PyXLL/xlwings, but slower                                             |
| **Debugging**         | Easy (attach to Excel process)            | Requires IDE setup (e.g., VS Code, PyCharm)                                         |

##### **2. Functionality & Use Cases**

| **Feature**                       | **ExcelDNA**                | **Python in Excel**                  |
| --------------------------------- | --------------------------- | ------------------------------------ |
| **User-Defined Functions (UDFs)** | Yes (high performance)      | Yes (slower, but flexible)           |
| **Macros & Automation**           | Yes (via `[ExcelMacro]`)    | Yes (xlwings, COM)                   |
| **Data Processing**               | Fast (direct .NET arrays)   | Slower (Pandas/NumPy marshaling)     |
| **Machine Learning**              | ML.NET, TensorFlow\.NET     | Full scikit-learn/TensorFlow/PyTorch |
| **Excel UI Control**              | Custom Ribbon, WinForms/WPF | Limited (depends on tool)            |
| **Cross-Platform**                | Windows-only                | Works on Mac (xlwings)               |

##### **3. Pros and Cons**

###### **ExcelDNA (.NET)**

 **Pros:**

* Blazing fast (native XLL performance).

* Direct access to Excel’s C API (low-level control).

* Strong typing (C#/F# reduces runtime errors).

* Easy deployment (single `.xll` file).

* Full .NET ecosystem (e.g., parallel computing, databases).

 **Cons:**

* Windows-only (no macOS support).

* Requires .NET knowledge.

* Only works with desktop version of Excel.

* Less popular for data science than Python.

###### **Python in Excel**

 **Pros:**

* **Built-in Python in Excel (Microsoft 365)**: No add-ins needed.

* **Huge ecosystem** (Pandas, NumPy, scikit-learn, etc.).

* **Better for prototyping** (Jupyter-like workflows).

* **Cross-platform** (xlwings works on Mac).

 **Cons:**

* **Slower** (Python interpreter + data marshaling).

* **Dependency hell** (conda/pip environments).

* **Limited real-time performance** (no RTD in pure Python).

* **Debugging is harder** (external IDE needed).

- - -

##### **4. When to Use Which?**

* **Use ExcelDNA if:**

  * You need **maximum performance** (financial models, real-time data).

  * You’re already using **.NET/C#**.

  * You need **deep Excel integration** (custom UI, RTD, async).

* **Use Python in Excel if:**

  * You’re doing **data science/ML** (Pandas, scikit-learn).

  * You prefer **quick prototyping** (Jupyter-style).

  * You need **cross-platform** support (Mac + Windows).


##### **5. Emerging Trends**

* **Microsoft’s built-in Python in Excel** (2023+):

  * Runs **Python in the cloud** (not locally).

  * Seamless grid integration (no add-ins).

  * Still early (limited libraries, no local execution).

* **Alternatives**:

  * **PyXLL**: Commercial, high-performance Python XLL.

  * **xlwings**: Free, but COM-based (slower).

##### **Final Verdict**

* **ExcelDNA** = **Speed + Control** (best for .NET devs).

* **Python in Excel** = **Flexibility + Ecosystem** (best for data scientists).

- - -

## Appendix B: Using Without eSharper

You do **not** need the eSharper add-in to use these Excel-DNA functions. They can be deployed as standard Excel add-ins using the following steps:

###  Requirements

* [Excel-DNA](https://excel-dna.net/)

* Visual Studio (recommended) or a text editor

* .NET Framework 4.7.2 or later _(for compatibility with most versions of Excel)_

* Excel (2010 or newer recommended)


###  Steps to Compile and Use the UDFs

#### 1. **Create or Use a `.dna` File**

Create a file named `MyAddIn.dna` with the following content:

``` xml
<DnaLibrary Name="MyExcelFunctions" RuntimeVersion="v4.0">
  <ExternalLibrary Path="MyFunctions.dll" />
</DnaLibrary>
```

* `MyFunctions.dll` is the compiled output of your `.cs` code (see next step).

* `RuntimeVersion` must match the .NET version used for compiling the DLL.

#### 2. **Compile Your `.cs` Code**

Compile your C# file into a class library (`.dll`). You can do this using:

* Visual Studio (File > New > Project > Class Library)

* Or with the command line:

``` bash

csc /target:library /out:MyFunctions.dll Custom-Excel-DNA-UDFs.cs
```

#### 3. **Download Excel-DNA Loader**

Download the latest [Excel-DNA binaries](https://github.com/Excel-DNA/ExcelDna/releases) and place the following in your project folder:

* `ExcelDna.Integration.dll`

* `ExcelDna.Loader.dll`

* `ExcelDna.xll` _(rename this to `MyAddIn.xll` for clarity)_

#### 4. **Build the `.xll` Add-In**

To link everything together, you should have:

``` bash

MyAddIn.dna
MyFunctions.dll
MyAddIn.xll (copied/renamed from ExcelDna.xll)
```

**Optional**: Use the Excel-DNA `Pack` utility to bundle the `.xll`, `.dll`, and `.dna` into a single file:

``` bash

ExcelDnaPack.exe MyAddIn.dna
```

This will create `MyAddIn-packed.xll`.

#### 5. **Load the Add-In in Excel**

* Open Excel.

* Go to `File > Options > Add-Ins`.

* At the bottom, select **Manage: Excel Add-ins**, and click **Go...**.

* Click **Browse**, find your `.xll` or `*-packed.xll` file, and open it.

* The UDFs will now be available as native Excel functions.

- - -

###  Notes

* Excel-DNA add-ins are fully portable and do not require administrator installation.

* You can distribute the `.xll` or `.xll + .dll` pair to other users.

* No COM registration is needed.

* You can sign your `.dll` for macro security compliance.

- - -
Excel-DNA is powerful and flexible, making it ideal for deploying managed-code add-ins without the overhead and complexity of COM registration or VSTO.
