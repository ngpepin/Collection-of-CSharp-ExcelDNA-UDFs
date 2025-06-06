# Excel-DNA Utility Functions

This repository provides a collection of high-performance, thread-safe User Defined Functions (UDFs) for Microsoft Excel, developed using [Excel-DNA](https://excel-dna.net/). These functions are designed to enhance Excel's capabilities, offering advanced features for power users and developers.

Author: Nicolas Pepin
Date: 2025-03-01
Version: 2.5.0
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

**What is Excel-DNA?**

Excel-DNA (Excel .NET Assembly) is an open-source library that allows you to create high-performance Excel add-ins using .NET languages such as C# and VB.NET. Developed by Govert van Drimmelen, it provides a bridge between Excel and the .NET CLR.

**Key Features:**
- Write UDFs, macros, and custom Ribbon interfaces.
- Embed all logic in a single `.xll` file.
- Support for both .NET Framework and .NET 6/8.
- Open-source and actively maintained.

**Performance:**
- Much faster than VBA for computation-heavy tasks.
- Comparable to VSTO for managed code scenarios but easier to deploy.
- Native integration with Excel using C API, which reduces interop overhead compared to VSTO.

**How it Works:**
- Uses unmanaged C++ loader to bootstrap a CLR host.
- UDFs and UI elements are declared using attributes and registered via reflection.
- Allows embedded .NET assemblies directly in `.xll`.

**Constraints:**
- No GUI threading — avoid UI operations from background threads.
- Some Excel COM interfaces behave differently between .NET Framework and .NET Core.
- No automatic garbage collection for Excel objects — be cautious with memory usage.

**Best Practices:**
- Use `ExcelAsyncUtil.QueueAsMacro` for thread-safe UI interaction.
- Mark volatile functions only when needed.
- Cache results to avoid excessive recalculations.

## Using Without eSharper

You do **not** need the eSharper add-in to use these Excel-DNA functions. They can be deployed as standard Excel add-ins using the following steps:

### 🧰 Requirements

* [Excel-DNA](https://excel-dna.net/)

* Visual Studio (recommended) or a text editor

* .NET Framework 4.7.2 or later _(for compatibility with most versions of Excel)_

* Excel (2010 or newer recommended)

- - -

### 🛠️ Steps to Compile and Use the UDFs

#### 1. **Create or Use a `.dna` File**

Create a file named `MyAddIn.dna` with the following content:

```
xml
CopyEdit
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

```
bash
CopyEdit
csc /target:library /out:MyFunctions.dll Custom-Excel-DNA-UDFs.cs
```

#### 3. **Download Excel-DNA Loader**

Download the latest [Excel-DNA binaries](https://github.com/Excel-DNA/ExcelDna/releases) and place the following in your project folder:

* `ExcelDna.Integration.dll`

* `ExcelDna.Loader.dll`

* `ExcelDna.xll` _(rename this to `MyAddIn.xll` for clarity)_

#### 4. **Build the `.xll` Add-In**

To link everything together, you should have:

```
bash
CopyEdit
MyAddIn.dna
MyFunctions.dll
MyAddIn.xll (copied/renamed from ExcelDna.xll)
```

**Optional**: Use the Excel-DNA `Pack` utility to bundle the `.xll`, `.dll`, and `.dna` into a single file:

```
bash
CopyEdit
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

### 🧩 Notes

* Excel-DNA add-ins are fully portable and do not require administrator installation.

* You can distribute the `.xll` or `.xll + .dll` pair to other users.

* No COM registration is needed.

* You can sign your `.dll` for macro security compliance.

- - -
Excel-DNA is powerful and flexible, making it ideal for deploying managed-code add-ins without the overhead and complexity of COM registration or VSTO.