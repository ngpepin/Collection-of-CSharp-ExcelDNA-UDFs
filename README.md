# Excel-DNA Utility Functions

This repository provides a collection of high-performance, thread-safe User Defined Functions (UDFs) for Microsoft Excel, developed using [Excel-DNA](https://excel-dna.net/). These functions are designed to enhance Excel's capabilities, offering advanced features for power users and developers.

Author: Nicolas Pepin
Date: 2026-07
Version: 3.9.0
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

23. **STRING_DIFF(s1, s2, minLength)**
    - Returns maximal differing substrings with a minimum length
    - **Usage**: `=STRING_DIFF("Hello there, how are you", "Hello there how are you", 1)`
    - **Returns**: Dynamic array of differing substrings from both inputs

24. **TEXT_BEFORE(text, delimiter, [instance])**
    - Returns text before a selected delimiter occurrence
    - **Usage**: `=TEXT_BEFORE("North|America|Canada", "|", 2)`
    - **Returns**: `"North|America"`; returns `#N/A` when the occurrence is absent

25. **TEXT_AFTER(text, delimiter, [instance])**
    - Returns text after a selected delimiter occurrence
    - **Usage**: `=TEXT_AFTER("North|America|Canada", "|", 2)`
    - **Returns**: `"Canada"`; returns `#N/A` when the occurrence is absent

26. **REGEX_ISMATCH(text, pattern, [ignoreCase])**
    - Tests text against a .NET regular expression with a one-second timeout
    - **Usage**: `=REGEX_ISMATCH("INV-2026-0042", "^INV-[0-9]{4}-[0-9]{4}$")`
    - **Returns**: `TRUE` for a match, `FALSE` otherwise, or `#VALUE!` for an invalid pattern

27. **REGEX_EXTRACT(text, pattern, [group])**
    - Returns the first regex match or a numbered or named capture group
    - **Usage**: `=REGEX_EXTRACT("Order 8472", "Order ([0-9]+)", 1)`
    - **Returns**: `"8472"`; returns `#N/A` when no match or group exists

28. **REGEX_REPLACE(text, pattern, replacement, [ignoreCase])**
    - Replaces every regex match using standard .NET replacement syntax
    - **Usage**: `=REGEX_REPLACE("A  B   C", "\s+", " ")`
    - **Returns**: `"A B C"`; returns `#VALUE!` for an invalid pattern

29. **ARRAY_UNIQUE(inputArray, [ignoreCase])**
    - Returns unique nonblank values as a vertical dynamic array in first-seen order
    - **Usage**: `=ARRAY_UNIQUE(A2:A100, TRUE)`
    - **Returns**: A spill range; `TRUE` makes text comparisons case-insensitive

30. **ARRAY_DISTINCT_COUNT(inputArray, [ignoreCase])**
    - Counts unique nonblank values using the same comparison rules as `ARRAY_UNIQUE`
    - **Usage**: `=ARRAY_DISTINCT_COUNT(A2:A100, TRUE)`
    - **Returns**: Integer distinct count

31. **NUM_CLAMP(value, minimum, maximum)**
    - Restricts a number to an inclusive lower and upper bound
    - **Usage**: `=NUM_CLAMP(125, 0, 100)`
    - **Returns**: `100`; returns `#VALUE!` when inputs are invalid or minimum exceeds maximum

### Machine Learning and AI vector UDFs

32. **VECTOR_DOT(vectorA, vectorB)**
    - Computes the dot product of equally sized row or column vectors
    - **Usage**: `=VECTOR_DOT(A2:A5, B2:B5)`
    - **Returns**: Scalar dot product

33. **VECTOR_NORM(vector, [p])**
    - Computes an L-p norm, with Euclidean norm (`p=2`) as the default
    - **Usage**: `=VECTOR_NORM(A2:A5, 2)`
    - **Returns**: Scalar norm for any positive finite `p`

34. **VECTOR_NORMALIZE(vector, [p])**
    - Normalizes a vector to unit L-p norm while preserving row or column orientation
    - **Usage**: `=VECTOR_NORMALIZE(A2:A5)`
    - **Returns**: Dynamic vector; a zero vector returns `#DIV/0!`

35. **VECTOR_COSINE_SIMILARITY(vectorA, vectorB)**
    - Measures directional similarity, commonly used for embeddings
    - **Usage**: `=VECTOR_COSINE_SIMILARITY(A2:A5, B2:B5)`
    - **Returns**: Similarity from approximately `-1` to `1`

36. **VECTOR_EUCLIDEAN_DISTANCE(vectorA, vectorB)**
    - Computes straight-line distance between feature vectors
    - **Usage**: `=VECTOR_EUCLIDEAN_DISTANCE(A2:A5, B2:B5)`
    - **Returns**: Nonnegative scalar distance

37. **VECTOR_MANHATTAN_DISTANCE(vectorA, vectorB)**
    - Computes L1 distance between feature vectors
    - **Usage**: `=VECTOR_MANHATTAN_DISTANCE(A2:A5, B2:B5)`
    - **Returns**: Sum of absolute element differences

38. **VECTOR_SOFTMAX(vector)**
    - Converts logits into probabilities using a numerically stable implementation
    - **Usage**: `=VECTOR_SOFTMAX(B2:B6)`
    - **Returns**: Dynamic vector with values summing to approximately `1`

39. **VECTOR_SIGMOID(vector)**
    - Applies logistic sigmoid activation element-wise
    - **Usage**: `=VECTOR_SIGMOID(B2:B6)`
    - **Returns**: Dynamic vector with values between `0` and `1`

40. **VECTOR_RELU(vector)**
    - Applies rectified linear activation element-wise
    - **Usage**: `=VECTOR_RELU(B2:B6)`
    - **Returns**: Dynamic vector with negative values replaced by zero

### Machine Learning and AI matrix UDFs

41. **MATRIX_STANDARDIZE_COLUMNS(matrix, [sample])**
    - Z-score standardizes each feature column, treating rows as observations
    - **Usage**: `=MATRIX_STANDARDIZE_COLUMNS(A2:D101)`
    - **Returns**: Dynamic matrix; constant columns become zeros

42. **MATRIX_MINMAX_SCALE_COLUMNS(matrix, [targetMin], [targetMax])**
    - Scales every feature column into a requested range, defaulting to `[0,1]`
    - **Usage**: `=MATRIX_MINMAX_SCALE_COLUMNS(A2:D101, -1, 1)`
    - **Returns**: Dynamic scaled matrix; constant columns use `targetMin`

43. **MATRIX_PAIRWISE_DISTANCE(matrix, [metric])**
    - Builds a square row-to-row distance matrix
    - **Usage**: `=MATRIX_PAIRWISE_DISTANCE(A2:D20, "cosine")`
    - **Returns**: Dynamic matrix using `euclidean`, `manhattan`, or `cosine` distance

44. **MATRIX_COVARIANCE(matrix, [sample])**
    - Computes covariance between feature columns
    - **Usage**: `=MATRIX_COVARIANCE(A2:D101, TRUE)`
    - **Returns**: Symmetric feature-by-feature covariance matrix

45. **MATRIX_ONE_HOT(labels, [classLabels])**
    - One-hot encodes a row or column label vector
    - **Usage**: `=MATRIX_ONE_HOT(A2:A101)`
    - **Returns**: Dynamic indicator matrix; inferred class order is first-seen unless supplied explicitly

46. **MATRIX_CONFUSION(actual, predicted, [classLabels])**
    - Builds a multiclass confusion matrix with actual classes as rows and predictions as columns
    - **Usage**: `=MATRIX_CONFUSION(A2:A101, B2:B101)`
    - **Returns**: Dynamic count matrix using explicit or first-seen class order


47. **VECTOR_LOG_SOFTMAX(vector)**
    - Converts logits to log probabilities using a stable log-sum-exp calculation
    - **Usage**: `=VECTOR_LOG_SOFTMAX(B2:B6)`
    - **Returns**: Dynamic row or column vector preserving the input orientation

48. **VECTOR_TOP_K(vector, k, [largest])**
    - Ranks a vector and returns source positions with their values
    - **Usage**: `=VECTOR_TOP_K(B2:B100, 5, TRUE)`
    - **Returns**: A `k`-by-2 spill array containing 1-based index and value; ties preserve source order

49. **MATRIX_LINEAR_PREDICT(matrix, weights, [bias])**
    - Applies a dense linear layer to row observations
    - **Usage**: `=MATRIX_LINEAR_PREDICT(A2:D100, F2:H5, J2:L2)`
    - **Returns**: Observation-by-output prediction matrix; bias may be omitted, scalar, or output-length

50. **MATRIX_CORRELATION(matrix)**
    - Computes Pearson correlation between feature columns
    - **Usage**: `=MATRIX_CORRELATION(A2:D100)`
    - **Returns**: Symmetric feature correlation matrix; constant columns return `#DIV/0!`

51. **MATRIX_KMEANS_ASSIGN(matrix, centroids, [metric])**
    - Assigns each observation to its nearest centroid
    - **Usage**: `=MATRIX_KMEANS_ASSIGN(A2:D100, F2:I6, "euclidean")`
    - **Returns**: Two-column spill array containing 1-based centroid index and distance

### Version 3.9.0 design notes

- Removed the four low-value or Excel-duplicative UDFs `TRIM_RIGHT`, `TRIM_LEFT`, `SAFE_DIVIDE`, and `DATE_ISBUSINESSDAY`.
- The ML/AI collection now contains 20 deterministic, side-effect-free UDFs, all registered as thread-safe.
- Numeric vectors must be a single row or column; numeric matrices treat rows as observations and columns as features.
- Spill results use correctly sized `object[,]` arrays and preserve vector orientation where relevant.
- Feature scaling, activations, similarities, distances, covariance, one-hot encoding, and confusion matrices require no external packages.
- Invalid shapes, nonnumeric cells, incompatible dimensions, unsupported metrics, and invalid class definitions return explicit Excel errors.

## Integration with eSharper

The standalone `Custom-Excel-DNA-UDFs-ESharper.cs` file remains compatible with the [eSharper](https://vlasovstudio.com/esharper/) Excel add-in container for rapid interactive development. The normal deployment path uses `deploy.sh` and does not require eSharper.

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

**Automated deployment:**
1. Install Mono (`mcs` and `mono`), `curl`, and Python 3.
2. Run `./deploy.sh`.
3. Copy the contents of `dist/` together and load `dist/CustomExcelDnaUdfs.xll` from Excel Add-ins.

The script downloads pinned Excel-DNA and Office Interop packages, compiles the managed assembly, creates the `.dna` manifest, and builds a standard Excel-DNA XLL deployment bundle. It defaults to 64-bit Excel; set `ARCH=x86` for 32-bit Excel. On Windows or under Wine, it also creates `CustomExcelDnaUdfs-packed.xll` when packing is available. Use `PACK_XLL=false` to skip packing or `PACK_XLL=true` to require it.

**Testing:** Run `./tests/run-tests.sh` for the command-line C# harness. Open `Tests.xlsx` with `dist/CustomExcelDnaUdfs.xll` loaded to run the worksheet assertions on the `AIML_UDF_Tests` sheet.

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
