# AGENT.md

## Project Overview

This project is a collection of high-performance, thread-safe User Defined Functions (UDFs) for Microsoft Excel, implemented using Excel-DNA in C#. The UDFs extend Excel's native capabilities, providing advanced data manipulation, control over calculation settings, and utility functions for power users and developers.

- **Main UDF file:** Custom-Excel-DNA-UDFs.cs
- **Documentation:** README.md (user-facing), AGENT.md (agent/coder-facing)
- **Language:** C# (up to version 10)
- **Excel-DNA:** https://excel-dna.net/

## Coding Practices & Proven Approaches

### 1. **Function Structure**
- All UDFs are implemented as `public static` methods in a single class (usually `C`).
- Each UDF is decorated with `[ExcelFunction]` and `[ExcelArgument]` attributes for Excel-DNA registration and argument documentation.
- Volatility is controlled globally via a `MaybeVolatile()` helper and a global flag, not per-function attributes.
- Input validation is performed at the start of each function, returning appropriate `ExcelError` values for invalid input.
- Functions are designed to be thread-safe and avoid side effects unless explicitly documented (e.g., INJECTVALUE).

### 2. **Argument Handling**
- Use helper methods (e.g., `TryGetInt`) to robustly parse Excel arguments (handles int, double, string, etc.).
- Always check for null, empty, or error values before processing arguments.
- For optional arguments, use `object` type and check for `ExcelMissing` or `ExcelEmpty`.

### 3. **Return Values**
- Return types are `object`, `string`, `bool`, or `object[,]` (for arrays).
- For errors, return `ExcelError.ExcelErrorValue`, `ExcelError.ExcelErrorNA`, etc.
- For dynamic arrays, build and return a properly sized `object[,]`.

### 4. **Volatility**
- Use `MaybeVolatile()` at the top of any function that should be volatile when the global flag is enabled.
- Do **not** use `[ExcelFunction(IsVolatile = true)]` directly; rely on the global switch for consistency and performance.

### 5. **Stateful Functions**
- Functions that maintain state (e.g., INJECTVALUE, PUTOBJECT/GETOBJECT) use static dictionaries for storage.
- State is cleared with PURGEOBJECTS() or when the workbook closes.

### 6. **Error Handling**
- Use try/catch blocks for any code that interacts with Excel interop or may throw exceptions.
- Return Excel error codes or descriptive error messages for user-facing functions.

### 7. **Documentation**
- All new UDFs must be documented in README.md (user-facing) and AGENT.md (developer-facing).
- Use XML comments (`/// <summary>...</summary>`) for all public methods.
- Update the summary of functions at the top of the main C# file.

### 8. **Naming Conventions**
- UDF names are ALL_CAPS with underscores (e.g., `TRIM_RIGHT`, `HASHARRAY`).
- Method names in C# use PascalCase (e.g., `TrimRight`).
- Arguments use camelCase or descriptive names.

### 9. **Compatibility**
- Only use C# 10 or earlier syntax/features.
- Avoid features from C# 11+ (e.g., list patterns, required properties, etc.).
- Ensure compatibility with .NET Framework 4.7.2+ and .NET 6+.

### 10. **Testing & Proven Patterns**
- New UDFs should follow the structure and validation patterns of existing, tested functions.
- Use helper methods for repeated logic (e.g., argument parsing, array building).
- Review similar UDFs for best practices before adding new ones.

## Adding New UDFs: Step-by-Step

1. **Define the function** in `Custom-Excel-DNA-UDFs.cs` as a `public static` method.
2. **Decorate** with `[ExcelFunction]` and `[ExcelArgument]` attributes.
3. **Add input validation** and call `MaybeVolatile()` if needed.
4. **Use helper methods** for argument parsing and error handling.
5. **Document** the function in README.md and AGENT.md.
6. **Update** the summary of functions at the top of the C# file.
7. **Test** the function in Excel for expected behavior and error handling.

## Example UDF Template

```csharp
/// <summary>
/// Brief description of the function.
/// </summary>
[ExcelFunction(Name = "UDF_NAME", Description = "Description for Excel", Category = "ExcelDNA Utilities")]
public static object UdfName(
    [ExcelArgument(Description = "Description of arg1")] string arg1,
    [ExcelArgument(Description = "Description of arg2")] object arg2)
{
    MaybeVolatile();
    // Input validation
    if (arg1 == null) return ExcelError.ExcelErrorNull;
    int n;
    if (!TryGetInt(arg2, out n) || n < 0) return ExcelError.ExcelErrorValue;
    // Function logic
    // ...
    return result;
}
```

## File Structure
- `Custom-Excel-DNA-UDFs.cs` — Main UDF implementations
- `README.md` — User-facing documentation
- `AGENT.md` — Developer/agent-facing documentation and coding standards
- `enhancements.md` — Ideas and notes for future improvements
- `Archive/` — Previous versions of the UDF file

## Notes for Coding Agents
- Always review recent changes and the summary of functions before adding new code.
- Maintain consistency in style, validation, and documentation.
- If in doubt, follow the pattern of the most similar existing UDF.
- Document any new helpers or patterns in this file.

---

_Last updated: 2026-02-02_
