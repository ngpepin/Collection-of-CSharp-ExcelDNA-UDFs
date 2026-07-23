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
- UDF names are ALL_CAPS with underscores (e.g., `VECTOR_NORMALIZE`, `HASHARRAY`).
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

### 11. **Pure Utility and ML UDF Pattern (Version 3.9.0)**
- Prefer deterministic, side-effect-free implementations for text, regex, array, numeric, and date utilities.
- Do not call `MaybeVolatile()` unless the result genuinely depends on external workbook or application state.
- For optional worksheet arguments, accept `object` and explicitly handle `ExcelMissing` and `ExcelEmpty`.
- Bound regular-expression execution with a timeout and convert invalid patterns or timeout failures to `ExcelError.ExcelErrorValue`.
- Preserve first-seen order for de-duplication UDFs; use a `HashSet<string>` only as the membership index.
- Normalize typed array values with stable keys so numbers, booleans, dates, errors, and text remain distinguishable.
- Return empty dynamic arrays as `new object[0, 0]`, matching the established collection convention.

## Version 3.9.0 Utility and ML/AI UDFs

The low-value or Excel-duplicative `TRIM_RIGHT`, `TRIM_LEFT`, `SAFE_DIVIDE`, and `DATE_ISBUSINESSDAY` functions were removed.

| UDF | Purpose | Key behavior |
| --- | --- | --- |
| `TEXT_BEFORE` | Text extraction | Uses a 1-based delimiter occurrence and returns `#N/A` when absent |
| `TEXT_AFTER` | Text extraction | Returns text following the selected delimiter occurrence |
| `REGEX_ISMATCH` | Regex validation | Supports optional case-insensitive matching and a one-second timeout |
| `REGEX_EXTRACT` | Regex extraction | Supports whole-match, numbered-group, and named-group output |
| `REGEX_REPLACE` | Regex transformation | Uses standard .NET replacement syntax and bounded execution |
| `ARRAY_UNIQUE` | Array de-duplication | Ignores blanks, preserves first-seen values, and spills vertically |
| `ARRAY_DISTINCT_COUNT` | Array summary | Reuses the same comparison rules as `ARRAY_UNIQUE` |
| `NUM_CLAMP` | Numeric guardrail | Validates bounds and clamps inclusively |
| `VECTOR_DOT` | Vector algebra | Requires equally sized row or column vectors |
| `VECTOR_NORM` | Vector magnitude | Supports positive finite L-p norms |
| `VECTOR_NORMALIZE` | Feature normalization | Preserves vector orientation and rejects zero norm |
| `VECTOR_COSINE_SIMILARITY` | Embedding similarity | Rejects zero vectors with `#DIV/0!` |
| `VECTOR_EUCLIDEAN_DISTANCE` | Geometric distance | Computes L2 distance |
| `VECTOR_MANHATTAN_DISTANCE` | Robust distance | Computes L1 distance |
| `VECTOR_SOFTMAX` | Probability activation | Uses max-shift numerical stabilization |
| `VECTOR_SIGMOID` | Neural activation | Uses branch-stable logistic evaluation |
| `VECTOR_RELU` | Neural activation | Replaces negative values with zero |
| `MATRIX_STANDARDIZE_COLUMNS` | Feature preprocessing | Rows are observations; constant columns become zero |
| `MATRIX_MINMAX_SCALE_COLUMNS` | Feature preprocessing | Supports custom target bounds |
| `MATRIX_PAIRWISE_DISTANCE` | Similarity analysis | Supports Euclidean, Manhattan, and cosine distance |
| `MATRIX_COVARIANCE` | Feature statistics | Returns sample or population covariance |
| `MATRIX_ONE_HOT` | Categorical encoding | Uses explicit or first-seen class order |
| `MATRIX_CONFUSION` | Classification evaluation | Actual classes are rows; predicted classes are columns |
| `VECTOR_LOG_SOFTMAX` | Probability activation | Uses stable log-sum-exp evaluation |
| `VECTOR_TOP_K` | Ranking | Returns stable 1-based indices and values |
| `MATRIX_LINEAR_PREDICT` | Dense inference | Supports one or many outputs plus scalar/vector bias |
| `MATRIX_CORRELATION` | Feature statistics | Returns Pearson feature correlations |
| `MATRIX_KMEANS_ASSIGN` | Clustering | Returns nearest centroid index and distance |

### ML/AI array implementation rules
- Numeric vectors must be exactly one row or one column; preserve that orientation in vector spill outputs.
- Numeric matrices use rows as observations and columns as features unless explicitly documented otherwise.
- Reject blanks, Excel errors, nonnumeric cells, mismatched vector lengths, and invalid matrix shapes rather than silently coercing them.
- Use max-shift stabilization for softmax and branch-stable sigmoid evaluation.
- For zero-variance feature columns, emit zero during standardization and `targetMin` during min-max scaling.
- In pairwise distance matrices, compute one triangle and mirror it to preserve symmetry and reduce work.
- For categorical encoders, use `BuildValueKey` so text, numbers, booleans, and dates remain type-distinct.
- When class labels are omitted, preserve deterministic first-seen order; when supplied, reject duplicates and return `#N/A` for observations outside the class list.
- Keep implementations dependency-free and compatible with C# 10, .NET Framework 4.7.2+, and .NET 6+.

### Shared helpers added for these UDFs
- `FindDelimiterOccurrence`, `TryGetOptionalPositiveInt`, and `GetOptionalBool` support optional utility arguments.
- `GetUniqueValues`, `BuildValueKey`, and `BuildObjectColumnArray` implement stable typed de-duplication.
- `TryGetNumericVector`, `TryGetNumericMatrix`, `BuildNumericVector`, and `BuildNumericMatrix` centralize strict spill-array conversion.
- `ComputeVectorNorm`, `TryGetOptionalDouble`, and `TryGetDistanceMetric` centralize vector and distance validation.
- `TryGetLabelVector`, `TryGetClassLabels`, and `BuildLabelIndex` centralize categorical encoding and evaluation.
- `TryGetBiasVector` validates scalar and vector biases for dense linear inference.
- `IndexedValue` provides stable value ranking for `VECTOR_TOP_K`.

### Formal testing and deployment
- `tests/ExcelDnaStubs.cs` provides the minimal Excel-DNA and Office Interop surface needed for command-line tests.
- `tests/AimlUdfTests.cs` exercises every ML/AI UDF plus shape and error cases.
- `tests/run-tests.sh` compiles the production source with the stubs using Mono and runs the suite.
- `Tests.xlsx` contains Excel-native worksheet assertions for every ML/AI UDF.
- `deploy.sh` builds the standalone Excel-DNA XLL deployment bundle from `Custom-Excel-DNA-UDFs.cs`; on Windows or under Wine it can additionally create a packed single-file XLL.
- `Custom-Excel-DNA-UDFs-ESharper.cs` is the synchronized standalone source for eSharper users and must remain C# 10 compatible.

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

_Last updated: 2026-07-23_
