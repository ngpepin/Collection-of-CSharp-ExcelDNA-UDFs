using System;
using System.Linq;
using System.Reflection;
using ExcelDna.Integration;

public static class AimlUdfTests
{
    private static int failures;

    private static object[,] Col(params object[] values)
    {
        object[,] result = new object[values.Length, 1];
        for (int i = 0; i < values.Length; i++) result[i, 0] = values[i];
        return result;
    }

    private static object[,] Row(params object[] values)
    {
        object[,] result = new object[1, values.Length];
        for (int i = 0; i < values.Length; i++) result[0, i] = values[i];
        return result;
    }

    private static void Pass(string name) { Console.WriteLine("PASS " + name); }

    private static void Fail(string name, string detail)
    {
        failures++;
        Console.WriteLine("FAIL " + name + ": " + detail);
    }

    private static void Near(string name, object actualObject, double expected, double tolerance = 1e-9)
    {
        if (!(actualObject is double)) { Fail(name, "not a double: " + actualObject); return; }
        double actual = (double)actualObject;
        if (Math.Abs(actual - expected) > tolerance) Fail(name, "expected " + expected + ", got " + actual);
        else Pass(name);
    }

    private static void Error(string name, object actual, ExcelError expected)
    {
        if (!object.Equals(actual, expected)) Fail(name, "expected " + expected + ", got " + actual);
        else Pass(name);
    }

    private static void MatrixNear(string name, object value, double[,] expected, double tolerance = 1e-9)
    {
        object[,] actual = value as object[,];
        if (actual == null) { Fail(name, "not an object[,] result"); return; }
        if (actual.GetLength(0) != expected.GetLength(0) || actual.GetLength(1) != expected.GetLength(1))
        {
            Fail(name, "shape mismatch");
            return;
        }

        for (int r = 0; r < expected.GetLength(0); r++)
            for (int c = 0; c < expected.GetLength(1); c++)
            {
                if (!(actual[r, c] is double) || Math.Abs((double)actual[r, c] - expected[r, c]) > tolerance)
                {
                    Fail(name, "mismatch at [" + r + "," + c + "]: expected " + expected[r, c] + ", got " + actual[r, c]);
                    return;
                }
            }
        Pass(name);
    }

    private static void RegistrationTests()
    {
        string[] names =
        {
            "VECTOR_DOT", "VECTOR_NORM", "VECTOR_NORMALIZE", "VECTOR_COSINE_SIMILARITY",
            "VECTOR_EUCLIDEAN_DISTANCE", "VECTOR_MANHATTAN_DISTANCE", "VECTOR_SOFTMAX",
            "VECTOR_SIGMOID", "VECTOR_RELU", "MATRIX_STANDARDIZE_COLUMNS",
            "MATRIX_MINMAX_SCALE_COLUMNS", "MATRIX_PAIRWISE_DISTANCE", "MATRIX_COVARIANCE",
            "MATRIX_ONE_HOT", "MATRIX_CONFUSION", "VECTOR_LOG_SOFTMAX", "VECTOR_TOP_K",
            "MATRIX_LINEAR_PREDICT", "MATRIX_CORRELATION", "MATRIX_KMEANS_ASSIGN"
        };

        var attributes = typeof(C).GetMethods(BindingFlags.Public | BindingFlags.Static)
            .Select(m => (ExcelFunctionAttribute)Attribute.GetCustomAttribute(m, typeof(ExcelFunctionAttribute)))
            .Where(a => a != null && names.Contains(a.Name)).ToList();

        foreach (string name in names)
        {
            var matching = attributes.Where(a => a.Name == name).ToList();
            if (matching.Count != 1) Fail("registration " + name, "expected one registration, got " + matching.Count);
            else if (!matching[0].IsThreadSafe) Fail("registration " + name, "not marked thread-safe");
            else Pass("registration " + name);
        }
    }

    public static int Main()
    {
        RegistrationTests();

        object[,] a = Col(1.0, 2.0, 3.0);
        object[,] b = Col(4.0, 5.0, 6.0);
        Near("VECTOR_DOT", C.VectorDot(a, b), 32.0);
        Near("VECTOR_NORM", C.VectorNorm(Col(3.0, 4.0), new ExcelMissing()), 5.0);
        MatrixNear("VECTOR_NORMALIZE", C.VectorNormalize(Row(3.0, 4.0), new ExcelMissing()), new double[,] { { 0.6, 0.8 } });
        Near("VECTOR_COSINE_SIMILARITY", C.VectorCosineSimilarity(Col(1.0, 0.0), Col(0.0, 1.0)), 0.0);
        Near("VECTOR_EUCLIDEAN_DISTANCE", C.VectorEuclideanDistance(Col(1.0, 2.0), Col(4.0, 6.0)), 5.0);
        Near("VECTOR_MANHATTAN_DISTANCE", C.VectorManhattanDistance(Col(1.0, 2.0), Col(4.0, 6.0)), 7.0);
        MatrixNear("VECTOR_SOFTMAX", C.VectorSoftmax(Col(0.0, 0.0)), new double[,] { { 0.5 }, { 0.5 } });
        MatrixNear("VECTOR_SIGMOID", C.VectorSigmoid(Col(0.0)), new double[,] { { 0.5 } });
        MatrixNear("VECTOR_RELU", C.VectorRelu(Row(-2.0, 0.0, 3.0)), new double[,] { { 0.0, 0.0, 3.0 } });

        object[,] matrix = new object[,] { { 1.0, 10.0 }, { 2.0, 20.0 }, { 3.0, 30.0 } };
        MatrixNear("MATRIX_STANDARDIZE_COLUMNS", C.MatrixStandardizeColumns(matrix, false),
            new double[,] { { -1.224744871391589, -1.224744871391589 }, { 0.0, 0.0 }, { 1.224744871391589, 1.224744871391589 } }, 1e-8);
        MatrixNear("MATRIX_MINMAX_SCALE_COLUMNS", C.MatrixMinMaxScaleColumns(matrix, new ExcelMissing(), new ExcelMissing()),
            new double[,] { { 0.0, 0.0 }, { 0.5, 0.5 }, { 1.0, 1.0 } });
        MatrixNear("MATRIX_PAIRWISE_DISTANCE", C.MatrixPairwiseDistance(new object[,] { { 0.0, 0.0 }, { 3.0, 4.0 } }, "euclidean"),
            new double[,] { { 0.0, 5.0 }, { 5.0, 0.0 } });
        MatrixNear("MATRIX_COVARIANCE", C.MatrixCovariance(new object[,] { { 1.0, 2.0 }, { 2.0, 4.0 }, { 3.0, 6.0 } }, true),
            new double[,] { { 1.0, 2.0 }, { 2.0, 4.0 } });
        MatrixNear("MATRIX_ONE_HOT", C.MatrixOneHot(Col("cat", "dog", "cat"), new ExcelMissing()),
            new double[,] { { 1.0, 0.0 }, { 0.0, 1.0 }, { 1.0, 0.0 } });
        MatrixNear("MATRIX_CONFUSION", C.MatrixConfusion(Col("cat", "dog", "cat"), Col("cat", "cat", "dog"), new ExcelMissing()),
            new double[,] { { 1.0, 1.0 }, { 1.0, 0.0 } });

        MatrixNear("VECTOR_LOG_SOFTMAX", C.VectorLogSoftmax(Row(0.0, 0.0)), new double[,] { { -Math.Log(2.0), -Math.Log(2.0) } });
        MatrixNear("VECTOR_TOP_K", C.VectorTopK(Col(10.0, 30.0, 20.0), 2, true), new double[,] { { 2.0, 30.0 }, { 3.0, 20.0 } });
        MatrixNear("MATRIX_LINEAR_PREDICT",
            C.MatrixLinearPredict(new object[,] { { 1.0, 2.0 }, { 3.0, 4.0 } }, new object[,] { { 1.0, 0.0 }, { 0.0, 2.0 } }, Row(1.0, 1.0)),
            new double[,] { { 2.0, 5.0 }, { 4.0, 9.0 } });
        MatrixNear("MATRIX_CORRELATION", C.MatrixCorrelation(new object[,] { { 1.0, 2.0 }, { 2.0, 4.0 }, { 3.0, 6.0 } }),
            new double[,] { { 1.0, 1.0 }, { 1.0, 1.0 } });
        MatrixNear("MATRIX_KMEANS_ASSIGN",
            C.MatrixKMeansAssign(new object[,] { { 0.0, 0.0 }, { 9.0, 9.0 }, { 1.0, 1.0 } }, new object[,] { { 0.0, 0.0 }, { 10.0, 10.0 } }, "euclidean"),
            new double[,] { { 1.0, 0.0 }, { 2.0, Math.Sqrt(2.0) }, { 1.0, Math.Sqrt(2.0) } });

        Error("vector shape", C.VectorNorm(new object[,] { { 1.0, 2.0 }, { 3.0, 4.0 } }, new ExcelMissing()), ExcelError.ExcelErrorValue);
        Error("vector length", C.VectorDot(Col(1.0), Col(1.0, 2.0)), ExcelError.ExcelErrorValue);
        Error("zero normalization", C.VectorNormalize(Col(0.0, 0.0), new ExcelMissing()), ExcelError.ExcelErrorDiv0);
        Error("invalid metric", C.MatrixPairwiseDistance(matrix, "chebyshev"), ExcelError.ExcelErrorValue);
        Error("invalid top k", C.VectorTopK(Col(1.0, 2.0), 3, true), ExcelError.ExcelErrorValue);
        Error("linear shape", C.MatrixLinearPredict(new object[,] { { 1.0, 2.0 } }, new object[,] { { 1.0 } }, new ExcelMissing()), ExcelError.ExcelErrorValue);
        Error("constant correlation", C.MatrixCorrelation(new object[,] { { 1.0, 2.0 }, { 1.0, 3.0 } }), ExcelError.ExcelErrorDiv0);
        Error("kmeans shape", C.MatrixKMeansAssign(new object[,] { { 1.0, 2.0 } }, new object[,] { { 1.0 } }, "euclidean"), ExcelError.ExcelErrorValue);

        Console.WriteLine(failures == 0 ? "ALL TESTS PASSED" : failures + " TEST(S) FAILED");
        return failures == 0 ? 0 : 1;
    }
}
