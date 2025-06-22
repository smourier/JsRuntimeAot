using System.Diagnostics;

namespace JsRuntimeAot.Tests;

internal class Program
{
    static unsafe void Main()
    {
        var input = "1+2";
        var sum = JsRuntime.Eval(input);
        Console.WriteLine($"{input} = {sum}");

        input = "eval(1+2)";
        sum = JsRuntime.Eval(input);
        Console.WriteLine($"{input} = {sum}");

        using var rt = new JsRuntime();
        rt.WithContext(ctx =>
        {
            rt.RunScript("function square(n) { return n * n; }");
            var result = ctx.GlobalObject.CallFunction("square", null, 5);
            Console.WriteLine(result);
        });

        rt.WithContext(ctx =>
        {
            rt.RunScript("function square(n) { return n * n; }");

            var sw = Stopwatch.StartNew();
            var glo = ctx.GlobalObject;
            for (var i = 0; i < 1_000_000; i++)
            {
                var result = glo.CallFunction("square", null, 5);
                //Console.WriteLine(i + ":" + result);
            }
            Console.WriteLine($"Elapsed: {sw}.");
        });

        GC.Collect();
        GC.WaitForPendingFinalizers();
    }
}
