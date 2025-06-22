namespace JsRuntimeAot.Tests;

internal class Program
{
    static void Main()
    {
        var input = "1+2";
        var sum = JsRuntime.Eval(input);
        Console.WriteLine($"{input} => {sum}");

        input = "eval(1+2)";
        sum = JsRuntime.Eval(input);
        Console.WriteLine($"{input} => {sum}");

        using var rt = new JsRuntime();
        rt.WithContext(ctx =>
        {
            input = "function hello(n) { return 'héééééllooooo'; }";
            rt.RunScript(input);
            var result = ctx.GlobalObject.CallFunction("hello");
            Console.WriteLine($"{input} => {result}");
        });

        rt.WithContext(ctx =>
        {
            input = "function square(n) { return n * n; }";
            rt.RunScript(input);
            var sw = Stopwatch.StartNew();
            var glo = ctx.GlobalObject;
            var max = 1_000_000;
            for (var i = 0; i < max; i++)
            {
                var result = glo.CallFunction("square", null, 5);
                //Console.WriteLine(i + ":" + result);
            }
            Console.WriteLine($"{input} * {max} elapsed => {sw}.");
        });

        GC.Collect();
        GC.WaitForPendingFinalizers();
    }
}
