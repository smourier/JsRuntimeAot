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

        var rt = new JsRuntime();
        rt.CreateContext().Execute(() =>
        {
            rt.RunScript("function square(n) { return n * n; }");
            var result = JsContext.Current!.GlobalObject.CallFunction("square", null, 5);
            Console.WriteLine(result);
        });
    }
}
