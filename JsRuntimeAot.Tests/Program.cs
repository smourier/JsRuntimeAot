using JsRt;

namespace JsRuntimeAot.Tests;

internal class Program
{
    static void Main(string[] args)
    {
        var sum = JsRuntime.Eval("[1 + 2]");
        Console.WriteLine($"1 + 2 = {sum}");
    }
}
