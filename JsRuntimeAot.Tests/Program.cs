namespace JsRuntimeAot.Tests;

internal class Program
{
    static void Main(string[] args)
    {
        var input = "2 / 3.2";
        var sum = JsRuntime.Eval(input);
        Console.WriteLine($"{input} = {sum}");
    }
}
