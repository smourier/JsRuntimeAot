namespace JsRuntimeAot.Tests;

internal class Program
{
    static void Main(string[] args)
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
            var obj = rt.ParseScript("function square(number) { return number * number; }")!;
            var names = obj.PropertyNames;
            Console.WriteLine(obj.Call(5));
        });
    }
}
