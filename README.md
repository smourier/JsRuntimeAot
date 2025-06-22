# JsRuntimeAot
A .NET9+ AOT compatible wrapper over the Microsoft "Chakra" JavaScript engine.

=> It's just one 2500 lines of C# file that allows you to run Javascript code from .NET.

# Why?
Obviously, the [Chakra Javascript Engine](https://en.wikipedia.org/wiki/Chakra_(JavaScript_engine)) is deprecated, however using it has some advantages:

* It still works fine for most use cases.
* It's performance is great.
* It's installed in Windows (x86, x64, Arm64), so you don't have to distribute any native binaries, contrary to [Clearscript](https://github.com/microsoft/ClearScript) for example.
* Although its deprecated, it's still maintained by Microsoft (I guess, at least for security fixes).
* It has VARIANT <=> Javascript object conversion, which [ChakraCore](https://github.com/chakra-core/ChakraCore) hasn't.
* This all allows for easier .NET AOT compatiblity, which is, AFAIK, not currently the case with 100% .NET Javascript implementations.

So if one just needs "some level of javascript support" closely integrated with a modern .NET application, with zero deployment impact, it can be very useful.
