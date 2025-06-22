# JsRuntimeAot
A .NET8+ AOT compatible wrapper over the Microsoft "Chakra" JavaScript engine. It allows you to run Javascript code from .NET.

# Why?
Obviously, the [Chakra Javascript Engine](https://en.wikipedia.org/wiki/Chakra_(JavaScript_engine)) is deprecated, however using it has some advantages:

* It's installed in Windows (x86, x64, Arm64), so you don't have to distribute any native binaries, contrary to [Clearscript](https://github.com/microsoft/ClearScript).
* It has VARIANT <=> Javascript object conversion, which [ChakraCore](https://github.com/chakra-core/ChakraCore) hasn't.
* This all allows for easier .NET AOT compatiblity, which is, AFAIK, not currently the case with 100% .NET Javascript implementations.
