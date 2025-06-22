//NOSONAR
/*
MIT License

Copyright (c) 2025 Simon Mourier

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
*/
/*
AssemblyVersion: 1.3.2.0
AssemblyFileVersion: 1.3.2.0
AssemblyInformationalVersion: 1.3.2.0
*/
global using global::JsRt.Interop;
global using global::System;
global using global::System.Collections;
global using global::System.Collections.Generic;
global using global::System.ComponentModel;
global using global::System.Globalization;
global using global::System.Linq;
global using global::System.Reflection;
global using global::System.Runtime.CompilerServices;
global using global::System.Runtime.InteropServices;
global using global::System.Runtime.InteropServices.Marshalling;
global using global::System.Runtime.Versioning;
global using global::System.Text;
global using global::System.Threading;
global using global::System.Threading.Tasks;

#pragma warning disable IDE0079 // Remove unnecessary suppression
#pragma warning disable IDE0130 // Namespace does not match folder structure

namespace JsRt
{
	public class JsContext : IDisposable
	{
	    private readonly Lazy<JsValue> _undefined = new(() => { JsRuntime.JsGetUndefinedValue(out var handle); return new(handle); });
	    private readonly Lazy<JsValue> _null = new(() => { JsRuntime.JsGetNullValue(out var handle); return new(handle); });
	    private readonly Lazy<JsValue> _true = new(() => { JsRuntime.JsGetTrueValue(out var handle); return new(handle); });
	    private readonly Lazy<JsValue> _false = new(() => { JsRuntime.JsGetFalseValue(out var handle); return new(handle); });
	    private readonly Lazy<JsValue> _go = new(() => { JsRuntime.Check(JsRuntime.JsGetGlobalObject(out var handle)); return new JsValue(handle); });
	    private nint _handle;
	
	    public JsContext(nint handle, bool addRef)
	    {
	        if (handle == 0)
	            throw new ArgumentException(null, nameof(handle));
	
	        _handle = handle;
	        if (addRef)
	        {
	            JsRuntime.AddRef(handle);
	        }
	    }
	
	    public nint Handle
	    {
	        get
	        {
	            var handle = _handle;
	            ObjectDisposedException.ThrowIf(handle == 0, nameof(JsValue));
	            return handle;
	        }
	    }
	
	    public JsValue GlobalObject => _go.Value;
	    public JsValue Undefined => _undefined.Value;
	    public JsValue True => _true.Value;
	    public JsValue False => _false.Value;
	    public JsValue Null => _null.Value;
	    public JsRuntime? Runtime
	    {
	        get
	        {
	            JsRuntime.Check(JsRuntime.JsGetRuntime(Handle, out var handle));
	            if (handle == 0)
	                return null;
	
	            return new JsRuntime(handle);
	        }
	    }
	
	    public Version? EngineVersion
	    {
	        get
	        {
	            var go = GlobalObject;
	            var major = go.CallFunction<int>("ScriptEngineMajorVersion");
	            var minor = go.CallFunction<int>("ScriptEngineMinorVersion");
	            var build = go.CallFunction<int>("ScriptEngineBuildVersion");
	            return new Version(major, minor, Environment.OSVersion.Version.Build, build);
	        }
	    }
	
	    public override string ToString() => Handle.ToString();
	
	    public virtual void AddGlobalObject(string name, object? value)
	    {
	        ArgumentNullException.ThrowIfNull(name);
	        if (value != null && !Marshal.IsTypeVisibleFromCom(value.GetType()) && !Marshal.IsComObject(value))
	            throw new ArgumentException("Argument type must be ComVisible.", nameof(value));
	
	        GlobalObject.SetProperty(name, value);
	    }
	
	    public virtual void Execute(Action action)
	    {
	        ArgumentNullException.ThrowIfNull(action);
	        var prev = Current;
	        Current = this;
	        try
	        {
	            action();
	        }
	        finally
	        {
	            Current = prev;
	        }
	    }
	
	    public virtual T Execute<T>(Func<T> action)
	    {
	        ArgumentNullException.ThrowIfNull(action);
	        var prev = Current;
	        Current = this;
	        try
	        {
	            return action();
	        }
	        finally
	        {
	            Current = prev;
	        }
	    }
	
	    public virtual async Task<T> Execute<T>(Func<Task<T>> action)
	    {
	        ArgumentNullException.ThrowIfNull(action);
	        var prev = Current;
	        Current = this;
	        try
	        {
	            return await action();
	        }
	        finally
	        {
	            Current = prev;
	        }
	    }
	
	    ~JsContext() { Dispose(disposing: false); }
	    public void Dispose() { Dispose(disposing: true); GC.SuppressFinalize(this); }
	    protected virtual void Dispose(bool disposing)
	    {
	        if (_go.IsValueCreated)
	        {
	            _go.Value?.Dispose();
	        }
	
	        if (_null.IsValueCreated)
	        {
	            _null.Value?.Dispose();
	        }
	
	        if (_true.IsValueCreated)
	        {
	            _true.Value?.Dispose();
	        }
	
	        if (_false.IsValueCreated)
	        {
	            _false.Value?.Dispose();
	        }
	
	        if (_undefined.IsValueCreated)
	        {
	            _undefined.Value?.Dispose();
	        }
	
	        var handle = Interlocked.Exchange(ref _handle, 0);
	        if (handle != 0)
	        {
	            JsRuntime.Release(handle);
	        }
	    }
	
	    public static JsContext? Current
	    {
	        get
	        {
	            JsRuntime.Check(JsRuntime.JsGetCurrentContext(out var handle));
	            return handle != 0 ? new JsContext(handle, false) : null;
	        }
	        set => JsRuntime.JsSetCurrentContext(value != null ? value.Handle : 0);
	    }
	}
	
	public enum JsErrorCode
	{
	    JsNoError = 0,
	    JsErrorCategoryUsage = 0x10000,
	    JsErrorInvalidArgument,
	    JsErrorNullArgument,
	    JsErrorNoCurrentContext,
	    JsErrorInExceptionState,
	    JsErrorNotImplemented,
	    JsErrorWrongThread,
	    JsErrorRuntimeInUse,
	    JsErrorBadSerializedScript,
	    JsErrorInDisabledState,
	    JsErrorCannotDisableExecution,
	    JsErrorHeapEnumInProgress,
	    JsErrorArgumentNotObject,
	    JsErrorInProfileCallback,
	    JsErrorInThreadServiceCallback,
	    JsErrorCannotSerializeDebugScript,
	    JsErrorAlreadyDebuggingContext,
	    JsErrorAlreadyProfilingContext,
	    JsErrorIdleNotEnabled,
	    JsCannotSetProjectionEnqueueCallback,
	    JsErrorCannotStartProjection,
	    JsErrorInObjectBeforeCollectCallback,
	    JsErrorObjectNotInspectable,
	    JsErrorPropertyNotSymbol,
	    JsErrorPropertyNotString,
	    JsErrorCategoryEngine = 0x20000,
	    JsErrorOutOfMemory,
	    JsErrorCategoryScript = 0x30000,
	    JsErrorScriptException,
	    JsErrorScriptCompile,
	    JsErrorScriptTerminated,
	    JsErrorScriptEvalDisabled,
	    JsErrorCategoryFatal = 0x40000,
	    JsErrorFatal,
	    JsErrorWrongRuntime,
	}
	
	public partial class JsRuntime : IDisposable
	{
	    public const string JsDll = "jscript9.dll";
	
	    private readonly Lock _lock = new();
	    private nint _handle;
	
	    public JsRuntime(JsRuntimeAttributes attributes, JsRuntimeVersion version)
	    {
	        Check(JsCreateRuntime(attributes, version, null, out _handle));
	    }
	
	    public JsRuntime()
	        : this(JsRuntimeAttributes.JsRuntimeAttributeNone, JsRuntimeVersion.JsRuntimeVersionEdge)
	    {
	    }
	
	    public JsRuntime(JsRuntimeVersion version)
	        : this(JsRuntimeAttributes.JsRuntimeAttributeNone, version)
	    {
	    }
	
	    public JsRuntime(JsRuntimeAttributes attributes)
	        : this(attributes, JsRuntimeVersion.JsRuntimeVersionEdge)
	    {
	    }
	
	    public JsRuntime(nint handle)
	    {
	        if (handle == 0)
	            throw new ArgumentException(null, nameof(handle));
	
	        _handle = handle;
	    }
	
	    public virtual bool CacheParsedScripts { get; set; }
	    public IDictionary<string, JsValue> ParsedScriptCache { get; } = new Dictionary<string, JsValue>();
	    public nint Handle
	    {
	        get
	        {
	            var handle = _handle;
	            ObjectDisposedException.ThrowIf(handle == 0, nameof(JsValue));
	            return handle;
	        }
	    }
	
	    public long MemoryUsage
	    {
	        get
	        {
	            CheckDisposed();
	            Check(JsGetRuntimeMemoryUsage(Handle, out var value));
	            return value.ToInt64();
	        }
	    }
	
	    public long MemoryLimit
	    {
	        get
	        {
	            CheckDisposed();
	            Check(JsGetRuntimeMemoryLimit(Handle, out var value));
	            return value.ToInt64();
	        }
	        set
	        {
	            CheckDisposed();
	            Check(JsSetRuntimeMemoryLimit(Handle, new nint(value)));
	        }
	    }
	
	    public virtual bool ExecutionEnabled
	    {
	        get
	        {
	            CheckDisposed();
	            Check(JsIsRuntimeExecutionDisabled(Handle, out var disabled));
	            return !disabled;
	        }
	        set
	        {
	            if (value == ExecutionEnabled)
	                return;
	
	            if (value)
	            {
	                Check(JsEnableRuntimeExecution(Handle));
	            }
	            else
	            {
	                Check(JsDisableRuntimeExecution(Handle));
	            }
	        }
	    }
	
	    public override string ToString() => Handle.ToString();
	    public virtual void CollectGarbage()
	    {
	        CheckDisposed();
	        Check(JsCollectGarbage(Handle));
	    }
	
	    private void CheckDisposed() => ObjectDisposedException.ThrowIf(_handle == 0, "Engine has been disposed.");
	
	    public virtual JsContext CreateContext()
	    {
	        CheckDisposed();
	        Check(JsCreateContext(Handle, 0, out var handle));
	        return new JsContext(handle, true);
	    }
	
	    public virtual object? RunScript(string script, string? sourceUrl = null)
	    {
	        if (!TryRunScript(script, sourceUrl, out var error, out var value))
	        {
	            if (error != null)
	                throw error;
	
	            return null;
	        }
	
	        if (value == null)
	            return null;
	
	        try
	        {
	            return value.Value;
	        }
	        finally
	        {
	            value.Dispose();
	        }
	    }
	
	    public bool TryRunScript(string script, out Exception? error, out JsValue? value) => TryRunScript(script, null, out error, out value);
	    public virtual bool TryRunScript(string script, string? sourceUrl, out Exception? error, out JsValue? value)
	    {
	        ArgumentNullException.ThrowIfNull(script);
	        sourceUrl ??= string.Empty;
	
	        lock (_lock)
	        {
	            if (CacheParsedScripts && JsContext.Current != null)
	            {
	                // cache per context
	                var key = JsContext.Current.Handle + "." + script;
	                if (!ParsedScriptCache.TryGetValue(key, out var ps))
	                {
	                    error = Check(JsParseScript(script, 0, sourceUrl, out var psHandle), false);
	                    if (error != null)
	                    {
	                        // errored scripts are not cached
	                        value = null;
	                        return false;
	                    }
	
	                    ps = new JsValue(psHandle);
	                    ParsedScriptCache.Add(key, ps);
	                }
	                return ps.TryCall(out error, out value);
	            }
	        }
	
	        error = Check(JsRunScript(script, 0, sourceUrl, out var result), false);
	        if (error != null)
	        {
	            value = null;
	            return false;
	        }
	
	        value = new JsValue(result);
	        return true;
	    }
	
	#pragma warning disable IDE0079 // Remove unnecessary suppression
	#pragma warning disable CA1822 // Mark members as static
	    public virtual JsValue? ParseScript(string script, string? sourceUrl = null)
	    {
	        if (!TryParseScript(script, sourceUrl, out var error, out var value))
	        {
	            if (error != null)
	                throw error;
	        }
	
	        return value;
	    }
	
	    public bool TryParseScript(string script, out Exception? error, out JsValue? parsedScript) => TryParseScript(script, null, out error, out parsedScript);
	    public virtual bool TryParseScript(string script, string? sourceUrl, out Exception? error, out JsValue? parsedScript)
	    {
	        ArgumentNullException.ThrowIfNull(script);
	        sourceUrl ??= string.Empty;
	
	        error = Check(JsParseScript(script, 0, sourceUrl, out var result), false);
	        if (error != null)
	        {
	            parsedScript = null;
	            return false;
	        }
	
	        parsedScript = new JsValue(result);
	        return true;
	    }
	#pragma warning restore CA1822 // Mark members as static
	#pragma warning restore IDE0079 // Remove unnecessary suppression
	
	    protected internal delegate void JsBackgroundWorkItemCallback(nint callbackData);
	
	#pragma warning disable IDE0079 // Remove unnecessary suppression
	#pragma warning disable CA1401 // P/Invokes should not be visible
	    [return: MarshalAs(UnmanagedType.U1)]
	    protected internal delegate bool JsThreadServiceCallback(nint callback, nint callbackData);
	
	    [LibraryImport(JsDll)]
	    protected internal static partial JsErrorCode JsCreateRuntime(JsRuntimeAttributes attributes, JsRuntimeVersion runtimeVersion, JsThreadServiceCallback? threadService, out nint runtime);
	
	    [LibraryImport(JsDll)]
	    protected internal static partial JsErrorCode JsDisposeRuntime(nint runtime);
	
	    [LibraryImport(JsDll)]
	    protected internal static partial JsErrorCode JsIdle(out uint nextIdleTick);
	
	    [LibraryImport(JsDll)]
	    protected internal static partial JsErrorCode JsCollectGarbage(nint runtime);
	
	    [LibraryImport(JsDll)]
	    protected internal static partial JsErrorCode JsGetRuntimeMemoryUsage(nint runtime, out nint memoryUsage);
	
	    [LibraryImport(JsDll)]
	    protected internal static partial JsErrorCode JsGetRuntimeMemoryLimit(nint runtime, out nint memoryLimit);
	
	    [LibraryImport(JsDll)]
	    protected internal static partial JsErrorCode JsSetRuntimeMemoryLimit(nint runtime, nint memoryLimit);
	
	    [LibraryImport(JsDll)]
	    protected internal static partial JsErrorCode JsGetRuntime(nint context, out nint runtime);
	
	    [LibraryImport(JsDll)]
	    protected internal static partial JsErrorCode JsDisableRuntimeExecution(nint runtime);
	
	    [LibraryImport(JsDll)]
	    protected internal static partial JsErrorCode JsEnableRuntimeExecution(nint runtime);
	
	    [LibraryImport(JsDll)]
	    protected internal static partial JsErrorCode JsIsRuntimeExecutionDisabled(nint runtime, [MarshalAs(UnmanagedType.U1)] out bool isDisabled);
	
	    [LibraryImport(JsDll)]
	    protected internal static partial JsErrorCode JsCreateContext(nint runtime, nint debugApplication, out nint newContext);
	
	    [LibraryImport(JsDll)]
	    protected internal static partial JsErrorCode JsSetCurrentContext(nint context);
	
	    [LibraryImport(JsDll)]
	    protected internal static partial JsErrorCode JsGetCurrentContext(out nint context);
	
	    [LibraryImport(JsDll, StringMarshalling = StringMarshalling.Utf16)]
	    protected internal static partial JsErrorCode JsParseScript(string script, nint sourceContext, string sourceUrl, out nint result);
	
	    [LibraryImport(JsDll, StringMarshalling = StringMarshalling.Utf16)]
	    protected internal static partial JsErrorCode JsRunScript(string script, nint sourceContext, string sourceUrl, out nint result);
	
	    [LibraryImport(JsDll, StringMarshalling = StringMarshalling.Utf16)]
	    protected internal static partial JsErrorCode JsGetPropertyIdFromName(string name, out nint propertyId);
	
	    [LibraryImport(JsDll)]
	    protected internal static partial JsErrorCode JsGetProperty(nint @object, nint propertyId, out nint value);
	
	    [LibraryImport(JsDll)]
	    protected internal static partial JsErrorCode JsSetProperty(nint @object, nint propertyId, nint value, [MarshalAs(UnmanagedType.Bool)] bool useStrictRules);
	
	    [LibraryImport(JsDll)]
	    protected internal static partial JsErrorCode JsVariantToValue(in VARIANT variant, out nint value);
	
	    [LibraryImport(JsDll)]
	    protected internal static partial JsErrorCode JsValueToVariant(nint value, out VARIANT variant);
	
	    [LibraryImport(JsDll)]
	    protected internal static partial JsErrorCode JsGetValueType(nint value, out JsValueType type);
	
	    [LibraryImport(JsDll)]
	    protected internal static partial JsErrorCode JsGetGlobalObject(out nint globalObject);
	
	    [LibraryImport(JsDll)]
	    protected internal static partial JsErrorCode JsGetAndClearException(out nint exception);
	
	    [LibraryImport(JsDll)]
	    protected internal static partial JsErrorCode JsGetOwnPropertyNames(nint @object, out nint propertyNames);
	
	    [LibraryImport(JsDll)]
	    protected internal static partial JsErrorCode JsGetOwnPropertyDescriptor(nint @object, nint propertyId, out nint propertyDescriptor);
	
	    [LibraryImport(JsDll)]
	    protected internal static partial JsErrorCode JsCallFunction(nint function, [In][MarshalUsing(CountElementName = nameof(argumentCount))] nint[] arguments, ushort argumentCount, out nint result);
	
	    [LibraryImport(JsDll)]
	    protected internal static partial JsErrorCode JsGetIndexedProperty(nint @object, nint index, out nint result);
	
	    [LibraryImport(JsDll)]
	    protected internal static partial JsErrorCode JsSetIndexedProperty(nint @object, nint index, nint value);
	
	    [LibraryImport(JsDll)]
	    protected internal static partial JsErrorCode JsGetPrototype(nint @object, out nint prototypeObject);
	
	    [LibraryImport(JsDll)]
	    protected internal static partial JsErrorCode JsAddRef(nint handle, out uint count);
	
	    [LibraryImport(JsDll)]
	    protected internal static partial JsErrorCode JsGetUndefinedValue(out nint handle);
	
	    [LibraryImport(JsDll)]
	    protected internal static partial JsErrorCode JsGetNullValue(out nint handle);
	
	    [LibraryImport(JsDll)]
	    protected internal static partial JsErrorCode JsGetFalseValue(out nint handle);
	
	    [LibraryImport(JsDll)]
	    protected internal static partial JsErrorCode JsGetTrueValue(out nint handle);
	
	    [LibraryImport(JsDll)]
	    protected internal static partial JsErrorCode JsConvertValueToString(nint value, out nint handle);
	
	    [LibraryImport(JsDll)]
	    protected internal static partial JsErrorCode JsRelease(nint handle, out uint count);
	#pragma warning restore CA1401 // P/Invokes should not be visible
	#pragma warning restore IDE0079 // Remove unnecessary suppression
	
	    ~JsRuntime() { Dispose(disposing: false); }
	    public void Dispose() { Dispose(disposing: true); GC.SuppressFinalize(this); }
	    protected virtual void Dispose(bool disposing)
	    {
	        var handle = Interlocked.Exchange(ref _handle, 0);
	        if (handle != 0)
	        {
	            JsDisposeRuntime(handle);
	        }
	    }
	
	    public static object? Eval(string script)
	    {
	        using var rt = new JsRuntime();
	        return rt.CreateContext().Execute(() => rt.RunScript(script));
	    }
	
	    public static uint Idle()
	    {
	        Check(JsIdle(out uint ticks));
	        return ticks;
	    }
	
	    internal protected static uint AddRef(nint handle, bool throwOnError = true) { Check(JsAddRef(handle, out var count), throwOnError); return count; }
	    internal protected static uint Release(nint handle, bool throwOnError = true) { Check(JsRelease(handle, out var count), throwOnError); return count; }
	    internal static Exception? Check(JsErrorCode code, bool throwOnError = true)
	    {
	        Exception? error = null;
	        if (code != JsErrorCode.JsNoError)
	        {
	            JsGetAndClearException(out var ex);
	            using var value = ex != 0 ? new JsValue(ex) : null;
	            error = new JsRuntimeException(code, value);
	        }
	
	        if (throwOnError && error != null)
	            throw error;
	
	        return error;
	    }
	}
	
	[Flags]
	public enum JsRuntimeAttributes
	{
	    JsRuntimeAttributeNone = 0x00000000,
	    JsRuntimeAttributeDisableBackgroundWork = 0x00000001,
	    JsRuntimeAttributeAllowScriptInterrupt = 0x00000002,
	    JsRuntimeAttributeEnableIdleProcessing = 0x00000004,
	    JsRuntimeAttributeDisableNativeCodeGeneration = 0x00000008,
	    JsRuntimeAttributeDisableEval = 0x00000010,
	    JsRuntimeAttributeEnableExperimentalFeatures = 0x00000020,
	    JsRuntimeAttributeDispatchSetExceptionsToDebugger = 0x00000040,
	    JsRuntimeAttributeDisableFatalOnOOM = 0x00000080,
	    JsRuntimeAttributeDisableExecutablePageAllocation = 0x00000100,
	}
	
	[Serializable]
	public partial class JsRuntimeException : Exception
	{
	    internal JsRuntimeException(JsErrorCode code, JsValue? error)
	        : base(GetMessage(code, error))
	    {
	        Code = code;
	        if (error != null)
	        {
	            Line = error.GetProperty<int>("line");
	            Column = error.GetProperty<int>("column");
	            SourceCode = error.GetProperty<string?>("source") ?? string.Empty;
	        }
	    }
	
	    public JsRuntimeException()
	        : base("JsRuntime error")
	    {
	    }
	
	    public JsRuntimeException(string message)
	        : base(message)
	    {
	    }
	
	    public JsRuntimeException(string message, Exception innerException)
	        : base(message, innerException)
	    {
	    }
	
	    public JsRuntimeException(Exception innerException)
	        : base(null, innerException)
	    {
	    }
	
	    public JsErrorCode Code { get; }
	    public int Line { get; } = -1;
	    public int Column { get; } = -1;
	    public string SourceCode { get; } = string.Empty;
	
	    public static string GetErrorText(JsErrorCode error) => error switch
	    {
	        JsErrorCode.JsErrorInvalidArgument => "An argument to a hosting API was invalid.",
	        JsErrorCode.JsErrorNullArgument => "An argument to a hosting API was null in a context where null is not allowed.",
	        JsErrorCode.JsErrorNoCurrentContext => "The hosting API requires that a context be current, but there is no current context.",
	        JsErrorCode.JsErrorInExceptionState => "The engine is in an exception state and no APIs can be called until the exception is cleared.",
	        JsErrorCode.JsErrorNotImplemented => "A hosting API is not yet implemented.",
	        JsErrorCode.JsErrorWrongThread => "A hosting API was called on the wrong thread.",
	        JsErrorCode.JsErrorRuntimeInUse => "A runtime that is still in use cannot be disposed.",
	        JsErrorCode.JsErrorBadSerializedScript => "A bad serialized script was used, or the serialized script was serialized by a different version of the Chakra engine.",
	        JsErrorCode.JsErrorInDisabledState => "The runtime is in a disabled state.",
	        JsErrorCode.JsErrorCannotDisableExecution => "Runtime does not support reliable script interruption.",
	        JsErrorCode.JsErrorHeapEnumInProgress => "A heap enumeration is currently underway in the script context.",
	        JsErrorCode.JsErrorArgumentNotObject => "A hosting API that operates on object values was called with a non-object value.",
	        JsErrorCode.JsErrorInProfileCallback => "A script context is in the middle of a profile callback.",
	        JsErrorCode.JsErrorInThreadServiceCallback => "A thread service callback is currently underway.",
	        JsErrorCode.JsErrorCannotSerializeDebugScript => "Scripts cannot be serialized in debug contexts.",
	        JsErrorCode.JsErrorAlreadyDebuggingContext => "The context cannot be put into a debug state because it is already in a debug state.",
	        JsErrorCode.JsErrorAlreadyProfilingContext => "The context cannot start profiling because it is already profiling.",
	        JsErrorCode.JsErrorIdleNotEnabled => "Idle notification given when the host did not enable idle processing.",
	        JsErrorCode.JsErrorOutOfMemory => "The Chakra engine has run out of memory.",
	        JsErrorCode.JsErrorScriptException => "A JavaScript exception occurred while running a script.",
	        JsErrorCode.JsErrorScriptCompile => "JavaScript failed to compile.",
	        JsErrorCode.JsErrorScriptTerminated => "A script was terminated due to a request to suspend a runtime.",
	        JsErrorCode.JsErrorScriptEvalDisabled => "A script was terminated because it tried to use 'eval' or 'function' and eval was disabled.",
	        JsErrorCode.JsErrorFatal => "A fatal error in the engine has occurred.",
	        JsErrorCode.JsNoError => "Success error code.",
	        JsErrorCode.JsErrorCategoryUsage => "Category of errors that relates to incorrect usage of the API itself.",
	        JsErrorCode.JsCannotSetProjectionEnqueueCallback => "The context did not accept the enqueue callback.",
	        JsErrorCode.JsErrorCannotStartProjection => "Failed to start projection.",
	        JsErrorCode.JsErrorInObjectBeforeCollectCallback => "The operation is not supported in an object before collect callback.",
	        JsErrorCode.JsErrorObjectNotInspectable => "Object cannot be unwrapped to IInspectable pointer.",
	        JsErrorCode.JsErrorPropertyNotSymbol => "A hosting API that operates on symbol property ids but was called with a non-symbol property id. The error code is returned by JsGetSymbolFromPropertyId if the function is called with non-symbol property id.",
	        JsErrorCode.JsErrorPropertyNotString => "A hosting API that operates on string property ids but was called with a non-string property id. The error code is returned by existing JsGetPropertyNamefromId if the function is called with non-string property id.",
	        JsErrorCode.JsErrorCategoryEngine => "Category of errors that relates to errors occurring within the engine itself.",
	        JsErrorCode.JsErrorCategoryScript => "Category of errors that relates to errors in a script.",
	        JsErrorCode.JsErrorCategoryFatal => "Category of errors that are fatal and signify failure of the engine",
	        JsErrorCode.JsErrorWrongRuntime => "A hosting API was called with object created on different javascript runtime.",
	        _ => "An unknown error in the engine has occurred.",
	    };
	
	    private static string GetMessage(JsErrorCode code, JsValue? error)
	    {
	        var text = new StringBuilder(GetErrorText(code));
	        if (error != null)
	        {
	            var line = error.GetProperty<int>("line");
	            var column = error.GetProperty<int>("column");
	            var source = error.GetProperty<string?>("source");
	            var errorText = error.GetProperty<string?>("message");
	
	            if (!string.IsNullOrEmpty(errorText))
	            {
	                text.Append(" " + errorText);
	            }
	
	            if (line >= 0)
	            {
	                text.Append(" at line " + line);
	            }
	
	            if (column >= 0)
	            {
	                text.Append(", column " + column);
	            }
	
	            if (!string.IsNullOrEmpty(source))
	            {
	                text.Append(", in text \"" + source + "\"");
	            }
	        }
	        return text.ToString();
	    }
	}
	
	public enum JsRuntimeVersion
	{
	    JsRuntimeVersion10 = 0,
	    JsRuntimeVersion11 = 1,
	    JsRuntimeVersionEdge = -1,
	}
	
	public class JsValue : IDisposable
	{
	    private nint _handle;
	
	    public JsValue(nint handle)
	    {
	        if (handle == 0)
	            throw new ArgumentException(null, nameof(handle));
	
	        _handle = handle;
	        JsRuntime.Check(JsRuntime.JsGetValueType(Handle, out var vt));
	        ValueType = vt;
	        JsRuntime.AddRef(handle);
	    }
	
	    public nint Handle
	    {
	        get
	        {
	            var handle = _handle;
	            ObjectDisposedException.ThrowIf(handle == 0, nameof(JsValue));
	            return handle;
	        }
	    }
	
	    public override string ToString()
	    {
	        if (ValueType == JsValueType.JsNull || ValueType == JsValueType.JsUndefined)
	            return ValueType.ToString();
	
	        return ValueType + ": " + string.Format("{0}", Value);
	    }
	
	    public JsValueType ValueType { get; protected set; }
	
	    public virtual unsafe object? Value
	    {
	        get
	        {
	            JsRuntime.Check(JsRuntime.JsValueToVariant(Handle, out var v), false);
	            using var variant = Variant.Attach(ref v);
	            var value = variant.Value;
	            return value;
	        }
	    }
	
	    public string? ConvertToString()
	    {
	        JsRuntime.Check(JsRuntime.JsConvertValueToString(Handle, out var handle), false);
	        if (handle == 0)
	            return null;
	
	        using var strValue = new JsValue(handle);
	        return strValue.Value?.ToString();
	    }
	
	    public virtual object? DetachValue()
	    {
	        var value = Value;
	        Dispose();
	        return value;
	    }
	
	    public JsValue? Prototype
	    {
	        get
	        {
	            JsRuntime.Check(JsRuntime.JsGetPrototype(Handle, out var handle), false);
	            if (handle == 0)
	                return null;
	
	            return new JsValue(handle);
	        }
	    }
	
	    public virtual IDictionary<string, JsValue?> PropertyValues
	    {
	        get
	        {
	            var names = PropertyNames;
	            if (names == null)
	                return new Dictionary<string, JsValue?>();
	
	            var props = new Dictionary<string, JsValue?>();
	            foreach (string name in names)
	            {
	                props.Add(name, GetProperty<JsValue?>(name));
	            }
	            return props;
	        }
	    }
	
	    public virtual IList<JsValue> PropertyDescriptors
	    {
	        get
	        {
	            var names = PropertyNames;
	            if (names == null)
	                return [];
	
	            var values = new List<JsValue>();
	            foreach (string name in names)
	            {
	                JsRuntime.JsGetPropertyIdFromName(name, out var id);
	                if (id != 0)
	                {
	                    JsRuntime.AddRef(id);
	                    try
	                    {
	                        JsRuntime.JsGetOwnPropertyDescriptor(Handle, id, out var descriptor);
	                        if (descriptor != 0)
	                        {
	                            values.Add(new JsValue(descriptor));
	                        }
	                    }
	                    finally
	                    {
	                        JsRuntime.Release(id);
	                    }
	                }
	            }
	            return values;
	        }
	    }
	
	    public virtual IList<string> PropertyNames
	    {
	        get
	        {
	            JsRuntime.Check(JsRuntime.JsGetOwnPropertyNames(Handle, out var names), false);
	            if (names == 0)
	                return [];
	
	            var props = new List<string>();
	            using var v = new JsValue(names);
	            var i = 0;
	            do
	            {
	                var name = v.GetProperty<string?>(i, null);
	                if (name == null)
	                    break;
	
	                props.Add(name);
	                i++;
	            }
	            while (true);
	            return props;
	        }
	    }
	
	    public bool SetProperty(string name, object? value) => TrySetProperty(name, value, false, out _);
	
	    public virtual bool TrySetProperty(string name, object? value, bool useStrictRules, out Exception? error)
	    {
	        ArgumentNullException.ThrowIfNull(name);
	
	        nint id = 0;
	        nint valueHandle = 0;
	        error = null;
	        try
	        {
	            error = JsRuntime.Check(JsRuntime.JsGetPropertyIdFromName(name, out id), false);
	            if (error != null)
	                return false;
	
	            if (id == 0)
	                throw new JsRuntimeException("JsGetPropertyIdFromName returned an incorrect value.");
	
	            JsRuntime.AddRef(id);
	            error = VariantToValue(value, false, out valueHandle);
	            if (error != null)
	                return false;
	
	            if (valueHandle == 0)
	                throw new JsRuntimeException("VariantToValue returned an incorrect value.");
	
	            JsRuntime.AddRef(valueHandle);
	            error = JsRuntime.Check(JsRuntime.JsSetProperty(Handle, id, valueHandle, useStrictRules), false);
	            return error == null;
	        }
	        finally
	        {
	            JsRuntime.Release(id);
	            JsRuntime.Release(valueHandle);
	        }
	    }
	
	    public bool SetProperty(int index, object value) => TrySetProperty(index, value, out _);
	    public virtual bool TrySetProperty(int index, object? value, out Exception? error)
	    {
	        error = VariantToValue(value, false, out var valueHandle);
	        if (error != null)
	            return false;
	
	        using var jsValue = FromObject(index);
	        error = JsRuntime.Check(JsRuntime.JsSetIndexedProperty(Handle, jsValue.Handle, valueHandle));
	        return error == null;
	    }
	
	    public virtual object? GetProperty(string name, object? defaultValue = default) => GetProperty<object?>(name, defaultValue);
	    public virtual T? GetProperty<T>(string name, T? defaultValue = default)
	    {
	        if (!TryGetProperty(name, out _, out var jsValue) || jsValue == null)
	            return defaultValue;
	
	        if (typeof(T) == typeof(JsValue))
	            return (T?)(object?)jsValue;
	
	        try
	        {
	            if (TryChangeType<T>(jsValue.Value, out var value))
	                return value;
	
	            return defaultValue;
	        }
	        finally
	        {
	            jsValue?.Dispose();
	        }
	    }
	
	    public virtual bool TryGetProperty(string name, out Exception? error, out JsValue? value)
	    {
	        ArgumentNullException.ThrowIfNull(name);
	        nint id = 0;
	        try
	        {
	            error = JsRuntime.Check(JsRuntime.JsGetPropertyIdFromName(name, out id), false);
	            if (error != null)
	            {
	                value = null;
	                return false;
	            }
	            JsRuntime.AddRef(id);
	
	            error = JsRuntime.Check(JsRuntime.JsGetProperty(Handle, id, out var valueHandle), false);
	            if (error != null)
	            {
	                value = null;
	                return false;
	            }
	
	            value = new JsValue(valueHandle);
	            return true;
	        }
	        finally
	        {
	            JsRuntime.Release(id);
	        }
	    }
	
	    public virtual T? GetProperty<T>(int index, T? defaultValue = default)
	    {
	        if (!TryGetProperty(index, out _, out var jsValue) || jsValue == null)
	            return defaultValue;
	
	        if (typeof(T) == typeof(JsValue))
	            return (T)(object)jsValue!; // don't dispose this one since we want the JsValue itself
	
	        try
	        {
	            if (TryChangeType<T>(jsValue.Value, out var value))
	                return value;
	
	            return defaultValue;
	        }
	        finally
	        {
	            jsValue?.Dispose();
	        }
	    }
	
	    public virtual bool TryGetProperty(int index, out Exception? error, out JsValue? value)
	    {
	        using var iv = FromObject(index);
	        error = JsRuntime.Check(JsRuntime.JsGetIndexedProperty(Handle, iv.Handle, out nint valueHandle), false);
	        if (error != null)
	        {
	            value = null;
	            return false;
	        }
	
	        value = new JsValue(valueHandle);
	        return true;
	    }
	
	    public virtual object? CallFunction(string name, params object[] arguments) => CallFunction<object?>(name, arguments);
	    public virtual T? CallFunction<T>(string name, params object[] arguments)
	    {
	        ArgumentNullException.ThrowIfNull(name);
	        if (!TryCallFunction(name, out T? value, arguments))
	            return default;
	
	        return value;
	    }
	
	    public virtual bool TryCallFunction<T>(string name, out T? value, params object[] arguments)
	    {
	        ArgumentNullException.ThrowIfNull(name);
	        using var fn = GetProperty<JsValue>(name);
	        if (fn == null)
	        {
	            value = default;
	            return false;
	        }
	
	        JsValue? jsValue = null;
	        try
	        {
	            if (!fn.TryCall(out var error, out jsValue, arguments) || jsValue == null)
	            {
	                value = default;
	                return false;
	            }
	
	            return TryChangeType(jsValue.Value, out value);
	        }
	        finally
	        {
	            jsValue?.Dispose();
	        }
	    }
	
	    public virtual object? Call(params object?[] arguments)
	    {
	        ArgumentNullException.ThrowIfNull(arguments);
	        var args = Convert(arguments);
	        try
	        {
	            return Call(args);
	        }
	        finally
	        {
	            Dispose(args);
	        }
	    }
	
	    public virtual object? Call(JsValue[] arguments)
	    {
	        ArgumentNullException.ThrowIfNull(arguments);
	        if (!TryCall(out var error, out var jsValue, arguments) || jsValue == null)
	        {
	            if (error != null)
	                throw error;
	
	            return null;
	        }
	
	        try
	        {
	            return jsValue.Value;
	        }
	        finally
	        {
	            jsValue.Dispose();
	        }
	    }
	
	    public virtual bool TryCall(out Exception? error, out JsValue? value, params object?[] arguments)
	    {
	        var args = Convert(arguments);
	        try
	        {
	            return TryCall(out error, out value, args);
	        }
	        finally
	        {
	            Dispose(args);
	        }
	    }
	
	    public virtual bool TryCall(out Exception? error, out JsValue? value, JsValue[] arguments, bool throwOnError = true)
	    {
	        ArgumentNullException.ThrowIfNull(arguments);
	        var args = new List<nint>();
	        foreach (var arg in arguments)
	        {
	            args.Add(arg.Handle);
	        }
	
	        error = JsRuntime.Check(JsRuntime.JsCallFunction(Handle, [.. args], (ushort)args.Count, out var handle), throwOnError);
	        if (error != null)
	        {
	            value = null;
	            return false;
	        }
	
	        value = new JsValue(handle);
	        return true;
	    }
	
	    ~JsValue() { Dispose(disposing: false); }
	    public void Dispose() { Dispose(disposing: true); GC.SuppressFinalize(this); }
	    protected virtual void Dispose(bool disposing)
	    {
	        var h = Interlocked.Exchange(ref _handle, 0);
	        if (h != 0)
	        {
	            JsRuntime.Release(h);
	        }
	    }
	
	    public static JsValue FromObject(object? value, bool throwOnError = true)
	    {
	        if (value is JsValue jsv)
	            return jsv;
	
	        var error = VariantToValue(value, true, out var handle);
	        if (error != null)
	        {
	            if (throwOnError)
	                throw new JsRuntimeException(error);
	
	            if (JsContext.Current == null)
	                throw new JsRuntimeException("No current context available to convert the value.");
	
	            return JsContext.Current.Null;
	        }
	
	        return new JsValue(handle);
	    }
	
	    public static bool IsUndefined(object? obj) => obj is JsValue jsv && jsv.ValueType == JsValueType.JsUndefined;
	
	    public static bool TryChangeType<T>(object? input, out T? value)
	    {
	        if (typeof(T) == typeof(object))
	        {
	            value = (T?)input;
	            return true;
	        }
	
	        if (input == null)
	        {
	            value = default;
	            return true;
	        }
	
	        if (input is T t)
	        {
	            value = t;
	            return true;
	        }
	
	        if (input is JsValue jsv)
	            return TryChangeType(jsv.Value, out value);
	
	        try
	        {
	            value = (T?)System.Convert.ChangeType(input, typeof(T), CultureInfo.InvariantCulture);
	            return true;
	
	        }
	        catch (Exception)
	        {
	            value = default;
	            return false;
	        }
	    }
	
	    private static Exception? VariantToValue(object? value, bool throwOnError, out nint handle)
	    {
	        using var v = new Variant(value);
	        var error = JsRuntime.Check(JsRuntime.JsVariantToValue(v.Detached, out handle), throwOnError);
	        return error;
	    }
	
	    private static void Dispose(IEnumerable<JsValue> arguments)
	    {
	        foreach (var arg in arguments)
	        {
	            arg.Dispose();
	        }
	    }
	
	    private static JsValue[] Convert(object?[] arguments)
	    {
	        var values = new JsValue[arguments.Length];
	        for (var i = 0; i < arguments.Length; i++)
	        {
	            values[i] = FromObject(arguments[i]);
	        }
	        return values;
	    }
	}
	
	public enum JsValueType
	{
	    JsUndefined = 0,
	    JsNull = 1,
	    JsNumber = 2,
	    JsString = 3,
	    JsBoolean = 4,
	    JsObject = 5,
	    JsFunction = 6,
	    JsError = 7,
	    JsArray = 8,
	    JsSymbol = 9,
	    JsArrayBuffer = 10,
	    JsTypedArray = 11,
	    JsDataView = 12,
	}
	
	public sealed class Variant : IDisposable
	{
	    private VARIANT _inner;
	
	    public VARIANT Detached => _inner;
	    public ref VARIANT RefDetached => ref _inner;
	
	    public static int Size { get; } = GetSizeOf32();
	    private static int GetSizeOf32() { unsafe { return sizeof(VARIANT); } }
	
	    internal Variant(VARIANT inner)
	    {
	        _inner = inner;
	    }
	
	    public Variant()
	    {
	        // it's a VT_EMPTY
	    }
	
	    public Variant(object? value, VARENUM? type = null)
	    {
	        if (value == null)
	        {
	            _inner.Anonymous.Anonymous.vt = VARENUM.VT_NULL;
	            return;
	        }
	
	        value = Unwrap(value);
	
	        if (value is nint ptr)
	        {
	            _inner.Anonymous.Anonymous.Anonymous.punkVal = ptr;
	            _inner.Anonymous.Anonymous.vt = type ?? VARENUM.VT_UNKNOWN;
	            return;
	        }
	
	        if (value is nuint uptr)
	        {
	            _inner.Anonymous.Anonymous.Anonymous.punkVal = (nint)uptr;
	            _inner.Anonymous.Anonymous.vt = type ?? VARENUM.VT_UNKNOWN;
	            return;
	        }
	
	        if (value is ComObject co)
	        {
	            var sw = new StrategyBasedComWrappers();
	            _inner.Anonymous.Anonymous.Anonymous.punkVal = sw.GetOrCreateComInterfaceForObject(co, CreateComInterfaceFlags.None);
	            _inner.Anonymous.Anonymous.vt = VARENUM.VT_UNKNOWN;
	            return;
	        }
	
	        if (value is char[] chars)
	        {
	            value = new string(chars);
	        }
	
	        if (value is char[][] charray)
	        {
	            var strings = new string[charray.GetLength(0)];
	            for (var i = 0; i < charray.Length; i++)
	            {
	                strings[i] = new string(charray[i]);
	            }
	            value = strings;
	        }
	
	        if (value is Array array)
	        {
	            ConstructArray(array, type);
	            return;
	        }
	
	        if (value is not string && value is IEnumerable enumerable)
	        {
	            ConstructEnumerable(enumerable, type);
	            return;
	        }
	
	        if (value == null)
	        {
	            _inner.Anonymous.Anonymous.vt = VARENUM.VT_NULL;
	            return;
	        }
	
	        var valueType = value.GetType();
	        var vt = FromType(valueType, type, true);
	        var tc = Type.GetTypeCode(valueType);
	        switch (tc)
	        {
	            case TypeCode.Boolean:
	                _inner.Anonymous.Anonymous.Anonymous.boolVal = new VARIANT_BOOL { Value = (bool)value ? (short)(-1) : (short)0 };
	                vt = VARENUM.VT_BOOL;
	                break;
	
	            case TypeCode.Byte:
	                _inner.Anonymous.Anonymous.Anonymous.bVal = (byte)value;
	                vt = VARENUM.VT_UI1;
	                break;
	
	            case TypeCode.Char:
	                chars = [(char)value];
	                // note: all strings (PWSTR, PSTR, BSTR) point to same place
	                _inner.Anonymous.Anonymous.Anonymous.bstrVal = new BSTR { Value = MarshalString(new string(chars), VARENUM.VT_BSTR) };
	                vt = VARENUM.VT_BSTR;
	                break;
	
	            case TypeCode.DateTime:
	                if (type == VARENUM.VT_FILETIME)
	                {
	                    var ft = ToPositiveFILETIME((DateTime)value);
	                    Functions.InitVariantFromFileTime(ft, out _inner);
	                    return;
	                }
	
	                var dt = (DateTime)value;
	                _inner.Anonymous.Anonymous.Anonymous.dblVal = dt.ToOADate();
	                vt = VARENUM.VT_DATE;
	                break;
	
	            case TypeCode.Empty:
	            case TypeCode.DBNull:
	                break;
	
	            case TypeCode.Decimal:
	                _inner.Anonymous.decVal = (decimal)value;
	                vt = VARENUM.VT_DECIMAL;
	                break;
	
	            case TypeCode.Double:
	                _inner.Anonymous.Anonymous.Anonymous.dblVal = (double)value;
	                vt = VARENUM.VT_R8;
	                break;
	
	            case TypeCode.Int16:
	                _inner.Anonymous.Anonymous.Anonymous.iVal = (short)value;
	                vt = VARENUM.VT_I2;
	                break;
	
	            case TypeCode.Int32:
	                _inner.Anonymous.Anonymous.Anonymous.lVal = (int)value;
	                vt = VARENUM.VT_I4;
	                break;
	
	            case TypeCode.Int64:
	                _inner.Anonymous.Anonymous.Anonymous.llVal = (long)value;
	                vt = VARENUM.VT_I8;
	                break;
	
	            case TypeCode.SByte:
	                _inner.Anonymous.Anonymous.Anonymous.cVal.Value = (sbyte)value;
	                vt = VARENUM.VT_I1;
	                break;
	
	            case TypeCode.Single:
	                _inner.Anonymous.Anonymous.Anonymous.fltVal = (float)value;
	                vt = VARENUM.VT_R4;
	                break;
	
	            case TypeCode.String:
	                // note: all strings (PWSTR, PSTR, BSTR) point to same place
	                _inner.Anonymous.Anonymous.Anonymous.bstrVal = new BSTR { Value = MarshalString((string)value, VARENUM.VT_BSTR) };
	                vt = VARENUM.VT_BSTR;
	                break;
	
	            case TypeCode.UInt16:
	                _inner.Anonymous.Anonymous.Anonymous.uiVal = (ushort)value;
	                vt = VARENUM.VT_UI2;
	                break;
	
	            case TypeCode.UInt32:
	                _inner.Anonymous.Anonymous.Anonymous.ulVal = (uint)value;
	                vt = VARENUM.VT_UI4;
	                break;
	
	            case TypeCode.UInt64:
	                _inner.Anonymous.Anonymous.Anonymous.ullVal = (ulong)value;
	                vt = VARENUM.VT_UI8;
	                break;
	
	            //case TypeCode.Object:
	            default:
	                if (value is Guid guid)
	                {
	                    _inner.Anonymous.Anonymous.Anonymous.bstrVal = new BSTR { Value = MarshalString(guid.ToString("B"), VARENUM.VT_BSTR) };
	                    vt = VARENUM.VT_BSTR;
	                    break;
	                }
	
	                if (value is DateTimeOffset dto)
	                {
	                    if (type == VARENUM.VT_FILETIME)
	                    {
	                        var ft = ToPositiveFILETIME(dto.DateTime);
	                        Functions.InitVariantFromFileTime(ft, out _inner);
	                        return;
	                    }
	
	                    _inner.Anonymous.Anonymous.Anonymous.dblVal = dto.DateTime.ToOADate();
	                    vt = VARENUM.VT_DATE;
	                    break;
	                }
	
	                throw new ArgumentException("Value of type '" + value.GetType().FullName + "' is not supported.", nameof(value));
	        }
	
	        _inner.Anonymous.Anonymous.vt = vt;
	    }
	
	    public VARENUM VarType { get => _inner.Anonymous.Anonymous.vt; }
	    public object? Value
	    {
	        get
	        {
	            switch (_inner.Anonymous.Anonymous.vt)
	            {
	                case VARENUM.VT_EMPTY:
	                case VARENUM.VT_NULL: // DbNull
	                    return null;
	
	                case VARENUM.VT_I1:
	                    return _inner.Anonymous.Anonymous.Anonymous.cVal.Value;
	
	                case VARENUM.VT_UI1:
	                    return _inner.Anonymous.Anonymous.Anonymous.bVal;
	
	                case VARENUM.VT_I2:
	                    return _inner.Anonymous.Anonymous.Anonymous.iVal;
	
	                case VARENUM.VT_UI2:
	                    return _inner.Anonymous.Anonymous.Anonymous.uiVal;
	
	                case VARENUM.VT_I4:
	                case VARENUM.VT_INT:
	                    return _inner.Anonymous.Anonymous.Anonymous.lVal;
	
	                case VARENUM.VT_UI4:
	                case VARENUM.VT_UINT:
	                    return _inner.Anonymous.Anonymous.Anonymous.ulVal;
	
	                case VARENUM.VT_I8:
	                    return _inner.Anonymous.Anonymous.Anonymous.llVal;
	
	                case VARENUM.VT_UI8:
	                    return _inner.Anonymous.Anonymous.Anonymous.ullVal;
	
	                case VARENUM.VT_R4:
	                    return _inner.Anonymous.Anonymous.Anonymous.fltVal;
	
	                case VARENUM.VT_R8:
	                    return _inner.Anonymous.Anonymous.Anonymous.dblVal;
	
	                case VARENUM.VT_BOOL:
	                    return _inner.Anonymous.Anonymous.Anonymous.boolVal.Value != 0;
	
	                case VARENUM.VT_ERROR:
	                    return _inner.Anonymous.Anonymous.Anonymous.scode;
	
	                case VARENUM.VT_CY:
	                    return _inner.Anonymous.decVal;
	
	                case VARENUM.VT_DATE:
	                    return DateTime.FromOADate(_inner.Anonymous.Anonymous.Anonymous.dblVal);
	
	                case VARENUM.VT_BSTR:
	                    return Marshal.PtrToStringBSTR(_inner.Anonymous.Anonymous.Anonymous.bstrVal.Value);
	
	                case VARENUM.VT_LPSTR:
	                    // all strings point to same place anyway
	                    return Marshal.PtrToStringAnsi(_inner.Anonymous.Anonymous.Anonymous.bstrVal.Value);
	
	                case VARENUM.VT_LPWSTR:
	                    // all strings point to same place anyway
	                    return Marshal.PtrToStringUni(_inner.Anonymous.Anonymous.Anonymous.bstrVal.Value);
	
	                case VARENUM.VT_UNKNOWN:
	                case VARENUM.VT_DISPATCH:
	                    var sw = new StrategyBasedComWrappers();
	                    return sw.GetOrCreateObjectForComInstance(_inner.Anonymous.Anonymous.Anonymous.punkVal, CreateObjectFlags.UniqueInstance);
	
	                case VARENUM.VT_DECIMAL:
	                    return _inner.Anonymous.decVal;
	
	                default:
	                    if (_inner.Anonymous.Anonymous.vt.HasFlag(VARENUM.VT_ARRAY))
	                    {
	                        var et = _inner.Anonymous.Anonymous.vt & ~VARENUM.VT_ARRAY;
	                        if (TryGetArrayValue(et, out var array))
	                            return array;
	                    }
	
	                    throw new NotSupportedException("Value of property type " + _inner.Anonymous.Anonymous.vt + " is not supported.");
	            }
	        }
	    }
	
	    public Variant? ChangeType(VARENUM type, bool throwOnError = true)
	    {
	        var inner = new VARIANT();
	        var hr = Functions.VariantChangeType(ref inner, _inner, 0, type).ThrowOnError(throwOnError);
	        if (hr.IsError)
	            return null;
	
	        return new Variant { _inner = inner };
	    }
	
	    public void CopyFrom(Variant source, bool throwOnError = true)
	    {
	        ArgumentNullException.ThrowIfNull(source);
	        if (source == this)
	            return;
	
	        Clear(throwOnError);
	        var inner = new VARIANT();
	        Functions.VariantCopy(ref inner, source._inner).ThrowOnError(throwOnError);
	        _inner = inner;
	    }
	
	    public Variant? Copy(bool throwOnError = true)
	    {
	        var inner = new VARIANT();
	        var hr = Functions.VariantCopy(ref inner, _inner).ThrowOnError(throwOnError);
	        if (hr.IsError)
	            return null;
	
	        return new Variant { _inner = inner };
	    }
	
	    public VARIANT Detach()
	    {
	        var pv = _inner;
	        Zero();
	        return pv;
	    }
	
	    public unsafe void DetachTo(nint variantPtr)
	    {
	        if (variantPtr == 0)
	            throw new ArgumentException(null, nameof(variantPtr));
	
	        var pv = _inner;
	        Zero();
	        *(VARIANT*)(variantPtr) = pv;
	    }
	
	    public static nint MarshalString(string? text, VARENUM vt)
	    {
	        if (text == null)
	            return 0;
	
	        return vt switch
	        {
	            VARENUM.VT_LPWSTR => Marshal.StringToCoTaskMemUni(text),
	            VARENUM.VT_BSTR => Marshal.StringToBSTR(text),
	            VARENUM.VT_LPSTR => Marshal.StringToCoTaskMemAnsi(text),
	            _ => throw new NotSupportedException("A string can only be of property type VT_LPWSTR, VT_LPSTR or VT_BSTR."),
	        };
	    }
	
	    public static string? PtrTostring(nint ptr, VARENUM vt)
	    {
	        if (ptr == 0)
	            return null;
	
	        return vt switch
	        {
	            VARENUM.VT_LPWSTR => Marshal.PtrToStringUni(ptr),
	            VARENUM.VT_BSTR => Marshal.PtrToStringBSTR(ptr),
	            VARENUM.VT_LPSTR => Marshal.PtrToStringAnsi(ptr),
	            _ => throw new NotSupportedException("A string can only be of property type VT_LPWSTR, VT_LPSTR or VT_BSTR."),
	        };
	    }
	
	    private static Type FromType(VARENUM type) => type switch
	    {
	        VARENUM.VT_I1 => typeof(sbyte),
	        VARENUM.VT_UI1 => typeof(byte),
	        VARENUM.VT_I2 => typeof(short),
	        VARENUM.VT_UI2 => typeof(ushort),
	        VARENUM.VT_UI4 or VARENUM.VT_UINT => typeof(uint),
	        VARENUM.VT_I8 => typeof(long),
	        VARENUM.VT_UI8 => typeof(ulong),
	        VARENUM.VT_R4 => typeof(float),
	        VARENUM.VT_R8 => typeof(double),
	        VARENUM.VT_BOOL => typeof(bool),
	        VARENUM.VT_I4 or VARENUM.VT_INT or VARENUM.VT_ERROR => typeof(int),
	        VARENUM.VT_DATE => typeof(DateTime),
	        VARENUM.VT_FILETIME => typeof(ulong),
	        VARENUM.VT_BLOB => typeof(byte[]),
	        VARENUM.VT_CLSID => typeof(Guid),
	        VARENUM.VT_BSTR or VARENUM.VT_LPSTR or VARENUM.VT_LPWSTR => typeof(string),
	        VARENUM.VT_UNKNOWN or VARENUM.VT_DISPATCH => typeof(object),
	        VARENUM.VT_CY or VARENUM.VT_DECIMAL => typeof(decimal),
	        _ => throw new ArgumentException("Property type " + type + " is not supported.", nameof(type)),
	    };
	
	    private static VARENUM FromType(Type type, VARENUM? vt, bool forVariant)
	    {
	        if (type == null)
	            return VARENUM.VT_NULL;
	
	        var tc = Type.GetTypeCode(type);
	        switch (tc)
	        {
	            case TypeCode.Boolean:
	                return VARENUM.VT_BOOL;
	
	            case TypeCode.Byte:
	                return VARENUM.VT_UI1;
	
	            case TypeCode.Char:
	                if (forVariant)
	                    return VARENUM.VT_BSTR;
	
	                return VARENUM.VT_LPWSTR;
	
	            case TypeCode.DateTime:
	                if (vt == VARENUM.VT_FILETIME)
	                    return VARENUM.VT_FILETIME;
	
	                return VARENUM.VT_DATE;
	
	            case TypeCode.DBNull:
	                return VARENUM.VT_NULL;
	
	            case TypeCode.Decimal:
	                return VARENUM.VT_DECIMAL;
	
	            case TypeCode.Double:
	                return VARENUM.VT_R8;
	
	            case TypeCode.Empty:
	                return VARENUM.VT_EMPTY;
	
	            case TypeCode.Int16:
	                return VARENUM.VT_I2;
	
	            case TypeCode.Int32:
	                return VARENUM.VT_I4;
	
	            case TypeCode.Int64:
	                return VARENUM.VT_I8;
	
	            case TypeCode.SByte:
	                return VARENUM.VT_I1;
	
	            case TypeCode.Single:
	                return VARENUM.VT_R4;
	
	            case TypeCode.String:
	                if (forVariant)
	                    return VARENUM.VT_BSTR;
	
	                if (!vt.HasValue)
	                    return VARENUM.VT_LPWSTR;
	
	                if (vt != VARENUM.VT_LPSTR && vt != VARENUM.VT_BSTR && vt != VARENUM.VT_LPWSTR)
	                    throw new ArgumentException("Property type " + vt + " is not supported for string.", nameof(type));
	
	                return vt.Value;
	
	            case TypeCode.UInt16:
	                return VARENUM.VT_UI2;
	
	            case TypeCode.UInt32:
	                return VARENUM.VT_UI4;
	
	            case TypeCode.UInt64:
	                return VARENUM.VT_UI8;
	
	            // case TypeCode.Object:
	            default:
	                if (type == typeof(Guid))
	                {
	                    if (forVariant)
	                        return VARENUM.VT_BSTR;
	
	                    return VARENUM.VT_CLSID;
	                }
	
	                if (type == typeof(FILETIME))
	                {
	                    if (forVariant)
	                        return VARENUM.VT_DATE;
	
	                    return VARENUM.VT_FILETIME;
	                }
	
	                if (type == typeof(byte))
	                {
	                    if (forVariant)
	                        return VARENUM.VT_UI1 | VARENUM.VT_ARRAY;
	
	                    if (!vt.HasValue)
	                        return VARENUM.VT_UI1 | VARENUM.VT_VECTOR;
	
	                    if (vt != VARENUM.VT_BLOB && vt != (VARENUM.VT_UI1 | VARENUM.VT_VECTOR))
	                        throw new ArgumentException("Property type " + vt + " is not supported for array of bytes.", nameof(type));
	
	                    return vt.Value;
	                }
	
	                if (type == typeof(object))
	                    return VARENUM.VT_VARIANT;
	
	                throw new ArgumentException("Value of type '" + type.FullName + "' is not supported.", nameof(type));
	        }
	    }
	
	    public static Variant Attach(ref VARIANT detached, bool zeroDetached = true)
	    {
	        var pv = new Variant { _inner = detached };
	        if (zeroDetached)
	        {
	            unsafe
	            {
	                var ptr = Unsafe.AsPointer(ref detached);
	                Functions.ZeroMemory((nint)ptr, Size);
	            }
	        }
	        return pv;
	    }
	
	    public static object? Unwrap(object? value)
	    {
	        if (value is Variant variant)
	            return Unwrap(variant.Value);
	
	        if (value is VARIANT v)
	        {
	            var v2 = Attach(ref v, false);
	            value = v2.Value;
	            v2.Detach();
	            return Unwrap(value);
	        }
	
	        return value;
	    }
	
	    public override string ToString()
	    {
	        var value = Value;
	        if (value == null)
	            return "<null>";
	
	        if (value is string svalue)
	            return "[" + VarType + "] `" + svalue + "`";
	
	        if (value is not byte[] && value is IEnumerable enumerable)
	            return "[" + VarType + "] " + string.Join(", ", enumerable.OfType<object>());
	
	        if (value is byte[] bytes)
	            return "[" + VarType + "] bytes[" + bytes.Length + "]";
	
	        return "[" + VarType + "] " + value;
	    }
	
	    ~Variant() => Dispose();
	    public void Dispose() { Clear(false); GC.SuppressFinalize(this); }
	
	    private void Zero()
	    {
	        unsafe
	        {
	            fixed (VARIANT* p = &_inner)
	            {
	                Functions.ZeroMemory((nint)p, Size);
	            }
	        }
	    }
	
	    private void ConstructEnumerable(IEnumerable enumerable, VARENUM? type = null)
	    {
	        var et = GetElementType(enumerable) ?? throw new ArgumentException("Enumerable type '" + enumerable.GetType().FullName + "' is not supported.", nameof(enumerable));
	        var count = GetCount(enumerable);
	#pragma warning disable IDE0079 // Remove unnecessary suppression
	#pragma warning disable IL3050 // Calling members annotated with 'RequiresDynamicCodeAttribute' may break functionality when AOT compiling.
	        var array = Array.CreateInstance(et, count);
	#pragma warning restore IL3050 // Calling members annotated with 'RequiresDynamicCodeAttribute' may break functionality when AOT compiling.
	#pragma warning restore IDE0079 // Remove unnecessary suppression
	        var i = 0;
	        foreach (var obj in enumerable)
	        {
	            array.SetValue(obj, i++);
	        }
	        ConstructArray(array, type);
	    }
	
	    private static int GetCount(IEnumerable enumerable)
	    {
	        if (enumerable is ICollection col)
	            return col.Count;
	
	        var count = 0;
	        var e = enumerable.GetEnumerator();
	        Using(e, () =>
	        {
	            while (e.MoveNext())
	            {
	                count++;
	            }
	        });
	        return count;
	
	        static void Using(object resource, Action action)
	        {
	            try
	            {
	                action();
	            }
	            finally
	            {
	                (resource as IDisposable)?.Dispose();
	            }
	        }
	    }
	
	    private static Type? GetElementType(IEnumerable enumerable)
	    {
	        var et = GetElementType(enumerable.GetType());
	        if (et != null)
	            return et;
	
	        foreach (var obj in enumerable)
	        {
	            return obj.GetType();
	        }
	        return null;
	    }
	
	    private static Type? GetElementType(Type collectionType)
	    {
	#pragma warning disable IDE0079 // Remove unnecessary suppression
	#pragma warning disable IL2070 // 'this' argument does not satisfy 'DynamicallyAccessedMembersAttribute' in call to target method. The parameter of method does not have matching annotations.
	        foreach (var iface in collectionType.GetInterfaces())
	        {
	            if (!iface.IsGenericType)
	                continue;
	
	            if (iface.GetGenericTypeDefinition() == typeof(IEnumerable<>))
	                return iface.GetGenericArguments()[0];
	
	            if (iface.GetGenericTypeDefinition() == typeof(ICollection<>))
	                return iface.GetGenericArguments()[0];
	
	            if (iface.GetGenericTypeDefinition() == typeof(IList<>))
	                return iface.GetGenericArguments()[0];
	        }
	#pragma warning restore IL2070 // 'this' argument does not satisfy 'DynamicallyAccessedMembersAttribute' in call to target method. The parameter of method does not have matching annotations.
	#pragma warning restore IDE0079 // Remove unnecessary suppression
	        return null;
	    }
	
	    private void ConstructArray(Array array, VARENUM? type = null)
	    {
	        // special case for bools which are shorts...
	        if (array is bool[] bools)
	        {
	            var shorts = new short[bools.Length];
	            for (var i = 0; i < bools.Length; i++)
	            {
	                shorts[i] = bools[i] ? ((short)(-1)) : ((short)0);
	            }
	            ConstructSafeArray(shorts, typeof(short), VARENUM.VT_BOOL);
	            return;
	        }
	
	        if (array is Guid[] guids)
	        {
	            var strings = new string[guids.Length];
	            for (var i = 0; i < strings.Length; i++)
	            {
	                strings[i] = guids[i].ToString("B");
	            }
	            ConstructSafeArray(strings, typeof(string), VARENUM.VT_BSTR);
	            return;
	        }
	
	        var et = array.GetType().GetElementType() ?? throw new NotSupportedException();
	        ConstructSafeArray(array, et, FromType(et, type, true));
	    }
	
	    private void ConstructSafeArray(Array array, Type type, VARENUM vt)
	    {
	        unsafe
	        {
	            var bounds = new SAFEARRAYBOUND { lLbound = 0, cElements = (uint)array.Length };
	            var sa = Functions.SafeArrayCreate(vt, 1, bounds);
	            if (sa == 0)
	                throw new OutOfMemoryException();
	
	            var psa = (SAFEARRAY*)sa;
	            Functions.SafeArrayAccessData(*psa, out var ptr).ThrowOnError();
	            try
	            {
	                if (type == typeof(string))
	                {
	                    for (var i = 0; i < array.Length; i++)
	                    {
	                        var str = MarshalString((string?)array.GetValue(i)!, vt);
	                        Marshal.WriteIntPtr(ptr, nint.Size * i, str);
	                    }
	                }
	                else if (type == typeof(object))
	                {
	                    for (var i = 0; i < array.Length; i++)
	                    {
	                        var variantValue = array.GetValue(i);
	                        using var variant = new Variant(variantValue);
	                        unsafe
	                        {
	                            var p = (VARIANT*)(ptr + Size * i);
	                            *p = variant.Detach();
	                        }
	                    }
	                }
	                else
	                {
	#pragma warning disable IDE0079 // Remove unnecessary suppression
	#pragma warning disable IL3050 // Calling members annotated with 'RequiresDynamicCodeAttribute' may break functionality when AOT compiling.
	#pragma warning disable CA1421 // This method uses runtime marshalling even when the 'DisableRuntimeMarshallingAttribute' is applied
	                    var size = Marshal.SizeOf(type) * array.Length;
	#pragma warning restore CA1421 // This method uses runtime marshalling even when the 'DisableRuntimeMarshallingAttribute' is applied
	#pragma warning restore IL3050 // Calling members annotated with 'RequiresDynamicCodeAttribute' may break functionality when AOT compiling.
	#pragma warning restore IDE0079 // Remove unnecessary suppression
	
	                    Functions.CopyMemory(ptr, Marshal.UnsafeAddrOfPinnedArrayElement(array, 0), size);
	                }
	            }
	            catch
	            {
	                Functions.SafeArrayDestroy(*psa);
	                throw;
	            }
	            finally
	            {
	                Functions.SafeArrayUnaccessData(*psa).ThrowOnError();
	            }
	
	            _inner.Anonymous.Anonymous.vt = vt | VARENUM.VT_ARRAY;
	            _inner.Anonymous.Anonymous.Anonymous.parray = sa;
	        }
	    }
	
	    private bool TryGetArrayValue(VARENUM vt, out object? value)
	    {
	        value = null;
	        if (_inner.Anonymous.Anonymous.Anonymous.parray == 0)
	            return false;
	
	        unsafe
	        {
	            var psa = (SAFEARRAY*)_inner.Anonymous.Anonymous.Anonymous.parray;
	            if (psa->cDims != 1)
	                return false;
	
	            Functions.SafeArrayGetLBound(*psa, 1, out var l).ThrowOnError();
	            Functions.SafeArrayGetUBound(*psa, 1, out var u).ThrowOnError();
	            var count = u - l + 1;
	
	            Functions.SafeArrayAccessData(*psa, out var ptr).ThrowOnError();
	            try
	            {
	                var ret = false;
	                uint size;
	                switch (vt)
	                {
	                    case VARENUM.VT_LPSTR:
	                    case VARENUM.VT_LPWSTR:
	                    case VARENUM.VT_BSTR:
	                        var strings = new string?[count];
	                        for (var i = 0; i < strings.Length; i++)
	                        {
	                            var str = Marshal.ReadIntPtr(ptr, (int)(psa->cbElements * i));
	                            strings[i] = PtrTostring(str, vt);
	                        }
	                        value = strings;
	                        ret = true;
	                        break;
	
	                    case VARENUM.VT_BOOL:
	                        var shorts = new short[count];
	                        size = (uint)(shorts.Length * sizeof(short));
	                        Functions.CopyMemory(Marshal.UnsafeAddrOfPinnedArrayElement(shorts, 0), ptr, (nint)size);
	                        var bools = new bool[shorts.Length];
	                        for (var i = 0; i < shorts.Length; i++)
	                        {
	                            bools[i] = shorts[i] != 0;
	                        }
	                        value = bools;
	                        ret = true;
	                        break;
	
	                    case VARENUM.VT_VARIANT:
	                        var variants = new object?[count];
	                        var variantSize = Size;
	                        for (var i = 0; i < variants.Length; i++)
	                        {
	                            var pv = ptr + Size * i;
	                            using var v = new Variant { _inner = *(VARIANT*)pv };
	                            variants[i] = v.Value;
	                            v.Detach();
	                        }
	                        value = variants;
	                        ret = true;
	                        break;
	
	                    case VARENUM.VT_I1:
	                    case VARENUM.VT_UI1:
	                    case VARENUM.VT_I2:
	                    case VARENUM.VT_UI2:
	                    case VARENUM.VT_I4:
	                    case VARENUM.VT_INT:
	                    case VARENUM.VT_UI4:
	                    case VARENUM.VT_UINT:
	                    case VARENUM.VT_I8:
	                    case VARENUM.VT_UI8:
	                    case VARENUM.VT_R4:
	                    case VARENUM.VT_R8:
	                    case VARENUM.VT_ERROR:
	                    case VARENUM.VT_CY:
	                    case VARENUM.VT_DATE:
	                    case VARENUM.VT_UNKNOWN:
	                    case VARENUM.VT_DISPATCH:
	                        var et = FromType(vt);
	#pragma warning disable IDE0079 // Remove unnecessary suppression
	#pragma warning disable IL3050 // Calling members annotated with 'RequiresDynamicCodeAttribute' may break functionality when AOT compiling.
	                        var values = Array.CreateInstance(et, psa->cbElements);
	#pragma warning disable CA1421 // This method uses runtime marshalling even when the 'DisableRuntimeMarshallingAttribute' is applied
	                        size = (uint)(values.Length * Marshal.SizeOf(et));
	#pragma warning restore CA1421 // This method uses runtime marshalling even when the 'DisableRuntimeMarshallingAttribute' is applied
	#pragma warning restore IL3050 // Calling members annotated with 'RequiresDynamicCodeAttribute' may break functionality when AOT compiling.
	#pragma warning restore IDE0079 // Remove unnecessary suppression
	                        Functions.CopyMemory(Marshal.UnsafeAddrOfPinnedArrayElement(values, 0), ptr, (nint)size);
	                        value = values;
	                        ret = true;
	                        break;
	                }
	                return ret;
	            }
	            finally
	            {
	                Functions.SafeArrayUnaccessData(*psa).ThrowOnError();
	            }
	        }
	    }
	
	    public void Clear(bool throwOnError = true) => Functions.VariantClear(ref _inner).ThrowOnError(throwOnError);
	
	    public static FILETIME ToPositiveFILETIME(DateTime dt) => ToFILETIME(ToPositiveFileTime(dt));
	    public static FILETIME ToFILETIME(long ft) => ToFILETIME((ulong)ft);
	    public static FILETIME ToFILETIME(ulong ft) => new() { dwLowDateTime = (uint)(ft & 0xFFFFFFFF), dwHighDateTime = (uint)(ft >> 32) };
	    public static long ToPositiveFileTime(DateTime dt)
	    {
	        var ft = dt.ToUniversalTime().ToFileTimeUtc();
	        return ft < 0 ? 0 : ft;
	    }
	}
}

namespace JsRt.Interop
{
	public partial struct BSTR // not disposable as we don't know here who allocated it
	{
	    public nint Value;
	
	    public static readonly BSTR Null = new();
	
	    public BSTR(nint value)
	    {
	        Value = value;
	    }
	
	    unsafe public BSTR(char* value)
	    {
	        Value = (nint)value;
	    }
	
	    public static void Dispose(ref BSTR bstr)
	    {
	        var value = Interlocked.Exchange(ref bstr.Value, 0);
	        if (value != 0)
	        {
	            Marshal.FreeBSTR(value);
	        }
	    }
	
	    public override readonly string? ToString() => Marshal.PtrToStringBSTR(Value);
	}
	
	public partial struct CHAR(sbyte value) : IEquatable<CHAR>
	{
	    public static readonly CHAR Null = new();
	
	    public sbyte Value = value;
	
	    public override readonly string ToString() => $"0x{Value:x}";
	
	    public override readonly bool Equals(object? obj) => obj is CHAR value && Equals(value);
	    public readonly bool Equals(CHAR other) => other.Value == Value;
	    public override readonly int GetHashCode() => Value.GetHashCode();
	    public static bool operator ==(CHAR left, CHAR right) => left.Equals(right);
	    public static bool operator !=(CHAR left, CHAR right) => !left.Equals(right);
	    public static implicit operator sbyte(CHAR value) => value.Value;
	    public static implicit operator CHAR(sbyte value) => new(value);
	}
	
	public partial struct FILETIME
	{
	    public uint dwLowDateTime;
	    public uint dwHighDateTime;
	}
	
	internal static partial class Functions
	{
	    [LibraryImport("OLEAUT32")]
	    [PreserveSig]
	    public static partial HRESULT VariantChangeType(ref VARIANT pvargDest, in VARIANT pvarSrc, VAR_CHANGE_FLAGS wFlags, VARENUM vt);
	
	    [LibraryImport("OLEAUT32")]
	    [PreserveSig]
	    public static partial HRESULT VariantCopy(ref VARIANT pvargDest, in VARIANT pvargSrc);
	
	    [LibraryImport("OLEAUT32")]
	    [PreserveSig]
	    public static partial HRESULT VariantClear(ref VARIANT pvarg);
	
	    [LibraryImport("PROPSYS")]
	    [PreserveSig]
	    public static partial HRESULT InitVariantFromFileTime(in FILETIME pft, out VARIANT pvar);
	
	    [LibraryImport("kernel32", EntryPoint = "RtlZeroMemory")]
	    [PreserveSig]
	    public static partial void ZeroMemory(nint pdst, nint cb);
	
	    [LibraryImport("kernel32", EntryPoint = "RtlMoveMemory")]
	    [PreserveSig]
	    public static partial void CopyMemory(nint pdst, nint psrc, nint cb);
	
	    [LibraryImport("OLEAUT32")]
	    [PreserveSig]
	    public static partial HRESULT SafeArrayAccessData(in SAFEARRAY psa, out nint ppvData);
	
	    [LibraryImport("OLEAUT32")]
	    [PreserveSig]
	    public static partial nint SafeArrayCreate(VARENUM vt, uint cDims, in SAFEARRAYBOUND rgsabound);
	
	    [LibraryImport("OLEAUT32")]
	    [PreserveSig]
	    public static partial HRESULT SafeArrayDestroy(in SAFEARRAY psa);
	
	    [LibraryImport("OLEAUT32")]
	    [PreserveSig]
	    public static partial HRESULT SafeArrayUnaccessData(in SAFEARRAY psa);
	
	    [LibraryImport("OLEAUT32")]
	    [PreserveSig]
	    public static partial HRESULT SafeArrayGetLBound(in SAFEARRAY psa, uint nDim, out int plLbound);
	
	    [LibraryImport("OLEAUT32")]
	    [PreserveSig]
	    public static partial HRESULT SafeArrayGetUBound(in SAFEARRAY psa, uint nDim, out int plUbound);
	}
	
	public partial struct HRESULT(int value) : IEquatable<HRESULT>, IFormattable
	{
	    public static readonly HRESULT Null = new();
	
	    public int Value = value;
	
	    public override readonly bool Equals(object? obj) => obj is HRESULT value && Equals(value);
	    public readonly bool Equals(HRESULT other) => other.Value == Value;
	    public override readonly int GetHashCode() => Value.GetHashCode();
	
	    public HRESULT(uint value)
	        : this((int)value)
	    {
	    }
	
	    public readonly uint UValue => (uint)Value;
	    public readonly string Name => ToString("n", null);
	    public readonly bool IsError => Value < 0;
	    public readonly bool IsSuccess => Value >= 0;
	    public readonly bool IsOk => Value == 0;
	    public readonly bool IsFalse => Value == 1;
	
	    public readonly HRESULT ThrowOnError(bool throwOnError = true)
	    {
	        if (!throwOnError)
	            return Value;
	
	        var exception = GetException();
	        if (exception != null)
	            throw exception;
	
	        return Value;
	    }
	
	    public readonly Exception? GetException()
	    {
	        if (Value < 0)
	            return new Win32Exception(Value);
	
	        return null;
	    }
	
	    public override readonly string ToString() => ToString(null, null);
	    public readonly string ToString(string? format, IFormatProvider? formatProvider) => (format?.ToLowerInvariant()) switch
	    {
	        "i" => Value.ToString(),
	        "u" => UValue.ToString(),
	        _ => "0x" + Value.ToString("X8"),
	    };
	
	    public static HRESULT FromWin32(uint error)
	    {
	        if (error >= 0x80000000)
	            return error;
	
	        return FromWin32((int)error);
	    }
	
	    public static HRESULT FromWin32(int error)
	    {
	        if (error < 0)
	            return error;
	
	        const int FACILITY_WIN32 = 7;
	        return (uint)(error & 0x0000FFFF) | (FACILITY_WIN32 << 16) | 0x80000000;
	    }
	
	    public static implicit operator HRESULT(uint result) => new(result);
	    public static explicit operator uint(HRESULT hr) => hr.UValue;
	    public static bool operator ==(HRESULT left, HRESULT right) => left.Equals(right);
	    public static bool operator !=(HRESULT left, HRESULT right) => !left.Equals(right);
	    public static implicit operator int(HRESULT value) => value.Value;
	    public static implicit operator HRESULT(int value) => new(value);
	}
	
	public partial struct PSTR // not disposable as we don't know here who allocated it
	{
	    public nint Value;
	
	    public static readonly PSTR Null = new();
	
	    public PSTR(nint value)
	    {
	        Value = value;
	    }
	
	    unsafe public PSTR(byte* value)
	    {
	        if (value != null)
	        {
	            Value = (nint)value;
	        }
	        else
	        {
	            Value = 0;
	        }
	    }
	
	    public unsafe static PSTR From(string? str, Encoding? encoding = null)
	    {
	        if (str == null)
	            return Null;
	
	        encoding ??= Encoding.Default;
	        var bytes = encoding.GetBytes(str);
	        fixed (byte* p = bytes)
	        {
	            return new PSTR(p);
	        }
	    }
	
	    public override readonly string? ToString() => Marshal.PtrToStringAnsi(Value);
	}
	
	public partial struct PWSTR // not disposable as we don't know here who allocated it
	{
	    public nint Value;
	
	    public static readonly PWSTR Null = new();
	
	    public PWSTR(nint value)
	    {
	        Value = value;
	    }
	
	    unsafe public PWSTR(char* value)
	    {
	        if (value != null)
	        {
	            Value = (nint)value;
	        }
	        else
	        {
	            Value = 0;
	        }
	    }
	
	    public unsafe static PWSTR From(string? str)
	    {
	        if (str == null)
	            return Null;
	
	        fixed (char* chars = str)
	        {
	            return new PWSTR(chars);
	        }
	    }
	
	    public override readonly string? ToString() => Marshal.PtrToStringUni(Value);
	}
	
	internal partial struct SAFEARRAY
	{
	    public ushort cDims;
	    public ushort fFeatures;
	    public uint cbElements;
	    public uint cLocks;
	    public nint pvData;
	    public SAFEARRAYBOUND rgsabound; // variable-length array placeholder
	}
	
	internal partial struct SAFEARRAYBOUND
	{
	    public uint cElements;
	    public int lLbound;
	}
	
	[Flags]
	public enum VARENUM : ushort
	{
	    VT_EMPTY = 0,
	    VT_NULL = 1,
	    VT_I2 = 2,
	    VT_I4 = 3,
	    VT_R4 = 4,
	    VT_R8 = 5,
	    VT_CY = 6,
	    VT_DATE = 7,
	    VT_BSTR = 8,
	    VT_DISPATCH = 9,
	    VT_ERROR = 10,
	    VT_BOOL = 11,
	    VT_VARIANT = 12,
	    VT_UNKNOWN = 13,
	    VT_DECIMAL = 14,
	    VT_I1 = 16,
	    VT_UI1 = 17,
	    VT_UI2 = 18,
	    VT_UI4 = 19,
	    VT_I8 = 20,
	    VT_UI8 = 21,
	    VT_INT = 22,
	    VT_UINT = 23,
	    VT_VOID = 24,
	    VT_HRESULT = 25,
	    VT_PTR = 26,
	    VT_SAFEARRAY = 27,
	    VT_CARRAY = 28,
	    VT_USERDEFINED = 29,
	    VT_LPSTR = 30,
	    VT_LPWSTR = 31,
	    VT_RECORD = 36,
	    VT_INT_PTR = 37,
	    VT_UINT_PTR = 38,
	    VT_FILETIME = 64,
	    VT_BLOB = 65,
	    VT_STREAM = 66,
	    VT_STORAGE = 67,
	    VT_STREAMED_OBJECT = 68,
	    VT_STORED_OBJECT = 69,
	    VT_BLOB_OBJECT = 70,
	    VT_CF = 71,
	    VT_CLSID = 72,
	    VT_VERSIONED_STREAM = 73,
	    VT_BSTR_BLOB = 4095,
	    VT_VECTOR = 4096,
	    VT_ARRAY = 8192,
	    VT_BYREF = 16384,
	    VT_RESERVED = 32768,
	    VT_ILLEGAL = ushort.MaxValue,
	    VT_ILLEGALMASKED = 4095,
	    VT_TYPEMASK = 4095,
	}
	
	public struct VARIANT
	{
	    [StructLayout(LayoutKind.Explicit)]
	    public struct AnonymousUnion
	    {
	        public struct AnonymousStruct
	        {
	            [StructLayout(LayoutKind.Explicit)]
	            public struct AnonymousUnion
	            {
	                public struct AnonymousStruct
	                {
	                    public nint pvRecord;
	                    public nint pRecInfo;
	                }
	
	                [FieldOffset(0)]
	                public long llVal;
	
	                [FieldOffset(0)]
	                public int lVal;
	
	                [FieldOffset(0)]
	                public byte bVal;
	
	                [FieldOffset(0)]
	                public short iVal;
	
	                [FieldOffset(0)]
	                public float fltVal;
	
	                [FieldOffset(0)]
	                public double dblVal;
	
	                [FieldOffset(0)]
	                public VARIANT_BOOL boolVal;
	
	                [FieldOffset(0)]
	                public int scode;
	
	                [FieldOffset(0)]
	                public long cyVal;
	
	                [FieldOffset(0)]
	                public double date;
	
	                [FieldOffset(0)]
	                public BSTR bstrVal;
	
	                [FieldOffset(0)]
	                public nint punkVal;
	
	                [FieldOffset(0)]
	                public nint pdispVal;
	
	                [FieldOffset(0)]
	                public nint parray;
	
	                [FieldOffset(0)]
	                public nint pbVal;
	
	                [FieldOffset(0)]
	                public nint piVal;
	
	                [FieldOffset(0)]
	                public nint plVal;
	
	                [FieldOffset(0)]
	                public nint pllVal;
	
	                [FieldOffset(0)]
	                public nint pfltVal;
	
	                [FieldOffset(0)]
	                public nint pdblVal;
	
	                [FieldOffset(0)]
	                public nint pboolVal;
	
	                [FieldOffset(0)]
	                public nint __OBSOLETE__VARIANT_PBOOL;
	
	                [FieldOffset(0)]
	                public nint pscode;
	
	                [FieldOffset(0)]
	                public nint pcyVal;
	
	                [FieldOffset(0)]
	                public nint pdate;
	
	                [FieldOffset(0)]
	                public nint pbstrVal;
	
	                [FieldOffset(0)]
	                public nint ppunkVal;
	
	                [FieldOffset(0)]
	                public nint ppdispVal;
	
	                [FieldOffset(0)]
	                public nint pparray;
	
	                [FieldOffset(0)]
	                public nint pvarVal;
	
	                [FieldOffset(0)]
	                public nint byref;
	
	                [FieldOffset(0)]
	                public CHAR cVal;
	
	                [FieldOffset(0)]
	                public ushort uiVal;
	
	                [FieldOffset(0)]
	                public uint ulVal;
	
	                [FieldOffset(0)]
	                public ulong ullVal;
	
	                [FieldOffset(0)]
	                public int intVal;
	
	                [FieldOffset(0)]
	                public uint uintVal;
	
	                [FieldOffset(0)]
	                public nint pdecVal;
	
	                [FieldOffset(0)]
	                public nint pcVal;
	
	                [FieldOffset(0)]
	                public nint puiVal;
	
	                [FieldOffset(0)]
	                public nint pulVal;
	
	                [FieldOffset(0)]
	                public nint pullVal;
	
	                [FieldOffset(0)]
	                public nint pintVal;
	
	                [FieldOffset(0)]
	                public nint puintVal;
	
	                [FieldOffset(0)]
	                public AnonymousStruct Anonymous;
	            }
	
	            public VARENUM vt;
	            public ushort wReserved1;
	            public ushort wReserved2;
	            public ushort wReserved3;
	            public AnonymousUnion Anonymous;
	        }
	
	        [FieldOffset(0)]
	        public AnonymousStruct Anonymous;
	
	        [FieldOffset(0)]
	        public decimal decVal;
	    }
	
	    public AnonymousUnion Anonymous;
	}
	
	public struct VARIANT_BOOL(short value) : IEquatable<VARIANT_BOOL>
	{
	    public static readonly VARIANT_BOOL Null = new();
	
	    public short Value = value;
	
	    public override readonly string ToString() => $"0x{Value:x}";
	
	    public override readonly bool Equals(object? obj) => obj is VARIANT_BOOL value && Equals(value);
	    public readonly bool Equals(VARIANT_BOOL other) => other.Value == Value;
	    public override readonly int GetHashCode() => Value.GetHashCode();
	    public static bool operator ==(VARIANT_BOOL left, VARIANT_BOOL right) => left.Equals(right);
	    public static bool operator !=(VARIANT_BOOL left, VARIANT_BOOL right) => !left.Equals(right);
	    public static implicit operator short(VARIANT_BOOL value) => value.Value;
	    public static implicit operator VARIANT_BOOL(short value) => new(value);
	}
	
	[Flags]
	internal enum VAR_CHANGE_FLAGS : ushort
	{
	    VARIANT_NOVALUEPROP = 1,
	    VARIANT_ALPHABOOL = 2,
	    VARIANT_NOUSEROVERRIDE = 4,
	    VARIANT_CALENDAR_HIJRI = 8,
	    VARIANT_LOCALBOOL = 16,
	    VARIANT_CALENDAR_THAI = 32,
	    VARIANT_CALENDAR_GREGORIAN = 64,
	    VARIANT_USE_NLS = 128,
	}
}

#pragma warning restore IDE0130 // Namespace does not match folder structure
#pragma warning restore IDE0079 // Remove unnecessary suppression
