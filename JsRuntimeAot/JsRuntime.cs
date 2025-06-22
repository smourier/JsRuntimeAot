namespace JsRt;

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
    public nint Handle => _handle;
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
    protected internal static partial JsErrorCode JsAddRef(nint handle, out int count);

    [LibraryImport(JsDll)]
    protected internal static partial JsErrorCode JsGetUndefinedValue(out nint handle);

    [LibraryImport(JsDll)]
    protected internal static partial JsErrorCode JsGetNullValue(out nint handle);

    [LibraryImport(JsDll)]
    protected internal static partial JsErrorCode JsGetFalseValue(out nint handle);

    [LibraryImport(JsDll)]
    protected internal static partial JsErrorCode JsGetTrueValue(out nint handle);

    [LibraryImport(JsDll)]
    protected internal static partial JsErrorCode JsRelease(nint handle, out int count);
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

    internal protected static Exception? AddRef(nint handle, bool throwOnError = true) => AddRef(handle, throwOnError, out _);
    internal protected static Exception? AddRef(nint handle, bool throwOnError, out int count)
    {
        if (handle == 0)
        {
            count = 0;
            return null;
        }

        return Check(JsAddRef(handle, out count), throwOnError);
    }

    internal protected static Exception? Release(nint handle, bool throwOnError = true) => Release(handle, throwOnError, out _);
    internal protected static Exception? Release(nint handle, bool throwOnError, out int count)
    {
        if (handle == 0)
        {
            count = 0;
            return null;
        }

        return Check(JsRelease(handle, out count), throwOnError);
    }

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