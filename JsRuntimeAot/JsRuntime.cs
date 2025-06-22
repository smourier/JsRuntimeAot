namespace JsRt;

public partial class JsRuntime : IDisposable, IEquatable<JsRuntime>
{
    public const string JsDll = "jscript9.dll";

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

    public bool Equals(JsRuntime? other) => other is not null && _handle == other._handle;
    public override bool Equals(object? obj) => obj is JsRuntime other && Equals(other);
    public override int GetHashCode() => _handle.GetHashCode();
    public static bool operator ==(JsRuntime? left, JsRuntime? right) => left?.Equals(right) ?? right is null;
    public static bool operator !=(JsRuntime? left, JsRuntime? right) => !(left == right);
    public override string ToString() => Handle.ToString();
    public virtual void CollectGarbage()
    {
        CheckDisposed();
        Check(JsCollectGarbage(Handle));
    }

    private void CheckDisposed() => ObjectDisposedException.ThrowIf(_handle == 0, nameof(JsRuntime));

    public virtual void WithContext(Action<JsContext> action)
    {
        using var ctx = CreateContext();
        JsContext.Current = ctx;
        try
        {
            action(ctx);
        }
        finally
        {
            JsContext.Current = null;
        }
    }

    public virtual T WithContext<T>(Func<JsContext, T> action)
    {
        using var ctx = CreateContext();
        JsContext.Current = ctx;
        try
        {
            return action(ctx);
        }
        finally
        {
            JsContext.Current = null;
        }
    }

    public virtual async Task<T> WithContext<T>(Func<JsContext, Task<T>> action)
    {
        using var ctx = CreateContext();
        JsContext.Current = ctx;
        try
        {
            return await action(ctx);
        }
        finally
        {
            JsContext.Current = null;
        }
    }

    public virtual JsContext CreateContext()
    {
        CheckDisposed();
        Check(JsCreateContext(Handle, 0, out var handle));
        return new JsContext(handle, true);
    }

    public virtual object? RunScript(string script, string? sourceUrl = null)
    {
        if (!TryRunScript(script, sourceUrl, out var error, out var jsValue) || jsValue == null)
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

    public bool TryRunScript(string script, out Exception? error, out JsValue? value) => TryRunScript(script, null, out error, out value);
    public virtual bool TryRunScript(string script, string? sourceUrl, out Exception? error, out JsValue? value)
    {
        ArgumentNullException.ThrowIfNull(script);
        sourceUrl ??= string.Empty;
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
        return rt.WithContext(ctx => rt.RunScript(script));
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