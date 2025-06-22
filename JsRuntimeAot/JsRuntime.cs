namespace JsRt;

public partial class JsRuntime : IDisposable
{
    public const string JsDll = "jscript9.dll";

    private JsValue? _go;
    private nint _handle;

    private static object? _syncObject;
    internal static object SyncObject
    {
        get
        {
            if (_syncObject == null)
            {
                var obj = new object();
                Interlocked.CompareExchange(ref _syncObject, obj, null);
            }
            return _syncObject;
        }
    }

    public JsRuntime(JsRuntimeAttributes attributes, JsRuntimeVersion version) => Check(JsCreateRuntime(attributes, version, null, out _handle));
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

    internal JsRuntime(nint handle)
    {
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

    public JsValue? GlobalObject
    {
        get
        {
            lock (SyncObject)
            {
                if (_go == null)
                {
                    JsGetGlobalObject(out nint go);
                    if (go == 0)
                        return null;

                    _go = new JsValue(go);
                }
                return _go;
            }
        }
    }

    public Version? EngineVersion
    {
        get
        {
            var go = GlobalObject;
            if (go == null)
                return null;

            var major = go.CallFunction("ScriptEngineMajorVersion", 0);
            var minor = go.CallFunction("ScriptEngineMinorVersion", 0);
            var build = go.CallFunction("ScriptEngineBuildVersion", 0);
            return new Version(major, minor, Environment.OSVersion.Version.Build, build);
        }
    }

    public virtual bool ExecutionEnabled
    {
        get
        {
            CheckDisposed();
            Check(JsIsRuntimeExecutionDisabled(Handle, out bool disabled));
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

    internal static Exception? Check(JsErrorCode code, bool throwOnError = true)
    {
        Exception? error = null;
        if (code != JsErrorCode.JsNoError)
        {
            JsGetAndClearException(out nint ex);
            using var value = ex != 0 ? new JsValue(ex) : null;
            error = new JsRuntimeException(code, value);
        }

        if (throwOnError && error != null)
            throw error;

        return error;
    }

    public void CollectGarbage()
    {
        CheckDisposed();
        Check(JsCollectGarbage(Handle));
    }

    public static uint Idle()
    {
        Check(JsIdle(out uint ticks));
        return ticks;
    }

    private void CheckDisposed() => ObjectDisposedException.ThrowIf(_handle == 0, "Engine has been disposed.");

    public JsContext CreateContext()
    {
        CheckDisposed();
        Check(JsCreateContext(Handle, 0, out nint handle));
        return new JsContext(handle, true);
    }

    public void AddGlobalObject(string name, object? value)
    {
        ArgumentNullException.ThrowIfNull(name);
        if (value != null && !Marshal.IsTypeVisibleFromCom(value.GetType()) && !Marshal.IsComObject(value))
            throw new ArgumentException("Argument type must be ComVisible.", nameof(value));

        var go = GlobalObject;
        if (go == null)
            return;

        go.SetProperty(name, value);
    }

    public static object? Eval(string script)
    {
        using var rt = new JsRuntime();
        return rt.CreateContext().Execute(() => rt.RunScript(script));
    }

    public object? RunScript(string script, string? sourceUrl = null)
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
    public bool TryRunScript(string script, string? sourceUrl, out Exception? error, out JsValue? value)
    {
        ArgumentNullException.ThrowIfNull(script);
        sourceUrl ??= string.Empty;

        lock (SyncObject)
        {
            if (CacheParsedScripts && JsContext.Current != null)
            {
                // cache per context
                string key = JsContext.Current.Handle + "\0" + script;
                if (!ParsedScriptCache.TryGetValue(key, out var ps))
                {
                    error = Check(JsParseScript(script, 0, sourceUrl, out nint psHandle), false);
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

        error = Check(JsRunScript(script, 0, sourceUrl, out nint result), false);
        if (error != null)
        {
            value = null;
            return false;
        }

        value = new JsValue(result);
        return true;
    }

    public static JsValue? ParseScript(string script, string? sourceUrl = null)
    {
        if (!TryParseScript(script, sourceUrl, out var error, out var value))
        {
            if (error != null)
                throw error;
        }

        return value;
    }

    public static bool TryParseScript(string script, out Exception? error, out JsValue? parsedScript) => TryParseScript(script, null, out error, out parsedScript);
    public static bool TryParseScript(string script, string? sourceUrl, out Exception? error, out JsValue? parsedScript)
    {
        ArgumentNullException.ThrowIfNull(script);
        sourceUrl ??= string.Empty;

        error = Check(JsParseScript(script, 0, sourceUrl, out nint result), false);
        if (error != null)
        {
            parsedScript = null;
            return false;
        }

        parsedScript = new JsValue(result);
        return true;
    }

    internal static Exception? AddRef(nint handle, bool throwOnError = true) => AddRef(handle, throwOnError, out _);
    internal static Exception? AddRef(nint handle, bool throwOnError, out int count)
    {
        if (handle == 0)
        {
            count = 0;
            return null;
        }

        return Check(JsAddRef(handle, out count), throwOnError);
    }

    internal static Exception? Release(nint handle, bool throwOnError = true) => Release(handle, throwOnError, out _);
    internal static Exception? Release(nint handle, bool throwOnError, out int count)
    {
        if (handle == 0)
        {
            count = 0;
            return null;
        }

        return Check(JsRelease(handle, out count), throwOnError);
    }

    private delegate void JsBackgroundWorkItemCallback(nint callbackData);
    private delegate bool JsThreadServiceCallback(JsBackgroundWorkItemCallback callback, nint callbackData);

    [LibraryImport(JsDll)]
    private static partial JsErrorCode JsCreateRuntime(JsRuntimeAttributes attributes, JsRuntimeVersion runtimeVersion, JsThreadServiceCallback? threadService, out nint runtime);

    [LibraryImport(JsDll)]
    private static partial JsErrorCode JsDisposeRuntime(nint runtime);

    [LibraryImport(JsDll)]
    private static partial JsErrorCode JsIdle(out uint nextIdleTick);

    [LibraryImport(JsDll)]
    private static partial JsErrorCode JsCollectGarbage(nint runtime);

    [LibraryImport(JsDll)]
    private static partial JsErrorCode JsGetRuntimeMemoryUsage(nint runtime, out nint memoryUsage);

    [LibraryImport(JsDll)]
    private static partial JsErrorCode JsGetRuntimeMemoryLimit(nint runtime, out nint memoryLimit);

    [LibraryImport(JsDll)]
    private static partial JsErrorCode JsSetRuntimeMemoryLimit(nint runtime, nint memoryLimit);

    [LibraryImport(JsDll)]
    internal static partial JsErrorCode JsGetRuntime(nint context, out nint runtime);

    [LibraryImport(JsDll)]
    private static partial JsErrorCode JsDisableRuntimeExecution(nint runtime);

    [LibraryImport(JsDll)]
    private static partial JsErrorCode JsEnableRuntimeExecution(nint runtime);

    [DllImport(JsDll)]
    private static extern JsErrorCode JsIsRuntimeExecutionDisabled(nint runtime, out bool isDisabled);

    [LibraryImport(JsDll)]
    private static partial JsErrorCode JsCreateContext(nint runtime, nint debugApplication, out nint newContext);

    [LibraryImport(JsDll)]
    internal static partial JsErrorCode JsSetCurrentContext(nint context);

    [LibraryImport(JsDll)]
    internal static partial JsErrorCode JsGetCurrentContext(out nint context);

    [LibraryImport(JsDll, StringMarshalling = StringMarshalling.Utf16)]
    private static partial JsErrorCode JsParseScript(string script, nint sourceContext, string sourceUrl, out nint result);

    [LibraryImport(JsDll, StringMarshalling = StringMarshalling.Utf16)]
    private static partial JsErrorCode JsRunScript(string script, nint sourceContext, string sourceUrl, out nint result);

    [LibraryImport(JsDll, StringMarshalling = StringMarshalling.Utf16)]
    internal static partial JsErrorCode JsGetPropertyIdFromName(string name, out nint propertyId);

    [LibraryImport(JsDll)]
    internal static partial JsErrorCode JsGetProperty(nint @object, nint propertyId, out nint value);

    [DllImport(JsDll)]
    internal static extern JsErrorCode JsSetProperty(nint @object, nint propertyId, nint value, bool useStrictRules);

    [DllImport(JsDll)]
    internal static extern JsErrorCode JsVariantToValue(ref object variant, out nint value);

    [DllImport(JsDll)]
    internal static extern JsErrorCode JsValueToVariant(nint value, out object variant);

    [LibraryImport(JsDll)]
    internal static partial JsErrorCode JsGetValueType(nint value, out JsValueType type);

    [LibraryImport(JsDll)]
    private static partial JsErrorCode JsGetGlobalObject(out nint globalObject);

    [LibraryImport(JsDll)]
    private static partial JsErrorCode JsGetAndClearException(out nint exception);

    [LibraryImport(JsDll)]
    internal static partial JsErrorCode JsGetOwnPropertyNames(nint @object, out nint propertyNames);

    [LibraryImport(JsDll)]
    internal static partial JsErrorCode JsGetOwnPropertyDescriptor(nint @object, nint propertyId, out nint propertyDescriptor);

    [LibraryImport(JsDll)]
    internal static partial JsErrorCode JsCallFunction(nint function, nint[]? arguments, ushort argumentCount, out nint result);

    [LibraryImport(JsDll)]
    internal static partial JsErrorCode JsGetIndexedProperty(nint @object, nint index, out nint result);

    [LibraryImport(JsDll)]
    internal static partial JsErrorCode JsSetIndexedProperty(nint @object, nint index, nint value);

    [LibraryImport(JsDll)]
    internal static partial JsErrorCode JsGetPrototype(nint @object, out nint prototypeObject);

    [LibraryImport(JsDll)]
    private static partial JsErrorCode JsAddRef(nint handle, out int count);

    [LibraryImport(JsDll)]
    private static partial JsErrorCode JsRelease(nint handle, out int count);

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
}