namespace JsRt;

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
            JsRuntime.AddRef(handle, true, out var count);
            if (count > MaxRefCount)
            {
                MaxRefCount = count;
            }
        }
    }

    public nint Handle => _handle;
    public JsValue GlobalObject => _go.Value;
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
        var handle = Interlocked.Exchange(ref _handle, 0);
        if (handle != 0)
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

            JsRuntime.Release(handle, true, out var count);
            if (count > MaxRefCount)
            {
                MaxRefCount = count;
            }
        }
    }

    public JsValue Undefined => _undefined.Value;
    public JsValue True => _true.Value;
    public JsValue False => _false.Value;
    public JsValue Null => _null.Value;

    public static int MaxRefCount { get; private set; }
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
