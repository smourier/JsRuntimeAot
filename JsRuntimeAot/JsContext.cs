namespace JsRt;

public class JsContext : IDisposable
{
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

#pragma warning disable IDE0079 // Remove unnecessary suppression
#pragma warning disable CA1822 // Mark members as static
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

    public JsValue Undefined { get { JsRuntime.JsGetUndefinedValue(out var handle); return new(handle); } }
    public JsValue Null { get { JsRuntime.JsGetNullValue(out var handle); return new(handle); } }
    public JsValue True { get { JsRuntime.JsGetTrueValue(out var handle); return new(handle); } }
    public JsValue False { get { JsRuntime.JsGetFalseValue(out var handle); return new(handle); } }
    public JsValue GlobalObject { get { JsRuntime.Check(JsRuntime.JsGetGlobalObject(out var handle)); return new JsValue(handle); } }
#pragma warning restore CA1822 // Mark members as static
#pragma warning restore IDE0079 // Remove unnecessary suppression

    public override string ToString() => Handle.ToString();

    public void AddGlobalObject(string name, object? value)
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

    public JsValue ObjectToValue(object? value, bool throwOnError = true)
    {
        if (value is JsValue jsv)
            return jsv;

        var error = VariantToValue(value, throwOnError, out var handle);
        if (error != null)
        {
            if (throwOnError)
                throw new JsRuntimeException(error);

            return Null;
        }

        return new JsValue(handle);
    }

    ~JsContext() { Dispose(disposing: false); }
    public void Dispose() { Dispose(disposing: true); GC.SuppressFinalize(this); }
    protected virtual void Dispose(bool disposing)
    {
        var handle = Interlocked.Exchange(ref _handle, 0);
        if (handle != 0)
        {
            JsRuntime.Release(handle);
        }
    }

    internal JsValue[] Convert(object?[]? arguments)
    {
        var values = new JsValue[arguments?.Length ?? 0];
        if (arguments != null)
        {
            for (var i = 0; i < arguments.Length; i++)
            {
                values[i] = ObjectToValue(arguments[i]);
            }
        }
        return values;
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

#pragma warning disable IDE0079 // Remove unnecessary suppression
#pragma warning disable CA1822 // Mark members as static
    internal Exception? VariantToValue(object? value, bool throwOnError, out nint handle)
    {
        using var v = new Variant(value);
        var error = JsRuntime.Check(JsRuntime.JsVariantToValue(v.Detached, out handle), throwOnError);
        return error;
    }

#pragma warning restore CA1822 // Mark members as static
#pragma warning restore IDE0079 // Remove unnecessary suppression
}
