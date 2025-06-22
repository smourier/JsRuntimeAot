namespace JsRt;

public class JsContext : IDisposable, IEquatable<JsContext>
{
    private readonly Lazy<JsValue> _globalObject = new(() => { JsRuntime.Check(JsRuntime.JsGetGlobalObject(out var handle)); return new JsValue(handle); });
    private readonly Lazy<JsValue> _undefined = new(() => { JsRuntime.JsGetUndefinedValue(out var handle); return new JsValue(handle); });
    private readonly Lazy<JsValue> _null = new(() => { JsRuntime.JsGetNullValue(out var handle); return new JsValue(handle); });
    private readonly Lazy<JsValue> _true = new(() => { JsRuntime.JsGetTrueValue(out var handle); return new JsValue(handle); });
    private readonly Lazy<JsValue> _false = new(() => { JsRuntime.JsGetFalseValue(out var handle); return new JsValue(handle); });

    private nint _handle;
    private static readonly Lock _currentLock = new();

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

    public JsValue Undefined => _undefined.Value;
    public JsValue Null => _null.Value;
    public JsValue True => _true.Value;
    public JsValue False => _false.Value;
    public JsValue GlobalObject => _globalObject.Value;
#pragma warning restore CA1822 // Mark members as static
#pragma warning restore IDE0079 // Remove unnecessary suppression

    public bool Equals(JsContext? other) => other is not null && _handle == other._handle;
    public override bool Equals(object? obj) => obj is JsContext other && Equals(other);
    public override int GetHashCode() => _handle.GetHashCode();
    public static bool operator ==(JsContext? left, JsContext? right) => left?.Equals(right) ?? right is null;
    public static bool operator !=(JsContext? left, JsContext? right) => !(left == right);
    public override string ToString() => Handle.ToString();

    public void AddGlobalObject(string name, object? value)
    {
        ArgumentNullException.ThrowIfNull(name);
        if (value != null && !Marshal.IsTypeVisibleFromCom(value.GetType()) && !Marshal.IsComObject(value))
            throw new ArgumentException("Argument type must be ComVisible.", nameof(value));

        GlobalObject.SetProperty(name, value);
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
            if (_null.IsValueCreated)
            {
                _null.Value.Dispose();
            }

            if (_undefined.IsValueCreated)
            {
                _undefined.Value.Dispose();
            }

            if (_true.IsValueCreated)
            {
                _true.Value.Dispose();
            }

            if (_false.IsValueCreated)
            {
                _false.Value.Dispose();
            }

            if (_globalObject.IsValueCreated)
            {
                _globalObject.Value.Dispose();
            }
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

    private static JsContext? _current;
    public static JsContext? Current
    {
        get
        {
            if (_current is null)
            {
                lock (_currentLock)
                {
                    JsRuntime.Check(JsRuntime.JsGetCurrentContext(out var handle));
                    if (handle != 0)
                    {
                        _current = new JsContext(handle, true);
                    }
                }
            }
            return _current;
        }
        set
        {
            if (_current == value)
                return;

            lock (_currentLock)
            {
                _current?.Dispose();
                _current = value;
                JsRuntime.Check(JsRuntime.JsSetCurrentContext(value != null ? value.Handle : 0));
            }
        }
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
