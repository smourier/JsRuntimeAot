namespace JsRt;

public sealed class JsContext : IDisposable
{
    public static int MaxRefCount { get; private set; }

    private nint _handle;

    internal JsContext(nint handle, bool addRef)
    {
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

    public void Dispose()
    {
        var handle = Interlocked.Exchange(ref _handle, 0);
        if (handle != 0)
        {
            JsRuntime.Release(handle, true, out var count);
            if (count > MaxRefCount)
            {
                MaxRefCount = count;
            }
        }
    }

    public override string ToString() => Handle.ToString();

    public static JsContext? Current
    {
        get
        {
            JsRuntime.Check(JsRuntime.JsGetCurrentContext(out var handle));
            return handle != 0 ? new JsContext(handle, false) : null;
        }

        set => JsRuntime.JsSetCurrentContext(value != null ? value.Handle : 0);
    }

    public void Execute(JsAction action)
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

    public T Execute<T>(JsAction<T> action)
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
}
