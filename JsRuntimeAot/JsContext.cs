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

    public override string ToString() => Handle.ToString();
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
            JsRuntime.Release(handle, true, out var count);
            if (count > MaxRefCount)
            {
                MaxRefCount = count;
            }
        }
    }

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
