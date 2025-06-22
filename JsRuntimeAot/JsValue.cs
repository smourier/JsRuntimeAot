namespace JsRt;

public class JsValue : IDisposable, IEquatable<JsValue>
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

    public bool Equals(JsValue? other) => other is not null && _handle == other._handle;
    public override bool Equals(object? obj) => obj is JsValue other && Equals(other);
    public override int GetHashCode() => _handle.GetHashCode();
    public static bool operator ==(JsValue? left, JsValue? right) => left?.Equals(right) ?? right is null;
    public static bool operator !=(JsValue? left, JsValue? right) => !(left == right);
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
        var context = Context;
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
            error = context.VariantToValue(value, false, out valueHandle);
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
        var context = Context;
        error = context.VariantToValue(value, false, out var valueHandle);
        if (error != null)
            return false;

        using var jsValue = context.ObjectToValue(index);
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
        var context = Context;
        using var iv = context.ObjectToValue(index);
        error = JsRuntime.Check(JsRuntime.JsGetIndexedProperty(Handle, iv.Handle, out nint valueHandle), false);
        if (error != null)
        {
            value = null;
            return false;
        }

        value = new JsValue(valueHandle);
        return true;
    }

    // don't forget the first argument is the function pointer or null if the function is global/static
    public virtual object? CallFunction(string name, params object?[]? arguments) => CallFunction<object?>(name, arguments);
    public virtual T? CallFunction<T>(string name, params object?[]? arguments)
    {
        ArgumentNullException.ThrowIfNull(name);
        if (!TryCallFunction(name, out T? value, arguments))
            return default;

        return value;
    }

    public virtual bool TryCallFunction<T>(string name, out T? value, params object?[]? arguments)
    {
        ArgumentNullException.ThrowIfNull(name);
        using var fn = GetProperty<JsValue>(name);
        if (fn == null)
        {
            value = default;
            return false;
        }

        if (!fn.TryCall(out _, out var jsValue, arguments) || jsValue == null)
        {
            value = default;
            return false;
        }

        try
        {
            return TryChangeType(jsValue.Value, out value);
        }
        finally
        {
            jsValue.Dispose();
        }
    }

    public virtual object? Call(params object?[] arguments)
    {
        ArgumentNullException.ThrowIfNull(arguments);
        var context = Context;
        var args = context.Convert(arguments);
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

    public virtual bool TryCall(out Exception? error, out JsValue? value, params object?[]? arguments)
    {
        var context = Context;
        var args = context.Convert(arguments);
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

#pragma warning disable IDE0079 // Remove unnecessary suppression
#pragma warning disable CA1822 // Mark members as static
    private JsContext Context => JsContext.Current ?? throw new InvalidOperationException("No active JavaScript context.");
#pragma warning restore CA1822 // Mark members as static
#pragma warning restore IDE0079 // Remove unnecessary suppression

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

    private static void Dispose(IEnumerable<JsValue> arguments)
    {
        foreach (var arg in arguments)
        {
            arg.Dispose();
        }
    }
}