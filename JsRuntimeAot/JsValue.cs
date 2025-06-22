namespace JsRt;

public sealed class JsValue : IDisposable
{
    private nint _handle;

    public JsValue(object? value)
    {
        VariantToValue(value, true, out var handle);
        _handle = handle;
        JsRuntime.Check(JsRuntime.JsGetValueType(Handle, out var vt));
        ValueType = vt;

        JsRuntime.AddRef(handle, true, out var count);
        if (count > MaxRefCount)
        {
            MaxRefCount = count;
        }
    }

    public static int MaxRefCount { get; set; }

    internal JsValue(nint handle)
    {
        _handle = handle;
        JsRuntime.Check(JsRuntime.JsGetValueType(Handle, out var vt));
        ValueType = vt;

        JsRuntime.AddRef(handle, true, out var count);
        if (count > MaxRefCount)
        {
            MaxRefCount = count;
        }
    }

    public nint Handle => _handle;

    public static bool IsUndefined(object obj) => obj is JsValue jsv && jsv.ValueType == JsValueType.JsUndefined;

    public override string ToString()
    {
        if (ValueType == JsValueType.JsNull || ValueType == JsValueType.JsUndefined)
            return ValueType.ToString();

        return ValueType + ": " + string.Format("{0}", Value);
    }

    public void Dispose()
    {
        var h = Interlocked.Exchange(ref _handle, 0);
        if (h != 0)
        {
            JsRuntime.Release(h, true, out var count);
            if (count > MaxRefCount)
            {
                MaxRefCount = count;
            }
        }
    }

    public JsValueType ValueType { get; private set; }

    public object Value
    {
        get
        {
            JsRuntime.Check(JsRuntime.JsValueToVariant(Handle, out var value), false);
            return value;
        }
    }

    public object DetachValue()
    {
        object value = Value;
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

    public IDictionary<string, JsValue?> PropertyValues
    {
        get
        {
            var names = PropertyNames;
            if (names == null)
                return new Dictionary<string, JsValue?>();

            var props = new Dictionary<string, JsValue?>();
            foreach (string name in names)
            {
                props.Add(name, GetProperty<JsValue?>(name, null));
            }
            return props;
        }
    }

    public IList<JsValue> PropertyDescriptors
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

    public IList<string> PropertyNames
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

    private static Exception? VariantToValue(object? variant, bool throwOnError, out nint handle)
    {
        var value = variant;
        var error = JsRuntime.Check(JsRuntime.JsVariantToValue(ref value, out handle), throwOnError);
        return error;
    }

    public bool TrySetProperty(string name, object? value, bool useStrictRules, out Exception? error)
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
    public bool TrySetProperty(int index, object? value, out Exception? error)
    {
        error = VariantToValue(value, false, out var valueHandle);
        if (error != null)
            return false;

        using (var jsValue = new JsValue(index))
        {
            error = JsRuntime.Check(JsRuntime.JsSetIndexedProperty(Handle, jsValue.Handle, valueHandle));
        }
        return error == null;
    }

    public T? GetProperty<T>(string name, T? defaultValue)
    {
        if (!TryGetProperty(name, out _, out var value) || value == null)
            return defaultValue;

        if (typeof(T) == typeof(JsValue))
            return (T)(object)value!;

        try
        {
            return ChangeType(value.Value, defaultValue);
        }
        finally
        {
            value?.Dispose();
        }
    }

    public bool TryGetProperty(string name, out Exception? error, out JsValue? value)
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

    public T GetProperty<T>(int index, T defaultValue)
    {
        if (!TryGetProperty(index, out _, out var value) || value == null)
            return defaultValue;

        if (typeof(T) == typeof(JsValue))
            return (T)(object)value!; // don't dispose this one since we want the JsValue itself

        try
        {
            return ChangeType(value.Value, defaultValue);
        }
        finally
        {
            value?.Dispose();
        }
    }

    public bool TryGetProperty(int index, out Exception? error, out JsValue? value)
    {
        nint valueHandle;
        using (var iv = new JsValue(index))
        {
            error = JsRuntime.Check(JsRuntime.JsGetIndexedProperty(Handle, iv.Handle, out valueHandle), false);
        }

        if (error != null)
        {
            value = null;
            return false;
        }
        value = new JsValue(valueHandle);
        return true;
    }

    private static void Dispose(IEnumerable<JsValue>? arguments)
    {
        if (arguments == null)
            return;

        foreach (var arg in arguments)
        {
            if (arg == null)
                continue;

            arg.Dispose();
        }
    }

    private static JsValue[]? Convert(object?[]? arguments)
    {
        if (arguments == null || arguments.Length == 0)
            return null;

        var values = new JsValue[arguments.Length];
        for (var i = 0; i < arguments.Length; i++)
        {
            values[i] = new JsValue(arguments[i]);
        }
        return values;
    }

    public T? CallFunction<T>(string name, T? defaultValue, params object[] arguments)
    {
        using var fn = GetProperty<JsValue>(name, null);
        if (fn == null || IsUndefined(fn))
            return defaultValue;

        JsValue? value = null;
        try
        {
            if (!fn.TryCall(out var error, out value, arguments) || value == null)
                return defaultValue;

            return ChangeType(value.Value, defaultValue);
        }
        finally
        {
            value?.Dispose();
        }
    }

    public object? Call(params object?[]? arguments)
    {
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

    public object? Call(JsValue?[]? arguments)
    {
        if (!TryCall(out var error, out var value, arguments) || value == null)
        {
            if (error != null)
                throw error;

            return null;
        }

        try
        {
            return value.Value;
        }
        finally
        {
            value.Dispose();
        }
    }

    public T? CallWithDefault<T>(T? defaultValue, params object[]? arguments)
    {
        var args = Convert(arguments);
        try
        {
            return CallWithDefault(defaultValue, args);
        }
        finally
        {
            Dispose(args);
        }
    }

    public T CallWithDefault<T>(T defaultValue, JsValue[]? arguments)
    {
        if (!TryCall(out _, out var value, arguments))
            return defaultValue;

        if (typeof(T) == typeof(JsValue))
            return (T)(object)value!;

        if (value == null)
            return defaultValue;

        try
        {
            return ChangeType(value.Value, defaultValue);
        }
        finally
        {
            value.Dispose();
        }
    }

    public bool TryCall(out Exception? error, out JsValue? value, params object?[]? arguments)
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

    public bool TryCall(out Exception? error, out JsValue? value, JsValue?[]? arguments)
    {
        nint[]? args = null;
        if (arguments != null && arguments.Length > 0)
        {
            args = new nint[arguments.Length];
            for (var i = 0; i < arguments.Length; i++)
            {
                args[i] = arguments[i]?.Handle ?? 0;
            }
        }

        error = JsRuntime.Check(JsRuntime.JsCallFunction(Handle, args, (ushort)(args != null ? args.Length : 0), out var result), false);
        if (error != null)
        {
            value = null;
            return false;
        }

        value = new JsValue(result);
        return true;
    }

    public static T ChangeType<T>(object value, T defaultValue)
    {
        if (value == null)
            return defaultValue;

        if (value is T t)
            return t;

        return (T)value;
    }
}


