namespace JsRt;

public sealed class Variant : IDisposable
{
    private VARIANT _inner;

    public VARIANT Detached => _inner;
    public ref VARIANT RefDetached => ref _inner;

    public static int Size { get; } = GetSizeOf32();
    private static int GetSizeOf32() { unsafe { return sizeof(VARIANT); } }

    internal Variant(VARIANT inner)
    {
        _inner = inner;
    }

    public Variant()
    {
        // it's a VT_EMPTY
    }

    public Variant(object? value, VARENUM? type = null)
    {
        if (value == null)
        {
            _inner.Anonymous.Anonymous.vt = VARENUM.VT_NULL;
            return;
        }

        value = Unwrap(value);

        if (value is nint ptr)
        {
            _inner.Anonymous.Anonymous.Anonymous.punkVal = ptr;
            _inner.Anonymous.Anonymous.vt = type ?? VARENUM.VT_UNKNOWN;
            return;
        }

        if (value is nuint uptr)
        {
            _inner.Anonymous.Anonymous.Anonymous.punkVal = (nint)uptr;
            _inner.Anonymous.Anonymous.vt = type ?? VARENUM.VT_UNKNOWN;
            return;
        }

        if (value is ComObject co)
        {
            var sw = new StrategyBasedComWrappers();
            _inner.Anonymous.Anonymous.Anonymous.punkVal = sw.GetOrCreateComInterfaceForObject(co, CreateComInterfaceFlags.None);
            _inner.Anonymous.Anonymous.vt = VARENUM.VT_UNKNOWN;
            return;
        }

        if (value is char[] chars)
        {
            value = new string(chars);
        }

        if (value is char[][] charray)
        {
            var strings = new string[charray.GetLength(0)];
            for (var i = 0; i < charray.Length; i++)
            {
                strings[i] = new string(charray[i]);
            }
            value = strings;
        }

        if (value is Array array)
        {
            ConstructArray(array, type);
            return;
        }

        if (value is not string && value is IEnumerable enumerable)
        {
            ConstructEnumerable(enumerable, type);
            return;
        }

        if (value == null)
        {
            _inner.Anonymous.Anonymous.vt = VARENUM.VT_NULL;
            return;
        }

        var valueType = value.GetType();
        var vt = FromType(valueType, type, true);
        var tc = Type.GetTypeCode(valueType);
        switch (tc)
        {
            case TypeCode.Boolean:
                _inner.Anonymous.Anonymous.Anonymous.boolVal = new VARIANT_BOOL { Value = (bool)value ? (short)(-1) : (short)0 };
                vt = VARENUM.VT_BOOL;
                break;

            case TypeCode.Byte:
                _inner.Anonymous.Anonymous.Anonymous.bVal = (byte)value;
                vt = VARENUM.VT_UI1;
                break;

            case TypeCode.Char:
                chars = [(char)value];
                // note: all strings (PWSTR, PSTR, BSTR) point to same place
                _inner.Anonymous.Anonymous.Anonymous.bstrVal = new BSTR { Value = MarshalString(new string(chars), VARENUM.VT_BSTR) };
                vt = VARENUM.VT_BSTR;
                break;

            case TypeCode.DateTime:
                if (type == VARENUM.VT_FILETIME)
                {
                    var ft = ToPositiveFILETIME((DateTime)value);
                    Functions.InitVariantFromFileTime(ft, out _inner);
                    return;
                }

                var dt = (DateTime)value;
                _inner.Anonymous.Anonymous.Anonymous.dblVal = dt.ToOADate();
                vt = VARENUM.VT_DATE;
                break;

            case TypeCode.Empty:
            case TypeCode.DBNull:
                break;

            case TypeCode.Decimal:
                _inner.Anonymous.decVal = (decimal)value;
                vt = VARENUM.VT_DECIMAL;
                break;

            case TypeCode.Double:
                _inner.Anonymous.Anonymous.Anonymous.dblVal = (double)value;
                vt = VARENUM.VT_R8;
                break;

            case TypeCode.Int16:
                _inner.Anonymous.Anonymous.Anonymous.iVal = (short)value;
                vt = VARENUM.VT_I2;
                break;

            case TypeCode.Int32:
                _inner.Anonymous.Anonymous.Anonymous.lVal = (int)value;
                vt = VARENUM.VT_I4;
                break;

            case TypeCode.Int64:
                _inner.Anonymous.Anonymous.Anonymous.llVal = (long)value;
                vt = VARENUM.VT_I8;
                break;

            case TypeCode.SByte:
                _inner.Anonymous.Anonymous.Anonymous.cVal.Value = (sbyte)value;
                vt = VARENUM.VT_I1;
                break;

            case TypeCode.Single:
                _inner.Anonymous.Anonymous.Anonymous.fltVal = (float)value;
                vt = VARENUM.VT_R4;
                break;

            case TypeCode.String:
                // note: all strings (PWSTR, PSTR, BSTR) point to same place
                _inner.Anonymous.Anonymous.Anonymous.bstrVal = new BSTR { Value = MarshalString((string)value, VARENUM.VT_BSTR) };
                vt = VARENUM.VT_BSTR;
                break;

            case TypeCode.UInt16:
                _inner.Anonymous.Anonymous.Anonymous.uiVal = (ushort)value;
                vt = VARENUM.VT_UI2;
                break;

            case TypeCode.UInt32:
                _inner.Anonymous.Anonymous.Anonymous.ulVal = (uint)value;
                vt = VARENUM.VT_UI4;
                break;

            case TypeCode.UInt64:
                _inner.Anonymous.Anonymous.Anonymous.ullVal = (ulong)value;
                vt = VARENUM.VT_UI8;
                break;

            //case TypeCode.Object:
            default:
                if (value is Guid guid)
                {
                    _inner.Anonymous.Anonymous.Anonymous.bstrVal = new BSTR { Value = MarshalString(guid.ToString("B"), VARENUM.VT_BSTR) };
                    vt = VARENUM.VT_BSTR;
                    break;
                }

                if (value is DateTimeOffset dto)
                {
                    if (type == VARENUM.VT_FILETIME)
                    {
                        var ft = ToPositiveFILETIME(dto.DateTime);
                        Functions.InitVariantFromFileTime(ft, out _inner);
                        return;
                    }

                    _inner.Anonymous.Anonymous.Anonymous.dblVal = dto.DateTime.ToOADate();
                    vt = VARENUM.VT_DATE;
                    break;
                }

                throw new ArgumentException("Value of type '" + value.GetType().FullName + "' is not supported.", nameof(value));
        }

        _inner.Anonymous.Anonymous.vt = vt;
    }

    public VARENUM VarType { get => _inner.Anonymous.Anonymous.vt; }
    public object? Value
    {
        get
        {
            switch (_inner.Anonymous.Anonymous.vt)
            {
                case VARENUM.VT_EMPTY:
                case VARENUM.VT_NULL: // DbNull
                    return null;

                case VARENUM.VT_I1:
                    return _inner.Anonymous.Anonymous.Anonymous.cVal.Value;

                case VARENUM.VT_UI1:
                    return _inner.Anonymous.Anonymous.Anonymous.bVal;

                case VARENUM.VT_I2:
                    return _inner.Anonymous.Anonymous.Anonymous.iVal;

                case VARENUM.VT_UI2:
                    return _inner.Anonymous.Anonymous.Anonymous.uiVal;

                case VARENUM.VT_I4:
                case VARENUM.VT_INT:
                    return _inner.Anonymous.Anonymous.Anonymous.lVal;

                case VARENUM.VT_UI4:
                case VARENUM.VT_UINT:
                    return _inner.Anonymous.Anonymous.Anonymous.ulVal;

                case VARENUM.VT_I8:
                    return _inner.Anonymous.Anonymous.Anonymous.llVal;

                case VARENUM.VT_UI8:
                    return _inner.Anonymous.Anonymous.Anonymous.ullVal;

                case VARENUM.VT_R4:
                    return _inner.Anonymous.Anonymous.Anonymous.fltVal;

                case VARENUM.VT_R8:
                    return _inner.Anonymous.Anonymous.Anonymous.dblVal;

                case VARENUM.VT_BOOL:
                    return _inner.Anonymous.Anonymous.Anonymous.boolVal.Value != 0;

                case VARENUM.VT_ERROR:
                    return _inner.Anonymous.Anonymous.Anonymous.scode;

                case VARENUM.VT_CY:
                    return _inner.Anonymous.decVal;

                case VARENUM.VT_DATE:
                    return DateTime.FromOADate(_inner.Anonymous.Anonymous.Anonymous.dblVal);

                case VARENUM.VT_BSTR:
                    return Marshal.PtrToStringBSTR(_inner.Anonymous.Anonymous.Anonymous.bstrVal.Value);

                case VARENUM.VT_LPSTR:
                    // all strings point to same place anyway
                    return Marshal.PtrToStringAnsi(_inner.Anonymous.Anonymous.Anonymous.bstrVal.Value);

                case VARENUM.VT_LPWSTR:
                    // all strings point to same place anyway
                    return Marshal.PtrToStringUni(_inner.Anonymous.Anonymous.Anonymous.bstrVal.Value);

                case VARENUM.VT_UNKNOWN:
                case VARENUM.VT_DISPATCH:
                    var sw = new StrategyBasedComWrappers();
                    return sw.GetOrCreateObjectForComInstance(_inner.Anonymous.Anonymous.Anonymous.punkVal, CreateObjectFlags.UniqueInstance);

                case VARENUM.VT_DECIMAL:
                    return _inner.Anonymous.decVal;

                default:
                    if (_inner.Anonymous.Anonymous.vt.HasFlag(VARENUM.VT_ARRAY))
                    {
                        var et = _inner.Anonymous.Anonymous.vt & ~VARENUM.VT_ARRAY;
                        if (TryGetArrayValue(et, out var array))
                            return array;
                    }

                    throw new NotSupportedException("Value of property type " + _inner.Anonymous.Anonymous.vt + " is not supported.");
            }
        }
    }

    public Variant? ChangeType(VARENUM type, bool throwOnError = true)
    {
        var inner = new VARIANT();
        var hr = Functions.VariantChangeType(ref inner, _inner, 0, type).ThrowOnError(throwOnError);
        if (hr.IsError)
            return null;

        return new Variant { _inner = inner };
    }

    public void CopyFrom(Variant source, bool throwOnError = true)
    {
        ArgumentNullException.ThrowIfNull(source);
        if (source == this)
            return;

        Clear(throwOnError);
        var inner = new VARIANT();
        Functions.VariantCopy(ref inner, source._inner).ThrowOnError(throwOnError);
        _inner = inner;
    }

    public Variant? Copy(bool throwOnError = true)
    {
        var inner = new VARIANT();
        var hr = Functions.VariantCopy(ref inner, _inner).ThrowOnError(throwOnError);
        if (hr.IsError)
            return null;

        return new Variant { _inner = inner };
    }

    public VARIANT Detach()
    {
        var pv = _inner;
        Zero();
        return pv;
    }

    public unsafe void DetachTo(nint variantPtr)
    {
        if (variantPtr == 0)
            throw new ArgumentException(null, nameof(variantPtr));

        var pv = _inner;
        Zero();
        *(VARIANT*)(variantPtr) = pv;
    }

    public static nint MarshalString(string? text, VARENUM vt)
    {
        if (text == null)
            return 0;

        return vt switch
        {
            VARENUM.VT_LPWSTR => Marshal.StringToCoTaskMemUni(text),
            VARENUM.VT_BSTR => Marshal.StringToBSTR(text),
            VARENUM.VT_LPSTR => Marshal.StringToCoTaskMemAnsi(text),
            _ => throw new NotSupportedException("A string can only be of property type VT_LPWSTR, VT_LPSTR or VT_BSTR."),
        };
    }

    public static string? PtrTostring(nint ptr, VARENUM vt)
    {
        if (ptr == 0)
            return null;

        return vt switch
        {
            VARENUM.VT_LPWSTR => Marshal.PtrToStringUni(ptr),
            VARENUM.VT_BSTR => Marshal.PtrToStringBSTR(ptr),
            VARENUM.VT_LPSTR => Marshal.PtrToStringAnsi(ptr),
            _ => throw new NotSupportedException("A string can only be of property type VT_LPWSTR, VT_LPSTR or VT_BSTR."),
        };
    }

    private static Type FromType(VARENUM type) => type switch
    {
        VARENUM.VT_I1 => typeof(sbyte),
        VARENUM.VT_UI1 => typeof(byte),
        VARENUM.VT_I2 => typeof(short),
        VARENUM.VT_UI2 => typeof(ushort),
        VARENUM.VT_UI4 or VARENUM.VT_UINT => typeof(uint),
        VARENUM.VT_I8 => typeof(long),
        VARENUM.VT_UI8 => typeof(ulong),
        VARENUM.VT_R4 => typeof(float),
        VARENUM.VT_R8 => typeof(double),
        VARENUM.VT_BOOL => typeof(bool),
        VARENUM.VT_I4 or VARENUM.VT_INT or VARENUM.VT_ERROR => typeof(int),
        VARENUM.VT_DATE => typeof(DateTime),
        VARENUM.VT_FILETIME => typeof(ulong),
        VARENUM.VT_BLOB => typeof(byte[]),
        VARENUM.VT_CLSID => typeof(Guid),
        VARENUM.VT_BSTR or VARENUM.VT_LPSTR or VARENUM.VT_LPWSTR => typeof(string),
        VARENUM.VT_UNKNOWN or VARENUM.VT_DISPATCH => typeof(object),
        VARENUM.VT_CY or VARENUM.VT_DECIMAL => typeof(decimal),
        _ => throw new ArgumentException("Property type " + type + " is not supported.", nameof(type)),
    };

    private static VARENUM FromType(Type type, VARENUM? vt, bool forVariant)
    {
        if (type == null)
            return VARENUM.VT_NULL;

        var tc = Type.GetTypeCode(type);
        switch (tc)
        {
            case TypeCode.Boolean:
                return VARENUM.VT_BOOL;

            case TypeCode.Byte:
                return VARENUM.VT_UI1;

            case TypeCode.Char:
                if (forVariant)
                    return VARENUM.VT_BSTR;

                return VARENUM.VT_LPWSTR;

            case TypeCode.DateTime:
                if (vt == VARENUM.VT_FILETIME)
                    return VARENUM.VT_FILETIME;

                return VARENUM.VT_DATE;

            case TypeCode.DBNull:
                return VARENUM.VT_NULL;

            case TypeCode.Decimal:
                return VARENUM.VT_DECIMAL;

            case TypeCode.Double:
                return VARENUM.VT_R8;

            case TypeCode.Empty:
                return VARENUM.VT_EMPTY;

            case TypeCode.Int16:
                return VARENUM.VT_I2;

            case TypeCode.Int32:
                return VARENUM.VT_I4;

            case TypeCode.Int64:
                return VARENUM.VT_I8;

            case TypeCode.SByte:
                return VARENUM.VT_I1;

            case TypeCode.Single:
                return VARENUM.VT_R4;

            case TypeCode.String:
                if (forVariant)
                    return VARENUM.VT_BSTR;

                if (!vt.HasValue)
                    return VARENUM.VT_LPWSTR;

                if (vt != VARENUM.VT_LPSTR && vt != VARENUM.VT_BSTR && vt != VARENUM.VT_LPWSTR)
                    throw new ArgumentException("Property type " + vt + " is not supported for string.", nameof(type));

                return vt.Value;

            case TypeCode.UInt16:
                return VARENUM.VT_UI2;

            case TypeCode.UInt32:
                return VARENUM.VT_UI4;

            case TypeCode.UInt64:
                return VARENUM.VT_UI8;

            // case TypeCode.Object:
            default:
                if (type == typeof(Guid))
                {
                    if (forVariant)
                        return VARENUM.VT_BSTR;

                    return VARENUM.VT_CLSID;
                }

                if (type == typeof(FILETIME))
                {
                    if (forVariant)
                        return VARENUM.VT_DATE;

                    return VARENUM.VT_FILETIME;
                }

                if (type == typeof(byte))
                {
                    if (forVariant)
                        return VARENUM.VT_UI1 | VARENUM.VT_ARRAY;

                    if (!vt.HasValue)
                        return VARENUM.VT_UI1 | VARENUM.VT_VECTOR;

                    if (vt != VARENUM.VT_BLOB && vt != (VARENUM.VT_UI1 | VARENUM.VT_VECTOR))
                        throw new ArgumentException("Property type " + vt + " is not supported for array of bytes.", nameof(type));

                    return vt.Value;
                }

                if (type == typeof(object))
                    return VARENUM.VT_VARIANT;

                throw new ArgumentException("Value of type '" + type.FullName + "' is not supported.", nameof(type));
        }
    }

    public static Variant Attach(ref VARIANT detached, bool zeroDetached = true)
    {
        var pv = new Variant { _inner = detached };
        if (zeroDetached)
        {
            unsafe
            {
                var ptr = Unsafe.AsPointer(ref detached);
                Functions.ZeroMemory((nint)ptr, Size);
            }
        }
        return pv;
    }

    public static object? Unwrap(object? value)
    {
        if (value is Variant variant)
            return Unwrap(variant.Value);

        if (value is VARIANT v)
        {
            var v2 = Attach(ref v, false);
            value = v2.Value;
            v2.Detach();
            return Unwrap(value);
        }

        return value;
    }

    public override string ToString()
    {
        var value = Value;
        if (value == null)
            return "<null>";

        if (value is string svalue)
            return "[" + VarType + "] `" + svalue + "`";

        if (value is not byte[] && value is IEnumerable enumerable)
            return "[" + VarType + "] " + string.Join(", ", enumerable.OfType<object>());

        if (value is byte[] bytes)
            return "[" + VarType + "] bytes[" + bytes.Length + "]";

        return "[" + VarType + "] " + value;
    }

    ~Variant() => Dispose();
    public void Dispose() { Clear(false); GC.SuppressFinalize(this); }

    private void Zero()
    {
        unsafe
        {
            fixed (VARIANT* p = &_inner)
            {
                Functions.ZeroMemory((nint)p, Size);
            }
        }
    }

    private void ConstructEnumerable(IEnumerable enumerable, VARENUM? type = null)
    {
        var et = GetElementType(enumerable) ?? throw new ArgumentException("Enumerable type '" + enumerable.GetType().FullName + "' is not supported.", nameof(enumerable));
        var count = GetCount(enumerable);
#pragma warning disable IDE0079 // Remove unnecessary suppression
#pragma warning disable IL3050 // Calling members annotated with 'RequiresDynamicCodeAttribute' may break functionality when AOT compiling.
        var array = Array.CreateInstance(et, count);
#pragma warning restore IL3050 // Calling members annotated with 'RequiresDynamicCodeAttribute' may break functionality when AOT compiling.
#pragma warning restore IDE0079 // Remove unnecessary suppression
        var i = 0;
        foreach (var obj in enumerable)
        {
            array.SetValue(obj, i++);
        }
        ConstructArray(array, type);
    }

    private static int GetCount(IEnumerable enumerable)
    {
        if (enumerable is ICollection col)
            return col.Count;

        var count = 0;
        var e = enumerable.GetEnumerator();
        Using(e, () =>
        {
            while (e.MoveNext())
            {
                count++;
            }
        });
        return count;

        static void Using(object resource, Action action)
        {
            try
            {
                action();
            }
            finally
            {
                (resource as IDisposable)?.Dispose();
            }
        }
    }

    private static Type? GetElementType(IEnumerable enumerable)
    {
        var et = GetElementType(enumerable.GetType());
        if (et != null)
            return et;

        foreach (var obj in enumerable)
        {
            return obj.GetType();
        }
        return null;
    }

    private static Type? GetElementType(Type collectionType)
    {
#pragma warning disable IDE0079 // Remove unnecessary suppression
#pragma warning disable IL2070 // 'this' argument does not satisfy 'DynamicallyAccessedMembersAttribute' in call to target method. The parameter of method does not have matching annotations.
        foreach (var iface in collectionType.GetInterfaces())
        {
            if (!iface.IsGenericType)
                continue;

            if (iface.GetGenericTypeDefinition() == typeof(IEnumerable<>))
                return iface.GetGenericArguments()[0];

            if (iface.GetGenericTypeDefinition() == typeof(ICollection<>))
                return iface.GetGenericArguments()[0];

            if (iface.GetGenericTypeDefinition() == typeof(IList<>))
                return iface.GetGenericArguments()[0];
        }
#pragma warning restore IL2070 // 'this' argument does not satisfy 'DynamicallyAccessedMembersAttribute' in call to target method. The parameter of method does not have matching annotations.
#pragma warning restore IDE0079 // Remove unnecessary suppression
        return null;
    }

    private void ConstructArray(Array array, VARENUM? type = null)
    {
        // special case for bools which are shorts...
        if (array is bool[] bools)
        {
            var shorts = new short[bools.Length];
            for (var i = 0; i < bools.Length; i++)
            {
                shorts[i] = bools[i] ? ((short)(-1)) : ((short)0);
            }
            ConstructSafeArray(shorts, typeof(short), VARENUM.VT_BOOL);
            return;
        }

        if (array is Guid[] guids)
        {
            var strings = new string[guids.Length];
            for (var i = 0; i < strings.Length; i++)
            {
                strings[i] = guids[i].ToString("B");
            }
            ConstructSafeArray(strings, typeof(string), VARENUM.VT_BSTR);
            return;
        }

        var et = array.GetType().GetElementType() ?? throw new NotSupportedException();
        ConstructSafeArray(array, et, FromType(et, type, true));
    }

    private void ConstructSafeArray(Array array, Type type, VARENUM vt)
    {
        unsafe
        {
            var bounds = new SAFEARRAYBOUND { lLbound = 0, cElements = (uint)array.Length };
            var sa = Functions.SafeArrayCreate(vt, 1, bounds);
            if (sa == 0)
                throw new OutOfMemoryException();

            var psa = (SAFEARRAY*)sa;
            Functions.SafeArrayAccessData(*psa, out var ptr).ThrowOnError();
            try
            {
                if (type == typeof(string))
                {
                    for (var i = 0; i < array.Length; i++)
                    {
                        var str = MarshalString((string?)array.GetValue(i)!, vt);
                        Marshal.WriteIntPtr(ptr, nint.Size * i, str);
                    }
                }
                else if (type == typeof(object))
                {
                    for (var i = 0; i < array.Length; i++)
                    {
                        var variantValue = array.GetValue(i);
                        using var variant = new Variant(variantValue);
                        unsafe
                        {
                            var p = (VARIANT*)(ptr + Size * i);
                            *p = variant.Detach();
                        }
                    }
                }
                else
                {
#pragma warning disable IDE0079 // Remove unnecessary suppression
#pragma warning disable IL3050 // Calling members annotated with 'RequiresDynamicCodeAttribute' may break functionality when AOT compiling.
#pragma warning disable CA1421 // This method uses runtime marshalling even when the 'DisableRuntimeMarshallingAttribute' is applied
                    var size = Marshal.SizeOf(type) * array.Length;
#pragma warning restore CA1421 // This method uses runtime marshalling even when the 'DisableRuntimeMarshallingAttribute' is applied
#pragma warning restore IL3050 // Calling members annotated with 'RequiresDynamicCodeAttribute' may break functionality when AOT compiling.
#pragma warning restore IDE0079 // Remove unnecessary suppression

                    Functions.CopyMemory(ptr, Marshal.UnsafeAddrOfPinnedArrayElement(array, 0), size);
                }
            }
            catch
            {
                Functions.SafeArrayDestroy(*psa);
                throw;
            }
            finally
            {
                Functions.SafeArrayUnaccessData(*psa).ThrowOnError();
            }

            _inner.Anonymous.Anonymous.vt = vt | VARENUM.VT_ARRAY;
            _inner.Anonymous.Anonymous.Anonymous.parray = sa;
        }
    }

    private bool TryGetArrayValue(VARENUM vt, out object? value)
    {
        value = null;
        if (_inner.Anonymous.Anonymous.Anonymous.parray == 0)
            return false;

        unsafe
        {
            var psa = (SAFEARRAY*)_inner.Anonymous.Anonymous.Anonymous.parray;
            if (psa->cDims != 1)
                return false;

            Functions.SafeArrayGetLBound(*psa, 1, out var l).ThrowOnError();
            Functions.SafeArrayGetUBound(*psa, 1, out var u).ThrowOnError();
            var count = u - l + 1;

            Functions.SafeArrayAccessData(*psa, out var ptr).ThrowOnError();
            try
            {
                var ret = false;
                uint size;
                switch (vt)
                {
                    case VARENUM.VT_LPSTR:
                    case VARENUM.VT_LPWSTR:
                    case VARENUM.VT_BSTR:
                        var strings = new string?[count];
                        for (var i = 0; i < strings.Length; i++)
                        {
                            var str = Marshal.ReadIntPtr(ptr, (int)(psa->cbElements * i));
                            strings[i] = PtrTostring(str, vt);
                        }
                        value = strings;
                        ret = true;
                        break;

                    case VARENUM.VT_BOOL:
                        var shorts = new short[count];
                        size = (uint)(shorts.Length * sizeof(short));
                        Functions.CopyMemory(Marshal.UnsafeAddrOfPinnedArrayElement(shorts, 0), ptr, (nint)size);
                        var bools = new bool[shorts.Length];
                        for (var i = 0; i < shorts.Length; i++)
                        {
                            bools[i] = shorts[i] != 0;
                        }
                        value = bools;
                        ret = true;
                        break;

                    case VARENUM.VT_VARIANT:
                        var variants = new object?[count];
                        var variantSize = Size;
                        for (var i = 0; i < variants.Length; i++)
                        {
                            var pv = ptr + Size * i;
                            using var v = new Variant { _inner = *(VARIANT*)pv };
                            variants[i] = v.Value;
                            v.Detach();
                        }
                        value = variants;
                        ret = true;
                        break;

                    case VARENUM.VT_I1:
                    case VARENUM.VT_UI1:
                    case VARENUM.VT_I2:
                    case VARENUM.VT_UI2:
                    case VARENUM.VT_I4:
                    case VARENUM.VT_INT:
                    case VARENUM.VT_UI4:
                    case VARENUM.VT_UINT:
                    case VARENUM.VT_I8:
                    case VARENUM.VT_UI8:
                    case VARENUM.VT_R4:
                    case VARENUM.VT_R8:
                    case VARENUM.VT_ERROR:
                    case VARENUM.VT_CY:
                    case VARENUM.VT_DATE:
                    case VARENUM.VT_UNKNOWN:
                    case VARENUM.VT_DISPATCH:
                        var et = FromType(vt);
#pragma warning disable IDE0079 // Remove unnecessary suppression
#pragma warning disable IL3050 // Calling members annotated with 'RequiresDynamicCodeAttribute' may break functionality when AOT compiling.
                        var values = Array.CreateInstance(et, psa->cbElements);
#pragma warning disable CA1421 // This method uses runtime marshalling even when the 'DisableRuntimeMarshallingAttribute' is applied
                        size = (uint)(values.Length * Marshal.SizeOf(et));
#pragma warning restore CA1421 // This method uses runtime marshalling even when the 'DisableRuntimeMarshallingAttribute' is applied
#pragma warning restore IL3050 // Calling members annotated with 'RequiresDynamicCodeAttribute' may break functionality when AOT compiling.
#pragma warning restore IDE0079 // Remove unnecessary suppression
                        Functions.CopyMemory(Marshal.UnsafeAddrOfPinnedArrayElement(values, 0), ptr, (nint)size);
                        value = values;
                        ret = true;
                        break;
                }
                return ret;
            }
            finally
            {
                Functions.SafeArrayUnaccessData(*psa).ThrowOnError();
            }
        }
    }

    public void Clear(bool throwOnError = true) => Functions.VariantClear(ref _inner).ThrowOnError(throwOnError);

    public static FILETIME ToPositiveFILETIME(DateTime dt) => ToFILETIME(ToPositiveFileTime(dt));
    public static FILETIME ToFILETIME(long ft) => ToFILETIME((ulong)ft);
    public static FILETIME ToFILETIME(ulong ft) => new() { dwLowDateTime = (uint)(ft & 0xFFFFFFFF), dwHighDateTime = (uint)(ft >> 32) };
    public static long ToPositiveFileTime(DateTime dt)
    {
        var ft = dt.ToUniversalTime().ToFileTimeUtc();
        return ft < 0 ? 0 : ft;
    }
}
