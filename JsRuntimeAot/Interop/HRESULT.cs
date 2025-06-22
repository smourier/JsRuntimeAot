namespace JsRt.Interop;

public partial struct HRESULT(int value) : IEquatable<HRESULT>, IFormattable
{
    public static readonly HRESULT Null = new();

    public int Value = value;

    public override readonly bool Equals(object? obj) => obj is HRESULT value && Equals(value);
    public readonly bool Equals(HRESULT other) => other.Value == Value;
    public override readonly int GetHashCode() => Value.GetHashCode();

    public HRESULT(uint value)
        : this((int)value)
    {
    }

    public readonly uint UValue => (uint)Value;
    public readonly string Name => ToString("n", null);
    public readonly bool IsError => Value < 0;
    public readonly bool IsSuccess => Value >= 0;
    public readonly bool IsOk => Value == 0;
    public readonly bool IsFalse => Value == 1;

    public readonly HRESULT ThrowOnError(bool throwOnError = true)
    {
        if (!throwOnError)
            return Value;

        var exception = GetException();
        if (exception != null)
            throw exception;

        return Value;
    }

    public readonly Exception? GetException()
    {
        if (Value < 0)
            return new Win32Exception(Value);

        return null;
    }

    public override readonly string ToString() => ToString(null, null);
    public readonly string ToString(string? format, IFormatProvider? formatProvider) => (format?.ToLowerInvariant()) switch
    {
        "i" => Value.ToString(),
        "u" => UValue.ToString(),
        _ => "0x" + Value.ToString("X8"),
    };

    public static HRESULT FromWin32(uint error)
    {
        if (error >= 0x80000000)
            return error;

        return FromWin32((int)error);
    }

    public static HRESULT FromWin32(int error)
    {
        if (error < 0)
            return error;

        const int FACILITY_WIN32 = 7;
        return (uint)(error & 0x0000FFFF) | (FACILITY_WIN32 << 16) | 0x80000000;
    }

    public static implicit operator HRESULT(uint result) => new(result);
    public static explicit operator uint(HRESULT hr) => hr.UValue;
    public static bool operator ==(HRESULT left, HRESULT right) => left.Equals(right);
    public static bool operator !=(HRESULT left, HRESULT right) => !left.Equals(right);
    public static implicit operator int(HRESULT value) => value.Value;
    public static implicit operator HRESULT(int value) => new(value);
}
