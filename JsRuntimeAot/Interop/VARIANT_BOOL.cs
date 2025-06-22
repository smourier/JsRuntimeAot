namespace JsRt.Interop;

public struct VARIANT_BOOL(short value) : IEquatable<VARIANT_BOOL>
{
    public static readonly VARIANT_BOOL Null = new();

    public short Value = value;

    public override readonly string ToString() => $"0x{Value:x}";

    public override readonly bool Equals(object? obj) => obj is VARIANT_BOOL value && Equals(value);
    public readonly bool Equals(VARIANT_BOOL other) => other.Value == Value;
    public override readonly int GetHashCode() => Value.GetHashCode();
    public static bool operator ==(VARIANT_BOOL left, VARIANT_BOOL right) => left.Equals(right);
    public static bool operator !=(VARIANT_BOOL left, VARIANT_BOOL right) => !left.Equals(right);
    public static implicit operator short(VARIANT_BOOL value) => value.Value;
    public static implicit operator VARIANT_BOOL(short value) => new(value);
}
