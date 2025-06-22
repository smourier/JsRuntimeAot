namespace JsRt.Interop;

internal partial struct SAFEARRAY
{
    public ushort cDims;
    public ushort fFeatures;
    public uint cbElements;
    public uint cLocks;
    public nint pvData;
    public SAFEARRAYBOUND rgsabound; // variable-length array placeholder
}
