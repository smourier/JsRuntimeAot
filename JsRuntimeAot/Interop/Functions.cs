namespace JsRt.Interop;

internal static partial class Functions
{
    [LibraryImport("OLEAUT32")]
    [PreserveSig]
    public static partial HRESULT VariantChangeType(ref VARIANT pvargDest, in VARIANT pvarSrc, VAR_CHANGE_FLAGS wFlags, VARENUM vt);

    [LibraryImport("OLEAUT32")]
    [PreserveSig]
    public static partial HRESULT VariantCopy(ref VARIANT pvargDest, in VARIANT pvargSrc);

    [LibraryImport("OLEAUT32")]
    [PreserveSig]
    public static partial HRESULT VariantClear(ref VARIANT pvarg);

    [LibraryImport("PROPSYS")]
    [PreserveSig]
    public static partial HRESULT InitVariantFromFileTime(in FILETIME pft, out VARIANT pvar);

    [LibraryImport("kernel32", EntryPoint = "RtlZeroMemory")]
    [PreserveSig]
    public static partial void ZeroMemory(nint pdst, nint cb);

    [LibraryImport("kernel32", EntryPoint = "RtlMoveMemory")]
    [PreserveSig]
    public static partial void CopyMemory(nint pdst, nint psrc, nint cb);

    [LibraryImport("OLEAUT32")]
    [PreserveSig]
    public static partial HRESULT SafeArrayAccessData(in SAFEARRAY psa, out nint ppvData);

    [LibraryImport("OLEAUT32")]
    [PreserveSig]
    public static partial nint SafeArrayCreate(VARENUM vt, uint cDims, in SAFEARRAYBOUND rgsabound);

    [LibraryImport("OLEAUT32")]
    [PreserveSig]
    public static partial HRESULT SafeArrayDestroy(in SAFEARRAY psa);

    [LibraryImport("OLEAUT32")]
    [PreserveSig]
    public static partial HRESULT SafeArrayUnaccessData(in SAFEARRAY psa);

    [LibraryImport("OLEAUT32")]
    [PreserveSig]
    public static partial HRESULT SafeArrayGetLBound(in SAFEARRAY psa, uint nDim, out int plLbound);

    [LibraryImport("OLEAUT32")]
    [PreserveSig]
    public static partial HRESULT SafeArrayGetUBound(in SAFEARRAY psa, uint nDim, out int plUbound);
}
