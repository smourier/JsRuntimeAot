namespace JsRt.Interop;

[GeneratedComInterface, Guid("00020400-0000-0000-c000-000000000046")]
internal partial interface IDispatch
{
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GetTypeInfoCount(out uint pctinfo);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GetTypeInfo(uint iTInfo, uint lcid, out nint ppTInfo);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GetIDsOfNames(in Guid riid, [In][MarshalUsing(CountElementName = nameof(cNames))] PWSTR[] rgszNames, uint cNames, uint lcid, [In][Out][MarshalUsing(CountElementName = nameof(cNames))] int[] rgDispId);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT Invoke(int dispIdMember, in Guid riid, uint lcid, DISPATCH_FLAGS wFlags, in DISPPARAMS pDispParams, nint /* optional VARIANT* */ pVarResult, nint /* optional EXCEPINFO* */ pExcepInfo, nint /* optional uint* */ puArgErr);
}
