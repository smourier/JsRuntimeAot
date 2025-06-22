namespace JsRt.Interop;

internal struct VARIANT
{
    [StructLayout(LayoutKind.Explicit)]
    public struct AnonymousUnion
    {
        public struct AnonymousStruct
        {
            [StructLayout(LayoutKind.Explicit)]
            public struct AnonymousUnion
            {
                public struct AnonymousStruct
                {
                    public nint pvRecord;
                    public nint pRecInfo;
                }

                [FieldOffset(0)]
                public long llVal;

                [FieldOffset(0)]
                public int lVal;

                [FieldOffset(0)]
                public byte bVal;

                [FieldOffset(0)]
                public short iVal;

                [FieldOffset(0)]
                public float fltVal;

                [FieldOffset(0)]
                public double dblVal;

                [FieldOffset(0)]
                public VARIANT_BOOL boolVal;

                [FieldOffset(0)]
                public int scode;

                [FieldOffset(0)]
                public long cyVal;

                [FieldOffset(0)]
                public double date;

                [FieldOffset(0)]
                public BSTR bstrVal;

                [FieldOffset(0)]
                public nint punkVal;

                [FieldOffset(0)]
                public nint pdispVal;

                [FieldOffset(0)]
                public nint parray;

                [FieldOffset(0)]
                public nint pbVal;

                [FieldOffset(0)]
                public nint piVal;

                [FieldOffset(0)]
                public nint plVal;

                [FieldOffset(0)]
                public nint pllVal;

                [FieldOffset(0)]
                public nint pfltVal;

                [FieldOffset(0)]
                public nint pdblVal;

                [FieldOffset(0)]
                public nint pboolVal;

                [FieldOffset(0)]
                public nint __OBSOLETE__VARIANT_PBOOL;

                [FieldOffset(0)]
                public nint pscode;

                [FieldOffset(0)]
                public nint pcyVal;

                [FieldOffset(0)]
                public nint pdate;

                [FieldOffset(0)]
                public nint pbstrVal;

                [FieldOffset(0)]
                public nint ppunkVal;

                [FieldOffset(0)]
                public nint ppdispVal;

                [FieldOffset(0)]
                public nint pparray;

                [FieldOffset(0)]
                public nint pvarVal;

                [FieldOffset(0)]
                public nint byref;

                [FieldOffset(0)]
                public CHAR cVal;

                [FieldOffset(0)]
                public ushort uiVal;

                [FieldOffset(0)]
                public uint ulVal;

                [FieldOffset(0)]
                public ulong ullVal;

                [FieldOffset(0)]
                public int intVal;

                [FieldOffset(0)]
                public uint uintVal;

                [FieldOffset(0)]
                public nint pdecVal;

                [FieldOffset(0)]
                public nint pcVal;

                [FieldOffset(0)]
                public nint puiVal;

                [FieldOffset(0)]
                public nint pulVal;

                [FieldOffset(0)]
                public nint pullVal;

                [FieldOffset(0)]
                public nint pintVal;

                [FieldOffset(0)]
                public nint puintVal;

                [FieldOffset(0)]
                public AnonymousStruct Anonymous;
            }

            public VARENUM vt;
            public ushort wReserved1;
            public ushort wReserved2;
            public ushort wReserved3;
            public AnonymousUnion Anonymous;
        }

        [FieldOffset(0)]
        public AnonymousStruct Anonymous;

        [FieldOffset(0)]
        public decimal decVal;
    }

    public AnonymousUnion Anonymous;
}
