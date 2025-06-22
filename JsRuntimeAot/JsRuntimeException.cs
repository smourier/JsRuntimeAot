namespace JsRt;

[Serializable]
public partial class JsRuntimeException : Exception
{
    internal JsRuntimeException(JsErrorCode code, JsValue error)
        : base(GetMessage(code, error))
    {
        Code = code;
        if (error != null)
        {
            Line = error.GetProperty("line", -1);
            Column = error.GetProperty("column", -1);
            SourceCode = error.GetProperty<string?>("source", null) ?? string.Empty;
        }
    }

    public JsRuntimeException()
        : base("JsRuntime error")
    {
    }

    public JsRuntimeException(string message)
        : base(message)
    {
    }

    public JsRuntimeException(string message, Exception innerException)
        : base(message, innerException)
    {
    }

    public JsRuntimeException(Exception innerException)
        : base(null, innerException)
    {
    }

    public JsErrorCode Code { get; }
    public int Line { get; } = -1;
    public int Column { get; } = -1;
    public string SourceCode { get; } = string.Empty;

    public static string GetErrorText(JsErrorCode error) => error switch
    {
        JsErrorCode.JsErrorInvalidArgument => "An argument to a hosting API was invalid.",
        JsErrorCode.JsErrorNullArgument => "An argument to a hosting API was null in a context where null is not allowed.",
        JsErrorCode.JsErrorNoCurrentContext => "The hosting API requires that a context be current, but there is no current context.",
        JsErrorCode.JsErrorInExceptionState => "The engine is in an exception state and no APIs can be called until the exception is cleared.",
        JsErrorCode.JsErrorNotImplemented => "A hosting API is not yet implemented.",
        JsErrorCode.JsErrorWrongThread => "A hosting API was called on the wrong thread.",
        JsErrorCode.JsErrorRuntimeInUse => "A runtime that is still in use cannot be disposed.",
        JsErrorCode.JsErrorBadSerializedScript => "A bad serialized script was used, or the serialized script was serialized by a different version of the Chakra engine.",
        JsErrorCode.JsErrorInDisabledState => "The runtime is in a disabled state.",
        JsErrorCode.JsErrorCannotDisableExecution => "Runtime does not support reliable script interruption.",
        JsErrorCode.JsErrorHeapEnumInProgress => "A heap enumeration is currently underway in the script context.",
        JsErrorCode.JsErrorArgumentNotObject => "A hosting API that operates on object values was called with a non-object value.",
        JsErrorCode.JsErrorInProfileCallback => "A script context is in the middle of a profile callback.",
        JsErrorCode.JsErrorInThreadServiceCallback => "A thread service callback is currently underway.",
        JsErrorCode.JsErrorCannotSerializeDebugScript => "Scripts cannot be serialized in debug contexts.",
        JsErrorCode.JsErrorAlreadyDebuggingContext => "The context cannot be put into a debug state because it is already in a debug state.",
        JsErrorCode.JsErrorAlreadyProfilingContext => "The context cannot start profiling because it is already profiling.",
        JsErrorCode.JsErrorIdleNotEnabled => "Idle notification given when the host did not enable idle processing.",
        JsErrorCode.JsErrorOutOfMemory => "The Chakra engine has run out of memory.",
        JsErrorCode.JsErrorScriptException => "A JavaScript exception occurred while running a script.",
        JsErrorCode.JsErrorScriptCompile => "JavaScript failed to compile.",
        JsErrorCode.JsErrorScriptTerminated => "A script was terminated due to a request to suspend a runtime.",
        JsErrorCode.JsErrorScriptEvalDisabled => "A script was terminated because it tried to use 'eval' or 'function' and eval was disabled.",
        JsErrorCode.JsErrorFatal => "A fatal error in the engine has occurred.",
        JsErrorCode.JsNoError => "Success error code.",
        JsErrorCode.JsErrorCategoryUsage => "Category of errors that relates to incorrect usage of the API itself.",
        JsErrorCode.JsCannotSetProjectionEnqueueCallback => "The context did not accept the enqueue callback.",
        JsErrorCode.JsErrorCannotStartProjection => "Failed to start projection.",
        JsErrorCode.JsErrorInObjectBeforeCollectCallback => "The operation is not supported in an object before collect callback.",
        JsErrorCode.JsErrorObjectNotInspectable => "Object cannot be unwrapped to IInspectable pointer.",
        JsErrorCode.JsErrorPropertyNotSymbol => "A hosting API that operates on symbol property ids but was called with a non-symbol property id. The error code is returned by JsGetSymbolFromPropertyId if the function is called with non-symbol property id.",
        JsErrorCode.JsErrorPropertyNotString => "A hosting API that operates on string property ids but was called with a non-string property id. The error code is returned by existing JsGetPropertyNamefromId if the function is called with non-string property id.",
        JsErrorCode.JsErrorCategoryEngine => "Category of errors that relates to errors occurring within the engine itself.",
        JsErrorCode.JsErrorCategoryScript => "Category of errors that relates to errors in a script.",
        JsErrorCode.JsErrorCategoryFatal => "Category of errors that are fatal and signify failure of the engine",
        JsErrorCode.JsErrorWrongRuntime => "A hosting API was called with object created on different javascript runtime.",
        _ => "An unknown error in the engine has occurred.",
    };

    private static string GetMessage(JsErrorCode code, JsValue error)
    {
        var text = new StringBuilder(GetErrorText(code));
        if (error != null)
        {
            var line = error.GetProperty("line", -1);
            var column = error.GetProperty("column", -1);
            var source = error.GetProperty<string?>("source", null);
            var errorText = error.GetProperty<string?>("message", null);

            if (!string.IsNullOrEmpty(errorText))
            {
                text.Append(" " + errorText);
            }

            if (line >= 0)
            {
                text.Append(" at line " + line);
            }

            if (column >= 0)
            {
                text.Append(", column " + column);
            }

            if (!string.IsNullOrEmpty(source))
            {
                text.Append(", in text \"" + source + "\"");
            }
        }
        return text.ToString();
    }
}
