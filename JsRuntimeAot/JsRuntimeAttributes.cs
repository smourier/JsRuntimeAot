namespace JsRt;

[Flags]
public enum JsRuntimeAttributes
{
    JsRuntimeAttributeNone = 0x00000000,
    JsRuntimeAttributeDisableBackgroundWork = 0x00000001,
    JsRuntimeAttributeAllowScriptInterrupt = 0x00000002,
    JsRuntimeAttributeEnableIdleProcessing = 0x00000004,
    JsRuntimeAttributeDisableNativeCodeGeneration = 0x00000008,
    JsRuntimeAttributeDisableEval = 0x00000010,
    JsRuntimeAttributeEnableExperimentalFeatures = 0x00000020,
    JsRuntimeAttributeDispatchSetExceptionsToDebugger = 0x00000040,
    JsRuntimeAttributeDisableFatalOnOOM = 0x00000080,
    JsRuntimeAttributeDisableExecutablePageAllocation = 0x00000100,
}