export type { ComBridge } from "./bridge.ts";
export * from "./types.ts";
export {
  ComError,
  RETRYABLE_KINDS,
  asciiJson,
  classifyHresult,
  comErrorFromAnswer,
  connectionLost,
  errorText,
  hresult,
  isComError,
  timeoutError,
} from "./errors.ts";
export { PowerShellComBridge } from "./powershell.ts";
export { FakeComBridge } from "./fake.ts";
export { OUTLOOK_HELPER_SCRIPT } from "./script.ts";
