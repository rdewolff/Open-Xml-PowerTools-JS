export class OpenXmlPowerToolsError extends Error {
  constructor(code, message, details) {
    super(message);
    this.name = "OpenXmlPowerToolsError";
    this.code = code;
    this.details = details;
  }
}

