export class HttpError extends Error {
  constructor(message, status = 500, retryable = false) {
    super(message);
    this.name = "HttpError";
    this.status = status;
    this.retryable = retryable;
  }
}
