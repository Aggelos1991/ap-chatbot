from fastapi import Request
from fastapi.responses import PlainTextResponse

@app.middleware("http")
async def allow_chunked_requests(request: Request, call_next):
    # If the client uses Transfer-Encoding: chunked (Power Automate does),
    # FastAPI/Uvicorn won't compute content length automatically.
    # This ensures the body is still properly read.
    if request.headers.get("transfer-encoding", "").lower() == "chunked":
        # Force reading the raw body manually
        body = await request.body()
        request._body = body  # inject the raw body back for downstream parsing
    response = await call_next(request)
    return response
