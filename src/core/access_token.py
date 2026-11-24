"""Optional shared-secret guard: set MCP_ACCESS_TOKEN and call /mcp?token=<value>."""

import hmac
import os
from urllib.parse import parse_qs


class AccessTokenMiddleware:
    def __init__(self, app):
        self.app = app
        self.token = os.environ.get("MCP_ACCESS_TOKEN", "").strip()

    async def __call__(self, scope, receive, send):
        if self.token and scope["type"] == "http" and scope.get("path", "").startswith("/mcp"):
            if scope.get("method") != "OPTIONS":
                qs = parse_qs(scope.get("query_string", b"").decode())
                given = (qs.get("token") or [""])[0]
                if not hmac.compare_digest(given, self.token):
                    await send({"type": "http.response.start", "status": 401,
                                "headers": [(b"content-type", b"text/plain")]})
                    await send({"type": "http.response.body", "body": b"unauthorized"})
                    return
        await self.app(scope, receive, send)
