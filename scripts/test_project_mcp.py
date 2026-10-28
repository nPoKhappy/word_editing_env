"""Initialize each project-configured MCP and make a bounded read-only tool call."""
import json
import os
from pathlib import Path
import queue
import subprocess
import threading
import time
import tomllib

ROOT = Path(__file__).resolve().parents[1]
REPORT = ROOT/"outputs/chat-replay-20261028/logs/project-mcp.json"


def check(name, config):
    env = {**os.environ, **config.get("env", {})}
    env["PYTHONDONTWRITEBYTECODE"] = "1"
    env["PIP_CACHE_DIR"] = str(ROOT/".tmp/pip-cache")
    proc = subprocess.Popen([config["command"], *config.get("args", [])], cwd=config["cwd"],
                            env=env, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                            text=True, encoding="utf-8", creationflags=subprocess.CREATE_NO_WINDOW)
    lines = queue.Queue()
    errors = []
    def output():
        for line in proc.stdout:
            try:
                lines.put(json.loads(line))
            except ValueError:
                continue
    threading.Thread(target=output, daemon=True).start()
    threading.Thread(target=lambda: errors.append(proc.stderr.read()), daemon=True).start()
    def rpc(identifier, method, params=None):
        req = {"jsonrpc": "2.0", "method": method}
        if identifier is not None:
            req["id"] = identifier
        if params is not None:
            req["params"] = params
        proc.stdin.write(json.dumps(req)+"\n")
        proc.stdin.flush()
        if identifier is None:
            return
        deadline = time.monotonic()+45
        while time.monotonic() < deadline:
            item = lines.get(timeout=max(.1, deadline-time.monotonic()))
            if item.get("id") == identifier:
                if "error" in item:
                    raise RuntimeError(item["error"])
                return item["result"]
        raise TimeoutError(method)
    try:
        initialization = rpc(1, "initialize", {"protocolVersion": "2024-11-05", "capabilities": {},
                                               "clientInfo": {"name": "officecli-project-test", "version": "1.0"}})
        rpc(None, "notifications/initialized")
        tools = rpc(2, "tools/list")["tools"]
        selected = {"officecli": ("officecli", {"command": "help docx paragraph --json"}),
                    "endnote-word-local": ("get_endnote_capabilities", {}),
                    "endnote-library-search": ("search_library", {"query": "Difficult"}),
                    "endnote-reference-check": ("citecheck_version", {})}[name]
        schemas = {t["name"]: t for t in tools}
        assert selected[0] in schemas
        required = schemas[selected[0]]["inputSchema"].get("required", [])
        assert all(k in selected[1] for k in required), (name, required)
        result = rpc(3, "tools/call", {"name": selected[0], "arguments": selected[1]})
        assert not result.get("isError"), result
        return {"server": name, "status": "PASS", "tools": len(tools),
                "call": selected[0], "server_info": initialization.get("serverInfo"),
                "command": config["command"], "cwd": config["cwd"]}
    except Exception as exc:
        return {"server": name, "status": "FAIL", "error": repr(exc), "stderr": "".join(errors)[-3000:]}
    finally:
        if proc.poll() is None:
            proc.terminate()
        proc.wait(timeout=10)


if __name__ == "__main__":
    config = tomllib.loads((ROOT/".codex/config.toml").read_text(encoding="utf-8-sig"))
    reports = []
    for name, settings in config["mcp_servers"].items():
        result = check(name, settings)
        reports.append(result)
        print(json.dumps(result, ensure_ascii=False), flush=True)
    REPORT.write_text(json.dumps(reports, ensure_ascii=False, indent=2), encoding="utf-8")
    raise SystemExit(0 if all(x["status"] == "PASS" for x in reports) else 1)
