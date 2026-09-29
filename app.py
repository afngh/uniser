"""Local PIN lookup. Run: python app.py  ->  http://localhost:8000
Loads an Excel/CSV sheet into an in-memory dict (O(1) lookups) and reloads
automatically whenever the file changes on disk."""
import csv, json, os, shutil, tempfile, threading, time, webbrowser
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs
from python_calamine import CalamineWorkbook  # fast Rust-based xlsx reader

HERE = os.path.dirname(os.path.abspath(__file__))
CFG_FILE = os.path.join(HERE, "config.json")
UPLOAD_DIR = os.path.join(HERE, "uploads")
S = {"cfg": None, "headers": [], "index": {}, "sig": None, "version": 0,
     "loaded_at": None, "error": None, "loading": False}


def norm(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def resolve_path(path):
    if not path:
        return path
    p = path.strip().strip("\"'")
    if not os.path.isabs(p):
        rel = os.path.join(HERE, p)
        if os.path.exists(rel):
            return rel
    return p


def read_table(path, sheet=None):
    path = resolve_path(path)
    # Read from a temp copy so Excel's file lock never blocks us
    fd, tmp = tempfile.mkstemp(suffix=os.path.splitext(path)[1])
    os.close(fd)
    try:
        shutil.copyfile(path, tmp)
        if path.lower().endswith((".csv", ".txt")):
            with open(tmp, newline="", encoding="utf-8-sig", errors="replace") as f:
                return [], list(csv.reader(f))
        wb = CalamineWorkbook.from_path(tmp)
        sheet_names = wb.sheet_names
        target_sheet = sheet if sheet and sheet in sheet_names else sheet_names[0]
        return sheet_names, wb.get_sheet_by_name(target_sheet).to_python()
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass


def stat_sig(path):
    path = resolve_path(path)
    st = os.stat(path)
    return (st.st_mtime_ns, st.st_size)


def load():
    cfg = S["cfg"]
    if not cfg:
        return
    S["loading"] = True
    try:
        path = resolve_path(cfg["path"])
        sig = stat_sig(path)
        _, rows = read_table(path, cfg.get("sheet"))
        if not rows:
            raise ValueError("File is empty or has no header row.")
        hdr = [norm(h) for h in rows[0]]
        key = cfg["key"]
        if key not in hdr:
            ki = 0
        else:
            ki = hdr.index(key)
        idx = {}
        for r in rows[1:]:
            if ki < len(r):
                k = norm(r[ki])
                if k:
                    idx[k] = r
        S.update(headers=hdr, index=idx, sig=sig, loaded_at=time.time(),
                 version=S["version"] + 1, error=None)  # atomic swap
    except Exception as e:
        S["error"] = f"{type(e).__name__}: {e}"
    finally:
        S["loading"] = False


def watcher():
    prev, last_try = None, 0
    while True:
        time.sleep(1)
        cfg = S["cfg"]
        if not cfg or S["loading"]:
            continue
        try:
            path = resolve_path(cfg["path"])
            sig = stat_sig(path)
        except OSError:
            continue
        # reload once the file has stopped changing (Excel may still be writing)
        if sig != S["sig"] and sig == prev and time.time() - last_try > 3:
            last_try = time.time()
            load()
        prev = sig


def search(pins):
    idx, n = S["index"], len(S["headers"])
    rows, missing = [], []
    for p in dict.fromkeys(norm(x) for x in pins if norm(x)):
        r = idx.get(p)
        if r is None:
            missing.append(p)
        else:
            rows.append([norm(c) for c in r[:n]] + [""] * (n - len(r)))
    return {"headers": S["headers"], "rows": rows, "missing": missing}


class H(BaseHTTPRequestHandler):
    def log_message(self, *a): pass

    def send(self, obj, code=200, ctype="application/json"):
        data = obj if isinstance(obj, bytes) else json.dumps(obj).encode()
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        u = urlparse(self.path)
        if u.path == "/":
            with open(os.path.join(HERE, "index.html"), "rb") as f:
                return self.send(f.read(), ctype="text/html; charset=utf-8")
        if u.path == "/api/status":
            return self.send({"configured": bool(S["cfg"]), "cfg": S["cfg"], "version": S["version"],
                              "count": len(S["index"]), "loaded_at": S["loaded_at"],
                              "loading": S["loading"], "error": S["error"]})
        if u.path == "/api/inspect":
            q = parse_qs(u.query)
            try:
                raw_path = q["path"][0].strip()
                path = resolve_path(raw_path)
                sheets, rows = read_table(path, q.get("sheet", [None])[0])
                if not rows:
                    return self.send({"error": "File is empty or contains no rows."}, 400)
                return self.send({"path": path, "sheets": sheets, "headers": [norm(h) for h in rows[0]]})
            except Exception as e:
                return self.send({"error": f"{type(e).__name__}: {e}"}, 400)
        self.send({"error": "not found"}, 404)

    def do_POST(self):
        if self.path.startswith("/api/upload"):
            u = urlparse(self.path)
            q = parse_qs(u.query)
            filename = q.get("name", ["uploaded_file.xlsx"])[0]
            filename = os.path.basename(filename)
            if not filename:
                filename = "uploaded_file.xlsx"

            os.makedirs(UPLOAD_DIR, exist_ok=True)
            saved_path = os.path.join(UPLOAD_DIR, filename)

            content_len = int(self.headers.get("Content-Length", 0))
            if content_len == 0:
                return self.send({"error": "No file content received"}, 400)

            bytes_left = content_len
            try:
                with open(saved_path, "wb") as f:
                    while bytes_left > 0:
                        chunk = self.rfile.read(min(bytes_left, 65536))
                        if not chunk:
                            break
                        f.write(chunk)
                        bytes_left -= len(chunk)

                sheets, rows = read_table(saved_path)
                if not rows:
                    return self.send({"error": "Uploaded file is empty or invalid."}, 400)
                headers = [norm(h) for h in rows[0]]
                return self.send({
                    "path": saved_path,
                    "filename": filename,
                    "sheets": sheets,
                    "headers": headers
                })
            except Exception as e:
                return self.send({"error": f"Failed to parse uploaded file: {type(e).__name__}: {e}"}, 400)

        body = json.loads(self.rfile.read(int(self.headers.get("Content-Length", 0))) or b"{}")
        if self.path == "/api/config":
            p = resolve_path(body["path"])
            S.update(cfg={"path": p, "sheet": body.get("sheet"), "key": body["key"]},
                     index={}, sig=None, error=None)
            with open(CFG_FILE, "w") as f:
                json.dump(S["cfg"], f)
            threading.Thread(target=load, daemon=True).start()
            return self.send({"ok": True})
        if self.path == "/api/search":
            return self.send(search(body.get("pins", [])))
        self.send({"error": "not found"}, 404)


if __name__ == "__main__":
    if os.path.exists(CFG_FILE):
        try:
            S["cfg"] = json.load(open(CFG_FILE))
            threading.Thread(target=load, daemon=True).start()
        except Exception:
            pass
    threading.Thread(target=watcher, daemon=True).start()
    srv = ThreadingHTTPServer(("127.0.0.1", 8000), H)
    print("Running at http://localhost:8000  (Ctrl+C to stop)")
    webbrowser.open("http://localhost:8000")
    srv.serve_forever()

