# -*- coding: utf-8 -*-
"""课表排空工具 —— 本地网页服务（手机/电脑同网访问）"""
import os, sys, json, uuid, shutil, zipfile, threading, traceback
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse, parse_qs

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(HERE, "vendor"))

import pipeline

ROOT = os.path.dirname(HERE)
JOBS = os.environ.get("JOBS_DIR") or os.path.join(ROOT, "jobs")
os.makedirs(JOBS, exist_ok=True)
STATIC = os.path.join(HERE, "static")

LOCK = threading.Lock()
DATA_CACHE = {}   # job_id -> 已解析的成员数据，避免生成时再解析一遍
INDEX_CACHE = {}  # job_id -> 查询索引（成员资料 + 排班明细）



def job_dir(jid):
    return os.path.join(JOBS, jid)


class Handler(BaseHTTPRequestHandler):
    server_version = "TimetableTool/1.0"

    def log_message(self, fmt, *args):
        pass

    def _json(self, obj, code=200):
        data = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _file(self, path, ctype, download_name=None):
        if not os.path.exists(path):
            self._json({"error": "文件不存在"}, 404)
            return
        with open(path, "rb") as f:
            data = f.read()
        self.send_response(200)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        if download_name:
            from urllib.parse import quote
            self.send_header("Content-Disposition",
                             "attachment; filename*=UTF-8''%s" % quote(download_name))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        u = urlparse(self.path)
        if u.path in ("/", "/index.html"):
            self._file(os.path.join(STATIC, "index.html"), "text/html; charset=utf-8")
            return
        if u.path == "/api/download":
            q = parse_qs(u.query)
            jid = (q.get("job") or [""])[0]
            kind = (q.get("kind") or [""])[0]
            names = {"empty": "空课表.zip", "duty": "值班表.zip"}
            if kind in names:
                self._file(os.path.join(job_dir(jid), names[kind]), "application/zip", names[kind])
                return
            if kind == "member":
                kw = (q.get("q") or [""])[0].strip()
                dept = (q.get("dept") or [""])[0].strip()
                idx = INDEX_CACHE.get(jid, [])
                hits = []
                for x in idx:
                    if dept and x["dept"] != dept:
                        continue
                    if kw:
                        if kw in x["name"] or kw in x["klass"] or kw in x["phone"]:
                            hits.append(x)
                    elif dept:
                        hits.append(x)
                hits.sort(key=lambda x: (x["name"] != kw, x["name"]))
                self._json({"matches": hits[:40], "depts": sorted({x["dept"] for x in idx})})
                return
            if kind == "pdf":
                name = (q.get("name") or [""])[0]
                idx = INDEX_CACHE.get(jid, [])
                hit = next((x for x in idx if x["name"] == name and x["pdf"]), None)
                if not hit:
                    self._json({"error": "没有这个人的课表文件"}, 404)
                    return
                base = os.path.abspath(os.path.join(job_dir(jid), "src"))
                full = os.path.abspath(os.path.join(base, hit["pdf"]))
                if not full.startswith(base) or not os.path.exists(full):
                    self._json({"error": "文件不存在"}, 404)
                    return
                self._file(full, "application/pdf", os.path.basename(full))
                return
            if kind == "detail":
                self._file(
                    os.path.join(job_dir(jid), "解析明细.xlsx"),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "解析明细.xlsx")
                return
            self._json({"error": "参数错误"}, 400)
            return
        self._json({"error": "not found"}, 404)

    def do_POST(self):
        u = urlparse(self.path)
        try:
            q = parse_qs(u.query)
            defer = (q.get("defer") or ["0"])[0] == "1"
            if u.path == "/api/upload":
                self.handle_upload(defer=defer)
            elif u.path == "/api/add":
                self.handle_upload((q.get("job") or [""])[0], defer=defer)
            elif u.path == "/api/check":
                self.handle_check((q.get("job") or [""])[0])
            elif u.path == "/api/generate":
                self.handle_generate()
            else:
                self._json({"error": "not found"}, 404)
        except Exception:
            traceback.print_exc()
            self._json({"error": "服务器出错：" + traceback.format_exc(limit=1)}, 500)

    MAX_UPLOAD = 120 * 1024 * 1024

    def handle_upload(self, jid=None, defer=False):
        n = int(self.headers.get("Content-Length") or 0)
        if n <= 0:
            self._json({"error": "没有收到文件"}, 400)
            return
        if n > self.MAX_UPLOAD:
            self._json({"error": "压缩包太大了（上限 120MB）"}, 413)
            return
        data = self.rfile.read(n)
        raw_name = self.headers.get("X-Filename") or ""
        from urllib.parse import unquote
        fname = unquote(raw_name).strip().replace("/", "_").replace("\\", "_")
        is_sheet = fname.lower().endswith((".xlsx", ".xlsm"))
        new_job = jid is None or not jid
        if new_job:
            jid = uuid.uuid4().hex[:12]
        d = job_dir(jid)
        os.makedirs(d, exist_ok=True)
        src = os.path.join(d, "src")
        os.makedirs(src, exist_ok=True)

        with LOCK:
            meta_path = os.path.join(d, "meta.json")
            meta = {"zips": [], "order": []}
            if os.path.exists(meta_path):
                try:
                    meta = json.load(open(meta_path, encoding="utf-8"))
                except Exception:
                    pass
            if is_sheet:
                # 第二种通道：直接传做好的空课表 xlsx
                target = os.path.join(src, fname or ("空课表_%02d.xlsx" % (len(meta["zips"]) + 1)))
                with open(target, "wb") as f:
                    f.write(data)
                meta["zips"].append(fname)
                meta["order"].append(os.path.basename(target))
                meta.setdefault("mode", "sheet")
            else:
                zpath = os.path.join(d, "upload_%02d.zip" % (len(meta["zips"]) + 1))
                with open(zpath, "wb") as f:
                    f.write(data)
                try:
                    with zipfile.ZipFile(zpath) as z:
                        z.extractall(src)
                        names = [os.path.basename(x) for x in z.namelist()
                                 if x.lower().endswith((".pdf", ".xlsx", ".xlsm"))]
                except zipfile.BadZipFile:
                    os.remove(zpath)
                    if new_job:
                        shutil.rmtree(d, ignore_errors=True)
                    self._json({"error": "这不是有效的 ZIP 压缩包（如果要传空课表，请直接选 .xlsx 文件）"}, 400)
                    return
                meta["zips"].append(os.path.basename(zpath))
                meta["order"].extend(names)
            with open(meta_path, "w", encoding="utf-8") as f:
                json.dump(meta, f, ensure_ascii=False)

        DATA_CACHE.pop(jid, None)   # 有新文件进来，缓存作废
        if defer:
            # 先把文件收下，等全部传完再统一解析一次（避免每个包都重跑一遍检查）
            self._json({"job": jid, "files": meta["zips"], "deferred": True})
            return

        report, data = pipeline.check(src, meta["order"])
        DATA_CACHE[jid] = data
        report["job"] = jid
        report["files"] = meta["zips"]
        self._json(report)

    def handle_check(self, jid):
        d = job_dir(jid)
        src = os.path.join(d, "src")
        meta_path = os.path.join(d, "meta.json")
        if not os.path.isdir(src) or not os.path.exists(meta_path):
            self._json({"error": "任务不存在，请重新上传"}, 400)
            return
        meta = json.load(open(meta_path, encoding="utf-8"))
        report, data = pipeline.check(src, meta.get("order"))
        DATA_CACHE[jid] = data
        report["job"] = jid
        report["files"] = meta.get("zips", [])
        self._json(report)

    def handle_generate(self):
        n = int(self.headers.get("Content-Length") or 0)
        body = json.loads(self.rfile.read(n) or b"{}")
        jid = body.get("job")
        w1 = int(body.get("start") or 7)
        w2 = int(body.get("end") or 18)
        if w1 > w2:
            w1, w2 = w2, w1
        d = job_dir(jid)
        src = os.path.join(d, "src")
        if not os.path.isdir(src):
            self._json({"error": "任务不存在，请重新上传"}, 400)
            return

        meta = {}
        meta_path = os.path.join(d, "meta.json")
        if os.path.exists(meta_path):
            try:
                meta = json.load(open(meta_path, encoding="utf-8"))
            except Exception:
                meta = {}
        mode, data = pipeline.load_any(src, meta.get("order"))
        if jid in DATA_CACHE:
            data = DATA_CACHE[jid]
        out_empty = os.path.join(d, "空课表")
        out_duty = os.path.join(d, "值班表")
        for p in (out_empty, out_duty):
            shutil.rmtree(p, ignore_errors=True)
            os.makedirs(p, exist_ok=True)

        warnings = []
        if mode == "pdf":
            for dept, members in data.items():
                if not members:
                    continue
                pipeline.build_empty(pipeline.EMPTY_TPL, dept, members, w1, w2,
                                     os.path.join(out_empty, "%s空课表.xlsx" % dept))

        import history as history_mod
        hist_counts, hist_weeks, hist_detail = history_mod.load_history(src)
        skipped = sorted(w for w in hist_weeks if w1 <= w <= w2)
        if skipped:
            warnings.append("已按上传的值班表跳过：" + "、".join("第%d周" % w for w in skipped)
                            + "（这些周已排的人次已计入后续排班）")
        boost = {}
        for k, v in (body.get("boost") or {}).items():
            try:
                n = int(v)
            except (TypeError, ValueError):
                continue
            if n > 0:
                boost[k] = n
        plan, notes, duty_stats = pipeline.plan_duty(
            data, w1, w2, seed_counts=hist_counts, skip_weeks=hist_weeks, boost=boost)
        duty_files = []
        for w in range(w1, w2 + 1):
            if w in hist_weeks:
                continue
            week_plan = plan[w]
            fn = "第%s周系办值班表.xlsx" % pipeline.cn_num(w)
            pipeline.build_duty(pipeline.DUTY_TPL, w, week_plan, notes.get(w, []),
                                os.path.join(out_duty, fn))
            duty_files.append(fn)
            for msg in notes.get(w, []):
                warnings.append("第%s周：%s" % (pipeline.cn_num(w), msg))

        counts = {}
        for wk in plan.values():
            for lst in wk.values():
                for p in lst:
                    counts[p.name] = counts.get(p.name, 0) + 1
        # 每个人的排班明细（供"查人"和解析明细使用）
        assign = {}
        # 先放上传的值班表里已经排好的（标记为已有）
        for w, slots in (hist_detail or {}).items():
            for (blk, day), names in slots.items():
                for n in names:
                    assign.setdefault(n, []).append(
                        {"week": w, "day": day, "block": blk, "history": True})
        for w in range(w1, w2 + 1):
            for (blk, day), lst in plan.get(w, {}).items():
                for p in lst:
                    assign.setdefault(p.name, []).append(
                        {"week": w, "day": day, "block": blk})

        want_detail = bool(body.get("detail"))
        detail_path = os.path.join(d, "解析明细.xlsx")
        if want_detail:
            pipeline.build_detail(data, w1, w2, detail_path, counts, assign)
        idx = []
        for dept, members in data.items():
            for m in members:
                rel = ""
                try:
                    if m.pdf_path and os.path.exists(m.pdf_path):
                        rel = os.path.relpath(m.pdf_path, src).replace("\\", "/")
                except Exception:
                    rel = ""
                idx.append({
                    "name": m.name, "klass": m.klass, "phone": m.phone,
                    "dept": dept, "minister": bool(m.is_minister),
                    "sid": getattr(m, "sid", ""), "pdf": rel,
                    "shifts": assign.get(m.name, []),
                })
        INDEX_CACHE[jid] = idx

        folders = [(out_duty, "值班表.zip", duty_files)]
        if mode == "pdf":
            folders.insert(0, (out_empty, "空课表.zip", sorted(os.listdir(out_empty))))
        for folder, zname, names in folders:
            zp = os.path.join(d, zname)
            with zipfile.ZipFile(zp, "w", zipfile.ZIP_DEFLATED) as zf:
                for f in names:
                    zf.write(os.path.join(folder, f), f)

        self._json({
            "job": jid,
            "mode": mode,
            "weeks": "%d-%d" % (w1, w2),
            "depts": list(data.keys()),
            "empty": ("/api/download?job=%s&kind=empty" % jid) if mode == "pdf" else None,
            "duty": "/api/download?job=%s&kind=duty" % jid,
            "detail": ("/api/download?job=%s&kind=detail" % jid) if want_detail else None,
            "skipped_weeks": skipped,
            "warnings": warnings[:50],
            "duty_weeks": len(duty_files),
            "stats": duty_stats,
        })


def cleanup_old_jobs(days=7):
    import time as _t
    now = _t.time()
    for name in os.listdir(JOBS):
        p = os.path.join(JOBS, name)
        try:
            if os.path.isdir(p) and now - os.path.getmtime(p) > days * 86400:
                shutil.rmtree(p, ignore_errors=True)
        except Exception:
            pass


def lan_ip():
    import socket
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("8.8.8.8", 80))
        return s.getsockname()[0]
    except Exception:
        return "127.0.0.1"
    finally:
        s.close()


def main():
    port = int(os.environ.get("PORT") or 7860)
    cleanup_old_jobs()
    srv = ThreadingHTTPServer(("0.0.0.0", port), Handler)
    print("本机访问: http://127.0.0.1:%d" % port, flush=True)
    print("手机访问: http://%s:%d  (同一 WiFi)" % (lan_ip(), port), flush=True)
    srv.serve_forever()


if __name__ == "__main__":
    main()
