# -*- coding: utf-8 -*-
"""push_via_api.py - 通过 GitHub REST API (api.github.com) 推送提交, 绕过被拦截的 git 协议"""
import base64
import json
import os
import sys
import urllib.request

TOKEN = os.environ["GH_TOKEN"]
REPO = "mew5748-alt/BatchPrinter"
API = f"https://api.github.com/repos/{REPO}"
HDRS = {
    "Authorization": f"Bearer {TOKEN}",
    "Accept": "application/vnd.github+json",
    "Content-Type": "application/json",
}


def req(method, url, data=None, raw=False, retries=4):
    import time
    last = None
    for i in range(retries):
        try:
            r = urllib.request.Request(url, method=method,
                                       data=json.dumps(data).encode() if data is not None else None,
                                       headers=HDRS)
            with urllib.request.urlopen(r, timeout=120) as resp:
                body = resp.read()
            if raw:
                return body
            return json.loads(body) if body else {}
        except urllib.error.HTTPError:
            raise
        except (urllib.error.URLError, OSError, TimeoutError) as e:
            last = e
            print(f"  网络抖动, 重试 {i+1}/{retries}: {e}")
            time.sleep(2 + i * 2)
    raise last


def main():
    root = os.path.dirname(os.path.abspath(__file__))
    # 直接读 git 已跟踪文件, 避免手写清单漏文件 (v1.2~v1.4 新增的 office_core/preview_ui 等)
    import subprocess
    out = subprocess.run(["git", "ls-files"], cwd=root, capture_output=True,
                         text=True, check=True).stdout
    files = [l.strip() for l in out.splitlines() if l.strip()]
    msg = sys.argv[1] if len(sys.argv) > 1 else "sync"

    # 仓库是否为空 (无 main 分支)
    try:
        ref = req("GET", f"{API}/git/ref/heads/main")
        base_sha = ref["object"]["sha"]
        print("远端已有 main, 基于其提交")
    except urllib.error.HTTPError as e:
        if e.code in (404, 409):   # 404=分支不存在, 409=空仓库
            base_sha = None
            print("空仓库, 先用 Contents API 播种...")
            with open(os.path.join(root, "README.md"), "rb") as fh:
                readme_b64 = base64.b64encode(fh.read()).decode()
            req("PUT", f"{API}/contents/README.md",
                {"message": "init", "content": readme_b64})
            ref = req("GET", f"{API}/git/ref/heads/main")
            base_sha = ref["object"]["sha"]
            print("  播种完成, base:", base_sha[:10])
        else:
            raise

    # 1. blobs
    tree = []
    for f in files:
        p = os.path.join(root, f)
        with open(p, "rb") as fh:
            content = base64.b64encode(fh.read()).decode()
        blob = req("POST", f"{API}/git/blobs", {"content": content, "encoding": "base64"})
        tree.append({"path": f, "mode": "100644", "type": "blob", "sha": blob["sha"]})
        print(f"  blob {f}")

    # 2. tree
    t = req("POST", f"{API}/git/trees", {"tree": tree})
    print("tree:", t["sha"][:10])

    # 3. commit
    commit_data = {"message": msg, "tree": t["sha"]}
    if base_sha:
        commit_data["parents"] = [base_sha]
    c = req("POST", f"{API}/git/commits", commit_data)
    print("commit:", c["sha"][:10])

    # 4. ref
    if base_sha:
        req("PATCH", f"{API}/git/refs/heads/main", {"sha": c["sha"], "force": False})
    else:
        req("POST", f"{API}/git/refs", {"ref": "refs/heads/main", "sha": c["sha"]})
    print("已推送到 main")


if __name__ == "__main__":
    main()
