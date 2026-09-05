#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
publish_gitee_release.py - 把 GitHub Release 镜像到 Gitee（本地直连上传）

背景：Gitee 源码由 sync-to-gitee.yml 自动同步，但 Release 附件走云端 runner
跨境上传 Gitee 稳定失败（见 DEVELOPMENT.md 坑 11）。本脚本在本机运行，
直连 Gitee 上传，绕开跨境链路。

用法:
  # 常规：tag + 本地 exe
  python scripts/publish_gitee_release.py v2.1.0 dist/ExcelSplitter-v2.1.0.exe

  # 省略 exe 路径时，先找 dist/ExcelSplitter-<tag>.exe，找不到再从
  # GitHub Release 下载（私有仓库，token 取自 GITHUB_TOKEN 或 git 凭证管理器）
  python scripts/publish_gitee_release.py v2.1.0

前置条件:
  - 环境变量 GITEE_TOKEN（Gitee 私人令牌，需勾选 projects 权限）
  - Release notes 自动从 GitHub Release 复制（需要 GITHUB_TOKEN 环境变量，
    或本机 git 凭证管理器里存有 github.com 的凭证；取不到则用兜底文案）

踩坑对齐（DEVELOPMENT.md 坑 11）:
  - Gitee OpenAPI 的 access_token 必须放 query 参数，放 JSON body 会 401
  - 创建 Release 必须显式传 target_commitish
"""
import argparse
import json
import os
import sys
import time
import uuid
import urllib.request
import urllib.error

GITEE_API = "https://gitee.com/api/v5"
OWNER_REPO = "wiggins-kong/ExcelSplitter"
GITHUB_API = "https://api.github.com/repos/" + OWNER_REPO
UPLOAD_TIMEOUT = 600  # 大文件上传给足超时


def die(msg, hint=None):
    print("[错误] " + msg)
    if hint:
        print(hint)
    sys.exit(1)


def log(msg):
    print("[发布Gitee] " + msg)


def http_json(method, url, token, payload=None):
    """发 JSON 请求，token 放 query 的 URL 已由调用方拼好。返回 (status, dict)。"""
    data = json.dumps(payload).encode("utf-8") if payload is not None else None
    req = urllib.request.Request(url, data=data, method=method)
    req.add_header("Content-Type", "application/json")
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            return resp.status, json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", "replace")
        try:
            return e.code, json.loads(body)
        except Exception:
            return e.code, {"raw": body}


def github_token():
    """GITHUB_TOKEN 优先，其次从 git 凭证管理器取 github.com 的凭证。"""
    tok = os.environ.get("GITHUB_TOKEN")
    if tok:
        return tok
    try:
        import subprocess
        out = subprocess.run(
            ["git", "credential", "fill"],
            input="protocol=https\nhost=github.com\n",
            capture_output=True, text=True, timeout=15)
        for line in out.stdout.splitlines():
            if line.startswith("password="):
                return line.split("=", 1)[1]
    except Exception:
        pass
    return None


def fetch_github_release(tag, token):
    """取 GitHub Release 的 name / body，失败返回 None（不阻塞发布）。"""
    if not token:
        return None
    req = urllib.request.Request(GITHUB_API + "/releases/tags/" + tag)
    req.add_header("Authorization", "token " + token)
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            rel = json.loads(resp.read().decode("utf-8"))
            return {"name": rel.get("name") or tag, "body": rel.get("body") or ""}
    except Exception as e:
        log("取 GitHub Release 描述失败（不影响继续）: %s" % e)
        return None


def download_github_asset(tag, token, dest):
    """从 GitHub Release 下载 ExcelSplitter-<tag>.exe 到 dest。"""
    req = urllib.request.Request(GITHUB_API + "/releases/tags/" + tag)
    req.add_header("Authorization", "token " + token)
    with urllib.request.urlopen(req, timeout=120) as resp:
        rel = json.loads(resp.read().decode("utf-8"))
    asset = next((a for a in rel.get("assets", [])
                  if a["name"] == "ExcelSplitter-%s.exe" % tag), None)
    if not asset:
        die("GitHub Release %s 上没有 ExcelSplitter-%s.exe 附件" % (tag, tag))
    log("从 GitHub 下载 %s（%.1f MB）..." % (asset["name"], asset["size"] / 1048576))
    req = urllib.request.Request(asset["url"])  # API URL，带 token 走 302
    req.add_header("Authorization", "token " + token)
    req.add_header("Accept", "application/octet-stream")
    with urllib.request.urlopen(req, timeout=600) as resp, open(dest, "wb") as f:
        while True:
            chunk = resp.read(1 << 20)
            if not chunk:
                break
            f.write(chunk)
    return dest


def find_gitee_release(tag, token):
    url = "%s/repos/%s/releases/tags/%s?access_token=%s" % (
        GITEE_API, OWNER_REPO, tag, token)
    status, rel = http_json("GET", url, token)
    if status == 200:
        return rel
    if status == 404:
        return None
    die("查询 Gitee Release 失败（HTTP %d）: %s" % (status, rel),
        hint="401 通常是 token 无效或没勾 projects 权限；token 必须以私人令牌（个人设置→私人令牌）生成")


def create_gitee_release(tag, name, body, token):
    """创建 Gitee Release。坑：access_token 放 query；target_commitish 必传。"""
    url = "%s/repos/%s/releases?access_token=%s" % (GITEE_API, OWNER_REPO, token)
    status, rel = http_json("POST", url, token, payload={
        "tag_name": tag,
        "name": name,
        "body": body,
        "target_commitish": tag,  # tag 已由 git push 同步到 Gitee，直接指过去
        "prerelease": False,
    })
    if status not in (200, 201):
        die("创建 Gitee Release 失败（HTTP %d）: %s" % (status, rel))
    return rel


def update_gitee_release(release_id, name, body, token):
    url = "%s/repos/%s/releases/%s?access_token=%s" % (
        GITEE_API, OWNER_REPO, release_id, token)
    status, rel = http_json("PATCH", url, token, payload={"name": name, "body": body})
    if status not in (200, 201):
        die("更新 Gitee Release 描述失败（HTTP %d）: %s" % (status, rel))


def upload_attachment(release_id, filepath, token, retries=3):
    """multipart/form-data 上传附件，token 放 query。带重试。"""
    filename = os.path.basename(filepath)
    size = os.path.getsize(filepath)
    with open(filepath, "rb") as f:
        data = f.read()
    boundary = "----ExcelSplitterBoundary" + uuid.uuid4().hex
    part = []
    part.append(("--%s\r\n" % boundary).encode())
    part.append(('Content-Disposition: form-data; name="file"; filename="%s"\r\n'
                 % filename).encode("utf-8"))
    part.append(b"Content-Type: application/octet-stream\r\n\r\n")
    part.append(data)
    part.append(("\r\n--%s--\r\n" % boundary).encode())
    body = b"".join(part)

    url = "%s/repos/%s/releases/%s/attach_files?access_token=%s" % (
        GITEE_API, OWNER_REPO, release_id, token)
    last_err = None
    for attempt in range(1, retries + 1):
        log("上传附件（第 %d/%d 次，%.1f MB）..." % (attempt, retries, size / 1048576))
        req = urllib.request.Request(url, data=body, method="POST")
        req.add_header("Content-Type", "multipart/form-data; boundary=" + boundary)
        req.add_header("Content-Length", str(len(body)))
        try:
            with urllib.request.urlopen(req, timeout=UPLOAD_TIMEOUT) as resp:
                resp.read()
            log("附件上传成功")
            return
        except urllib.error.HTTPError as e:
            last_err = "HTTP %d: %s" % (e.code, e.read().decode("utf-8", "replace"))
        except Exception as e:
            last_err = str(e)
        log("第 %d 次失败: %s" % (attempt, last_err))
        time.sleep(3 * attempt)
    die("附件上传失败（已重试 %d 次）" % retries, hint=last_err)


def main():
    parser = argparse.ArgumentParser(description="把 GitHub Release 镜像到 Gitee")
    parser.add_argument("tag", help="版本 tag，如 v2.1.0（源码 tag 须已随 push 同步到 Gitee）")
    parser.add_argument("exe", nargs="?", help="exe 路径；省略则先找 dist/ 再从 GitHub 下载")
    args = parser.parse_args()

    tag = args.tag
    if not tag.startswith("v"):
        die("tag 必须以 v 开头，如 v2.1.0")

    gitee_token = os.environ.get("GITEE_TOKEN")
    if not gitee_token:
        die("未设置 GITEE_TOKEN 环境变量",
            hint="在系统环境变量里配置 Gitee 私人令牌（个人设置→私人令牌→勾选 projects）")

    # ① 定位 exe
    exe = args.exe
    if exe is None:
        local = os.path.join("dist", "ExcelSplitter-%s.exe" % tag)
        if os.path.exists(local):
            exe = local
        else:
            gh_token = github_token()
            if not gh_token:
                die("dist/ 下没有 %s，也没有可用的 GitHub 凭证来自动下载" % local)
            exe = os.path.join("dist", "ExcelSplitter-%s.exe" % tag)
            os.makedirs("dist", exist_ok=True)
            download_github_asset(tag, gh_token, exe)
    if not os.path.exists(exe):
        die("找不到 exe 文件: %s" % exe)
    log("待上传: %s（%.1f MB）" % (exe, os.path.getsize(exe) / 1048576))

    # ② Release notes 从 GitHub 复制（取不到用兜底文案）
    gh_rel = fetch_github_release(tag, github_token())
    name = gh_rel["name"] if gh_rel else "ExcelSplitter %s" % tag
    body = gh_rel["body"] if gh_rel else (
        "ExcelSplitter %s 附件镜像。\n\n完整说明见 GitHub Release：" % tag
        + "\nhttps://github.com/%s/releases/tag/%s" % (OWNER_REPO, tag))

    # ③ 建 Release（已存在则更新描述，避免重复建）
    rel = find_gitee_release(tag, gitee_token)
    if rel:
        log("Gitee 已有 %s 的 Release（id=%s），更新描述" % (tag, rel["id"]))
        update_gitee_release(rel["id"], name, body, gitee_token)
        release_id = rel["id"]
    else:
        log("创建 Gitee Release: %s" % tag)
        release_id = create_gitee_release(tag, name, body, gitee_token)["id"]

    # ④ 传附件；同名附件已存在时提示手动删（Gitee 无删除附件的 API）
    existing = [a.get("name") for a in (rel or {}).get("assets", [])]
    asset_name = os.path.basename(exe)
    if asset_name in existing:
        print("[提示] Gitee Release 上已存在附件 %s，Gitee API 不支持删除附件，"
              "请到网页「发行版→编辑发布」手动删除旧附件后重跑本脚本" % asset_name)
        sys.exit(2)
    upload_attachment(release_id, exe, gitee_token)

    log("完成！查看: https://gitee.com/%s/releases/tag/%s" % (OWNER_REPO, tag))


if __name__ == "__main__":
    main()
