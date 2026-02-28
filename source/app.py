from pathlib import Path
import os
import tempfile
import uuid
import zipfile
import json
from datetime import datetime, timezone

from fastapi import FastAPI, File, UploadFile, HTTPException, Request, Form
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, HttpUrl
import httpx

import excel_book_matcher


BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_DIR = BASE_DIR / "assets"
NEW_TEMPLATE_PATH = TEMPLATE_DIR / "新表模板.xlsx"

# 临时 zip 下载：token -> zip 字节，扣子用短链接代替超长 base64
_download_cache: dict[str, bytes] = {}


class MatchByUrlRequest(BaseModel):
    old_file: HttpUrl
    new_file: HttpUrl


class MatchByUrlResponse(BaseModel):
    status: str
    report_url: str  # 临时下载链接，例如 https://xxx.zeabur.app/download/abc123


app = FastAPI(title="Excel 图书匹配服务")

# Serve shared UI assets (CSS, etc.)
STATIC_DIR = BASE_DIR / "static"
app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


@app.get("/", response_class=HTMLResponse)
async def index() -> str:
    """简单上传页面，方便同事在浏览器里直接使用。"""
    return """
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Excel 图书匹配工具</title>
    <link rel="stylesheet" href="/static/app.css" />
</head>
<body>
    <div class="container">
        <header class="topbar">
            <div class="brand">
                <div class="logo" aria-hidden="true"></div>
                <div>
                    <div class="brand-title">Excel 图书匹配工具</div>
                    <div class="brand-sub">上传新旧表，自动生成匹配结果 ZIP</div>
                </div>
            </div>
            <nav class="nav" aria-label="导航">
                <a href="/" aria-current="page">单次匹配</a>
                <a href="/upload_to_oss">批量上传到 OSS</a>
                <a href="/download_results">下载处理结果</a>
            </nav>
        </header>

        <section class="hero">
            <h1>单次匹配</h1>
            <p>
                选择一份旧表和一份新表，点击“开始匹配”后会自动下载结果压缩包（包含匹配明细、未匹配表、匹配原始表）。
            </p>
        </section>

        <div class="grid">
            <section class="card">
                <div class="card-body">
                    <div class="card-title">上传文件</div>
                    <form class="form" action="/match" method="post" enctype="multipart/form-data">
                        <div class="field">
                            <label class="label" for="old_file">旧表 Excel</label>
                            <input class="control" type="file" id="old_file" name="old_file" accept=".xlsx,.xls" required />
                        </div>
                        <div class="field">
                            <label class="label" for="new_file">新表 Excel</label>
                            <input class="control" type="file" id="new_file" name="new_file" accept=".xlsx,.xls" required />
                        </div>
                        <div class="actions">
                            <button class="btn primary" type="submit">开始匹配并下载结果</button>
                            <a class="btn link" href="/template/new" download>下载新表模板</a>
                        </div>
                        <div class="small muted">支持格式：.xlsx / .xls。处理在服务器端完成。</div>
                    </form>
                </div>
            </section>

            <aside class="card">
                <div class="card-body">
                    <div class="card-title">使用小贴士</div>
                    <div class="callout">
                        <strong>新表格式不确定？</strong> 先下载模板查看列结构，再按模板填写即可。
                        <ul>
                            <li>匹配结果会以 ZIP 形式下载</li>
                            <li>ZIP 中包含 3 张 Excel 结果表</li>
                            <li>遇到报错通常是列名/格式不符合</li>
                        </ul>
                    </div>
                    <div class="footer">如需批量处理，请使用“批量上传到 OSS”。</div>
                </div>
            </aside>
        </div>
    </div>
</body>
</html>
    """


@app.get("/template/new")
async def download_new_template() -> FileResponse:
    """
    下载新表 Excel 模板，供用户了解并填写新表结构。
    """
    if not NEW_TEMPLATE_PATH.exists():
        raise HTTPException(status_code=404, detail="新表模板文件不存在，请联系管理员上传。")

    return FileResponse(
        path=NEW_TEMPLATE_PATH,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="新表模板.xlsx",
    )


@app.post("/match")
async def match_excels(
    old_file: UploadFile = File(...),
    new_file: UploadFile = File(...),
) -> FileResponse:
    """
    接收两份 Excel，调用 excel_book_matcher.run_matching，返回压缩包下载。
    """
    if not old_file.filename or not new_file.filename:
        raise HTTPException(status_code=400, detail="请同时上传旧表和新表 Excel 文件。")

    # 使用持久一点的临时目录存放上传文件和生成结果（不会在函数结束时立刻删除）
    tmpdir = Path(tempfile.mkdtemp())

    old_path = tmpdir / (old_file.filename or "old.xlsx")
    new_path = tmpdir / (new_file.filename or "new.xlsx")

    try:
        old_bytes = await old_file.read()
        new_bytes = await new_file.read()
        old_path.write_bytes(old_bytes)
        new_path.write_bytes(new_bytes)
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(status_code=500, detail=f"保存上传文件失败: {exc}") from exc

    # 结果文件路径（中文文件名）
    matched_path = tmpdir / "已匹配数据表.xlsx"
    unmatched_path = tmpdir / "未匹配数据表.xlsx"
    matched_original_path = tmpdir / "已匹配数据原始表.xlsx"

    try:
        excel_book_matcher.run_matching(
            old_path=old_path,
            new_path=new_path,
            matched_path=matched_path,
            unmatched_path=unmatched_path,
            matched_original_path=matched_original_path,
        )
    except FileNotFoundError as exc:
        raise HTTPException(status_code=400, detail=f"读取文件失败，请检查格式: {exc}") from exc
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception as exc:  # noqa: BLE001
        print("UNEXPECTED_ERROR_IN_MATCH:", repr(exc))
        raise HTTPException(status_code=500, detail=f"处理 Excel 时出错: {exc}") from exc

    # 打包为 zip 返回
    zip_path = tmpdir / "match_results.zip"
    try:
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.write(matched_path, arcname=matched_path.name)
            zf.write(unmatched_path, arcname=unmatched_path.name)
            zf.write(matched_original_path, arcname=matched_original_path.name)
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(status_code=500, detail=f"打包结果文件失败: {exc}") from exc

    return FileResponse(
        path=zip_path,
        media_type="application/zip",
        filename="match_results.zip",
    )


@app.get("/download/{token}")
async def download_report(token: str) -> Response:
    """
    根据 match_by_url 返回的临时 token 下载 zip，仅可下载一次。
    """
    zip_bytes = _download_cache.pop(token, None)
    if zip_bytes is None:
        raise HTTPException(status_code=404, detail="链接已失效或已下载过，请重新发起匹配。")
    return Response(
        content=zip_bytes,
        media_type="application/zip",
        headers={"Content-Disposition": 'attachment; filename="match_results.zip"'},
    )


@app.post("/match_by_url")
async def match_by_url(request: Request, payload: MatchByUrlRequest) -> JSONResponse:
    """
    扣子等场景使用：传入两个 Excel 的下载 URL，由服务端拉取后再跑匹配。
    返回 JSON，report_url 为临时下载链接（短链接，避免超长 base64 导致扣子报错）。
    """
    tmpdir = Path(tempfile.mkdtemp())

    old_path = tmpdir / "old.xlsx"
    new_path = tmpdir / "new.xlsx"

    # 扣子等 CDN 可能要求浏览器式请求头才允许下载
    download_headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "application/octet-stream,*/*",
    }

    try:
        async with httpx.AsyncClient(
            timeout=60.0,
            follow_redirects=True,
            headers=download_headers,
        ) as client:
            old_resp = await client.get(str(payload.old_file))
            old_resp.raise_for_status()
            new_resp = await client.get(str(payload.new_file))
            new_resp.raise_for_status()

        old_path.write_bytes(old_resp.content)
        new_path.write_bytes(new_resp.content)
    except httpx.HTTPStatusError as exc:
        msg = f"下载 Excel 失败: HTTP {exc.response.status_code}"
        if exc.response.text:
            msg += f" - {exc.response.text[:200]}"
        raise HTTPException(status_code=400, detail=msg) from exc
    except httpx.RequestError as exc:
        raise HTTPException(status_code=400, detail=f"下载 Excel 请求失败: {exc!s}") from exc
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(status_code=400, detail=f"下载 Excel 文件失败: {exc!s}") from exc

    matched_path = tmpdir / "已匹配数据表.xlsx"
    unmatched_path = tmpdir / "未匹配数据表.xlsx"
    matched_original_path = tmpdir / "已匹配数据原始表.xlsx"

    try:
        excel_book_matcher.run_matching(
            old_path=old_path,
            new_path=new_path,
            matched_path=matched_path,
            unmatched_path=unmatched_path,
            matched_original_path=matched_original_path,
        )
    except FileNotFoundError as exc:
        raise HTTPException(status_code=400, detail=f"读取文件失败，请检查格式: {exc}") from exc
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception as exc:  # noqa: BLE001
        print("UNEXPECTED_ERROR_IN_MATCH_BY_URL:", repr(exc))
        raise HTTPException(status_code=500, detail=f"处理 Excel 时出错: {exc}") from exc

    zip_path = tmpdir / "match_results.zip"
    try:
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.write(matched_path, arcname=matched_path.name)
            zf.write(unmatched_path, arcname=unmatched_path.name)
            zf.write(matched_original_path, arcname=matched_original_path.name)
    except Exception as exc:  # noqa: BLE001
        raise HTTPException(status_code=500, detail=f"打包结果文件失败: {exc}") from exc

    # 生成临时下载链接，避免在 JSON 里塞整段 base64 导致扣子报错或超长
    zip_bytes = zip_path.read_bytes()
    token = uuid.uuid4().hex
    _download_cache[token] = zip_bytes
    base_url = str(request.base_url).rstrip("/")
    report_url = f"{base_url}/download/{token}"
    body = MatchByUrlResponse(status="ok", report_url=report_url)
    return JSONResponse(content=body.model_dump())


@app.get("/upload_to_oss", response_class=HTMLResponse)
async def upload_to_oss_page() -> str:
    """批量上传页面，可一次性上传多个 Excel 到 OSS"""
    return """
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>批量上传到 OSS</title>
    <link rel="stylesheet" href="/static/app.css" />
</head>
<body>
    <div class="container">
        <header class="topbar">
            <div class="brand">
                <div class="logo" aria-hidden="true"></div>
                <div>
                    <div class="brand-title">Excel 图书匹配工具</div>
                    <div class="brand-sub">批量上传到 OSS，自动分组处理</div>
                </div>
            </div>
            <nav class="nav" aria-label="导航">
                <a href="/">单次匹配</a>
                <a href="/upload_to_oss" aria-current="page">批量上传到 OSS</a>
                <a href="/download_results">下载处理结果</a>
            </nav>
        </header>

        <section class="hero">
            <h1>批量上传到 OSS</h1>
            <p>
                选择您的姓名，然后一次性上传多个 Excel 文件。系统会根据文件名中的“新表/旧表”自动识别并分组处理。
            </p>
        </section>

        <div class="grid">
            <section class="card">
                <div class="card-body">
                    <div class="card-title">上传文件</div>
                    <form class="form" action="/upload_to_oss" method="post" enctype="multipart/form-data">
                        <div class="field">
                            <label class="label" for="person">选择您的姓名</label>
                            <select class="control" id="person" name="person" required>
                                <option value="">-- 请选择 --</option>
                                <option value="张">张（管理员）</option>
                                <option value="徐">徐</option>
                                <option value="章">章</option>
                                <option value="李">李</option>
                                <option value="姜">姜（备用）</option>
                            </select>
                        </div>
                        <div class="field">
                            <label class="label" for="files">选择 Excel 文件（可多选）</label>
                            <input class="control" type="file" id="files" name="files" accept=".xlsx,.xls" multiple required />
                        </div>
                        <div class="actions">
                            <button class="btn success" type="submit">上传到 OSS</button>
                            <a class="btn link" href="/download_results">我已上传，去下载结果</a>
                        </div>
                        <div class="small muted">上传完成后会在后台自动处理，通常 1 分钟内生成结果。</div>
                    </form>
                </div>
            </section>

            <aside class="card">
                <div class="card-body">
                    <div class="card-title">文件命名规则</div>
                    <div class="callout">
                        <strong>务必包含关键字</strong>
                        <ul>
                            <li>新表文件名必须包含“新表”</li>
                            <li>旧表文件名必须包含“旧表”</li>
                            <li>“新表/旧表”前面的部分作为学校简称（如“玉环”“石夫人”）</li>
                            <li>同一学校的新表和旧表会自动配对处理</li>
                        </ul>
                        <div class="small muted" style="margin-top:8px;">
                            示例：<span class="mono">玉环新表2025.xlsx</span>、<span class="mono">玉环旧表.xlsx</span>
                        </div>
                    </div>
                    <div class="footer">如果页面提示 OSS 配置缺失，请联系管理员设置环境变量。</div>
                </div>
            </aside>
        </div>
    </div>
</body>
</html>
    """


@app.post("/upload_to_oss")
async def upload_to_oss_handler(
    person: str = Form(...),
    files: list[UploadFile] = File(...),
) -> JSONResponse:
    """接收多个文件并上传到 OSS"""
    if not person or not files:
        raise HTTPException(status_code=400, detail="请选择姓名并至少上传一个文件")

    # 读取 OSS 配置
    oss_ak = os.getenv("OSS_ACCESS_KEY_ID")
    oss_sk = os.getenv("OSS_ACCESS_KEY_SECRET")
    oss_endpoint = os.getenv("OSS_ENDPOINT", "https://oss-cn-hangzhou.aliyuncs.com")
    oss_bucket = os.getenv("OSS_BUCKET", "book-company-excel-uploads-v2")

    if not oss_ak or not oss_sk:
        raise HTTPException(
            status_code=500,
            detail="OSS 配置缺失，请联系管理员配置 OSS_ACCESS_KEY_ID 和 OSS_ACCESS_KEY_SECRET"
        )

    try:
        import oss2
        auth = oss2.Auth(oss_ak, oss_sk)
        bucket = oss2.Bucket(auth, oss_endpoint, oss_bucket)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"连接 OSS 失败: {exc}") from exc

    # 生成一个唯一的 task_id（基于时间戳）
    task_id = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
    uploaded_files = []
    errors = []

    for file in files:
        if not file.filename:
            continue

        try:
            # 读取文件内容
            content = await file.read()
            # 上传到 OSS: tasks/<person>/<任意命名>/文件名
            # 这里直接用 task_id 作为子目录，保证每次批量上传是一组
            oss_key = f"tasks/{person}/{file.filename}"
            bucket.put_object(oss_key, content)
            uploaded_files.append({"filename": file.filename, "oss_key": oss_key})
        except Exception as exc:
            errors.append({"filename": file.filename or "未知", "error": str(exc)})

    if errors:
        return JSONResponse(
            status_code=207,
            content={
                "status": "partial_success",
                "uploaded": uploaded_files,
                "errors": errors,
                "message": f"部分文件上传失败。成功: {len(uploaded_files)}, 失败: {len(errors)}"
            }
        )

    return JSONResponse(
        content={
            "status": "success",
            "uploaded": uploaded_files,
            "message": f"成功上传 {len(uploaded_files)} 个文件到 OSS，预计 1分钟内完成处理，届时会通过钉钉通知您。"
        }
    )


@app.get("/download_results", response_class=HTMLResponse)
async def download_results_page(person: str | None = None) -> str:
    """下载结果页面：选择用户或显示该用户的所有处理结果"""
    
    # 如果没有指定 person，显示选择页面
    if not person:
        return """
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>下载处理结果</title>
    <link rel="stylesheet" href="/static/app.css" />
</head>
<body>
    <div class="container">
        <header class="topbar">
            <div class="brand">
                <div class="logo" aria-hidden="true"></div>
                <div>
                    <div class="brand-title">Excel 图书匹配工具</div>
                    <div class="brand-sub">查看并下载已完成结果</div>
                </div>
            </div>
            <nav class="nav" aria-label="导航">
                <a href="/">单次匹配</a>
                <a href="/upload_to_oss">批量上传到 OSS</a>
                <a href="/download_results" aria-current="page">下载处理结果</a>
            </nav>
        </header>

        <section class="hero">
            <h1>下载处理结果</h1>
            <p>选择您的姓名，查看并下载所有已完成的数据处理结果（下载链接默认 1 小时有效）。</p>
        </section>

        <div class="grid">
            <section class="card">
                <div class="card-body">
                    <div class="card-title">选择姓名</div>
                    <form class="form" action="/download_results" method="get">
                        <div class="field">
                            <label class="label" for="person">您的姓名</label>
                            <select class="control" id="person" name="person" required>
                                <option value="">-- 请选择 --</option>
                                <option value="张">张（管理员）</option>
                                <option value="徐">徐</option>
                                <option value="章">章</option>
                                <option value="李">李</option>
                                <option value="姜">姜（备用）</option>
                            </select>
                        </div>
                        <div class="actions">
                            <button class="btn primary" type="submit">查看我的结果</button>
                            <a class="btn link" href="/upload_to_oss">去上传文件</a>
                        </div>
                    </form>
                </div>
            </section>

            <aside class="card">
                <div class="card-body">
                    <div class="card-title">说明</div>
                    <div class="callout">
                        <strong>下载链接有效期</strong>
                        <ul>
                            <li>链接默认 1 小时有效，过期刷新页面可重新生成</li>
                            <li>如果没有结果，请先去“批量上传到 OSS”上传文件</li>
                        </ul>
                    </div>
                </div>
            </aside>
        </div>
    </div>
</body>
</html>
        """
    
    # 如果指定了 person，列出该用户的所有结果
    oss_ak = os.getenv("OSS_ACCESS_KEY_ID")
    oss_sk = os.getenv("OSS_ACCESS_KEY_SECRET")
    oss_endpoint = os.getenv("OSS_ENDPOINT", "https://oss-cn-hangzhou.aliyuncs.com")
    oss_bucket = os.getenv("OSS_BUCKET", "book-company-excel-uploads-v2")
    
    if not oss_ak or not oss_sk:
        raise HTTPException(
            status_code=500,
            detail="OSS 配置缺失，请联系管理员配置 OSS_ACCESS_KEY_ID 和 OSS_ACCESS_KEY_SECRET"
        )
    
    try:
        import oss2
        auth = oss2.Auth(oss_ak, oss_sk)
        bucket = oss2.Bucket(auth, oss_endpoint, oss_bucket)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"连接 OSS 失败: {exc}") from exc
    
    # 列举该用户的所有结果目录
    # 这里通过查找所有 DONE.json 来确定有哪些任务，而不是依赖 OSS 的“子目录”特性，
    # 兼容不同版本 SDK / 控制台的行为。
    results_prefix = f"results/{person}/"
    tasks: list[dict[str, str]] = []

    try:
        task_dirs: set[str] = set()

        # 先找到所有 DONE.json，提取出 task_id
        for obj in oss2.ObjectIterator(bucket, prefix=results_prefix):
            key = obj.key
            if not key.endswith("DONE.json"):
                continue
            # 形如 results/张/东阳技校/DONE.json -> 取中间的“东阳技校”
            relative = key[len(results_prefix) :]
            parts = relative.split("/", 1)
            if not parts or not parts[0]:
                continue
            task_dirs.add(parts[0])

        # 对每个 task_id，读取 DONE.json 并生成下载链接
        for task_id in sorted(task_dirs, reverse=True):  # 最新的在前
            done_key = f"{results_prefix}{task_id}/DONE.json"
            try:
                done_obj = bucket.get_object(done_key)
                done_data = json.loads(done_obj.read().decode("utf-8"))

                # 生成 zip 文件的临时签名 URL（1小时有效期）
                zip_key = f"{results_prefix}{task_id}/match_results_{task_id}.zip"
                download_url = bucket.sign_url("GET", zip_key, 3600)

                tasks.append(
                    {
                        "task_id": task_id,
                        "school": done_data.get("school", "未知"),
                        "time": done_data.get("time", "未知"),
                        "download_url": download_url,
                    }
                )
            except Exception:
                # 如果没有 DONE.json 或读取失败，跳过该任务
                continue
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"列举结果失败: {exc}") from exc
    
    # 生成 HTML 页面显示结果列表
    if not tasks:
        tasks_html = '<p class="muted">暂无处理结果，请先上传文件进行处理。</p>'
    else:
        tasks_html = '<table class="table">'
        tasks_html += '''
            <thead>
                <tr>
                    <th>学校</th>
                    <th>任务ID</th>
                    <th>处理时间</th>
                    <th class="center">操作</th>
                </tr>
            </thead>
            <tbody>
        '''
        for task in tasks:
            time_str = task["time"].split('T')[0] if 'T' in task["time"] else task["time"]
            tasks_html += f'''
                <tr>
                    <td>{task["school"]}</td>
                    <td class="mono">{task["task_id"]}</td>
                    <td>{time_str}</td>
                    <td class="center">
                        <a class="btn success link" href="{task["download_url"]}">下载结果</a>
                    </td>
                </tr>
            '''
        tasks_html += '</tbody></table>'
    
    return f"""
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>{person} 的处理结果</title>
    <link rel="stylesheet" href="/static/app.css" />
</head>
<body>
    <div class="container">
        <header class="topbar">
            <div class="brand">
                <div class="logo" aria-hidden="true"></div>
                <div>
                    <div class="brand-title">Excel 图书匹配工具</div>
                    <div class="brand-sub">{person} 的处理结果</div>
                </div>
            </div>
            <nav class="nav" aria-label="导航">
                <a href="/">单次匹配</a>
                <a href="/upload_to_oss">批量上传到 OSS</a>
                <a href="/download_results" aria-current="page">下载处理结果</a>
            </nav>
        </header>

        <section class="hero">
            <h1>{person} 的处理结果</h1>
            <p>
                点击“下载结果”即可下载对应 ZIP（含匹配明细、未匹配表、匹配原始表）。下载链接有效期为 1 小时，过期后刷新页面即可重新生成。
            </p>
        </section>

        <section class="card">
            <div class="card-body">
                <div class="actions" style="margin-top:0;">
                    <a class="btn link" href="/download_results">← 返回选择页面</a>
                </div>
                <div style="margin-top:12px;">
                    {tasks_html}
                </div>
            </div>
        </section>

        <div class="footer">提示：如果结果为空，请先确认是否已成功上传并完成处理。</div>
    </div>
</body>
</html>
    """


@app.get("/health")
async def health() -> dict[str, str]:
    return {"status": "ok"}

