from pathlib import Path
import os
import tempfile
import uuid
import zipfile
from datetime import datetime, timezone

from fastapi import FastAPI, File, UploadFile, HTTPException, Request, Form
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, Response
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


@app.get("/", response_class=HTMLResponse)
async def index() -> str:
    """简单上传页面，方便同事在浏览器里直接使用。"""
    return """
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8" />
    <title>Excel 图书匹配工具</title>
    <style>
        body { font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; margin: 40px; }
        h1 { margin-bottom: 0.5rem; }
        form { margin-top: 1.5rem; display: flex; flex-direction: column; gap: 1rem; max-width: 480px; }
        label { font-weight: 600; }
        input[type="file"] { padding: 0.25rem 0; }
        button {
            padding: 0.6rem 1.2rem;
            border-radius: 4px;
            border: none;
            background-color: #2563eb;
            color: white;
            font-size: 1rem;
            cursor: pointer;
        }
        button:hover { background-color: #1d4ed8; }
        .hint { font-size: 0.9rem; color: #4b5563; }
        .template-link {
            display: inline-block;
            margin-top: 0.75rem;
            font-size: 0.9rem;
        }
        .template-link a {
            color: #2563eb;
            text-decoration: none;
        }
        .template-link a:hover {
            text-decoration: underline;
        }
        nav { margin-bottom: 2rem; padding-bottom: 1rem; border-bottom: 1px solid #e5e7eb; }
        nav a { margin-right: 1rem; color: #2563eb; text-decoration: none; }
        nav a:hover { text-decoration: underline; }
    </style>
</head>
<body>
    <nav>
        <a href="/">单次匹配</a>
        <a href="/upload_to_oss">批量上传到 OSS</a>
        <a href="/download_results">下载处理结果</a>
    </nav>
    <h1>Excel 图书匹配工具</h1>
    <p class="hint">
        选择一份旧表和一份新表，点击“开始匹配”后，系统会自动生成结果压缩包（包含匹配明细、未匹配表、匹配原始表）。
    </p>
    <p class="template-link">
        不确定新表格式？可以先
        <a href="/template/new" download>下载新表模板</a>
        查看或填写。
    </p>
    <form action="/match" method="post" enctype="multipart/form-data">
        <div>
            <label for="old_file">旧表 Excel：</label><br />
            <input type="file" id="old_file" name="old_file" accept=".xlsx,.xls" required />
        </div>
        <div>
            <label for="new_file">新表 Excel：</label><br />
            <input type="file" id="new_file" name="new_file" accept=".xlsx,.xls" required />
        </div>
        <button type="submit">开始匹配并下载结果</button>
    </form>
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
    <title>批量上传到 OSS</title>
    <style>
        body { font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; margin: 40px; }
        h1 { margin-bottom: 0.5rem; color: #1f2937; }
        .hint { font-size: 0.9rem; color: #4b5563; margin-bottom: 1.5rem; line-height: 1.6; }
        form { display: flex; flex-direction: column; gap: 1.2rem; max-width: 600px; }
        label { font-weight: 600; color: #374151; }
        select, input[type="file"] { padding: 0.5rem; font-size: 1rem; border: 1px solid #d1d5db; border-radius: 4px; }
        button {
            padding: 0.8rem 1.5rem;
            border-radius: 6px;
            border: none;
            background-color: #10b981;
            color: white;
            font-size: 1rem;
            font-weight: 600;
            cursor: pointer;
            transition: background-color 0.2s;
        }
        button:hover { background-color: #059669; }
        .warning { background-color: #fef3c7; border-left: 4px solid #f59e0b; padding: 1rem; margin-top: 1rem; }
        .warning strong { color: #92400e; }
        nav { margin-bottom: 2rem; padding-bottom: 1rem; border-bottom: 1px solid #e5e7eb; }
        nav a { margin-right: 1rem; color: #2563eb; text-decoration: none; }
        nav a:hover { text-decoration: underline; }
    </style>
</head>
<body>
    <nav>
        <a href="/">单次匹配</a>
        <a href="/upload_to_oss">批量上传到 OSS</a>
        <a href="/download_results">下载处理结果</a>
    </nav>
    <h1>批量上传 Excel 到阿里云 OSS</h1>
    <p class="hint">
        选择您的姓名，然后一次性上传多个 Excel 文件。<br>
        <strong>重要提示：</strong>文件名必须包含"新表"或"旧表"关键字，系统会自动识别并分组。<br>
        例如：<code>玉环新表2025.xlsx</code>、<code>玉环旧表.xlsx</code>、<code>石夫人新表2025秋.xlsx</code> 等。
    </p>
    <form action="/upload_to_oss" method="post" enctype="multipart/form-data">
        <div>
            <label for="person">选择您的姓名：</label>
            <select id="person" name="person" required>
                <option value="">-- 请选择 --</option>
                <option value="张">张（管理员）</option>
                <option value="徐">徐</option>
                <option value="章">章</option>
                <option value="李">李</option>
            </select>
        </div>
        <div>
            <label for="files">选择 Excel 文件（可多选）：</label>
            <input type="file" id="files" name="files" accept=".xlsx,.xls" multiple required />
        </div>
        <button type="submit">上传到 OSS</button>
    </form>
    <div class="warning">
        <strong>文件命名规则：</strong>
        <ul style="margin: 0.5rem 0 0 1.5rem;">
            <li>新表文件名必须包含"新表"二字</li>
            <li>旧表文件名必须包含"旧表"二字</li>
            <li>"新表"/"旧表"前面的部分作为学校简称（如"玉环"、"石夫人"）</li>
            <li>同一学校的新表和旧表会自动配对处理</li>
        </ul>
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
            "message": f"成功上传 {len(uploaded_files)} 个文件到 OSS，预计 10-20 分钟内完成处理，届时会通过钉钉通知您。"
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
    <title>下载处理结果</title>
    <style>
        body { font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; margin: 40px; }
        h1 { margin-bottom: 0.5rem; color: #1f2937; }
        .hint { font-size: 0.9rem; color: #4b5563; margin-bottom: 1.5rem; line-height: 1.6; }
        form { display: flex; flex-direction: column; gap: 1.2rem; max-width: 600px; }
        label { font-weight: 600; color: #374151; }
        select { padding: 0.5rem; font-size: 1rem; border: 1px solid #d1d5db; border-radius: 4px; }
        button {
            padding: 0.8rem 1.5rem;
            border-radius: 6px;
            border: none;
            background-color: #3b82f6;
            color: white;
            font-size: 1rem;
            font-weight: 600;
            cursor: pointer;
            transition: background-color 0.2s;
        }
        button:hover { background-color: #2563eb; }
        nav { margin-bottom: 2rem; padding-bottom: 1rem; border-bottom: 1px solid #e5e7eb; }
        nav a { margin-right: 1rem; color: #2563eb; text-decoration: none; }
        nav a:hover { text-decoration: underline; }
    </style>
</head>
<body>
    <nav>
        <a href="/">单次匹配</a>
        <a href="/upload_to_oss">批量上传到 OSS</a>
        <a href="/download_results">下载处理结果</a>
    </nav>
    <h1>下载处理结果</h1>
    <p class="hint">
        选择您的姓名，查看并下载所有已完成的数据处理结果。
    </p>
    <form action="/download_results" method="get">
        <div>
            <label for="person">选择您的姓名：</label>
            <select id="person" name="person" required>
                <option value="">-- 请选择 --</option>
                <option value="张">张（管理员）</option>
                <option value="徐">徐</option>
                <option value="章">章</option>
                <option value="李">李</option>
            </select>
        </div>
        <button type="submit">查看我的结果</button>
    </form>
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
    results_prefix = f"results/{person}/"
    tasks = []
    
    try:
        # 列举所有子目录（task_id）
        task_dirs = set()
        for obj in oss2.ObjectIterator(bucket, prefix=results_prefix, delimiter='/'):
            if obj.is_prefix():
                # 提取 task_id
                task_id = obj.key.rstrip('/').split('/')[-1]
                task_dirs.add(task_id)
        
        # 对每个 task_id，读取 DONE.json 并生成下载链接
        for task_id in sorted(task_dirs, reverse=True):  # 最新的在前
            done_key = f"{results_prefix}{task_id}/DONE.json"
            try:
                done_obj = bucket.get_object(done_key)
                done_data = json.loads(done_obj.read().decode('utf-8'))
                
                # 生成 zip 文件的临时签名 URL（1小时有效期）
                zip_key = f"{results_prefix}{task_id}/match_results_{task_id}.zip"
                download_url = bucket.sign_url('GET', zip_key, 3600)
                
                tasks.append({
                    "task_id": task_id,
                    "school": done_data.get("school", "未知"),
                    "time": done_data.get("time", "未知"),
                    "download_url": download_url,
                })
            except Exception:
                # 如果没有 DONE.json 或读取失败，跳过该任务
                continue
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"列举结果失败: {exc}") from exc
    
    # 生成 HTML 页面显示结果列表
    if not tasks:
        tasks_html = '<p style="color: #6b7280;">暂无处理结果，请先上传文件进行处理。</p>'
    else:
        tasks_html = '<table style="width: 100%; border-collapse: collapse;">'
        tasks_html += '''
            <thead>
                <tr style="background-color: #f3f4f6; border-bottom: 2px solid #e5e7eb;">
                    <th style="padding: 0.75rem; text-align: left;">学校</th>
                    <th style="padding: 0.75rem; text-align: left;">任务ID</th>
                    <th style="padding: 0.75rem; text-align: left;">处理时间</th>
                    <th style="padding: 0.75rem; text-align: center;">操作</th>
                </tr>
            </thead>
            <tbody>
        '''
        for task in tasks:
            time_str = task["time"].split('T')[0] if 'T' in task["time"] else task["time"]
            tasks_html += f'''
                <tr style="border-bottom: 1px solid #e5e7eb;">
                    <td style="padding: 0.75rem;">{task["school"]}</td>
                    <td style="padding: 0.75rem; font-family: monospace; font-size: 0.9rem;">{task["task_id"]}</td>
                    <td style="padding: 0.75rem;">{time_str}</td>
                    <td style="padding: 0.75rem; text-align: center;">
                        <a href="{task["download_url"]}" 
                           style="background-color: #10b981; color: white; padding: 0.5rem 1rem; 
                                  border-radius: 4px; text-decoration: none; font-size: 0.9rem;">
                            下载结果
                        </a>
                    </td>
                </tr>
            '''
        tasks_html += '</tbody></table>'
    
    return f"""
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8" />
    <title>{person} 的处理结果</title>
    <style>
        body {{ font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; margin: 40px; }}
        h1 {{ margin-bottom: 0.5rem; color: #1f2937; }}
        .hint {{ font-size: 0.9rem; color: #4b5563; margin-bottom: 1.5rem; line-height: 1.6; }}
        .back-link {{ display: inline-block; margin-bottom: 1.5rem; color: #2563eb; text-decoration: none; }}
        .back-link:hover {{ text-decoration: underline; }}
        table {{ margin-top: 1rem; }}
        nav {{ margin-bottom: 2rem; padding-bottom: 1rem; border-bottom: 1px solid #e5e7eb; }}
        nav a {{ margin-right: 1rem; color: #2563eb; text-decoration: none; }}
        nav a:hover {{ text-decoration: underline; }}
    </style>
</head>
<body>
    <nav>
        <a href="/">单次匹配</a>
        <a href="/upload_to_oss">批量上传到 OSS</a>
        <a href="/download_results">下载处理结果</a>
    </nav>
    <a href="/download_results" class="back-link">← 返回选择页面</a>
    <h1>{person} 的处理结果</h1>
    <p class="hint">
        以下是您所有已完成的数据处理结果。点击"下载结果"按钮即可下载对应的 ZIP 文件（包含匹配明细、未匹配表、匹配原始表）。<br>
        <strong>注意：</strong>下载链接有效期为 1 小时，过期后请刷新页面重新生成。
    </p>
    {tasks_html}
</body>
</html>
    """


@app.get("/health")
async def health() -> dict[str, str]:
    return {"status": "ok"}

