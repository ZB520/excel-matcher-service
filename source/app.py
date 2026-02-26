from pathlib import Path
import tempfile
import zipfile

from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import FileResponse, HTMLResponse

import excel_book_matcher


BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_DIR = BASE_DIR / "assets"
NEW_TEMPLATE_PATH = TEMPLATE_DIR / "新表模板.xlsx"

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
    </style>
</head>
<body>
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


@app.get("/health")
async def health() -> dict[str, str]:
    return {"status": "ok"}

