# Excel 图书匹配服务（FastAPI 版）

## 1. 本地运行（开发调试）

- 安装依赖：

```bash
pip install -r requirements.txt
```

- 启动 Web 服务：

```bash
uvicorn source.app:app --host 0.0.0.0 --port 8000
```

- 打开浏览器访问：
  - 本机：`http://127.0.0.1:8000/`

上传“旧表”和“新表” Excel 后，会下载一个 `match_results.zip`，里面包含三张结果表：

- `已匹配数据表.xlsx`
- `未匹配数据表.xlsx`
- `已匹配数据原始表.xlsx`

## 2. 在 Zeabur 上部署步骤

1. **准备代码仓库**
   - 确保下面这些文件都在项目根目录（例如一个名为 `excel-matcher-service` 的仓库根目录）：
     - `excel_book_matcher.py`
     - `source/app.py`
     - `source/__init__.py`
     - `requirements.txt`
     - `Dockerfile`

2. **推送到 Git 平台**
   - 将项目推到 GitHub / Gitee 任意一个代码托管平台（Zeabur 支持从这些平台拉取代码）。

3. **在 Zeabur 创建服务**
   - 登录 Zeabur 控制台，新建一个 Service。
   - 选择刚才的 Git 仓库作为代码来源。
   - 运行环境选择：
     - 如果使用 **Python 模式**：
       - 依赖文件：填写 `requirements.txt`
       - 启动命令：`uvicorn source.app:app --host 0.0.0.0 --port 8000`
     - 如果使用 **Docker 模式**：
       - 直接使用仓库中的 `Dockerfile`，Zeabur 会自动构建镜像并运行。

4. **端口和健康检查**
   - 暴露端口：`8000`（Zeabur 外部会自动映射到 HTTP 端口）。
   - 可以将 `/health` 配置为健康检查路径（返回简单的 `{"status": "ok"}`）。

5. **验证部署并分享链接**
   - 部署成功后，Zeabur 会给出一个 URL，例如：`https://excel-matcher.zeabur.app`。
   - 在浏览器打开该地址，应能看到上传页面：
     - 上传一份旧表和一份新表，点击按钮后，浏览器会下载 `match_results.zip`。
   - 确认结果无误后，把该 URL 发给公司同事使用，并简单说明：
     - 工具仅用于内部教学用书统计；
     - 上传文件只在服务的临时目录中短暂存在，处理完成后会被自动删除。

