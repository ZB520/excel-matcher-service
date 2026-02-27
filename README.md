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

## 2. 阿里云 OSS 自动化处理（同事上传即可自动生成结果）

适用场景：**同事不懂计算机，只会上传/下载**；每位同事在自己的名字文件夹下，直接放多组“学校简称 + 新表/旧表 + 版本(可选)”的 Excel，系统会自动分组并生成结果。

### 2.1 OSS 目录规范（按人 + 学校自动分组）

- **Bucket**：`book-company-excel-uploads-v2`
- **Region**：`cn-hangzhou`（Endpoint 建议设置为 `https://oss-cn-hangzhou.aliyuncs.com`）

- **输入（同事上传）**：
  - 每位同事一个目录：`tasks/<person>/`
    - 示例：`tasks/石夫人/`、`tasks/张三/`、`tasks/李四/`
  - 所有需要处理的 Excel 文件**直接放在自己目录下**，例如：
    - `tasks/张三/玉环新表2025.xlsx`
    - `tasks/张三/玉环旧表.xlsx`
    - `tasks/张三/石夫人新表2025秋.xlsx`
    - `tasks/张三/石夫人旧表.xlsx`
    - `tasks/张三/义务新表_A.xlsx`
    - `tasks/张三/义务旧表_A.xlsx`

- **命名规则（非常重要）**：
  - 文件名结构：`学校简称 + 新表/旧表 + 版本信息(可选)`
    - 例如：`玉环新表2025.xlsx`、`玉环旧表.xlsx`
    - 例如：`石夫人新表2025秋.xlsx`、`石夫人旧表.xlsx`
    - 例如：`义务新表_A.xlsx`、`义务旧表_A.xlsx`
  - 识别逻辑：
    - 含有“新表” → 视为新表
    - 含有“旧表” → 视为旧表
    - “新表/旧表”前面的部分作为**学校简称**（如 `玉环`、`石夫人`、`义务`）

> 只要老师们记住：**文件名里一定要出现“新表”或“旧表”这两个词，并把学校名写在前面**，系统就能自动把不同学校的一对新旧表配好。

### 2.2 输出（函数计算自动写回）

- **输出目录**：`results/<person>/<task_id>/`
  - `<person>`：同事名字（和 `tasks/<person>/` 一致）
  - `<task_id>`：由学校简称和新表文件名中的“版本信息”拼出来，例如：
    - 新表：`玉环新表2025.xlsx` → 任务 ID：`玉环_2025`
    - 新表：`石夫人新表2025秋.xlsx` → 任务 ID：`石夫人_2025秋`
    - 新表：`义务新表_A.xlsx` → 任务 ID：`义务__A`（版本信息是 `_A`）

- **输出文件**：
  - `match_results_<task_id>.zip`
  - `已匹配数据表.xlsx`
  - `未匹配数据表.xlsx`
  - `已匹配数据原始表.xlsx`
  - `DONE.json`：处理成功标记，里面会写清楚：
    - person（同事姓名）
    - school（学校简称）
    - version_suffix（版本信息）
    - inputs：新旧表的 OSS key
    - outputs：结果文件的 OSS key
  - `ERROR.json`：处理失败或无法识别新旧/缺少成对文件时生成，例如：
    - 某学校只有新表没有旧表
    - 某些文件名里既没有“新表”也没有“旧表”

### 2.3 阿里云函数计算（FC）部署要点（给管理员）

- **函数入口**：`fc_handler.handler`
- **触发器**：OSS Bucket 配置 `ObjectCreated:*`，前缀 `tasks/`，后缀 `.xlsx`
- **权限**：给函数绑定可访问 OSS 的 RAM Role（推荐，避免写死 AK/SK）
- **环境变量（至少其一）**：
  - `OSS_ENDPOINT`：例如 `https://oss-cn-hangzhou.aliyuncs.com`
  - 或让事件里带 `region` 时自动拼接（仍建议显式配置 `OSS_ENDPOINT` 更稳）
- **依赖**：本项目 `requirements.txt` 已包含 `oss2`，FC 构建时安装即可

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

