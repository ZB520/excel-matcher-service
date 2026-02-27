import json
import os
import tempfile
import uuid
import zipfile
import hmac
import hashlib
import base64
import time
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable
from urllib.parse import unquote, quote_plus

import excel_book_matcher


try:
    import oss2  # type: ignore
except Exception:  # noqa: BLE001
    oss2 = None  # type: ignore


@dataclass(frozen=True)
class OssLocation:
    bucket: str
    region: str | None
    key: str


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def _send_dingtalk_notification(person: str, task_id: str, school: str, bucket_name: str, region: str | None) -> None:
    """
    发送钉钉通知，@ 对应的人，并包含 OSS 临时下载链接
    """
    webhook = _env("DINGTALK_WEBHOOK")
    secret = _env("DINGTALK_SECRET")
    
    if not webhook:
        print(f"DINGTALK_WEBHOOK not configured, skip notification for {person}/{task_id}")
        return
    
    # 手机号映射（测试阶段）
    phone_map = {
        "张": "17737759937",
        "徐": "17737759937",  # TODO: 替换为实际手机号
        "章": "17737759937",  # TODO: 替换为实际手机号
        "李": "17737759937",  # TODO: 替换为实际手机号
    }
    
    phone = phone_map.get(person, "")
    if not phone:
        print(f"No phone mapping for person: {person}")
        return
    
    # 生成 OSS 临时下载链接
    download_url = ""
    try:
        bucket = _bucket_client(bucket_name, region)
        zip_key = f"results/{person}/{task_id}/match_results_{task_id}.zip"
        download_url = bucket.sign_url('GET', zip_key, 86400)  # 24小时有效期
    except Exception as exc:
        print(f"Failed to generate OSS download URL: {exc}")
        # 如果生成失败，继续发送通知，只是不包含下载链接
    
    # 构造消息
    at_mobiles = [phone]
    
    # 根据是否成功生成下载链接，构造不同的消息
    if download_url:
        text = f"@{phone} 您好，您的数据处理任务已完成！\n\n任务信息：\n- 学校：{school}\n- 任务ID：{task_id}\n\n点击下载结果（24小时内有效）：\n{download_url}\n\n或访问网页查看所有结果：\nhttps://你的域名.zeabur.app/download_results?person={person}"
    else:
        text = f"@{phone} 您好，您的数据处理任务已完成！\n\n任务信息：\n- 学校：{school}\n- 任务ID：{task_id}\n\n请访问网页下载结果：\nhttps://你的域名.zeabur.app/download_results?person={person}"
    
    message = {
        "msgtype": "text",
        "text": {"content": text},
        "at": {"atMobiles": at_mobiles, "isAtAll": False}
    }
    
    # 如果配置了加签，生成签名
    url = webhook
    if secret:
        timestamp = str(int(time.time() * 1000))
        string_to_sign = f"{timestamp}\n{secret}"
        hmac_code = hmac.new(
            secret.encode('utf-8'),
            string_to_sign.encode('utf-8'),
            digestmod=hashlib.sha256
        ).digest()
        sign = quote_plus(base64.b64encode(hmac_code))
        url = f"{webhook}&timestamp={timestamp}&sign={sign}"
    
    # 发送请求
    try:
        import urllib.request
        req = urllib.request.Request(
            url,
            data=json.dumps(message).encode('utf-8'),
            headers={'Content-Type': 'application/json'}
        )
        with urllib.request.urlopen(req, timeout=10) as response:
            result = response.read().decode('utf-8')
            print(f"DingTalk notification sent: {result}")
    except Exception as exc:
        print(f"Failed to send DingTalk notification: {exc}")


def _env(name: str, default: str | None = None) -> str | None:
    v = os.getenv(name)
    if v is None or v == "":
        return default
    return v


def _build_endpoint(region: str | None) -> str | None:
    explicit = _env("OSS_ENDPOINT") or _env("ALIBABA_CLOUD_OSS_ENDPOINT")
    if explicit:
        return explicit
    if not region:
        return None
    return f"https://oss-{region}.aliyuncs.com"


def _build_auth() -> Any:
    """
    Prefer STS credentials if present (FC RAM Role), fallback to AK/SK env vars.
    """
    if oss2 is None:
        raise RuntimeError("Missing dependency 'oss2'. Add it to requirements.txt.")

    ak = (
        _env("ALIBABA_CLOUD_ACCESS_KEY_ID")
        or _env("OSS_ACCESS_KEY_ID")
        or _env("ACCESS_KEY_ID")
    )
    sk = (
        _env("ALIBABA_CLOUD_ACCESS_KEY_SECRET")
        or _env("OSS_ACCESS_KEY_SECRET")
        or _env("ACCESS_KEY_SECRET")
    )
    token = _env("ALIBABA_CLOUD_SECURITY_TOKEN") or _env("SECURITY_TOKEN")

    if ak and sk and token:
        return oss2.StsAuth(ak, sk, token)
    if ak and sk:
        return oss2.Auth(ak, sk)

    raise RuntimeError(
        "No OSS credentials found. Set RAM Role (STS) or provide "
        "ALIBABA_CLOUD_ACCESS_KEY_ID/ALIBABA_CLOUD_ACCESS_KEY_SECRET "
        "(optional ALIBABA_CLOUD_SECURITY_TOKEN)."
    )


def _bucket_client(bucket_name: str, region: str | None) -> Any:
    if oss2 is None:
        raise RuntimeError("Missing dependency 'oss2'. Add it to requirements.txt.")
    endpoint = _build_endpoint(region)
    if not endpoint:
        raise RuntimeError("Cannot determine OSS endpoint. Set OSS_ENDPOINT.")
    return oss2.Bucket(_build_auth(), endpoint, bucket_name)


def _parse_oss_event(event: Any) -> list[OssLocation]:
    """
    FC OSS trigger event format:
      {"events":[{"region":"cn-hangzhou","oss":{"bucket":{"name":"..."},"object":{"key":"..."}}, ...}]}
    """
    if isinstance(event, (bytes, bytearray)):
        event = event.decode("utf-8", errors="replace")
    if isinstance(event, str):
        payload = json.loads(event)
    elif isinstance(event, dict):
        payload = event
    else:
        raise ValueError(f"Unsupported event type: {type(event)!r}")

    events = payload.get("events")
    if not isinstance(events, list):
        raise ValueError("Event payload missing 'events' list.")

    out: list[OssLocation] = []
    for e in events:
        if not isinstance(e, dict):
            continue
        oss_info = e.get("oss") or {}
        bucket = ((oss_info.get("bucket") or {}).get("name")) if isinstance(oss_info, dict) else None
        obj_key = ((oss_info.get("object") or {}).get("key")) if isinstance(oss_info, dict) else None
        region = e.get("region") if isinstance(e.get("region"), str) else None
        if not isinstance(bucket, str) or not isinstance(obj_key, str):
            continue
        out.append(OssLocation(bucket=bucket, region=region, key=unquote(obj_key)))
    return out


def _iter_xlsx_keys(bucket: Any, prefix: str) -> list[str]:
    if oss2 is None:
        raise RuntimeError("Missing dependency 'oss2'. Add it to requirements.txt.")
    keys: list[str] = []
    for obj in oss2.ObjectIterator(bucket, prefix=prefix):
        key = getattr(obj, "key", None)
        if isinstance(key, str) and key.lower().endswith(".xlsx"):
            keys.append(key)
    return keys


def _filename_from_key(key: str) -> str:
    return key.rsplit("/", 1)[-1]


def _parse_person_and_filename(key: str) -> tuple[str, str]:
    """
    Expect: tasks/<person>/<filename>.xlsx (不再要求 task_id 子目录).
    """
    parts = [p for p in key.split("/") if p]
    if len(parts) < 3 or parts[0] != "tasks":
        raise ValueError(f"Key not under tasks/<person>/: {key}")
    person = parts[1]
    filename = parts[-1]
    return person, filename


def _parse_school_and_version(filename: str) -> tuple[str, bool, bool, str]:
    """
    从文件名中解析出：
    - 学校简称 school_key：出现在“新表/旧表”之前的部分
    - is_new / is_old：是否为新表或旧表
    - version_suffix：去掉 <school>新表/旧表 和扩展名后的剩余部分（可为空字符串）
    """
    base = filename
    if base.lower().endswith(".xlsx"):
        base = base[:-5]
    elif base.lower().endswith(".xls"):
        base = base[:-4]

    idx_new = base.find("新表")
    idx_old = base.find("旧表")

    if idx_new == -1 and idx_old == -1:
        raise ValueError("文件名中未找到“新表”或“旧表”关键字")

    is_new = idx_new != -1
    is_old = idx_old != -1

    # 优先看“新表”，否则看“旧表”
    if is_new:
        idx = idx_new
        marker = "新表"
    else:
        idx = idx_old
        marker = "旧表"

    school = base[:idx]
    version_suffix = base[idx + len(marker) :]

    return school, is_new, is_old, version_suffix


def _exists(bucket: Any, key: str) -> bool:
    try:
        bucket.head_object(key)
        return True
    except Exception:  # noqa: BLE001
        return False


def _put_json(bucket: Any, key: str, body: dict[str, Any]) -> None:
    data = json.dumps(body, ensure_ascii=False, indent=2).encode("utf-8")
    bucket.put_object(key, data)


def _upload_file(bucket: Any, key: str, local_path: Path) -> None:
    bucket.put_object_from_file(key, str(local_path))


def _zip_files(zip_path: Path, files: Iterable[Path]) -> None:
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for p in files:
            zf.write(p, arcname=p.name)


def _process_one(bucket_name: str, region: str | None, key: str) -> dict[str, Any]:
    """
    根据触发事件涉及的 key，按 person 维度对 tasks/<person>/ 下的
    所有 xlsx 进行分组，并对每个学校+新表版本执行匹配。
    """
    bucket = _bucket_client(bucket_name, region)

    try:
        person, _ = _parse_person_and_filename(key)
    except Exception as exc:  # noqa: BLE001
        return {"status": "ignored", "reason": str(exc), "key": key}

    dir_prefix = f"tasks/{person}/"

    # 收集该同事目录下所有 xlsx，并按学校 key 分组
    xlsx_keys = _iter_xlsx_keys(bucket, dir_prefix)
    groups: dict[str, dict[str, list[dict[str, Any]]]] = {}
    invalid_files: list[dict[str, Any]] = []

    for obj_key in xlsx_keys:
        filename = _filename_from_key(obj_key)
        try:
            school, is_new, is_old, version_suffix = _parse_school_and_version(filename)
        except Exception as exc:  # noqa: BLE001
            invalid_files.append({"key": obj_key, "filename": filename, "reason": str(exc)})
            continue

        grp = groups.setdefault(school, {"new": [], "old": []})
        if is_new:
            grp["new"].append(
                {
                    "key": obj_key,
                    "filename": filename,
                    "version_suffix": version_suffix,
                }
            )
        if is_old:
            grp["old"].append({"key": obj_key, "filename": filename})

    # 若有非法命名文件，写一个汇总错误文件方便排查
    if invalid_files:
        _put_json(
            bucket,
            f"results/{person}/INVALID_FILES.json",
            {
                "status": "error",
                "time": _now_iso(),
                "person": person,
                "reason": "以下文件名中未正确包含“新表”或“旧表”关键字",
                "files": invalid_files,
            },
        )

    task_summaries: list[dict[str, Any]] = []

    for school, parts in groups.items():
        new_list = parts["new"]
        old_list = parts["old"]

        if not new_list or not old_list:
            # 缺少新表或旧表，按学校维度写错误
            err_key = f"results/{person}/{school}_no_pair/ERROR.json"
            _put_json(
                bucket,
                err_key,
                {
                    "status": "error",
                    "time": _now_iso(),
                    "person": person,
                    "school": school,
                    "reason": "缺少新表或旧表，请检查命名中是否包含“新表”/“旧表”",
                    "new_files": [n["filename"] for n in new_list],
                    "old_files": [o["filename"] for o in old_list],
                },
            )
            continue

        # 简单策略：对每个新表都跑一遍，旧表选第一个
        chosen_old = old_list[0]["key"]

        for new_info in new_list:
            new_key = new_info["key"]
            version_suffix = new_info["version_suffix"]
            if version_suffix:
                task_id = f"{school}_{version_suffix}"
            else:
                task_id = school

            results_prefix = f"results/{person}/{task_id}/"
            done_key = f"{results_prefix}DONE.json"
            error_key = f"{results_prefix}ERROR.json"

            if _exists(bucket, done_key):
                task_summaries.append(
                    {
                        "status": "skipped",
                        "reason": "DONE already exists",
                        "person": person,
                        "school": school,
                        "task_id": task_id,
                        "new_key": new_key,
                        "old_key": chosen_old,
                    }
                )
                continue

            workdir = Path(tempfile.mkdtemp(prefix="excel_match_")) / uuid.uuid4().hex
            workdir.mkdir(parents=True, exist_ok=True)

            local_old = workdir / "old.xlsx"
            local_new = workdir / "new.xlsx"

            try:
                bucket.get_object_to_file(chosen_old, str(local_old))
                bucket.get_object_to_file(new_key, str(local_new))
            except Exception as exc:  # noqa: BLE001
                _put_json(
                    bucket,
                    error_key,
                    {
                        "status": "error",
                        "time": _now_iso(),
                        "person": person,
                        "school": school,
                        "task_id": task_id,
                        "reason": f"Download OSS objects failed: {exc}",
                        "old_key": chosen_old,
                        "new_key": new_key,
                    },
                )
                task_summaries.append(
                    {
                        "status": "error",
                        "person": person,
                        "school": school,
                        "task_id": task_id,
                        "reason": f"download_failed: {exc}",
                    }
                )
                continue

            matched_path = workdir / "已匹配数据表.xlsx"
            unmatched_path = workdir / "未匹配数据表.xlsx"
            matched_original_path = workdir / "已匹配数据原始表.xlsx"

            try:
                excel_book_matcher.run_matching(
                    old_path=local_old,
                    new_path=local_new,
                    matched_path=matched_path,
                    unmatched_path=unmatched_path,
                    matched_original_path=matched_original_path,
                )
            except Exception as exc:  # noqa: BLE001
                _put_json(
                    bucket,
                    error_key,
                    {
                        "status": "error",
                        "time": _now_iso(),
                        "person": person,
                        "school": school,
                        "task_id": task_id,
                        "reason": f"Matching failed: {exc}",
                        "old_key": chosen_old,
                        "new_key": new_key,
                    },
                )
                task_summaries.append(
                    {
                        "status": "error",
                        "person": person,
                        "school": school,
                        "task_id": task_id,
                        "reason": f"match_failed: {exc}",
                    }
                )
                continue

            zip_name = f"match_results_{task_id}.zip"
            zip_path = workdir / zip_name
            try:
                _zip_files(zip_path, [matched_path, unmatched_path, matched_original_path])
            except Exception as exc:  # noqa: BLE001
                _put_json(
                    bucket,
                    error_key,
                    {
                        "status": "error",
                        "time": _now_iso(),
                        "person": person,
                        "school": school,
                        "task_id": task_id,
                        "reason": f"Zip failed: {exc}",
                    },
                )
                task_summaries.append(
                    {
                        "status": "error",
                        "person": person,
                        "school": school,
                        "task_id": task_id,
                        "reason": f"zip_failed: {exc}",
                    }
                )
                continue

            output_keys = {
                "zip": f"{results_prefix}{zip_name}",
                "matched_details": f"{results_prefix}{matched_path.name}",
                "unmatched_items": f"{results_prefix}{unmatched_path.name}",
                "matched_original": f"{results_prefix}{matched_original_path.name}",
            }

            try:
                _upload_file(bucket, output_keys["zip"], zip_path)
                _upload_file(bucket, output_keys["matched_details"], matched_path)
                _upload_file(bucket, output_keys["unmatched_items"], unmatched_path)
                _upload_file(bucket, output_keys["matched_original"], matched_original_path)
            except Exception as exc:  # noqa: BLE001
                _put_json(
                    bucket,
                    error_key,
                    {
                        "status": "error",
                        "time": _now_iso(),
                        "person": person,
                        "school": school,
                        "task_id": task_id,
                        "reason": f"Upload results failed: {exc}",
                        "output_keys": output_keys,
                    },
                )
                task_summaries.append(
                    {
                        "status": "error",
                        "person": person,
                        "school": school,
                        "task_id": task_id,
                        "reason": f"upload_failed: {exc}",
                    }
                )
                continue

            _put_json(
                bucket,
                done_key,
                {
                    "status": "ok",
                    "time": _now_iso(),
                    "person": person,
                    "school": school,
                    "task_id": task_id,
                    "bucket": bucket_name,
                    "region": region,
                    "inputs": {"old_key": chosen_old, "new_key": new_key},
                    "outputs": output_keys,
                },
            )

            # 发送钉钉通知
            try:
                _send_dingtalk_notification(person, task_id, school, bucket_name, region)
            except Exception as exc:  # noqa: BLE001
                print(f"DingTalk notification failed for {person}/{task_id}: {exc}")

            task_summaries.append(
                {
                    "status": "ok",
                    "person": person,
                    "school": school,
                    "task_id": task_id,
                    "outputs": output_keys,
                }
            )

    return {"status": "ok", "person": person, "tasks": task_summaries}


def handler(event: Any, context: Any) -> str:
    """
    Function Compute entrypoint.
    """
    try:
        locations = _parse_oss_event(event)
    except Exception as exc:  # noqa: BLE001
        return json.dumps({"status": "error", "reason": str(exc)}, ensure_ascii=False)

    results: list[dict[str, Any]] = []
    for loc in locations:
        results.append(_process_one(loc.bucket, loc.region, loc.key))
    return json.dumps({"status": "ok", "results": results}, ensure_ascii=False)

