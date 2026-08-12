from __future__ import annotations

from dataclasses import dataclass
import json
import re
import urllib.error
import urllib.parse
import urllib.request

import pandas as pd

from poi_importer import ImportResult, build_mapping, transform_source_to_poi


FEISHU_BASE_URL = "https://open.feishu.cn/open-apis"


@dataclass
class FeishuSyncResult:
    import_result: ImportResult
    source_dataframe: pd.DataFrame
    record_count: int


def clean_feishu_value(value):
    """Convert Feishu Bitable cell values into import-friendly text."""
    if value is None:
        return ""
    if isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, list):
        parts = [clean_feishu_value(item) for item in value]
        parts = [str(part).strip() for part in parts if str(part).strip()]
        return ";".join(parts)
    if isinstance(value, dict):
        if "text" in value:
            return value.get("text") or ""
        if "name" in value:
            return value.get("name") or ""
        if "url" in value:
            return value.get("url") or ""
        if "link" in value:
            return value.get("link") or ""
        if "file_token" in value:
            return value.get("name") or value.get("file_token") or ""
        return json.dumps(value, ensure_ascii=False)
    return str(value)


def feishu_json_request(url: str, payload: dict | None = None, token: str | None = None, method: str = "GET") -> dict:
    data = json.dumps(payload, ensure_ascii=False).encode("utf-8") if payload is not None else None
    headers = {"Content-Type": "application/json; charset=utf-8"}
    if token:
        headers["Authorization"] = f"Bearer {token}"
    request = urllib.request.Request(url, data=data, headers=headers, method=method)
    try:
        with urllib.request.urlopen(request, timeout=30) as response:
            return json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="ignore")
        raise RuntimeError(f"飞书接口返回错误：HTTP {exc.code} {detail}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"无法连接飞书接口：{exc.reason}") from exc


def get_tenant_access_token(app_id: str, app_secret: str) -> str:
    url = f"{FEISHU_BASE_URL}/auth/v3/tenant_access_token/internal"
    response = feishu_json_request(url, {"app_id": app_id, "app_secret": app_secret}, method="POST")
    if response.get("code") != 0:
        raise RuntimeError(f"获取飞书访问令牌失败：{response.get('msg') or response}")
    token = response.get("tenant_access_token")
    if not token:
        raise RuntimeError("飞书没有返回 tenant_access_token，请检查 App ID / App Secret。")
    return token


def extract_wiki_node_token(value: str) -> str:
    text = (value or "").strip()
    if not text:
        return ""
    match = re.search(r"/wiki/([^/?#]+)", text)
    if match:
        return match.group(1)
    if re.fullmatch(r"[A-Za-z0-9_-]{12,}", text) and not text.startswith(("base", "bascn", "app")):
        return text
    return ""


def extract_bitable_app_token(value: str) -> str:
    text = (value or "").strip()
    if not text:
        return ""
    match = re.search(r"/base/([^/?#]+)", text)
    if match:
        return match.group(1)
    match = re.search(r"/(?:base|bitable)/([^/?#]+)", text)
    if match:
        return match.group(1)
    if text.startswith(("base", "bascn", "app")) and "/" not in text and "?" not in text:
        return text
    return ""


def resolve_bitable_app_token(app_token_or_url: str, tenant_access_token: str) -> str:
    direct_token = extract_bitable_app_token(app_token_or_url)
    if direct_token:
        return direct_token
    wiki_token = extract_wiki_node_token(app_token_or_url)
    if not wiki_token:
        return (app_token_or_url or "").strip()

    # Feishu wiki nodes store the real document token in obj_token. For a Bitable node, obj_token is the Bitable app_token.
    url = f"{FEISHU_BASE_URL}/wiki/v2/spaces/get_node?token={urllib.parse.quote(wiki_token)}"
    response = feishu_json_request(url, token=tenant_access_token)
    if response.get("code") != 0:
        raise RuntimeError(f"解析飞书 wiki 链接失败：{response.get('msg') or response}")
    data = response.get("data") or {}
    node = data.get("node") or data
    obj_type = str(node.get("obj_type") or node.get("node_type") or "").lower()
    obj_token = node.get("obj_token") or node.get("obj_token_id") or ""
    if obj_type and "bitable" not in obj_type and "base" not in obj_type:
        raise RuntimeError(f"该 wiki 节点不是多维表格，当前类型为 {obj_type}。请确认复制的是嵌入的多维表格链接。")
    if not obj_token:
        raise RuntimeError("已读取 wiki 节点，但没有拿到多维表格 obj_token。请确认应用有知识库读取权限。")
    return obj_token


def fetch_bitable_records(app_token: str, table_id: str, tenant_access_token: str, view_id: str = "", page_size: int = 500) -> list[dict]:
    records: list[dict] = []
    page_token = ""
    while True:
        params = {"page_size": str(page_size)}
        if page_token:
            params["page_token"] = page_token
        if view_id:
            params["view_id"] = view_id
        query = urllib.parse.urlencode(params)
        url = f"{FEISHU_BASE_URL}/bitable/v1/apps/{app_token}/tables/{table_id}/records?{query}"
        response = feishu_json_request(url, token=tenant_access_token)
        if response.get("code") != 0:
            raise RuntimeError(f"读取飞书多维表格失败：{response.get('msg') or response}")
        data = response.get("data") or {}
        records.extend(data.get("items") or [])
        if not data.get("has_more"):
            break
        page_token = data.get("page_token") or ""
        if not page_token:
            break
    return records


def records_to_dataframe(records: list[dict]) -> pd.DataFrame:
    rows = []
    for record in records:
        fields = record.get("fields") or {}
        row = {field_name: clean_feishu_value(value) for field_name, value in fields.items()}
        row.setdefault("飞书记录ID", record.get("record_id", ""))
        rows.append(row)
    return pd.DataFrame(rows)


def sync_feishu_pois(app_id: str, app_secret: str, app_token: str, table_id: str, view_id: str = "", custom_mapping: dict[str, str] | None = None) -> FeishuSyncResult:
    if not all([app_id, app_secret, app_token, table_id]):
        raise ValueError("请先配置飞书 App ID、App Secret、app_token 或 wiki 链接，以及 table_id。")
    tenant_token = get_tenant_access_token(app_id.strip(), app_secret.strip())
    resolved_app_token = resolve_bitable_app_token(app_token.strip(), tenant_token)
    records = fetch_bitable_records(resolved_app_token, table_id.strip(), tenant_token, view_id=view_id.strip())
    source_df = records_to_dataframe(records)
    import_result = transform_source_to_poi(source_df, custom_mapping=custom_mapping)
    return FeishuSyncResult(import_result=import_result, source_dataframe=source_df, record_count=len(records))
