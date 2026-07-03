from __future__ import annotations

import argparse
import re
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path
from uuid import uuid4

import pandas as pd

POI_COLUMNS = [
    "id",
    "poi_id",
    "code",
    "name_cn",
    "name_en",
    "category",
    "city",
    "district",
    "duration_hours",
    "price_min",
    "price_max",
    "currency",
    "tags",
    "suitable_for",
    "description",
    "description_cn",
    "description_en",
    "recommended_reason",
    "highlights",
    "target_users",
    "image_placeholder",
    "image_notes",
    "image_source",
    "image_alt",
    "image_url",
    "image_path",
    "images",
    "address",
    "opening_hours",
    "reservation_info",
    "notes",
    "internal_remark",
    "status",
]

FIELD_LABELS = {
    "id": "内部ID",
    "poi_id": "点位编号",
    "code": "点位编码",
    "name_cn": "中文名称",
    "name_en": "英文名称",
    "category": "点位类型",
    "city": "城市",
    "district": "区域",
    "duration_hours": "建议游览时长（小时）",
    "price_min": "最低成本价",
    "price_max": "最高成本价",
    "currency": "币种",
    "tags": "标签",
    "suitable_for": "适合人群",
    "description": "综合描述",
    "description_cn": "中文介绍",
    "description_en": "英文介绍",
    "recommended_reason": "推荐理由",
    "highlights": "亮点",
    "target_users": "目标客群",
    "image_placeholder": "图片占位符或图片路径",
    "image_notes": "图片备注",
    "image_source": "图片来源",
    "image_alt": "图片说明",
    "image_url": "网络图片地址",
    "image_path": "本地图片路径",
    "images": "图片组",
    "address": "地址",
    "opening_hours": "开放时间",
    "reservation_info": "预约信息",
    "notes": "内部备注",
    "internal_remark": "内部备注补充",
    "status": "状态",
}

ALIASES = {
    "id": ["id", "ID", "内部ID", "唯一ID"],
    "poi_id": ["poi_id", "poi id", "POI_ID", "编号", "点位编号", "poi编号"],
    "code": ["code", "点位编码", "编码"],
    "name_cn": ["点位名称", "名称", "POI名称", "poi名称", "中文名称", "中文名"],
    "name_en": ["英文名", "英文名称", "英文名称/英文名", "name_en", "English Name"],
    "category": ["类型", "分类", "产品类型", "点位类型", "一级模块", "二级类别", "模块"],
    "city": ["城市", "地点", "所在城市", "目的地"],
    "district": ["区域", "所在区域", "商圈", "区县"],
    "duration_hours": ["时长_小时", "建议游览时长", "推荐时长", "时长", "游览时长", "体验时长"],
    "price_min": ["成本下限_元", "最低成本价", "成本", "成本价", "价格", "门票价格", "成本价_元"],
    "price_max": ["成本上限_元", "最高成本价", "成本", "成本价", "价格", "门票价格", "成本价_元"],
    "currency": ["币种", "货币"],
    "tags": ["标签", "多维标签", "适合场景", "组合建议标签", "兴趣类型", "线路定位"],
    "suitable_for": ["人群", "适合人群", "客户类型", "所属线路/适用场景", "适用客群"],
    "description": ["综合描述", "description"],
    "description_cn": ["简介", "中文介绍", "介绍", "价格说明", "线路特色"],
    "description_en": ["英文介绍", "英文简介", "description_en"],
    "recommended_reason": ["推荐理由", "推荐原因", "recommended_reason"],
    "highlights": ["亮点", "项目亮点", "highlights"],
    "target_users": ["目标客群", "target_users", "适合客户"],
    "image_placeholder": ["图片占位符", "image_placeholder"],
    "image_path": ["图片", "图片路径", "本地图片", "本地图片路径", "图片地址", "image", "image_path", "photo", "photo_path"],
    "image_url": ["图片网络地址", "网络图片", "网络图片地址", "image_url", "photo_url", "url"],
    "image_alt": ["图片说明", "图片英文说明", "image_alt", "alt"],
    "image_source": ["图片来源", "image_source", "source"],
    "image_notes": ["图片备注", "图片用途", "image_notes"],
    "images": ["图片组", "images"],
    "address": ["地址", "详细地址", "地点"],
    "opening_hours": ["开放时间", "营业时间", "opening_hours"],
    "reservation_info": ["预约信息", "预订信息", "reservation_info"],
    "notes": ["备注", "内部备注", "来源", "价格说明"],
    "internal_remark": ["内部备注补充", "internal_remark", "内部说明"],
    "status": ["状态", "status"],
}

EXTRA_TAG_COLUMNS = ["组合建议标签", "兴趣类型", "节奏", "线路定位", "时长筛选", "医疗级别/体检等级", "检查范围", "是否必选/可选"]
EXTRA_NOTE_COLUMNS = ["备注", "来源", "价格说明", "价格状态", "价格显示", "计价单位", "子项/机构"]


@dataclass
class ImportResult:
    dataframe: pd.DataFrame
    mapping: dict[str, str]
    source_columns: list[str]
    missing_required_count: int
    empty_price_count: int
    duplicate_count: int


def clean_text(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.lower() in {"nan", "none", "null"}:
        return ""
    return text


def normalized_header(value) -> str:
    return re.sub(r"[\s_\-\/（）()]+", "", clean_text(value).lower())


def choose_sheet(path_or_buffer) -> str | int:
    xl = pd.ExcelFile(path_or_buffer)
    for preferred in ["POI数据库", "POI 数据库", "源点位表"]:
        if preferred in xl.sheet_names:
            return preferred
    for sheet_name in xl.sheet_names:
        if "readme" not in sheet_name.lower() and "字段" not in sheet_name:
            return sheet_name
    return 0


def read_source_table(path_or_buffer, filename: str | None = None) -> pd.DataFrame:
    suffix = Path(filename or str(path_or_buffer)).suffix.lower()
    if suffix in [".xlsx", ".xls"]:
        sheet = choose_sheet(path_or_buffer)
        return pd.read_excel(path_or_buffer, sheet_name=sheet)
    return pd.read_csv(path_or_buffer)


def build_mapping(columns) -> dict[str, str]:
    columns = [clean_text(col) for col in columns]
    normalized_to_original = {normalized_header(col): col for col in columns}
    mapping: dict[str, str] = {}
    for target, aliases in ALIASES.items():
        for alias in aliases:
            key = normalized_header(alias)
            if key in normalized_to_original:
                mapping[target] = normalized_to_original[key]
                break
    return mapping


def get_value(row, column_name: str | None):
    if not column_name or column_name not in row.index:
        return ""
    return row[column_name]


def parse_duration(value) -> float | None:
    text = clean_text(value)
    if not text:
        return None
    if "半日" in text and not re.search(r"\d+(?:\.\d+)?", text):
        return 4
    match = re.search(r"\d+(?:\.\d+)?", text)
    if not match:
        return None
    number = float(match.group())
    if "半日" in text:
        return 4 if number <= 0.5 else number
    if "日" in text and "小时" not in text and "h" not in text.lower():
        return number * 8
    return number


def parse_price_values(*values) -> tuple[float | None, float | None]:
    text = " ".join(clean_text(value) for value in values if clean_text(value))
    if not text:
        return None, None
    text = text.replace(",", "")
    numbers = [float(item) for item in re.findall(r"\d+(?:\.\d+)?", text)]
    if not numbers:
        return None, None
    if len(numbers) == 1:
        return numbers[0], numbers[0]
    return min(numbers), max(numbers)


def split_image_reference(value) -> tuple[str, str]:
    text = clean_text(value)
    if not text:
        return "", ""
    if text.lower().startswith(("http://", "https://")):
        return "", text
    return text, ""


def normalize_multi_value(value) -> str:
    text = clean_text(value)
    if not text:
        return ""
    for sep in ["、", "，", ",", "/", "|", "；", ";", "\n"]:
        text = text.replace(sep, ";")
    parts = []
    for item in text.split(";"):
        item = item.strip()
        if item and item not in parts:
            parts.append(item)
    return ";".join(parts)


def join_fields(row, columns) -> str:
    parts = []
    for col in columns:
        if col in row.index:
            value = clean_text(row[col])
            if value and value not in parts:
                parts.append(value)
    return ";".join(parts)


def split_location(value) -> tuple[str, str]:
    text = clean_text(value)
    if not text:
        return "", ""
    parts = [part.strip() for part in re.split(r"[/,，;；]", text) if part.strip()]
    if not parts:
        return text, ""
    return parts[0], ";".join(parts[1:])


def normalize_poi_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for column in POI_COLUMNS:
        if column not in df.columns:
            df[column] = ""
    df = df[POI_COLUMNS].fillna("")
    for column in ["duration_hours", "price_min", "price_max"]:
        df[column] = pd.to_numeric(df[column], errors="coerce")
    df["currency"] = df["currency"].apply(lambda value: clean_text(value) or "CNY")
    df["tags"] = df["tags"].apply(normalize_multi_value)
    df["suitable_for"] = df["suitable_for"].apply(normalize_multi_value)
    missing_id = df["poi_id"].astype(str).str.strip() == ""
    code_as_id = missing_id & (df["code"].astype(str).str.strip() != "")
    df.loc[code_as_id, "poi_id"] = df.loc[code_as_id, "code"]
    missing_id = df["poi_id"].astype(str).str.strip() == ""
    df.loc[missing_id, "poi_id"] = [f"POI-{uuid4().hex[:8].upper()}" for _ in range(missing_id.sum())]
    df["id"] = df["id"].astype(str).str.strip()
    missing_internal_id = df["id"] == ""
    df.loc[missing_internal_id, "id"] = df.loc[missing_internal_id, "poi_id"]
    df["code"] = df["code"].astype(str).str.strip()
    missing_code = df["code"] == ""
    df.loc[missing_code, "code"] = df.loc[missing_code, "poi_id"]
    df["status"] = df["status"].apply(lambda value: "deleted" if clean_text(value).lower() == "deleted" else "active")
    missing_images = df["images"].astype(str).str.strip() == ""
    df.loc[missing_images, "images"] = "[]"
    return df


def transform_source_to_poi(source_df: pd.DataFrame) -> ImportResult:
    source_df = source_df.dropna(how="all").copy()
    mapping = build_mapping(source_df.columns)
    rows = []
    seen_keys = set()
    duplicate_count = 0

    for _, row in source_df.iterrows():
        city, district_from_location = split_location(get_value(row, mapping.get("city")))
        explicit_district = clean_text(get_value(row, mapping.get("district")))
        price_min, price_max = parse_price_values(
            get_value(row, mapping.get("price_min")),
            get_value(row, mapping.get("price_max")),
            row.get("价格显示", "") if "价格显示" in row.index else "",
        )
        duration = parse_duration(get_value(row, mapping.get("duration_hours")))
        category = clean_text(get_value(row, mapping.get("category")))
        if not category and "一级模块" in row.index:
            category = clean_text(row["一级模块"])

        name_cn = clean_text(get_value(row, mapping.get("name_cn")))
        sub_item = clean_text(row.get("子项/机构", "")) if "子项/机构" in row.index else ""
        if not name_cn and sub_item:
            name_cn = sub_item

        tags = normalize_multi_value(";".join([clean_text(get_value(row, mapping.get("tags"))), join_fields(row, EXTRA_TAG_COLUMNS)]))
        suitable_for = normalize_multi_value(clean_text(get_value(row, mapping.get("suitable_for"))))
        description_cn = clean_text(get_value(row, mapping.get("description_cn")))
        notes = normalize_multi_value(";".join([clean_text(get_value(row, mapping.get("notes"))), join_fields(row, EXTRA_NOTE_COLUMNS)]))

        raw_image_path = clean_text(get_value(row, mapping.get("image_path")))
        path_from_generic, url_from_generic = split_image_reference(raw_image_path)
        explicit_image_url = clean_text(get_value(row, mapping.get("image_url")))

        poi_id = clean_text(get_value(row, mapping.get("poi_id"))) or clean_text(get_value(row, mapping.get("code"))) or f"POI-{uuid4().hex[:8].upper()}"
        poi = {
            "id": clean_text(get_value(row, mapping.get("id"))) or poi_id,
            "poi_id": poi_id,
            "code": clean_text(get_value(row, mapping.get("code"))) or poi_id,
            "name_cn": name_cn,
            "name_en": clean_text(get_value(row, mapping.get("name_en"))),
            "category": category,
            "city": city,
            "district": explicit_district or district_from_location,
            "duration_hours": duration,
            "price_min": price_min,
            "price_max": price_max,
            "currency": clean_text(get_value(row, mapping.get("currency"))) or "CNY",
            "tags": tags,
            "suitable_for": suitable_for,
            "description": clean_text(get_value(row, mapping.get("description"))) or description_cn,
            "description_cn": description_cn,
            "description_en": clean_text(get_value(row, mapping.get("description_en"))),
            "recommended_reason": clean_text(get_value(row, mapping.get("recommended_reason"))),
            "highlights": clean_text(get_value(row, mapping.get("highlights"))),
            "target_users": clean_text(get_value(row, mapping.get("target_users"))) or suitable_for,
            "image_placeholder": clean_text(get_value(row, mapping.get("image_placeholder"))),
            "image_path": path_from_generic,
            "image_url": explicit_image_url or url_from_generic,
            "image_alt": clean_text(get_value(row, mapping.get("image_alt"))),
            "image_source": clean_text(get_value(row, mapping.get("image_source"))),
            "image_notes": clean_text(get_value(row, mapping.get("image_notes"))),
            "images": clean_text(get_value(row, mapping.get("images"))) or "[]",
            "address": clean_text(get_value(row, mapping.get("address"))) or clean_text(get_value(row, mapping.get("city"))),
            "opening_hours": clean_text(get_value(row, mapping.get("opening_hours"))),
            "reservation_info": clean_text(get_value(row, mapping.get("reservation_info"))),
            "notes": notes,
            "internal_remark": clean_text(get_value(row, mapping.get("internal_remark"))),
            "status": clean_text(get_value(row, mapping.get("status"))) or "active",
        }

        dedupe_key = (poi["poi_id"], poi["name_cn"], poi["city"], poi["category"])
        if dedupe_key in seen_keys:
            duplicate_count += 1
            continue
        seen_keys.add(dedupe_key)
        rows.append(poi)

    output_df = normalize_poi_df(pd.DataFrame(rows))
    missing_required_count = int((output_df["name_cn"].astype(str).str.strip() == "").sum())
    empty_price_count = int(output_df["price_min"].isna().sum())
    return ImportResult(
        dataframe=output_df,
        mapping=mapping,
        source_columns=[clean_text(col) for col in source_df.columns],
        missing_required_count=missing_required_count,
        empty_price_count=empty_price_count,
        duplicate_count=duplicate_count,
    )


def analyze_poi_file(path_or_bytes, filename: str | None = None) -> ImportResult:
    if isinstance(path_or_bytes, (bytes, bytearray)):
        source = BytesIO(path_or_bytes)
    else:
        source = path_or_bytes
    source_df = read_source_table(source, filename=filename)
    return transform_source_to_poi(source_df)


def merge_poi_data(existing_df: pd.DataFrame, imported_df: pd.DataFrame, overwrite: bool = False) -> pd.DataFrame:
    existing = normalize_poi_df(existing_df)
    imported = normalize_poi_df(imported_df)
    if overwrite:
        combined = pd.concat([existing, imported], ignore_index=True)
        combined = combined.drop_duplicates(subset=["poi_id"], keep="last")
    else:
        existing_ids = set(existing["poi_id"].astype(str))
        imported = imported[~imported["poi_id"].astype(str).isin(existing_ids)]
        combined = pd.concat([existing, imported], ignore_index=True)
    return normalize_poi_df(combined)


def main():
    parser = argparse.ArgumentParser(description="导入 HZH POI Excel/CSV 数据到工具数据库")
    parser.add_argument("source", help="Excel 或 CSV 文件路径")
    parser.add_argument("--output", default="data/poi_database.csv", help="输出 CSV 数据库路径")
    parser.add_argument("--mode", choices=["replace", "append", "overwrite"], default="replace")
    args = parser.parse_args()

    result = analyze_poi_file(args.source, filename=args.source)
    output_path = Path(args.output)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    if output_path.exists() and args.mode in {"append", "overwrite"}:
        existing = pd.read_csv(output_path)
        final_df = merge_poi_data(existing, result.dataframe, overwrite=args.mode == "overwrite")
    else:
        final_df = result.dataframe

    final_df.to_csv(output_path, index=False)
    print(f"source_rows={len(result.dataframe)}")
    print(f"saved_rows={len(final_df)}")
    print(f"missing_required_count={result.missing_required_count}")
    print(f"empty_price_count={result.empty_price_count}")
    print(f"duplicate_count={result.duplicate_count}")
    print("mapping=")
    for target, source in result.mapping.items():
        print(f"  {target} <- {source}")


if __name__ == "__main__":
    main()
