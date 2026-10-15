"""Resolve the shared Flask delivery profile and embed it into the frontend."""

from __future__ import annotations

import argparse
from fnmatch import fnmatchcase
import html
import json
from pathlib import Path
import re
import sys
from typing import Any


DEFAULT_CONFIG = Path(__file__).with_name("distribution-profiles.json")
PROFILE_FIELDS = {
    "name": "name",
    "enabledSystems": "enabledSystems",
    "artifactPrefix": "artifactPrefix",
    "outputSuffix": "outputSuffix",
    "debPackage": "debPackage",
    "debConflicts": "debConflicts",
    "productScope": "productScope",
}


def load_config(path: Path) -> dict[str, Any]:
    try:
        config = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError(f"无法读取交付配置 {path}: {exc}") from exc

    if not isinstance(config, dict):
        raise ValueError("交付配置根节点必须是 JSON 对象。")
    if config.get("schemaVersion") != 1:
        raise ValueError("distribution-profiles.json 的 schemaVersion 必须为 1。")
    systems = config.get("systems")
    profiles = config.get("profiles")
    if not isinstance(systems, list) or not isinstance(profiles, dict) or not profiles:
        raise ValueError("交付配置必须包含 systems 数组和 profiles 对象。")

    if any(not isinstance(item, dict) for item in systems):
        raise ValueError("systems 必须全部由对象组成。")
    system_ids = [item.get("id") for item in systems]
    if any(not isinstance(value, str) for value in system_ids) or len(set(system_ids)) != len(system_ids):
        raise ValueError("systems 必须是 id 不重复的对象数组。")
    if any(not isinstance(value, str) or not re.fullmatch(r"S[1-4]", value) for value in system_ids):
        raise ValueError("系统编号只支持 S1、S2、S3、S4。")
    pages = [item.get("page") for item in systems]
    if any(not isinstance(value, str) or not re.fullmatch(r"[a-z][a-z0-9-]*", value) for value in pages):
        raise ValueError("每个系统都必须填写有效的前端 page 标识。")
    if len(set(pages)) != len(pages):
        raise ValueError("systems 中的 page 标识不能重复。")
    for item in systems:
        for key in ("faqCategories", "contentKeywords", "uiComponents"):
            values = item.get(key, [])
            if not isinstance(values, list) or any(not isinstance(value, str) for value in values):
                raise ValueError(f"{item['id']} 的 {key} 必须是字符串数组。")

    known = set(system_ids)
    for profile_id, profile in profiles.items():
        if not re.fullmatch(r"[a-z0-9][a-z0-9-]*", profile_id):
            raise ValueError(f"交付方案标识不合法：{profile_id!r}。")
        enabled = profile.get("enabledSystems") if isinstance(profile, dict) else None
        if not isinstance(enabled, list) or any(not isinstance(value, str) for value in enabled):
            raise ValueError(f"交付方案 {profile_id} 的 enabledSystems 必须是无重复数组。")
        if len(enabled) != len(set(enabled)):
            raise ValueError(f"交付方案 {profile_id} 的 enabledSystems 必须是无重复数组。")
        unknown = set(enabled) - known
        if unknown:
            raise ValueError(f"交付方案 {profile_id} 引用了未知系统：{', '.join(sorted(unknown))}。")
        for key in ("name", "artifactPrefix", "debPackage", "productScope"):
            if not isinstance(profile.get(key), str) or not profile[key]:
                raise ValueError(f"交付方案 {profile_id} 缺少有效字段：{key}。")
        if not isinstance(profile.get("outputSuffix"), str):
            raise ValueError(f"交付方案 {profile_id} 的 outputSuffix 必须是字符串。")
        patterns = profile.get("excludedConfigPatterns", [])
        if not isinstance(patterns, list) or any(
            not isinstance(pattern, str) or not pattern or "/" in pattern or "\\" in pattern
            for pattern in patterns
        ):
            raise ValueError(f"交付方案 {profile_id} 的 excludedConfigPatterns 必须是文件名模式数组。")
        if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9_-]*", profile["artifactPrefix"]):
            raise ValueError(f"交付方案 {profile_id} 的 artifactPrefix 只能包含英文字母、数字、下划线和连字符。")
        if not re.fullmatch(r"[A-Za-z0-9_-]*", profile["outputSuffix"]):
            raise ValueError(f"交付方案 {profile_id} 的 outputSuffix 只能包含英文字母、数字、下划线和连字符。")
        if not re.fullmatch(r"[a-z0-9][a-z0-9.+-]*", profile["debPackage"]):
            raise ValueError(f"交付方案 {profile_id} 的 debPackage 不符合 Debian 包名格式。")
        conflicts = profile.get("debConflicts")
        if not isinstance(conflicts, list) or not conflicts or any(
            not isinstance(value, str) or not re.fullmatch(r"[a-z0-9][a-z0-9.+-]*", value)
            for value in conflicts
        ):
            raise ValueError(f"交付方案 {profile_id} 的 debConflicts 必须是 Debian 包名数组。")

    return config


def get_profile(config: dict[str, Any], profile_id: str) -> dict[str, Any]:
    try:
        return config["profiles"][profile_id]
    except KeyError as exc:
        choices = ", ".join(sorted(config["profiles"]))
        raise ValueError(f"未知交付方案 {profile_id!r}；可用方案：{choices}。") from exc


def field_value(profile: dict[str, Any], field: str) -> str:
    if field not in PROFILE_FIELDS:
        raise ValueError(f"未知 profile 字段：{field}。")
    value = profile[PROFILE_FIELDS[field]]
    if isinstance(value, list):
        return ", ".join(str(item) for item in value)
    return str(value)


def filter_release_config(directory: Path, config: dict[str, Any], profile_id: str) -> None:
    """仅清理发行副本；递归涵盖默认配置和备份目录，绝不修改源配置。"""
    patterns = get_profile(config, profile_id).get("excludedConfigPatterns", [])
    if not patterns:
        return
    repository = DEFAULT_CONFIG.parent.parent
    root = directory.resolve()
    if directory.name != "config" or directory.is_symlink() or root in {
        (repository / "config").resolve(), (repository / "config-real").resolve()
    }:
        raise ValueError(f"只能筛选发行副本的 config 目录：{directory}")
    if not root.is_dir():
        raise ValueError(f"发行配置目录不存在：{directory}")
    excluded = [path for path in root.rglob("*") if path.is_file()
                and any(fnmatchcase(path.name, pattern) for pattern in patterns)]
    for path in excluded:
        if root not in path.resolve().parents or path.is_symlink():
            raise ValueError(f"发行配置包含越界链接：{path}")
    for path in excluded:
        path.unlink()
        print(f"[配置排除] {path.relative_to(root)}")


def embed_profile(html_path: Path, config: dict[str, Any], profile_id: str) -> None:
    profile = get_profile(config, profile_id)
    source = html_path.read_text(encoding="utf-8")
    source = re.sub(
        r'<script id="audit-distribution-config" type="application/json">.*?</script>\s*',
        "",
        source,
        flags=re.DOTALL,
    )

    systems = config["systems"]
    enabled = profile["enabledSystems"]
    expected_pages = {item["page"] for item in systems}
    menu_pages = set(re.findall(r'<button\b[^>]*\bdata-page="([^"]+)"', source))
    missing = expected_pages - menu_pages
    if missing:
        raise ValueError(f"前端缺少交付配置声明的菜单页面：{', '.join(sorted(missing))}。")

    payload = {
        "profile": profile_id,
        "enabledSystems": enabled,
        "systems": systems,
    }
    serialized = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    serialized = serialized.replace("<", "\\u003c")
    body = re.search(r"<body(?:\s[^>]*)?>", source, flags=re.IGNORECASE)
    if body is None:
        raise ValueError("前端 HTML 缺少 <body> 标签。")
    body_tag = body.group(0)
    body_tag = re.sub(r'\sdata-enabled-systems="[^"]*"', "", body_tag)
    body_tag = re.sub(r'\sdata-distribution-profile="[^"]*"', "", body_tag)
    body_tag = body_tag[:-1] + (
        f' data-enabled-systems="{html.escape(" ".join(enabled), quote=True)}"'
        f' data-distribution-profile="{html.escape(profile_id, quote=True)}">'
    )
    injected = body_tag + f'\n<script id="audit-distribution-config" type="application/json">{serialized}</script>'
    source = source[: body.start()] + injected + source[body.end() :]

    for system in systems:
        page = re.escape(system["page"])
        pattern = re.compile(rf'(<button\b[^>]*\bdata-page="{page}"[^>]*)(>)')
        source, count = pattern.subn(
            lambda match, item=system: re.sub(
                r'\sdata-system="[^"]*"', "", match.group(1)
            ) + f' data-system="{item["id"]}"' + match.group(2),
            source,
            count=1,
        )
        if count != 1:
            raise ValueError(f"无法将 {system['id']} 标记到前端导航菜单。")

    html_path.write_text(source, encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("command", choices=("field", "embed-html", "filter-config"))
    parser.add_argument("profile", help="distribution-profiles.json 中的方案标识")
    parser.add_argument("value", help="字段名、HTML 路径，或 filter-config 使用的发行 config 目录")
    parser.add_argument("--config", type=Path, default=DEFAULT_CONFIG)
    args = parser.parse_args()

    try:
        config = load_config(args.config)
        profile = get_profile(config, args.profile)
        if args.command == "field":
            print(field_value(profile, args.value))
        elif args.command == "filter-config":
            filter_release_config(Path(args.value), config, args.profile)
        else:
            embed_profile(Path(args.value), config, args.profile)
    except (OSError, ValueError) as exc:
        print(f"[错误] {exc}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
