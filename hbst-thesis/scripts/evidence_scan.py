#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
evidence_scan.py — 阶段一：源码证据扫描辅助（hbst-thesis skill）

对给定源码目录做轻量证据梳理，输出 paper-context/evidence.json。
定位是"辅助记忆与核对"，不替代逐文件通读；高价值证据仍需人工/Agent 读源码确认。

用法:
    python evidence_scan.py <project-dir> [--out <evidence.json>] [--max-depth N] [--skip 目录名...]

输出 JSON 结构:
{
  "tech_stack": [...],
  "modules": [...],
  "roles": [...],
  "api_endpoints": [...],
  "db_tables": [...],
  "screenshots": [...],
  "test_artifacts": [...],
  "reads": [...]
}
"""
import argparse
import json
import os
import re
from collections import Counter

# 常见后端技术关键词 -> 归一化技术名
TECH_KEYWORDS = {
    "spring boot": "Spring Boot", "springboot": "Spring Boot", "spring": "Spring",
    "spring cloud": "Spring Cloud", "mybatis": "MyBatis", "mybatis-plus": "MyBatis-Plus",
    "mybatisplus": "MyBatis-Plus", "hibernate": "Hibernate", "jpa": "Spring Data JPA",
    "vue": "Vue", "vue2": "Vue2", "vue3": "Vue3", "element-ui": "Element UI",
    "element ui": "Element UI", "react": "React", "uniapp": "uni-app", "uni-app": "uni-app",
    "微信小程序": "微信小程序", "小程序": "微信小程序", "wxml": "微信小程序",
    "flask": "Flask", "django": "Django", "fastapi": "FastAPI", "express": "Express",
    "node.js": "Node.js", "mysql": "MySQL", "redis": "Redis", "mongodb": "MongoDB",
    "postgresql": "PostgreSQL", "sqlite": "SQLite", "oracle": "Oracle",
    "jwt": "JWT", "shiro": "Apache Shiro", "oauth2": "OAuth2", "sa-token": "Sa-Token",
    "docker": "Docker", "nacos": "Nacos", "rabbitmq": "RabbitMQ", "kafka": "Kafka",
    "maven": "Maven", "gradle": "Gradle", "npm": "npm", "weui": "WeUI",
    "vant": "Vant", "echarts": "ECharts", "coze": "Coze", "openai": "OpenAI",
    "llm": "LLM", "大模型": "大模型",
}

EXT_KIND = {
    ".java": "java", ".kt": "kotlin", ".py": "python", ".js": "js", ".ts": "ts",
    ".vue": "vue", ".wxml": "wxml", ".wxss": "wxss", ".json": "json", ".xml": "xml",
    ".yml": "yml", ".yaml": "yaml", ".properties": "properties", ".sql": "sql",
    ".go": "go", ".php": "php", ".cs": "csharp", ".c": "c", ".cpp": "cpp", ".html": "html",
    ".css": "css", ".md": "md", ".sh": "shell", ".ps1": "powershell",
}

# 常见图片扩展名（截图登记）
IMG_EXT = {".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp"}


def scan_files(root, skip_dirs, max_depth):
    """yield (path, rel, ext) 按深度限制遍历，跳过 .git/node_modules/target 等。"""
    skip_dirs = set(skip_dirs or []) | {
        ".git", ".idea", ".vscode", "node_modules", "target", "dist", "build",
        "out", ".venv", "venv", "__pycache__", ".next", ".nuxt", "logs",
    }
    root = os.path.abspath(root)
    for base, dirs, files in os.walk(root):
        rel = os.path.relpath(base, root)
        depth = 0 if rel == "." else rel.count(os.sep) + 1
        if depth >= max_depth:
            dirs[:] = []
            continue
        dirs[:] = [d for d in dirs if d not in skip_dirs]
        for fn in files:
            ext = os.path.splitext(fn)[1].lower()
            yield os.path.join(base, fn), os.path.relpath(os.path.join(base, fn), root), ext


def read_small(path, limit=200 * 1024):
    try:
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            return f.read(limit)
    except Exception:
        return ""


def detect_tech(files, root):
    tech = Counter()
    text_buf = []
    for path, rel, ext in files:
        if ext not in EXT_KIND:
            continue
        if ext in (".sql", ".xml", ".yml", ".yaml", ".properties", ".json"):
            text_buf.append((rel, read_small(path)))
    # 从依赖清单文件优先
    dep_files = ["pom.xml", "build.gradle", "package.json", "requirements.txt",
                 "go.mod", "composer.json", "pom.xml"]
    seen_dep = set()
    for name in dep_files:
        if name in seen_dep:
            continue
        # 直接按 basename 找（可能嵌套在子模块）
        continue
    # 汇总文本检测
    blob = "\n".join(t for _, t in text_buf)
    for kw, name in TECH_KEYWORDS.items():
        if kw in blob:
            tech[name] += blob.count(kw)
    # pom.xml / package.json 定向读取
    for path, rel, ext in files:
        base = os.path.basename(rel)
        if base in ("pom.xml", "package.json", "requirements.txt", "build.gradle"):
            content = read_small(path)
            for kw, name in TECH_KEYWORDS.items():
                if kw in content:
                    tech[name] += 1
    return tech


def find_controllers_and_routes(files, root):
    """粗略识别 controller/路由 与 url 映射。"""
    routes = []
    ctrl_re = re.compile(r"@(?:RequestMapping|GetMapping|PostMapping|PutMapping|DeleteMapping)\s*\(\s*[\"']([^\"']+)", re.I)
    flask_re = re.compile(r"@\w+\.route\s*\(\s*[\"']([^\"']+)", re.I)
    for path, rel, ext in files:
        if ext in (".java", ".py", ".js", ".ts"):
            content = read_small(path)
            hits = ctrl_re.findall(content) + flask_re.findall(content)
            for h in hits:
                if h.startswith("/"):
                    routes.append(h)
                else:
                    routes.append("/" + h if not h.startswith("/") else h)
    return sorted(set(routes))[:500]


def find_db_tables(files, root):
    """从 .sql 建表语句与实体注解提取表名。"""
    tables = []
    sql_re = re.compile(r"create\s+table\s+(?:if\s+not\s+exists\s+)?`?(\w+)`?", re.I)
    ann_re = re.compile(r"@TableName\s*\(\s*[\"']([^\"']+)", re.I)
    for path, rel, ext in files:
        if ext == ".sql":
            content = read_small(path)
            tables += sql_re.findall(content)
        elif ext == ".java":
            content = read_small(path)
            tables += ann_re.findall(content)
    return sorted(set(tables))[:300]


def find_screenshots(files, root):
    imgs = []
    for path, rel, ext in files:
        if ext in IMG_EXT:
            imgs.append(rel)
    return imgs


def find_test_artifacts(files, root):
    """识别测试代码与测试报告。"""
    tests = []
    pat = re.compile(r"(test|spec)", re.I)
    for path, rel, ext in files:
        if ext not in EXT_KIND:
            continue
        if ext == ".md" and ("test" in rel.lower() or "测试" in rel):
            tests.append({"kind": "report", "path": rel})
        elif ext in (".java", ".py", ".js", ".ts"):
            low = rel.lower()
            if "/test/" in "/" + low or low.startswith("test") or "spec." in low:
                tests.append({"kind": "code", "path": rel})
    return tests[:200]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("project_dir")
    ap.add_argument("--out", default="evidence.json")
    ap.add_argument("--max-depth", type=int, default=8)
    ap.add_argument("--skip", nargs="*", default=[])
    args = ap.parse_args()

    files = list(scan_files(args.project_dir, args.skip, args.max_depth))
    tech = detect_tech(files, args.project_dir)
    result = {
        "tech_stack": [{"name": k, "hits": v} for k, v in tech.most_common()],
        "modules": [],  # 需要 Agent 通读后手工补充
        "roles": [],
        "api_endpoints": find_controllers_and_routes(files, args.project_dir),
        "db_tables": find_db_tables(files, args.project_dir),
        "screenshots": find_screenshots(files, args.project_dir),
        "test_artifacts": find_test_artifacts(files, args.project_dir),
        "file_stats": {
            "total_files": len(files),
            "by_ext": dict(Counter(ext for _, _, ext in files).most_common(20)),
        },
    }
    os.makedirs(os.path.dirname(os.path.abspath(args.out)) or ".", exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f"[evidence_scan] wrote {args.out}")
    print(f"[evidence_scan] tech_stack={result['tech_stack'][:12]}")
    print(f"[evidence_scan] routes={len(result['api_endpoints'])} db_tables={len(result['db_tables'])} "
          f"screenshots={len(result['screenshots'])} test_artifacts={len(result['test_artifacts'])}")


if __name__ == "__main__":
    main()
