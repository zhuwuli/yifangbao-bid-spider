from __future__ import annotations

import argparse
import base64
import copy
import json
import os
import re
import shutil
import sys
import time
import uuid
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any
from urllib.parse import urlencode
from urllib.request import Request, urlopen

from openpyxl import load_workbook
from openpyxl.styles import Font


BASE = "https://qiye.qianlima.com/new_qd_yfbsite/api"
REFERER = "https://qiye.qianlima.com/new_qd_yfbsite/#/infoCenter/search"
AREA_IDS = "1738,1740"  # 济南, 莱芜
KEYWORD = "基坑监测"
SEARCH_KEYWORDS = ("监测", "水土保持", "测绘", "测量")
FILTER_CONDITIONS = (1, 2)  # 1=全文检索，2=标题检索；两种都跑，合并去重。
SEARCH_TYPES = (1, 2)  # 1=智能检索，2=精准检索；两种都跑，合并去重。
SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_XLSX = SCRIPT_DIR / "2026-乙方宝招标信息统计.xlsx"
DETAIL_SHEET_NAME = "公告详情"
QUALIFICATION_PLACEHOLDER = "公告资质"
REMARK_PLACEHOLDER = "备注：未完整获取公告正文或未识别到明确资格要求，请人工核对。公告内容"


class YfbAuthError(RuntimeError):
    pass


def request_json(path: str, params: dict[str, Any], headers: dict[str, str], timeout: int = 25) -> dict[str, Any]:
    url = f"{BASE}{path}?{urlencode({k: v for k, v in params.items() if v is not None})}"
    last_error: Exception | None = None
    for attempt in range(3):
        req = Request(url, headers=headers)
        try:
            with urlopen(req, timeout=timeout) as resp:
                text = resp.read().decode("utf-8", "ignore")
            data = json.loads(text)
            if data.get("code") == 401:
                raise YfbAuthError(data.get("msg") or "乙方宝接口认证失败，请提供登录态")
            if data.get("code") not in (None, 200):
                raise RuntimeError(f"接口返回异常: {data}")
            return data
        except YfbAuthError:
            raise
        except Exception as exc:
            last_error = exc
            if attempt < 2:
                time.sleep(1.5 * (attempt + 1))
    raise RuntimeError(f"请求失败，已重试 3 次: {last_error}")



def request_official_json(path: str, form: dict[str, str]) -> dict[str, Any]:
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept": "application/json, text/plain, */*",
        "Origin": "https://cg.95306.cn",
        "Referer": "https://cg.95306.cn/",
        "X-Requested-With": "XMLHttpRequest",
    }
    url = f"https://cg.95306.cn/proxy/portal/elasticSearch{path}"
    body = urlencode(form).encode("utf-8")
    with urlopen(Request(url, data=body, headers=headers), timeout=25) as response:
        data = json.loads(response.read().decode("utf-8", "ignore"))
    if not data.get("success"):
        raise RuntimeError(data.get("msg") or "国铁采购平台查询失败")
    return data


def fetch_official_content(title: str) -> str:
    code_match = re.search(r"20\d{2}(?:-[A-Z0-9]+){5,}", clean_html(title))
    if not code_match:
        return ""
    for attempt in range(5):
        mh_id = uuid.uuid4().hex
        query = {
            "mhId": mh_id,
            "Authorization": "",
            "projBidType": "",
            "bidType": "",
            "noticeType": "000",
            "unitType": "",
            "wzType": "",
            "title": code_match.group(0),
            "inforCode": "",
            "startDate": "",
            "endDate": "",
            "pageNum": "1",
            "projType": "",
            "createPeopUnit": "",
        }
        try:
            result = request_official_json("/queryProcurementNoticeList", query)
            items = result.get("data", {}).get("resultData", {}).get("result", [])
            notice_id = items[0].get("id") if items else ""
            if not notice_id:
                return ""
            detail = request_official_json("/indexView", {
                "noticeId": str(notice_id),
                "mhId": mh_id,
                "Authorization": "",
            })
            content = str(detail.get("data", {}).get("noticeContent", {}).get("notCont") or "")
            if content:
                return content
        except Exception:
            if attempt < 2:
                time.sleep(attempt + 1)
    return ""



def infer_openid(cookie: str) -> str:
    token_match = re.search(r"(?:^|;\s*)Admin-Token=([^;]+)", cookie)
    if not token_match:
        return ""
    try:
        payload = token_match.group(1).split(".")[1]
        payload += "=" * (-len(payload) % 4)
        data = json.loads(base64.urlsafe_b64decode(payload).decode("utf-8"))
        return str(data.get("ei") or "")
    except Exception:
        return ""


def build_headers(args: argparse.Namespace) -> dict[str, str]:
    token = args.token or os.getenv("YFB_TOKEN", "")
    cookie = args.cookie or os.getenv("YFB_COOKIE", "")
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept": "application/json, text/plain, */*",
        "Referer": REFERER,
    }
    if token:
        headers["Authorization"] = f"Bearer {token}"
    if cookie:
        headers["Cookie"] = cookie
    return headers


def parse_date(value: Any) -> datetime | None:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        try:
            if value > 10_000_000_000:
                value /= 1000
            return datetime.fromtimestamp(value)
        except Exception:
            return None
    text = str(value).strip()
    if not text:
        return None
    if text in ("今天", "今日"):
        now = datetime.now()
        return datetime(now.year, now.month, now.day)
    if text == "昨天":
        day = datetime.now() - timedelta(days=1)
        return datetime(day.year, day.month, day.day)
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S", "%Y-%m-%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(text[: len(fmt)], fmt)
        except ValueError:
            pass
    m = re.search(r"(20\d{2})\D+(\d{1,2})\D+(\d{1,2})", text)
    if m:
        return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return None


def first_value(obj: Any, keys: tuple[str, ...]) -> Any:
    if isinstance(obj, dict):
        for key in keys:
            if obj.get(key):
                return obj[key]
        for value in obj.values():
            found = first_value(value, keys)
            if found:
                return found
    elif isinstance(obj, list):
        for item in obj:
            found = first_value(item, keys)
            if found:
                return found
    return None


def flatten_text(obj: Any) -> str:
    if isinstance(obj, dict):
        return "\n".join(flatten_text(v) for v in obj.values() if v is not None)
    if isinstance(obj, list):
        return "\n".join(flatten_text(v) for v in obj)
    return str(obj)


def clean_html(text: Any) -> str:
    text = "" if text is None else str(text)
    text = re.sub(r"<(?:br|/p|/div|/li|/tr|/h[1-6])\b[^>]*>", "\n", text, flags=re.I)
    text = re.sub(r"<[^>]+>", "", text)
    text = re.sub(r"&nbsp;?", " ", text)
    text = re.sub(r"[ \t]+", " ", text)
    return re.sub(r"\n{3,}", "\n\n", text).strip()


def infer_unit_from_text(title: str, content: str) -> str:
    text = clean_html(content)
    title_text = clean_html(title)
    patterns = (
        r"(?:招标人名称|采购人名称|建设单位名称)\s*[：:]?\s*([\u4e00-\u9fa5A-Za-z0-9（）()·・\-]{2,80})",
        r"(?:采购人|招标人|发包人|建设单位|采购单位|招标单位)\s*(?:为|：|:)\s*([^，,。；;\n]{2,80})",
        r"(?:采购人信息|招标人信息)[\s\S]{0,80}?名称\s*[：:]\s*([^，,。；;\n]{2,80})",
        r"(?:凡对本次(?:采购|招标).*?联系|对本次(?:采购|招标).*?询问)[\s\S]{0,160}?名称\s*[：:]\s*([^，,。；;\n]{2,80})",
    )
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            unit = clean_html(match.group(1)).strip(" ：:，,。；;")
            unit = re.split(r"\s+(?:地址|电话|联系方式|联系人)", unit, 1)[0].strip()
            if 2 <= len(unit) <= 80 and not any(bad in unit for bad in ("采购代理", "代理机构", "项目联系人")):
                return unit
    title_patterns = (
        r"^([\u4e00-\u9fa5]{2,50}(?:厅|局|委员会|管理委员会|管理部|办事处|集团|有限公司|研究院|医院|学院|中心|公司))",
        r"^([\u4e00-\u9fa5]{2,50}自然资源局)",
    )
    for pattern in title_patterns:
        match = re.search(pattern, title_text)
        if match:
            unit = match.group(1).strip()
            if "项目" not in unit and len(unit) <= 60:
                return unit
    for marker in ("厅", "自然资源局", "管理委员会城市管理部", "管理委员会建设管理部", "集团"):
        idx = title_text.find(marker)
        if 1 < idx < 40:
            return title_text[: idx + len(marker)]
    return ""


def text_for_keys(obj: Any, keys: tuple[str, ...]) -> str:
    """Find the first usable text field recursively in a detail response."""
    if isinstance(obj, dict):
        for key in keys:
            if obj.get(key):
                return clean_html(obj[key])
        for value in obj.values():
            found = text_for_keys(value, keys)
            if found:
                return found
    elif isinstance(obj, list):
        for value in obj:
            found = text_for_keys(value, keys)
            if found:
                return found
    return ""


def extract_qualification(announcement: str, api_value: str = "") -> str:
    """Extract the qualification section that is commonly embedded in announcement HTML."""
    headings = (
        "竞标人资格要求", "供应商资格要求", "投标人资格要求", "响应供应商资格要求",
        "申请人资格要求", "报名资格要求", "资格条件", "资格要求",
    )
    parts = [clean_html(api_value)] if api_value else []
    normalized = re.sub(r"\r\n?", "\n", announcement)
    heading_pattern = "|".join(map(re.escape, headings))
    matches = []
    for match in re.finditer(heading_pattern, normalized):
        if match.group() in ("资格要求", "资格条件"):
            prefix = normalized[:match.start()]
            if not re.search(r"(?:^|\n)\s*(?:\d+[.、]|[一二三四五六七八九十]+、)?\s*$", prefix):
                continue
        matches.append(match)
    for index, match in enumerate(matches):
        end = matches[index + 1].start() if index + 1 < len(matches) else len(normalized)
        section = normalized[match.start():end]
        stop = re.search(
            r"\n\s*(?:[一二三四五六七八九十]+、|\d+[.、]|[（(]\d+[)）])?\s*"
            r"(?:报名|获取|谈判文件|采购文件|响应文件|开标|递交|公告期限|联系方式|项目概况)",
            section,
        )
        if stop and stop.start() > 30:
            section = section[:stop.start()]
        section = section.strip()
        if section and section not in parts:
            parts.append(section)
    return "\n".join(parts)[:12000]


def extract_signup_time(announcement: str, api_value: str = "") -> str:
    parts = [clean_html(api_value)] if api_value else []
    normalized = clean_html(re.sub(r"\r\n?", "\n", announcement))
    date_range = re.compile(
        r"(?:20\d{2}年\s*\d{1,2}月\s*\d{1,2}日|\d{1,2}月\s*\d{1,2}日|20\d{2}[-/]\d{1,2}[-/]\d{1,2})"
        r"[\s\S]{0,80}?(?:至|到|-)\s*"
        r"(?:20\d{2}年\s*)?\d{1,2}(?:月|[-/])\s*\d{1,2}(?:日)?(?:\s*\d{1,2}[：:]\d{2})?"
    )
    for match in date_range.finditer(normalized):
        start = max(0, normalized.rfind("\n", 0, match.start()))
        end = normalized.find("\n", match.end())
        if end == -1:
            end = min(len(normalized), match.end() + 160)
        line = normalized[start:end].strip(" ：:；;，,\n")
        line = re.split(r"(?:获取方式|报名所需|5\.|五、|并将|将以下|邮件|电话通知)", line, 1)[0].strip(" ，,。；;")
        line = re.sub(r"\s+", "", line)
        if any(key in line for key in ("获取", "报名", "发售", "领取", "下载", "请于", "凡有意", "文件")):
            if line and line not in parts:
                parts.append(line)
    return "\n".join(parts)[:1200]


def extract_bid_deadline(announcement: str) -> str:
    normalized = clean_html(re.sub(r"\r\n?", "\n", announcement))
    patterns = (
        r"(?:投标|响应|竞标)文件(?:递交|提交|上传)?(?:的)?(?:截止时间|递交截止时间)[^。；;\n]{0,120}",
        r"(?:投标|响应|竞标)(?:截止时间|截止日期)[^。；;\n]{0,120}",
        r"递交截止时间[^。；;\n]{0,120}",
        r"截止时间(?:为|：|:)?[^。；;\n]{0,120}",
    )
    date_time = re.compile(
        r"20\d{2}年\s*\d{1,2}月\s*\d{1,2}日(?:\s*(?:上午|下午)?\s*\d{1,2}(?:[：:]\d{2}|时\d{2}分?))?"
        r"|20\d{2}[-/]\d{1,2}[-/]\d{1,2}(?:\s*(?:上午|下午)?\s*\d{1,2}[：:]\d{2})?"
    )
    for pattern in patterns:
        for match in re.finditer(pattern, normalized):
            text = re.sub(r"\s+", "", match.group(0)).strip(" ：:，,。；;")
            dt = date_time.search(text)
            if not dt:
                continue
            label = "投标截止时间"
            if "响应" in text:
                label = "响应截止时间"
            elif "竞标" in text:
                label = "竞标截止时间"
            elif "递交" in text:
                label = "递交截止时间"
            return f"{label}：{dt.group(0)}"
    return ""


PROTECTED_TITLE_TERMS = (
    "水土保持", "测绘服务", "测绘项目", "国土测绘", "基础测绘", "地形图", "规划核实",
    "房产实测", "国土变更调查", "用地预审", "确权登记", "土地复垦", "竣工测量",
    "工程测量", "多测合一", "变形监测", "基坑监测", "第三方监测", "专项监测",
    "水土保持监测", "监测、验收", "水土保持验收", "水土保持方案", "遥感监管", "遥感专项监测",
)

NOISE_TITLE_TERMS = (
    "闲置车位", "车位使用权", "房间招标", "宾馆", "招租", "废旧资产", "机械密封",
    "水电户表", "金属栏杆", "格栅", "爬梯", "中央空调室内机", "加工件",
    "UPS系统采购", "电子反拍", "住宅、储藏室", "山沙一宗", "设备车间设置职工舒缓室",
    "充电站建设项目设计服务", "校园校舍建筑设施安全", "供水管网检测项目", "地基检测项目",
    "测绘仪器无人机采购", "无人机招标公告", "设备健康监测感知层设备与材料采购",
    "交易公告", "转让122套住宅", "联勤宾馆", "联勤宾馆迎宾楼", "双口峪村山沙",
    "莱芜基地用机械密封", "华电国际电力股份有限公司莱城发电厂", "设备租赁",
    "全站仪设备租赁", "监控", "摄像头", "观察孔", "维修件", "仪器维修", "设备采购",
    "过滤器", "水泵", "传动链", "钢丝", "预埋铁座", "模具", "塑料预埋件",
    "商品混凝土", "钢模板", "箱梁模板", "砂石料棚", "钢筋加工棚", "沥青", "河砂",
    "机制砂", "碎石", "铝合金门窗", "锚杆", "施工围挡", "钢丝网护栏", "防护网",
    "管路配件", "抗裂剂", "防火密封胶", "絮凝剂", "战略采购", "地名编制",
    "地图编制", "政务工作用图", "电子地图",
)

STRICT_NOISE_TITLE_TERMS = (
    "测绘仪器无人机采购", "UPS系统采购", "设备健康监测感知层设备与材料采购",
)


def normalize_title_for_dedupe(title: str) -> str:
    text = clean_html(title)
    text = re.sub(r"[（(]原标题[:：].*?[）)]", "", text)
    text = re.sub(r"第\d+次延期", "", text)
    text = re.sub(r"第一次变更公告|第一次更正公示|变更公告|更正公告|二次招标|二次|--.*$", "", text)
    return re.sub(r"\s+", "", text)


def is_noise_title(title: str) -> bool:
    title_text = clean_html(title)
    if any(term in title_text for term in STRICT_NOISE_TITLE_TERMS):
        return True
    if any(term in title_text for term in PROTECTED_TITLE_TERMS):
        return False
    return any(term in title_text for term in NOISE_TITLE_TERMS)

def relevance_level(keyword: str, title: str, text: str) -> str:
    title_text = clean_html(title)
    haystack = clean_html("\n".join(x for x in (title, text) if x))
    if not haystack or is_noise_title(title_text):
        return ""
    if keyword == "水土保持":
        return "明确相关" if "水土保持" in title_text else "疑似相关" if "水土保持" in haystack else ""
    if keyword == "监测":
        title_terms = ("监测", "变形监测", "基坑监测", "沉降观测", "第三方监测", "专项监测")
        body_terms = title_terms + ("监测服务", "监测工作", "监测项目", "深基坑", "基坑支护")
        if any(term in title_text for term in title_terms):
            return "明确相关"
        return "疑似相关" if any(term in haystack for term in body_terms) else ""
    if keyword == "测绘":
        title_terms = ("测绘", "地形图", "规划核实", "房产实测", "国土变更调查", "不动产")
        body_terms = title_terms + ("测绘服务", "测绘成果", "测绘项目")
        if any(term in title_text for term in title_terms):
            return "明确相关"
        return "疑似相关" if any(term in haystack for term in body_terms) else ""
    if keyword == "测量":
        title_terms = ("测量", "工程测量", "竣工测量", "房产实测", "规划核实", "地形图")
        body_terms = title_terms + ("测量服务", "测量工作", "测量项目")
        if any(term in title_text for term in title_terms):
            return "明确相关"
        return "疑似相关" if any(term in haystack for term in body_terms) else ""
    return "明确相关" if keyword in title_text else "疑似相关" if keyword in haystack else ""

def is_relevant_keyword_match(keyword: str, title: str, text: str) -> bool:
    return bool(relevance_level(keyword, title, text))


def is_relevant_monitoring_title(title: str, keyword: str = "", blob: str = "") -> bool:
    return is_relevant_keyword_match(keyword, title, blob)


def one_month_ago_start(today: datetime | None = None) -> datetime:
    today = today or datetime.now()
    month = today.month - 1
    year = today.year
    if month == 0:
        month = 12
        year -= 1
    day = min(today.day, 28 if month == 2 else 30 if month in (4, 6, 9, 11) else 31)
    return datetime(year, month, day)


def fetch_list(headers: dict[str, str], openid: str, days: int | None) -> list[dict[str, Any]]:
    today = datetime.now()
    cutoff = datetime(today.year, today.month, today.day) - timedelta(days=days) if days else one_month_ago_start(today)
    rows: list[dict[str, Any]] = []
    seen: set[str] = set()
    seen_titles: set[str] = set()
    time_option = 5 if days and days > 30 else 4
    for keyword in SEARCH_KEYWORDS:
        for filter_condition in FILTER_CONDITIONS:
            for search_type in SEARCH_TYPES:
                for page in range(1, 31):
                    params = {
                        "pageSize": 100,
                        "pageNum": page,
                        "pageFrom": "zhaobiao",
                        "keyword": keyword,
                        "areaIds": AREA_IDS,
                        "filterCondition": filter_condition,
                        "searchType": search_type,
                        # 长范围用“近三个月”(5)减少无关历史页；短范围用“全部”(4)，再由脚本本地过滤。
                        "timeOption": time_option,
                        "viewMonitor": "false",
                        "defTimeFlag": 0,
                        "openid": openid or None,
                    }
                    try:
                        data = request_json("/subZhaobiao/queryZBInfo", params, headers).get("data") or {}
                    except Exception as exc:
                        print(
                            f"列表页请求失败，跳过 keyword={keyword} filterCondition={filter_condition} "
                            f"searchType={search_type} page={page}: {exc}",
                            file=sys.stderr,
                        )
                        break
                    items = data.get("resultList") or data.get("realTimeList") or data.get("resultSet") or []
                    if not items:
                        break
                    old_count = 0
                    for item in items:
                        dt = parse_date(
                            first_value(item, ("publishDate", "releaseDate", "createTime", "addTime", "updateDate"))
                        )
                        area = str(first_value(item, ("areaName", "area", "areaText")) or "")
                        blob = flatten_text(item)
                        title = clean_html(first_value(item, ("title", "projectName", "name")))
                        title_key = normalize_title_for_dedupe(title)
                        content_id = str(item.get("contentId") or item.get("id") or title_key or title)
                        if dt and dt < cutoff:
                            old_count += 1
                            continue
                        if content_id in seen or (title_key and title_key in seen_titles):
                            continue
                        if "济南" not in area and "莱芜" not in area and "济南" not in blob and "莱芜" not in blob:
                            continue
                        level = relevance_level(keyword, title, blob)
                        if not level:
                            continue
                        candidate = dict(item)
                        candidate["_searchKeyword"] = keyword
                        candidate["_filterCondition"] = filter_condition
                        candidate["_searchType"] = search_type
                        candidate["_relevanceLevel"] = level
                        rows.append(candidate)
                        seen.add(content_id)
                        if title_key:
                            seen_titles.add(title_key)
                    if old_count == len(items):
                        break
    rows.sort(
        key=lambda item: parse_date(
            first_value(item, ("publishDate", "releaseDate", "createTime", "addTime", "updateDate"))
        ) or datetime.min,
        reverse=True,
    )
    return rows

def fetch_detail(item: dict[str, Any], headers: dict[str, str], openid: str) -> dict[str, Any]:
    content_id = item.get("contentId") or item.get("id")
    area_id = item.get("areaId") or item.get("area")
    if not content_id:
        return {}
    params = {"contentId": content_id, "areaId": area_id, "pageFrom": "search", "openid": openid or None}
    try:
        detail = request_json("/subZhaobiao/zbDetail", params, headers, timeout=8).get("data") or {}
    except Exception:
        detail = {}
    detail_content = text_for_keys(detail, ("content", "noticeContent", "htmlContent", "detailContent", "text"))
    item_content = text_for_keys(item, ("content", "summary", "noticeContent", "htmlContent", "text"))
    qualification = text_for_keys(detail, ("qualification", "aptitude", "qualificationRequirement"))
    combined_content = "\n".join(x for x in (detail_content, item_content) if x)
    if not extract_qualification(combined_content, qualification):
        official_content = fetch_official_content(str(item.get("title") or ""))
        if official_content:
            detail["officialContent"] = official_content
    return detail


def row_from_item(item: dict[str, Any], detail: dict[str, Any]) -> list[Any]:
    merged = {"list": item, "detail": detail}
    dt = parse_date(first_value(merged, ("publishDate", "releaseDate", "createTime", "addTime", "updateDate")))
    title = clean_html(first_value(merged, ("title", "projectName", "name")))
    unit = clean_html(first_value(merged, ("zhaoBiaoUnit", "zhaoBiaoRen", "tenderer", "buyerName", "purchaseUnit")))
    area = clean_html(first_value(merged, ("areaName", "areaText", "area")))
    content = text_for_keys(detail, ("officialContent", "content", "noticeContent", "htmlContent", "detailContent", "text"))
    if not content:
        content = text_for_keys(item, ("content", "summary", "noticeContent", "htmlContent", "text"))
    if not unit:
        unit = infer_unit_from_text(title, content)
    if not unit:
        match = re.match(r"(.{2,40}?有限公司)", title)
        if match:
            unit = match.group(1).strip()
    qualification = text_for_keys(detail, ("qualification", "aptitude", "qualificationRequirement"))
    signup_time = clean_html(first_value(merged, ("signUpTime", "registrationTime", "tenderEndTimeStr", "bidEndTime")))
    qualification_text = extract_qualification(content, qualification)
    time_text = extract_signup_time(content, signup_time)
    deadline_text = extract_bid_deadline(content)
    url = f"https://qiye.qianlima.com/new_qd_yfbsite/#/infoCenter/infoDetail/{item.get('contentId')}/{item.get('areaId')}/zhaobiao"
    return [
        None,
        dt.strftime("%m.%d") if dt else "",
        title,
        unit,
        area,
        qualification_text or "未在公告正文中识别到明确的资格要求",
        time_text,
        deadline_text,
        "",
        item.get("_relevanceLevel") or "明确相关",
        url,
    ]



def ensure_relevance_column(ws: Any) -> None:
    if ws.cell(1, 10).value == "相关性":
        return
    ws.insert_cols(10)
    ws.cell(1, 10, "相关性")
    for row in range(1, ws.max_row + 1):
        src = ws.cell(row, 9)
        dst = ws.cell(row, 10)
        if src.has_style:
            dst._style = copy.copy(src._style)
        if src.number_format:
            dst.number_format = src.number_format
        if src.alignment:
            dst.alignment = copy.copy(src.alignment)


def ensure_detail_sheet(wb: Any) -> Any:
    if DETAIL_SHEET_NAME in wb.sheetnames:
        detail_ws = wb[DETAIL_SHEET_NAME]
    else:
        detail_ws = wb.create_sheet(DETAIL_SHEET_NAME)
        detail_ws.append(["主表序号", "项目名称", "内容类型", "详细内容", "返回主表"])
    headers = ["主表序号", "项目名称", "内容类型", "详细内容", "返回主表"]
    for col, header in enumerate(headers, start=1):
        detail_ws.cell(1, col, header)
        detail_ws.cell(1, col).font = Font(bold=True)
    widths = {1: 12, 2: 70, 3: 16, 4: 120, 5: 18}
    for column, width in widths.items():
        letter = detail_ws.cell(1, column).column_letter
        detail_ws.column_dimensions[letter].width = width
    detail_ws.freeze_panes = "A2"
    return detail_ws


def short_remark_label(text: str) -> str:
    text = clean_html(text)
    if "未完整获取公告正文" in text or "未识别到明确资格要求" in text or "公告内容" in text:
        return REMARK_PLACEHOLDER
    if "疑似相关" in text:
        return "备注：疑似相关，需人工核对"
    return "公告备注"


def add_detail_link(ws: Any, detail_ws: Any, row: int, col: int, title: str, kind: str, full_text: str, label: str) -> None:
    full_text = clean_html(full_text)
    if not full_text:
        return
    detail_row = detail_ws.max_row + 1
    main_cell = ws.cell(row, col)
    detail_ws.cell(detail_row, 1, ws.cell(row, 1).value)
    detail_ws.cell(detail_row, 2, title)
    detail_ws.cell(detail_row, 3, kind)
    detail_ws.cell(detail_row, 4, full_text)
    back_cell = detail_ws.cell(detail_row, 5, "返回主表")
    back_cell.hyperlink = f"#'{ws.title}'!{main_cell.coordinate}"
    back_cell.style = "Hyperlink"
    for c in range(1, 6):
        cell = detail_ws.cell(detail_row, c)
        alignment = copy.copy(cell.alignment)
        alignment.wrap_text = True
        alignment.vertical = "top"
        cell.alignment = alignment
    detail_ws.row_dimensions[detail_row].height = min(409, max(45, len(full_text) // 80 * 15))
    main_cell.value = label
    main_cell.hyperlink = f"#'{DETAIL_SHEET_NAME}'!D{detail_row}"
    main_cell.style = "Hyperlink"
    alignment = copy.copy(main_cell.alignment)
    alignment.wrap_text = True
    alignment.vertical = "top"
    main_cell.alignment = alignment


def move_long_text_to_detail(ws: Any, detail_ws: Any, row: int) -> None:
    title = str(ws.cell(row, 3).value or "")
    qualification = str(ws.cell(row, 6).value or "")
    remark = str(ws.cell(row, 11).value or "") if ws.max_column >= 11 else ""
    if qualification and qualification != QUALIFICATION_PLACEHOLDER:
        add_detail_link(ws, detail_ws, row, 6, title, "资质", qualification, QUALIFICATION_PLACEHOLDER)
    if remark and not remark.startswith("http") and remark not in ("公告备注", REMARK_PLACEHOLDER, "备注：疑似相关，需人工核对"):
        add_detail_link(ws, detail_ws, row, 11, title, "备注", remark, short_remark_label(remark))


def format_notice_rows(ws: Any, row_numbers: list[int]) -> None:
    widths = {1: 9, 2: 13, 3: 62, 4: 37, 5: 18, 6: 60, 7: 34, 8: 30, 9: 30, 10: 14, 11: 55}
    line_widths = {3: 45, 4: 27, 5: 16, 6: 60, 7: 30, 8: 26, 9: 30, 10: 10, 11: 55}
    for column, width in widths.items():
        letter = ws.cell(1, column).column_letter
        current = ws.column_dimensions[letter].width or 0
        ws.column_dimensions[letter].width = max(current, width)
    for row in row_numbers:
        estimated_lines = 1
        for column in range(1, 12):
            cell = ws.cell(row, column)
            alignment = copy.copy(cell.alignment)
            alignment.wrap_text = True
            alignment.vertical = "top"
            cell.alignment = alignment
            value = str(cell.value or "")
            width = line_widths.get(column, 30)
            estimated_lines = max(
                estimated_lines,
                sum(max(1, (len(line) + width - 1) // width) for line in value.splitlines() or [""]),
            )
        ws.row_dimensions[row].height = min(409, max(30, estimated_lines * 15))


def append_to_workbook(path: Path, rows: list[list[Any]], dry_run: bool) -> None:
    if not rows:
        print("未抓到符合条件的新数据，Excel 未修改。")
        return
    wb = load_workbook(path)
    ws = wb.active
    ensure_relevance_column(ws)
    detail_ws = ensure_detail_sheet(wb)
    while ws.max_row > 2 and all(ws.cell(ws.max_row, c).value in (None, "") for c in range(1, ws.max_column + 1)):
        ws.delete_rows(ws.max_row)
    existing_titles = {
        normalize_title_for_dedupe(str(ws.cell(r, 3).value or ""))
        for r in range(2, ws.max_row + 1)
        if str(ws.cell(r, 3).value or "").strip()
    }
    filtered_rows = []
    seen_new_titles: set[str] = set()
    for row in rows:
        title_key = normalize_title_for_dedupe(str(row[2] or ""))
        if title_key and (title_key in existing_titles or title_key in seen_new_titles):
            continue
        filtered_rows.append(row)
        if title_key:
            seen_new_titles.add(title_key)
    rows = filtered_rows
    if not rows:
        print("抓到的数据均已存在，Excel 未修改。")
        return
    template_row = ws.max_row == 2 and all(
        ws.cell(2, c).value in (None, "") for c in range(1, ws.max_column + 1)
    )
    start = 2 if template_row else ws.max_row + 1
    last_no = max([ws.cell(r, 1).value for r in range(2, ws.max_row + 1) if isinstance(ws.cell(r, 1).value, int)] or [0])
    style_row = 2 if template_row else ws.max_row
    for offset, row in enumerate(rows, start=0):
        target = start + offset
        row[0] = last_no + offset + 1
        for col, value in enumerate(row, start=1):
            cell = ws.cell(target, col, value)
            src = ws.cell(style_row, col)
            if src.has_style:
                cell._style = copy.copy(src._style)
            if src.number_format:
                cell.number_format = src.number_format
            if src.alignment:
                cell.alignment = copy.copy(src.alignment)
        move_long_text_to_detail(ws, detail_ws, target)
    format_notice_rows(ws, list(range(start, start + len(rows))))
    if dry_run:
        print(f"演练模式：将追加 {len(rows)} 条，Excel 未修改。")
        return
    backup = path.with_suffix(f".backup-{datetime.now():%Y%m%d-%H%M%S}.xlsx")
    shutil.copy2(path, backup)
    wb.save(path)
    print(f"已追加 {len(rows)} 条，备份文件：{backup}")


def main() -> int:
    parser = argparse.ArgumentParser(description="搜索乙方宝近 N 天山东济南/莱芜监测、水土保持、测绘、测量招标并追加到 Excel")
    parser.add_argument("--xlsx", type=Path, default=DEFAULT_XLSX)
    parser.add_argument("--days", type=int, default=7, help="按最近 N 天筛选，默认 7 天")
    parser.add_argument("--token", default="")
    parser.add_argument("--cookie", default="")
    parser.add_argument("--openid", default=os.getenv("YFB_OPENID", ""))
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    headers = build_headers(args)
    openid = args.openid or infer_openid(args.cookie or os.getenv("YFB_COOKIE", ""))
    try:
        existing_titles: set[str] = set()
        if args.xlsx.exists():
            wb_existing = load_workbook(args.xlsx, read_only=True, data_only=True)
            ws_existing = wb_existing.active
            existing_titles = {
                normalize_title_for_dedupe(str(ws_existing.cell(r, 3).value or ""))
                for r in range(2, ws_existing.max_row + 1)
                if str(ws_existing.cell(r, 3).value or "").strip()
            }
        items = fetch_list(headers, openid, args.days)
        rows = []
        for item in items:
            list_title = clean_html(first_value(item, ("title", "projectName", "name")))
            if list_title and normalize_title_for_dedupe(list_title) in existing_titles:
                continue
            detail = fetch_detail(item, headers, openid)
            title = clean_html(first_value({"list": item, "detail": detail}, ("title", "projectName", "name")))
            content = text_for_keys(detail, ("officialContent", "content", "noticeContent", "htmlContent", "detailContent", "text"))
            if not content:
                content = text_for_keys(item, ("content", "summary", "noticeContent", "htmlContent", "text"))
            keyword = str(item.get("_searchKeyword") or "")
            level = relevance_level(keyword, title, content) if keyword else str(item.get("_relevanceLevel") or "")
            if not level:
                print(f"跳过详情正文未匹配 {keyword} 的项目：{title}", file=sys.stderr)
                continue
            item["_relevanceLevel"] = level
            row = row_from_item(item, detail)
            remark_parts = []
            if level == "疑似相关":
                remark_parts.append("备注：疑似相关，需人工核对：关键词未在标题中明确命中，但公告正文或列表摘要中出现相关内容。")
            if row[5] == "未在公告正文中识别到明确的资格要求":
                print(f"写入但标记为需人工核对：{row[2]}", file=sys.stderr)
                notice_text = clean_html(content) or clean_html(flatten_text(item))
                remark_parts.append("备注：未完整获取公告正文或未识别到明确资格要求，请人工核对。\n公告内容：\n" + notice_text[:10000])
            if remark_parts:
                row[10] = "\n".join(remark_parts) + "\n链接：" + str(row[10] or "")
            rows.append(row)
        append_to_workbook(args.xlsx, rows, args.dry_run)
    except YfbAuthError as exc:
        print(f"认证失败：{exc}", file=sys.stderr)
        print("请登录乙方宝后提供 YFB_TOKEN，或设置 YFB_COOKIE / YFB_OPENID 再运行。", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())




