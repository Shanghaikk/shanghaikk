#!/usr/bin/env python3
"""
从 OpenClaw 会话记录提取 API 用量数据
按天/小时汇总：请求次数、tokens、消费金额
"""

import os, sys, json
from collections import defaultdict
from datetime import datetime, timezone, timedelta

SESSIONS_DIR = os.path.expanduser("~/.openclaw/agents/main/sessions/")
OUTPUT_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "repo/deepseek/usage-data.json")
TZ = timezone(timedelta(hours=8))

# DeepSeek V4 Flash 定价
RATES = {"input": 0.14, "output": 0.28, "cache": 0.028}  # 每百万 tokens


def extract():
    users_req = {}
    today = datetime.now(TZ).strftime("%Y-%m-%d")
    
    if not os.path.isdir(SESSIONS_DIR):
        return {"error": "sessions dir not found"}
    
    for fname in os.listdir(SESSIONS_DIR):
        if not fname.endswith(".jsonl"):
            continue
        fpath = os.path.join(SESSIONS_DIR, fname)
        try:
            with open(fpath) as f:
                for line in f:
                    try:
                        d = json.loads(line)
                        msg = d.get("message", {})
                        usage = msg.get("usage", {}) or {}
                        cost = usage.get("cost", {}) or {}
                        ts = d.get("timestamp", "")
                        if not ts or not cost.get("total"):
                            continue
                        day = ts[:10]
                        h = ts[11:13]
                        if day not in users_req:
                            users_req[day] = defaultdict(lambda: {
                                "calls": 0, "input": 0, "output": 0,
                                "cache": 0, "cost": 0.0, "hours": {}
                            })
                        users_req[day]["total"]["calls"] += 1
                        users_req[day]["total"]["input"] += usage.get("input", 0)
                        users_req[day]["total"]["output"] += usage.get("output", 0)
                        users_req[day]["total"]["cache"] += usage.get("cacheRead", 0)
                        users_req[day]["total"]["cost"] += cost.get("total", 0)
                        
                        if h not in users_req[day]["total"]["hours"]:
                            users_req[day]["total"]["hours"][h] = {"calls": 0, "input": 0, "output": 0, "cache": 0, "cost": 0.0}
                        users_req[day]["total"]["hours"][h]["calls"] += 1
                        users_req[day]["total"]["hours"][h]["input"] += usage.get("input", 0)
                        users_req[day]["total"]["hours"][h]["output"] += usage.get("output", 0)
                        users_req[day]["total"]["hours"][h]["cache"] += usage.get("cacheRead", 0)
                        users_req[day]["total"]["hours"][h]["cost"] += cost.get("total", 0)
                    except: pass
        except: pass
    
    # 格式化输出
    result = {"usdRate": 7.3, "days": {}}
    for day in sorted(users_req.keys()):
        h = users_req[day]["total"]["hours"]
        result["days"][day] = {
            "calls": users_req[day]["total"]["calls"],
            "input": users_req[day]["total"]["input"],
            "output": users_req[day]["total"]["output"],
            "cache": users_req[day]["total"]["cache"],
            "cost_usd": round(users_req[day]["total"]["cost"], 6),
            "cost_cny": round(users_req[day]["total"]["cost"] * 7.3, 2),
            "hours": {h: v for h, v in sorted(h.items())}
        }
    
    return result


if __name__ == "__main__":
    data = extract()
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, "w") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    total = data.get("days", {}).get(max(data.get("days", {}) or [""]), {})
    if total:
        print(f"今日: 请求{total.get('calls',0)}次  input{total.get('input',0):,}  output{total.get('output',0):,}  ¥{total.get('cost_cny',0):.2f}")
