import os
import json
import time
import re
import random
import requests
from datetime import datetime, timedelta
from email.utils import parsedate_to_datetime

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import gspread

# ✅ 対象キーワードとスプレッドシートID（MAZDA用）
KEYWORDS = ["MAZDA", "mazda", "マツダ"]
SPREADSHEET_ID = "1RyflHrdqh9X9NhSq24ZX5rPJCVE5WoJuM_7U_h6Y4kU"

# ✅ 日時を統一フォーマットに変換
def format_datetime(dt_obj):
    return dt_obj.strftime("%Y/%m/%d %H:%M")

# ✅ 相対時間を日時に変換
def parse_relative_time(pub_label: str, base_time: datetime) -> str:
    pub_label = pub_label.strip().lower()
    try:
        if "分前" in pub_label or "minute" in pub_label:
            m = re.search(r"(\\d+)", pub_label)
            if m:
                dt = base_time - timedelta(minutes=int(m.group(1)))
                return format_datetime(dt)
        elif "時間前" in pub_label or "hour" in pub_label:
            h = re.search(r"(\\d+)", pub_label)
            if m:
                dt = base_time - timedelta(hours=int(h.group(1)))
                return format_datetime(dt)
        elif "日前" in pub_label or "day" in pub_label:
            d = re.search(r"(\\d+)", pub_label)
            if d:
                dt = base_time - timedelta(days=int(d.group(1)))
                return format_datetime(dt)
    except:
        pass
    return format_datetime(base_time)

# ✅ Google News取得（Selenium）
def get_google_news_with_selenium(keyword):
    # ...（略：既存の処理をそのまま使用）
    return []

# ✅ Yahooニュース取得（Selenium）
def get_yahoo_news_with_selenium(keyword):
    # ...（略：既存の処理をそのまま使用）
    return []

# ✅ MSNニュース取得（Selenium）
def get_msn_news_with_selenium(keyword):
    # ...（略：既存の処理をそのまま使用）
    return []

# ✅ スプレッドシート書き込み処理
def write_to_spreadsheet(articles, spreadsheet_id, source):
    credentials_json_str = os.getenv("GCP_SERVICE_ACCOUNT_KEY")
    credentials = json.loads(credentials_json_str)
    gc = gspread.service_account_from_dict(credentials)
    sh = gc.open_by_key(spreadsheet_id)

    sheet_name = datetime.now().strftime("%y%m%d") + f"_{source}"
    try:
        worksheet = sh.worksheet(sheet_name)
        worksheet.clear()
    except:
        worksheet = sh.add_worksheet(title=sheet_name, rows="1000", cols="20")

    worksheet.append_row(["タイトル", "投稿日時", "URL", "引用元"])
    for row in articles:
        worksheet.append_row(row)
    print(f"✅ {source}ニュースを{len(articles)}件書き込み完了")

# ✅ メイン処理
if __name__ == "__main__":
    for kw in KEYWORDS:
        print(f"\n=== キーワード: {kw} ===")
        for source, fetch_func in {
            "Google": get_google_news_with_selenium,
            "Yahoo": get_yahoo_news_with_selenium,
            "MSN": get_msn_news_with_selenium
        }.items():
            print(f"\n--- {source} News ---")
            articles = fetch_func(kw)
            if articles:
                write_to_spreadsheet(articles, SPREADSHEET_ID, source)
