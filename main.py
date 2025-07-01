import os
import json
import time
import re
import random
import requests
import unicodedata
from datetime import datetime, timedelta
from email.utils import parsedate_to_datetime

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import gspread

KEYWORDS = ["MAZDA", "mazda", "マツダ"]
SPREADSHEET_ID = "1RyflHrdqh9X9NhSq24ZX5rPJCVE5WoJuM_7U_h6Y4kU"

def format_datetime(dt_obj):
    return dt_obj.strftime("%Y/%m/%d %H:%M")

def parse_relative_time(pub_label: str, base_time: datetime) -> str:
    pub_label = pub_label.strip().lower()
    try:
        if "分前" in pub_label or "minute" in pub_label:
            m = re.search(r"(\d+)", pub_label)
            if m:
                dt = base_time - timedelta(minutes=int(m.group(1)))
                return format_datetime(dt)
        elif "時間前" in pub_label or "hour" in pub_label:
            h = re.search(r"(\d+)", pub_label)
            if h:
                dt = base_time - timedelta(hours=int(h.group(1)))
                return format_datetime(dt)
        elif "日前" in pub_label or "day" in pub_label:
            d = re.search(r"(\d+)", pub_label)
            if d:
                dt = base_time - timedelta(days=int(d.group(1)))
                return format_datetime(dt)
    except:
        pass
    return format_datetime(base_time)

def get_last_modified_datetime(url):
    try:
        response = requests.head(url, timeout=5)
        if 'Last-Modified' in response.headers:
            dt = parsedate_to_datetime(response.headers['Last-Modified'])
            jst = dt.astimezone(tz=timedelta(hours=9))
            return format_datetime(jst)
    except:
        pass
    return "取得不可"

def is_japanese(text):
    return any(
        'CJK UNIFIED' in unicodedata.name(ch, '') or
        'HIRAGANA' in unicodedata.name(ch, '') or
        'KATAKANA' in unicodedata.name(ch, '') for ch in text
    )

def get_google_news_with_selenium(keyword):
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    url = f"https://news.google.com/search?q={keyword}&hl=ja&gl=JP&ceid=JP:ja"
    driver.get(url)
    time.sleep(5)
    for _ in range(3):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)
    soup = BeautifulSoup(driver.page_source, "html.parser")
    driver.quit()

    data = []
    for article in soup.find_all("article"):
        try:
            a_tag = article.select_one("a.JtKRv")
            time_tag = article.select_one("time.hvbAAd")
            source_tag = article.select_one("div.vr1PYe")
            title = a_tag.text.strip()
            href = a_tag.get("href")
            url = "https://news.google.com" + href[1:] if href.startswith("./") else href
            dt = datetime.strptime(time_tag.get("datetime"), "%Y-%m-%dT%H:%M:%SZ") + timedelta(hours=9)
            pub_date = format_datetime(dt)
            source = source_tag.text.strip() if source_tag else "N/A"
            data.append([title, pub_date, url, source])
        except:
            continue
    print(f"✅ Googleニュース件数: {len(data)} 件（{keyword}）")
    return data

def get_yahoo_news_with_selenium(keyword):
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    search_url = f"https://news.yahoo.co.jp/search?p={keyword}&ei=utf-8&categories=domestic,world,business,it,science,life,local"
    driver.get(search_url)
    time.sleep(5)
    soup = BeautifulSoup(driver.page_source, "html.parser")
    driver.quit()

    data = []
    for article in soup.find_all("li", class_=re.compile("sc-1u4589e-0")):
        try:
            title_tag = article.find("div", class_=re.compile("sc-3ls169-0"))
            link_tag = article.find("a", href=True)
            time_tag = article.find("time")
            source_tag = article.find("div", class_="sc-n3vj8g-0 yoLqH")

            title = title_tag.text.strip() if title_tag else ""
            url = link_tag["href"] if link_tag else ""
            date_str = time_tag.text.strip() if time_tag else ""
            formatted_date = date_str
            if date_str:
                try:
                    dt = datetime.strptime(re.sub(r'\([\u6708\u706b\u6c34\u6728\u91d1\u571f\u65e5]\)', '', date_str), "%Y/%m/%d %H:%M")
                    formatted_date = format_datetime(dt)
                except:
                    pass
            source_text = source_tag.text.strip() if source_tag else "Yahoo"
            if title and url:
                data.append([title, formatted_date, url, source_text])
        except:
            continue
    print(f"✅ Yahooニュース件数: {len(data)} 件（{keyword}）")
    return data

def get_msn_news_with_selenium(keyword):
    now = datetime.utcnow() + timedelta(hours=9)
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    url = f"https://www.bing.com/news/search?q={keyword}+site:.jp&qft=sortbydate%3d'1'&form=YFNR"
    driver.get(url)
    time.sleep(5)
    soup = BeautifulSoup(driver.page_source, "html.parser")
    driver.quit()

    data = []
    for card in soup.select("div.news-card"):
        try:
            title = card.get("data-title", "").strip()
            url = card.get("data-url", "").strip()
            source = card.get("data-author", "").strip()

            pub_tag = card.find("span", attrs={"aria-label": True})
            pub_label = pub_tag["aria-label"].strip() if pub_tag and pub_tag.has_attr("aria-label") else ""
            pub_date = parse_relative_time(pub_label, now)

            if pub_date == "取得不可" and url:
                pub_date = get_last_modified_datetime(url)

            if title and url and is_japanese(title):  # ✅ 日本語フィルタ
                data.append([title, pub_date, url, source or "MSN"])
        except:
            continue
    print(f"✅ MSNニュース件数: {len(data)} 件（{keyword}）")
    return data

def write_to_spreadsheet(articles, spreadsheet_id, source):
    credentials_json_str = os.getenv("GCP_SERVICE_ACCOUNT_KEY")
    credentials = json.loads(credentials_json_str)
    gc = gspread.service_account_from_dict(credentials)
    sh = gc.open_by_key(spreadsheet_id)

    sheet_name = source  # Google, Yahoo, MSN 固定名
    try:
        worksheet = sh.worksheet(sheet_name)
        worksheet.clear()  # 上書き
    except:
        worksheet = sh.add_worksheet(title=sheet_name, rows="1000", cols="20")

    values = [["タイトル", "投稿日時", "URL", "引用元"]] + articles
    worksheet.append_rows(values)
    print(f"✅ {source}ニュースを{len(articles)}件書き込み完了")

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
