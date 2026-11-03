import os
import re
import time
from collections import deque
from urllib.parse import urljoin, urlparse, urldefrag

import requests
from bs4 import BeautifulSoup
from markdownify import markdownify as md
from playwright.sync_api import sync_playwright


START_URL = "https://docs.hyland.com/r/Alfresco/Alfresco-Content-Services/26.1/Alfresco-Content-Services"
OUTPUT_DIR = "hyland_docs"
IMAGES_DIR = os.path.join(OUTPUT_DIR, "images")
DELAY_SECONDS = 1.0
MAX_PAGES = 200

ALLOWED_DOMAIN = "docs.hyland.com"
ALLOWED_PATH_PREFIX = "/r/Alfresco/Alfresco-Content-Services/26.1/"

# Try common sidebar / TOC containers first
LINK_CONTAINER_SELECTORS = [
    "nav",
    "aside",
    ".sidebar",
    ".toc",
    ".table-of-contents",
    ".toc-nav",
    ".docs-sidebar",
    ".nav-sidebar",
    ".sidenav",
    "[aria-label='Table of contents']",
    "[role='navigation']",
]

HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; DocsScraper/1.0)"
}

session = requests.Session()
session.headers.update(HEADERS)


def normalize_url(url: str) -> str:
    url, _ = urldefrag(url)
    return url.rstrip("/")


def allowed_url(url: str) -> bool:
    parsed = urlparse(url)
    return (
        parsed.netloc == ALLOWED_DOMAIN
        and parsed.path.startswith(ALLOWED_PATH_PREFIX)
    )


def safe_md_filename_from_url(url: str) -> str:
    parsed = urlparse(url)
    path = parsed.path.strip("/") or "index"
    filename = f"{parsed.netloc}_{path}"

    if parsed.query:
        filename += "_" + parsed.query

    filename = re.sub(r"[^\w\-_.]+", "_", filename)

    if not filename.lower().endswith(".md"):
        filename += ".md"

    return filename


def safe_image_filename_from_url(url: str) -> str:
    parsed = urlparse(url)
    path = parsed.path.strip("/") or "image"
    filename = f"{parsed.netloc}_{path}"

    if parsed.query:
        filename += "_" + parsed.query

    filename = re.sub(r"[^\w\-_.]+", "_", filename)

    root, ext = os.path.splitext(filename)
    if not ext:
        filename += ".bin"

    return filename


def extract_main_content(html: str) -> str:
    soup = BeautifulSoup(html, "html.parser")

    for tag in soup(["script", "style", "nav", "header", "footer", "aside", "noscript"]):
        tag.decompose()

    selectors = [
        "main",
        "article",
        '[role="main"]',
        ".content",
        "#content",
        ".documentation",
        ".docs-content",
        ".page-content",
        ".article-content",
        ".entry-content",
        ".topic-content",
        ".topic",
        ".body",
    ]

    best = None
    best_len = 0

    for selector in selectors:
        el = soup.select_one(selector)
        if el:
            text_len = len(el.get_text(" ", strip=True))
            if text_len > best_len:
                best = el
                best_len = text_len

    if best and best_len > 100:
        return str(best)

    return str(soup.body) if soup.body else html


def get_best_image_url(img_tag, page_url: str) -> str | None:
    candidates = [
        img_tag.get("src"),
        img_tag.get("data-src"),
        img_tag.get("data-lazy-src"),
        img_tag.get("data-original"),
        img_tag.get("data-url"),
    ]

    srcset = img_tag.get("srcset")
    if srcset:
        first = srcset.split(",")[0].strip().split(" ")[0]
        candidates.append(first)

    for candidate in candidates:
        if candidate and candidate.strip():
            return normalize_url(urljoin(page_url, candidate.strip()))

    return None


def download_image(image_url: str, page_url: str) -> str | None:
    try:
        abs_url = normalize_url(urljoin(page_url, image_url))
        parsed = urlparse(abs_url)

        if not parsed.scheme.startswith("http"):
            return None

        os.makedirs(IMAGES_DIR, exist_ok=True)

        filename = safe_image_filename_from_url(abs_url)
        local_path = os.path.join(IMAGES_DIR, filename)

        if os.path.exists(local_path):
            return os.path.join("images", filename)

        resp = session.get(abs_url, timeout=30)
        resp.raise_for_status()

        with open(local_path, "wb") as f:
            f.write(resp.content)

        return os.path.join("images", filename)

    except Exception as e:
        print(f"Failed to download image {image_url}: {e}")
        return None


def rewrite_images_and_convert(html: str, page_url: str) -> str:
    soup = BeautifulSoup(html, "html.parser")

    for img in soup.find_all("img"):
        best_img_url = get_best_image_url(img, page_url)
        if not best_img_url:
            continue

        local_img = download_image(best_img_url, page_url)
        if local_img:
            img["src"] = local_img

        # remove lazy-load attrs so markdownify uses the rewritten src
        for attr in ["data-src", "data-lazy-src", "data-original", "data-url", "srcset"]:
            if attr in img.attrs:
                del img.attrs[attr]

    return md(str(soup), heading_style="ATX")


def save_markdown(url: str, markdown: str) -> str:
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    filename = safe_md_filename_from_url(url)
    path = os.path.join(OUTPUT_DIR, filename)

    with open(path, "w", encoding="utf-8") as f:
        f.write(markdown)

    return path


def fetch_rendered_page(page, url: str) -> tuple[str, str]:
    page.goto(url, wait_until="networkidle", timeout=60000)
    page.wait_for_timeout(1500)
    final_url = normalize_url(page.url)
    html = page.content()
    return html, final_url


def collect_sidebar_links(html: str, page_url: str) -> list[str]:
    soup = BeautifulSoup(html, "html.parser")
    links = set()

    for selector in LINK_CONTAINER_SELECTORS:
        container = soup.select_one(selector)
        if not container:
            continue

        for a in container.find_all("a", href=True):
            href = a["href"].strip()
            if not href:
                continue

            if href.startswith(("mailto:", "tel:", "javascript:")):
                continue

            abs_url = normalize_url(urljoin(page_url, href))
            if allowed_url(abs_url):
                links.add(abs_url)

        # stop after first matching container with links
        if links:
            break

    return sorted(links)


def scrape_page(page, url: str) -> tuple[str, list[str], str]:
    html, final_url = fetch_rendered_page(page, url)
    print(f"\nFetched: {final_url}")
    print(f"HTML length: {len(html)}")

    links = collect_sidebar_links(html, final_url)
    print(f"Collected sidebar links: {len(links)}")

    main_html = extract_main_content(html)
    markdown_body = rewrite_images_and_convert(main_html, final_url).strip()

    print(f"Markdown length: {len(markdown_body)}")

    if len(markdown_body) < 20:
        markdown_body = "[No extractable content found]\n"

    markdown = f"<!-- Source: {final_url} -->\n\n{markdown_body}\n"
    return markdown, links, final_url


def crawl(start_url: str, max_pages: int):
    visited = set()
    queue = deque([normalize_url(start_url)])

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(user_agent=HEADERS["User-Agent"])
        page = context.new_page()

        while queue and len(visited) < max_pages:
            url = queue.popleft()

            if url in visited:
                continue

            if not allowed_url(url):
                continue

            visited.add(url)

            try:
                print(f"\nScraping: {url}")
                markdown, links, final_url = scrape_page(page, url)
                save_path = save_markdown(final_url, markdown)
                print(f"Saved: {save_path}")

                for link in links:
                    if link not in visited:
                        queue.append(link)

                time.sleep(DELAY_SECONDS)

            except Exception as e:
                print(f"Failed: {url} -> {e}")

        context.close()
        browser.close()

    print("\nCrawl finished.")


if __name__ == "__main__":
    crawl(START_URL, MAX_PAGES)