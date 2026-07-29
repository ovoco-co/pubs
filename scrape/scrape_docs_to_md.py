import os
import re
import time
from urllib.parse import urljoin, urlparse, urldefrag

import requests
from bs4 import BeautifulSoup
from markdownify import markdownify as md


START_URL = "https://www.itophub.io/"
MAX_DEPTH = 4
OUTPUT_DIR = "downloaded_docs"
DELAY_SECONDS = 1.0

HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; DocsScraper/1.0)"
}

session = requests.Session()
session.headers.update(HEADERS)


def safe_filename_from_url(url: str) -> str:
    parsed = urlparse(url)
    path = parsed.path.strip("/")

    if not path:
        path = "index"

    filename = f"{parsed.netloc}_{path}"
    filename = re.sub(r"[^\w\-_.]+", "_", filename)

    if parsed.query:
        filename += "_" + re.sub(r"[^\w\-_.]+", "_", parsed.query)

    if not filename.endswith(".md"):
        filename += ".md"

    return filename


def normalize_url(url: str) -> str:
    # Remove fragment identifiers (#...)
    url, _ = urldefrag(url)
    return url.rstrip("/")


def is_same_site(url: str, allowed_netloc: str) -> bool:
    parsed = urlparse(url)
    return parsed.netloc == allowed_netloc


def extract_main_content(html: str) -> str:
    soup = BeautifulSoup(html, "html.parser")

    for tag in soup(["script", "style", "nav", "header", "footer", "aside"]):
        tag.decompose()

    # Try likely content containers first
    selectors = [
        "main",
        "article",
        ".content",
        "#content",
        ".wiki-content",
        ".page-content",
        ".documentation",
        ".docs-content",
    ]

    for selector in selectors:
        element = soup.select_one(selector)
        if element:
            return str(element)

    body = soup.body
    return str(body) if body else html


def html_to_markdown(html: str) -> str:
    return md(html, heading_style="ATX")


def save_markdown(output_dir: str, url: str, markdown: str) -> str:
    os.makedirs(output_dir, exist_ok=True)
    filename = safe_filename_from_url(url)
    path = os.path.join(output_dir, filename)

    with open(path, "w", encoding="utf-8") as f:
        f.write(markdown)

    return path


def scrape_page(url: str) -> tuple[str, list[str], str]:
    response = session.get(url, timeout=30)
    response.raise_for_status()

    final_url = normalize_url(response.url)
    html = response.text

    soup = BeautifulSoup(html, "html.parser")

    # Collect links before stripping elements
    links = set()
    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        abs_url = normalize_url(urljoin(final_url, href))
        links.add(abs_url)

    main_html = extract_main_content(html)
    markdown = html_to_markdown(main_html)
    markdown = f"<!-- Source: {final_url} -->\n\n{markdown.strip()}\n"

    return markdown, sorted(links), final_url


def crawl(url: str, max_depth: int):
    allowed_netloc = urlparse(url).netloc
    visited = set()
    queue = [(normalize_url(url), 0)]

    while queue:
        current_url, depth = queue.pop(0)

        if current_url in visited:
            continue

        if depth > max_depth:
            continue

        if not is_same_site(current_url, allowed_netloc):
            continue

        visited.add(current_url)

        try:
            print(f"[depth={depth}] Scraping: {current_url}")
            markdown, links, final_url = scrape_page(current_url)
            save_path = save_markdown(OUTPUT_DIR, final_url, markdown)
            print(f"Saved: {save_path}")

            if depth < max_depth:
                for link in links:
                    if link not in visited and is_same_site(link, allowed_netloc):
                        queue.append((link, depth + 1))

            time.sleep(DELAY_SECONDS)

        except Exception as e:
            print(f"Failed: {current_url} -> {e}")


if __name__ == "__main__":
    crawl(START_URL, MAX_DEPTH)