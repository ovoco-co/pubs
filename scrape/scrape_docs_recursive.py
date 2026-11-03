import os
import re
import time
from urllib.parse import urljoin, urlparse, urldefrag

import requests
from bs4 import BeautifulSoup
from markdownify import markdownify as md


START_URL = "https://www.itophub.io/"
MAX_DEPTH = 2
OUTPUT_DIR = "downloaded_docs"
DELAY_SECONDS = 1.0

# Only crawl links found inside this container
LINK_CONTAINER_SELECTOR = "#chapter-menu"

# Optional: keep crawling restricted to the same path prefix
ALLOWED_PATH_PREFIX = "/"

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
    url, _ = urldefrag(url)
    return url.rstrip("/")


def is_same_site(url: str, allowed_netloc: str) -> bool:
    parsed = urlparse(url)
    return parsed.netloc == allowed_netloc


def is_allowed_path(url: str) -> bool:
    parsed = urlparse(url)
    return parsed.path.startswith(ALLOWED_PATH_PREFIX)


def extract_main_content(html: str) -> str:
    soup = BeautifulSoup(html, "html.parser")

    # Remove obvious non-content elements
    for tag in soup(["script", "style", "nav", "header", "footer", "aside", "noscript"]):
        tag.decompose()

    # Try common content containers
    selectors = [
        "article",
        "main",
        '[role="main"]',
        ".content",
        "#content",
        ".page-content",
        ".entry-content",
        ".post-content",
        ".article-content",
        ".docs-content",
        ".wiki-content",
        ".markdown-body",
        ".field-name-body",
        ".region-content",
        ".view-content",
        "#primary",
    ]

    best = None
    best_len = 0

    for selector in selectors:
        element = soup.select_one(selector)
        if element:
            text_len = len(element.get_text(" ", strip=True))
            if text_len > best_len:
                best = element
                best_len = text_len

    if best and best_len > 100:
        return str(best)

    # Fallback to body if it exists
    if soup.body:
        return str(soup.body)

    return html


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

    # Collect links only from the chosen container
    links = set()
    container = soup.select_one(LINK_CONTAINER_SELECTOR)

    if container:
        for a in container.find_all("a", href=True):
            href = a["href"].strip()
            if not href:
                continue

            abs_url = normalize_url(urljoin(final_url, href))

            if is_same_site(abs_url, urlparse(START_URL).netloc) and is_allowed_path(abs_url):
                links.add(abs_url)
    else:
        print(f"Warning: link container not found on {final_url}")

    main_html = extract_main_content(html)
    markdown_body = html_to_markdown(main_html).strip()

    # Skip blank or near-blank pages
    if len(markdown_body) < 20:
        markdown_body = "[No extractable content found]\n"

    markdown = f"<!-- Source: {final_url} -->\n\n{markdown_body}\n"

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

        if not is_allowed_path(current_url):
            continue

        visited.add(current_url)

        try:
            print(f"[depth={depth}] Scraping: {current_url}")
            markdown, links, final_url = scrape_page(current_url)
            save_path = save_markdown(OUTPUT_DIR, final_url, markdown)
            print(f"Saved: {save_path}")

            if depth < max_depth:
                for link in links:
                    if link not in visited:
                        queue.append((link, depth + 1))

            time.sleep(DELAY_SECONDS)

        except Exception as e:
            print(f"Failed: {current_url} -> {e}")


if __name__ == "__main__":
    crawl(START_URL, MAX_DEPTH)