import os
import re
import time
from typing import List, Tuple
from urllib.parse import urljoin, urlparse, urldefrag

import requests
from bs4 import BeautifulSoup
from markdownify import markdownify as md


START_URL = "https://www.itophub.io/wiki/page?id=3_3_0:start"
MAX_DEPTH = 4
OUTPUT_DIR = "/home/artist/ovoco/ovoco/documentation/itop3-3"
IMAGES_DIR = os.path.join(OUTPUT_DIR, "images")
DELAY_SECONDS = 3.0

# For now, collect links from the whole page
LINK_CONTAINER_SELECTOR = "body"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; DocsScraper/1.0)"
}

session = requests.Session()
session.headers.update(HEADERS)


def safe_filename_from_url(url: str, default_ext: str = ".bin") -> str:
    parsed = urlparse(url)
    path = parsed.path.strip("/")

    if not path:
        path = "index"

    # A wiki page is named by its id, "3_3_0:start", not by the site and the script that served
    # it. A URL with no id keeps the last path segment and the query.
    query = dict(q.split("=", 1) for q in parsed.query.split("&") if "=" in q)
    if "id" in query:
        filename = query["id"]
    else:
        filename = path.rsplit("/", 1)[-1] + ("_" + parsed.query if parsed.query else "")

    filename = re.sub(r"[^\w\-_.]+", "_", filename)

    # The site name carries a dot, so splitext read it as an extension and never added one.
    if not filename.lower().endswith(default_ext):
        filename += default_ext

    return filename


def normalize_url(url: str) -> str:
    url, _ = urldefrag(url)
    return url.rstrip("/")


def is_same_site(url: str, allowed_netloc: str) -> bool:
    return urlparse(url).netloc == allowed_netloc


def already_downloaded(url: str) -> bool:
    filename = safe_filename_from_url(url, ".md")
    path = os.path.join(OUTPUT_DIR, filename)
    return os.path.exists(path)


def extract_main_content(html: str) -> str:
    soup = BeautifulSoup(html, "html.parser")

    for tag in soup(["script", "style", "nav", "header", "footer", "aside", "noscript"]):
        tag.decompose()

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

    if soup.body:
        return str(soup.body)

    return html


def download_image(image_url: str, page_url: str) -> str | None:
    """
    Download an image and return the local relative path to use in markdown.
    """
    try:
        abs_url = normalize_url(urljoin(page_url, image_url))
        parsed = urlparse(abs_url)

        if not parsed.scheme.startswith("http"):
            return None

        os.makedirs(IMAGES_DIR, exist_ok=True)

        # The wiki serves every image through fetch.php?media=<id>, so the path's last segment
        # is the same for all of them; the media id is the name. A plain image URL keeps its own.
        query = dict(q.split("=", 1) for q in parsed.query.split("&") if "=" in q)
        path = parsed.path.strip("/")
        base_name = query.get("media") or (path.split("/")[-1] if path else "image")
        base_name = re.sub(r"[^\w\-_.]+", "_", base_name)

        if not os.path.splitext(base_name)[1]:
            base_name += ".bin"

        filename = base_name
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


def html_to_markdown_with_images(html: str, page_url: str) -> str:
    """
    Convert HTML to markdown after downloading images and rewriting img src values.
    """
    soup = BeautifulSoup(html, "html.parser")

    # Download and rewrite image sources
    for img in soup.find_all("img"):
        src = img.get("src")
        if not src:
            continue

        local_img = download_image(src, page_url)
        if local_img:
            img["src"] = local_img

    return md(str(soup), heading_style="ATX")


def save_markdown(output_dir: str, url: str, markdown: str) -> str:
    os.makedirs(output_dir, exist_ok=True)
    filename = safe_filename_from_url(url, ".md")
    path = os.path.join(output_dir, filename)

    with open(path, "w", encoding="utf-8") as f:
        f.write(markdown)

    return path


def scrape_page(url: str) -> Tuple[str, List[str], str]:
    response = session.get(url, timeout=30)
    response.raise_for_status()

    final_url = normalize_url(response.url)
    html = response.text

    print(f"\nFetched: {final_url}")
    print(f"HTML length: {len(html)}")

    soup = BeautifulSoup(html, "html.parser")

    # Collect links
    links = set()
    container = soup.select_one(LINK_CONTAINER_SELECTOR)

    if container:
        print(f"Found link container: {LINK_CONTAINER_SELECTOR}")
        for a in container.find_all("a", href=True):
            href = a["href"].strip()
            if not href:
                continue

            abs_url = normalize_url(urljoin(final_url, href))

            if is_same_site(abs_url, urlparse(START_URL).netloc):
                links.add(abs_url)
    else:
        print(f"Link container NOT found: {LINK_CONTAINER_SELECTOR}")

    print(f"Collected {len(links)} links")

    main_html = extract_main_content(html)
    markdown_body = html_to_markdown_with_images(main_html, final_url).strip()

    print(f"Markdown length: {len(markdown_body)}")

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
        print(f"\nQueue size: {len(queue)} | Visiting: {current_url} | depth={depth}")

        if current_url in visited:
            print("Already visited, skipping")
            continue

        if depth > max_depth:
            print("Too deep, skipping")
            continue

        if not is_same_site(current_url, allowed_netloc):
            print("Different site, skipping")
            continue

        if already_downloaded(current_url):
            print(f"Already downloaded, skipping: {current_url}")
            visited.add(current_url)
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

    print("\nCrawl finished.")


if __name__ == "__main__":
    crawl(START_URL, MAX_DEPTH)