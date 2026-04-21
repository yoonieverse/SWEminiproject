import time
import math
import requests
from dotenv import load_dotenv
import os
from datetime import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

load_dotenv()

LIVE_MODE = True
API_KEY = os.getenv("GEMINI_API_KEY")
WIKIPEDIA_API = "https://en.wikipedia.org/w/api.php"
HEADERS = {"User-Agent": "BookVerifier/1.0 (educational project)"}

# -----------------------------
# 1. CONFIDENCE INTERVAL ENGINE
# -----------------------------
def compute_confidence_interval(results):
    total = len(results)
    verified = sum(1 for r in results if r["found"])

    if total == 0:
        return dict(total=0, verified=0, not_verified=0,
                    proportion=0, margin_error=0, lower=0, upper=0)

    p = verified / total
    z = 1.96

    margin_error = z * math.sqrt((p * (1 - p)) / total)

    return {
        "total": total,
        "verified": verified,
        "not_verified": total - verified,
        "proportion": p,
        "margin_error": margin_error,
        "lower": max(0, p - margin_error),
        "upper": min(1, p + margin_error)
    }


# -----------------------------
# 2. WIKIPEDIA CHECKER
# -----------------------------
def check_live(title):
    try:
        params = {
            "action": "query",
            "list": "search",
            "srsearch": f'"{title}" book',
            "srlimit": 3,
            "format": "json",
        }

        resp = requests.get(WIKIPEDIA_API, params=params, headers=HEADERS, timeout=8)
        data = resp.json()
        results = data.get("query", {}).get("search", [])

        if not results:
            return False, "Not found on Wikipedia", ""

        top = results[0]["title"]
        url = f"https://en.wikipedia.org/wiki/{top.replace(' ', '_')}"

        return True, top, url

    except Exception as e:
        return False, f"Error: {e}", ""


def check_wikipedia(title):
    return check_live(title)


# -----------------------------
# 3. INPUT LOADER
# -----------------------------
def load_books():
    wb = load_workbook("books_input.xlsx")
    ws = wb.active

    books = []

    for row in ws.iter_rows(min_row=1, values_only=True):
        if not row or not row[1]:
            continue

        books.append({
            "num": row[0],
            "title": row[1],
            "author": row[2],
            "genre": row[3]
        })

    return books


# -----------------------------
# 4. RUN CHECKS
# -----------------------------
def run_checks(books):
    results = []

    for book in books:
        found, wiki_title, url = check_wikipedia(book["title"])

        results.append({
            **book,
            "found": found,
            "wiki_title": wiki_title,
            "url": url if found else "N/A"
        })

        print(("OK" if found else "XX"), book["title"])

        if LIVE_MODE:
            time.sleep(0.3)

    return results


# -----------------------------
# 5. EXPORT EXCEL
# -----------------------------
def export_excel(results, stats):
    wb = Workbook()

    thin = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    # ---------------- SHEET 1 ----------------
    ws = wb.active
    ws.title = "Results"

    headers = ["#", "Title", "Author", "Genre", "Status", "Wikipedia Title", "URL"]
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = Font(bold=True)
        cell.border = thin

    for i, r in enumerate(results, 2):
        ws.cell(i, 1, r["num"])
        ws.cell(i, 2, r["title"])
        ws.cell(i, 3, r["author"])
        ws.cell(i, 4, r["genre"])
        ws.cell(i, 5, "VERIFIED" if r["found"] else "NOT FOUND")
        ws.cell(i, 6, r["wiki_title"])
        ws.cell(i, 7, r["url"])

    # ---------------- SHEET 2 ----------------
    ws2 = wb.create_sheet("Stats")

    summary = [
        ("Total", stats["total"]),
        ("Verified", stats["verified"]),
        ("Not Verified", stats["not_verified"]),
        ("Proportion", f"{stats['proportion']*100:.2f}%"),
        ("Margin of Error", f"{stats['margin_error']*100:.2f}%"),
        ("Lower Bound", f"{stats['lower']*100:.2f}%"),
        ("Upper Bound", f"{stats['upper']*100:.2f}%"),
    ]

    for i, (k, v) in enumerate(summary, 1):
        ws2.cell(i, 1, k)
        ws2.cell(i, 2, v)

    wb.save("books_verified.xlsx")


# -----------------------------
# MAIN
# -----------------------------
def main():
    books = load_books()
    print(f"Loaded {len(books)} books")

    results = run_checks(books)

    stats = compute_confidence_interval(results)

    export_excel(results, stats)

    print("\nDONE")
    print(stats)


if __name__ == "__main__":
    main()