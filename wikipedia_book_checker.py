import time
import math
import requests
from dotenv import load_dotenv
import os
from datetime import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, Border, Side

load_dotenv()

# Information in Other.txt if LIVE_MODE = FALSE
LIVE_MODE = True
API_KEY = os.getenv("GEMINI_API_KEY")
WIKIPEDIA_API = "https://en.wikipedia.org/w/api.php"
HEADERS = {"User-Agent": "BookVerifier/1.0 (educational project)"}

# 1. CONFIDENCE INTERVAL ENGINE
def compute_confidence_interval(results):
    total = len(results)
    verified = sum(1 for r in results if r["found"])

    if total == 0:
        return {
            "total": 0,
            "verified": 0,
            "not_verified": 0,
            "proportion": 0,
            "margin_error": 0,
            "lower": 0,
            "upper": 0
        }

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


# 2. WIKIPEDIA CHECKER
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


# 3. INPUT LOADER
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

# 4. TERMINAL OUTPUT
def print_header():
    print("\n" + "=" * 60)
    print("BOOK VERIFICATION")
    print("=" * 60 + "\n")


def print_result(book, found):
    status = "VERIFIED" if found else "NOT FOUND"
    print(f"{status:<15} | {book['title']}")


def print_summary(stats):
    print("\n" + "=" * 60)
    print("FINAL STATISTICS")
    print("=" * 60)

    print(f"Total Books        : {stats['total']}")
    print(f"Verified Books     : {stats['verified']}")
    print(f"Not Verified       : {stats['not_verified']}")
    print(f"Proportion Found   : {stats['proportion']*100:.2f}%")
    print(f"Margin of Error    : ±{stats['margin_error']*100:.2f}%")
    print(f"95% CI Range       : {stats['lower']*100:.2f}% → {stats['upper']*100:.2f}%")


# 5. RUN CHECKS
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

        print_result(book, found)

        if LIVE_MODE:
            time.sleep(0.3)

    return results


# 6. EXPORT EXCEL
def export_excel(results, stats):
    wb = Workbook()

    # ---------------- SHEET 1 ----------------
    ws = wb.active
    ws.title = "Results"

    headers = ["#", "Title", "Author", "Genre", "Status", "Wikipedia Title", "URL"]

    for c, h in enumerate(headers, 1):
        ws.cell(1, c, h).font = Font(bold=True)

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

    ws2["A1"] = "BOOK VERIFICATION STATISTICS"
    ws2["A1"].font = Font(bold=True, size=14)

    summary = [
        ("Total Books", stats["total"]),
        ("Verified Books", stats["verified"]),
        ("Not Verified Books", stats["not_verified"]),
        ("Proportion Verified", f"{stats['proportion']*100:.2f}%"),
        ("Margin of Error", f"±{stats['margin_error']*100:.2f}%"),
        ("Lower Bound (95%)", f"{stats['lower']*100:.2f}%"),
        ("Upper Bound (95%)", f"{stats['upper']*100:.2f}%"),
    ]

    row_start = 3

    for i, (k, v) in enumerate(summary, row_start):
        ws2.cell(i, 1, k).font = Font(bold=True)
        ws2.cell(i, 2, v)

    wb.save("books_verified.xlsx")


# MAIN
def main():
    print_header()

    books = load_books()
    print(f"Loaded {len(books)} books\n")

    results = run_checks(books)

    stats = compute_confidence_interval(results)

    print_summary(stats)

    export_excel(results, stats)

    print()     
    print("Verification Finished!")
    print("File saved as books_verified.xlsx")


if __name__ == "__main__":
    main()