import re
import sys
import time
from pathlib import Path
from difflib import SequenceMatcher

FILE_PATH = "twan_updated.xlsx"
OUTPUT_FILE = "twan_fully_updated.xlsx"
NOT_FOUND_FILE = "not_found.txt"
UNCERTAIN_FILE = "uncertain_matches.txt"
BASE_URL = "https://www.outsystems.com"
SEARCH_URL = f"{BASE_URL}/forge/list?q={{query}}"
REQUEST_TIMEOUT = 20
MAX_RESULTS = 5
SIMILARITY_THRESHOLD = 0.6

NOT_FOUND = []
UNCERTAIN = []


def similarity(a: str, b: str) -> float:
    return SequenceMatcher(None, a.lower(), b.lower()).ratio()


def get_search_results(name: str):
    import requests
    from bs4 import BeautifulSoup

    query = name.replace(" ", "+")
    res = requests.get(SEARCH_URL.format(query=query), timeout=REQUEST_TIMEOUT)
    res.raise_for_status()

    soup = BeautifulSoup(res.text, "html.parser")
    results = []

    for link in soup.select("a[href*='/forge/component-overview/']")[:MAX_RESULTS]:
        title = link.get_text(strip=True)
        href = link.get("href", "")
        if not href:
            continue
        if href.startswith("http"):
            url = href
        else:
            url = BASE_URL + href
        results.append((title, url))

    return results


def get_version_from_page(url: str):
    import requests
    from bs4 import BeautifulSoup

    try:
        res = requests.get(url, timeout=REQUEST_TIMEOUT)
        res.raise_for_status()
    except requests.RequestException:
        return None

    soup = BeautifulSoup(res.text, "html.parser")
    text = "\n".join(line.strip() for line in soup.get_text("\n").splitlines() if line.strip())

    match = re.search(r"Version\s*[:\-]?\s*([0-9]+(?:\.[0-9]+){0,3})", text, flags=re.IGNORECASE)
    if match:
        return match.group(1)

    return None


def compare_versions(current: str, latest: str) -> bool:
    return str(current).strip() != str(latest).strip()


def format_current_version(row) -> str:
    major = row.get("Forge Major Version", "")
    minor = row.get("Forge Minor Version", "")
    revision = row.get("Forge Revision Version", "")

    parts = [major, minor, revision]
    cleaned = [str(part).strip() for part in parts if str(part).strip() and str(part).strip() != "nan"]
    return ".".join(cleaned)


def main():
    input_path = Path(FILE_PATH).resolve()
    output_path = Path(OUTPUT_FILE).resolve()
    not_found_path = Path(NOT_FOUND_FILE).resolve()
    uncertain_path = Path(UNCERTAIN_FILE).resolve()

    print(f"Invoer Excel verwacht op: {input_path}")
    print(f"Output Excel komt op: {output_path}")

    try:
        import pandas as pd
    except ModuleNotFoundError:
        print("Missing dependency: pandas. Install it with: pip install pandas openpyxl")
        sys.exit(1)

    if not input_path.exists():
        print(f"Invoerbestand niet gevonden: {input_path}")
        print("Plaats je Excel-bestand als 'twan_updated.xlsx' in dezelfde map als dit script.")
        sys.exit(1)

    try:
        import requests  # noqa: F401
        from bs4 import BeautifulSoup  # noqa: F401
    except ModuleNotFoundError as exc:
        missing_module = exc.name or "required package"
        print(f"Missing dependency: {missing_module}. Install with: pip install requests beautifulsoup4")
        sys.exit(1)

    try:
        df = pd.read_excel(input_path)
    except ImportError:
        print("Missing dependency: openpyxl. Install it with: pip install openpyxl")
        sys.exit(1)

    for index, row in df.iterrows():
        name = str(row.get("Name", "")).strip()
        if not name:
            continue

        current_version = format_current_version(row)
        print(f"\nChecking: {name}")

        try:
            results = get_search_results(name)
        except requests.RequestException:
            NOT_FOUND.append(name)
            continue

        if not results:
            NOT_FOUND.append(name)
            continue

        best_match = None
        best_score = 0.0

        for title, url in results:
            score = similarity(name, title)
            if score > best_score:
                best_score = score
                best_match = (title, url)

        if not best_match:
            NOT_FOUND.append(name)
            continue

        title, url = best_match
        latest_version = get_version_from_page(url)

        if not latest_version:
            NOT_FOUND.append(name)
            continue

        df.at[index, "Laatste Forge versie"] = latest_version
        df.at[index, "Update beschikbaar (ja/nee)"] = "ja" if compare_versions(current_version, latest_version) else "nee"

        if best_score < SIMILARITY_THRESHOLD:
            UNCERTAIN.append((name, title))

        time.sleep(1)

    df.to_excel(output_path, index=False)

    with open(not_found_path, "w", encoding="utf-8") as f:
        for item in NOT_FOUND:
            f.write(item + "\n")

    with open(uncertain_path, "w", encoding="utf-8") as f:
        for orig, match in UNCERTAIN:
            f.write(f"{orig} -> {match}\n")

    print("\nKlaar!")
    print(f"Output Excel: {output_path}")
    print(f"Niet gevonden-lijst: {not_found_path}")
    print(f"Twijfelachtige matches-lijst: {uncertain_path}")
    print(f"Niet gevonden: {len(NOT_FOUND)}")
    print(f"Twijfelachtig: {len(UNCERTAIN)}")


if __name__ == "__main__":
    main()
