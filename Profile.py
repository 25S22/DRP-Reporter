
import re
import time
import random
import sys
import inspect
import pandas as pd
from googlesearch import search

# ── Terminal Input ─────────────────────────────────────────────────────────────
def prompt(msg, default=None):
    suffix = f" [{default}]" if default is not None else ""
    val = input(f"{msg}{suffix}: ").strip()
    return val if val else (str(default) if default is not None else "")

def get_char_rule(field_label):
    while True:
        val = input(f"  {field_label} chars to use ('all' or a number): ").strip().lower()
        if val == "all":
            return "all"
        try:
            n = int(val)
            if n > 0:
                return n
            print("  [!] Enter a positive integer or 'all'.")
        except ValueError:
            print("  [!] Invalid input. Type 'all' or a positive integer.")

def banner():
    print("""
╔══════════════════════════════════════════════════════╗
║          LinkedIn OSINT Automation Tool              ║
║          Searches Google → Parses → Excel            ║
╚══════════════════════════════════════════════════════╝
""")

def collect_inputs():
    banner()
    print("── Step 1: Target Organisation ───────────────────────────────────")
    org_name = ""
    while not org_name:
        org_name = input("  Company / Organisation name: ").strip()
        if not org_name:
            print("  [!] Organisation name cannot be empty.")

    print("\n── Step 2: Search Volume ─────────────────────────────────────────")
    while True:
        raw = prompt("  Max results to retrieve (max 10000)", default=10000)
        try:
            num_results = int(raw)
            if 0 < num_results <= 10000:
                break
            print("  [!] Must be between 1 and 10000.")
        except ValueError:
            print("  [!] Please enter a valid integer.")

    print("\n── Step 3: Pseudo-Username Rules ─────────────────────────────────")
    print("  These rules control how usernames like  john_smi  are built.")
    first_name_chars = get_char_rule("First name")
    last_name_chars  = get_char_rule("Last  name")

    print("\n── Step 4: Username Format ───────────────────────────────────────")
    print("  Separator between first and last part:")
    print("  [1] Underscore  →  john_smi")
    print("  [2] Dot         →  john.smi")
    print("  [3] None        →  johnsmi")
    sep_choice = ""
    while sep_choice not in ("1", "2", "3"):
        sep_choice = input("  Choose [1/2/3]: ").strip()
    separator = {"1": "_", "2": ".", "3": ""}[sep_choice]

    print("\n── Configuration Summary ─────────────────────────────────────────")
    print(f"  Organisation  : {org_name}")
    print(f"  Max results   : {num_results}")
    print(f"  First name    : {'ALL chars' if first_name_chars == 'all' else str(first_name_chars) + ' char(s)'}")
    print(f"  Last name     : {'ALL chars' if last_name_chars  == 'all' else str(last_name_chars)  + ' char(s)'}")
    print(f"  Separator     : '{separator}'")
    confirm = input("\n  Proceed? [Y/n]: ").strip().lower()
    if confirm == "n":
        print("  Aborted.")
        sys.exit(0)

    return org_name, num_results, first_name_chars, last_name_chars, separator

# ── Username Builder ───────────────────────────────────────────────────────────
def build_username(first, last, first_chars, last_chars, sep):
    fn = first.lower()
    ln = last.lower()
    fn_part = fn  if str(first_chars) == "all" else fn[:int(first_chars)]
    ln_part = ln  if str(last_chars)  == "all" else ln[:int(last_chars)]
    return f"{fn_part}{sep}{ln_part}"

# ── Title Parser ───────────────────────────────────────────────────────────────
TITLE_RE = re.compile(
    r"^(?P<name>[^|\-]+?)\s*[-–]\s*(?P<position>[^|\-][^|]+?)\s*(?:[-–|])",
    re.UNICODE
)

def parse_title(title):
    if not title:
        return None, None
    cleaned = re.sub(r"\|\s*LinkedIn\s*$", "", title, flags=re.IGNORECASE).strip()
    m = TITLE_RE.match(cleaned)
    if not m:
        return None, None
    name     = m.group("name").strip()
    position = m.group("position").strip()
    if not name or not position or len(name) < 2:
        return None, None
    return name, position

# ── Google Search Collector ────────────────────────────────────────────────────
def collect_results(org_name, num_results):
    query = f'site:linkedin.com/in "{org_name}"'
    print(f"\n[*] Query        : {query}")
    print(f"[*] Target count : {num_results}\n")

    collected      = {}
    BATCH_SIZE     = 100
    TOTAL_ATTEMPTS = 2
    RETRY_WAIT_RANGE = (12, 20)
    QUERY_SLEEP_RANGE = (2.5, 5.0)
    BATCH_SLEEP_RANGE = (4.0, 8.0)
    start          = 0
    dupe_streak    = 0
    MAX_DUPE_STREAK = 3
    search_params = inspect.signature(search).parameters
    supports_start = "start" in search_params
    supports_start_num = "start_num" in search_params
    supports_sleep_interval = "sleep_interval" in search_params
    supports_advanced = "advanced" in search_params
    supports_timeout = "timeout" in search_params
    can_paginate = supports_start or supports_start_num

    while len(collected) < num_results:
        want = min(BATCH_SIZE, num_results - len(collected))
        print(f"  → Batch start={start:>5}  want={want:<4}  "
              f"collected so far={len(collected)}", end="  ", flush=True)

        batch = []
        requested = want if can_paginate else min(num_results, len(collected) + want)
        for attempt in range(TOTAL_ATTEMPTS):
            try:
                kwargs = {"num_results": requested}
                if supports_advanced:
                    kwargs["advanced"] = True
                if supports_sleep_interval:
                    kwargs["sleep_interval"] = random.uniform(*QUERY_SLEEP_RANGE)
                if supports_timeout:
                    kwargs["timeout"] = 12
                if supports_start:
                    kwargs["start"] = start
                elif supports_start_num:
                    kwargs["start_num"] = start
                batch = list(search(query, **kwargs))
                break
            except Exception as e:
                if attempt < (TOTAL_ATTEMPTS - 1):
                    wait_s = random.uniform(*RETRY_WAIT_RANGE)
                    print(f"\n  [WARN] Error: {e}. Retrying in {wait_s:.1f} s …")
                    time.sleep(wait_s)
                else:
                    print(f"\n  [ERROR] Retry failed: {e}. Stopping.")
                    return collected

        if not batch:
            print("\n  [*] Google returned empty batch – no more results available.")
            break

        new_count = 0
        for sr in batch:
            url = getattr(sr, "url", "") or ""
            if url and url not in collected:
                collected[url] = sr
                new_count += 1

        print(f"new={new_count}")

        if new_count == 0:
            dupe_streak += 1
            print(f"  [WARN] No new URLs ({dupe_streak}/{MAX_DUPE_STREAK} "
                  "consecutive empty batches)")
            if dupe_streak >= MAX_DUPE_STREAK:
                print("  [*] Google exhausted for this query. Stopping.")
                break
        else:
            dupe_streak = 0

        start = start + len(batch) if can_paginate else len(collected)
        sleep_s = random.uniform(*BATCH_SLEEP_RANGE)
        print(f"       sleeping {sleep_s:.1f} s …")
        time.sleep(sleep_s)

    print(f"\n[*] Collection done. Unique URLs: {len(collected)}")
    return collected

# ── Build Rows ─────────────────────────────────────────────────────────────────
def build_rows(collected, first_name_chars, last_name_chars, separator):
    rows = []
    for url, sr in collected.items():
        title    = getattr(sr, "title", "") or ""
        desc     = getattr(sr, "description", "") or ""
        name, position = parse_title(title)

        pseudo = ""
        if name:
            parts = name.split()
            if len(parts) >= 2:
                try:
                    pseudo = build_username(parts[0], parts[-1],
                                            first_name_chars, last_name_chars,
                                            separator)
                except Exception:
                    pass
            elif len(parts) == 1:
                try:
                    pseudo = build_username(parts[0], "",
                                            first_name_chars, last_name_chars,
                                            separator)
                except Exception:
                    pass

        rows.append({
            "Name":            name     or "",
            "Position":        position or "",
            "Pseudo_Username": pseudo,
            "URL":             url,
            "Description":     desc,
            "Raw_Title":       title,
        })
    return rows

# ── Export Excel ───────────────────────────────────────────────────────────────
def auto_fit(ws):
    for col_cells in ws.columns:
        max_len = max(
            (len(str(c.value)) for c in col_cells if c.value is not None),
            default=10,
        )
        ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 4, 80)

def export_excel(rows, org_name):
    df_all      = pd.DataFrame(rows, columns=["Name", "Position",
                                               "Pseudo_Username", "URL",
                                               "Description", "Raw_Title"])
    df_parsed   = df_all[df_all["Name"] != ""].drop(
                    columns=["Raw_Title"]).reset_index(drop=True)
    df_unparsed = df_all[df_all["Name"] == ""][
                    ["URL", "Raw_Title", "Description"]].reset_index(drop=True)

    output_file = f"OSINT_{org_name.replace(' ', '_')}.xlsx"
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df_parsed.to_excel(writer,   index=False, sheet_name="Parsed")
        df_unparsed.to_excel(writer, index=False, sheet_name="Unparsed_Raw")
        auto_fit(writer.sheets["Parsed"])
        auto_fit(writer.sheets["Unparsed_Raw"])

    print(f"\n[✓] Saved  →  {output_file}")
    print(f"    Parsed rows    : {len(df_parsed)}")
    print(f"    Unparsed rows  : {len(df_unparsed)}")
    return output_file

# ── Main ───────────────────────────────────────────────────────────────────────
def main():
    org_name, num_results, first_name_chars, last_name_chars, separator = collect_inputs()

    collected = collect_results(org_name, num_results)

    if not collected:
        print("[!] No results collected. Exiting.")
        sys.exit(1)

    print("\n[*] Parsing titles and building usernames …")
    rows = build_rows(collected, first_name_chars, last_name_chars, separator)

    export_excel(rows, org_name)

if __name__ == "__main__":
    main()
