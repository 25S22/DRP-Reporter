```python
import yaml
import re
import time
import random
import pandas as pd
from googlesearch import search

# ── Load config ────────────────────────────────────────────────────────────────
with open("config.yaml", "r") as f:
    config = yaml.safe_load(f)

org_name         = str(config["org_name"])
num_results      = int(config.get("num_results", 1000))
first_name_chars = config["first_name_chars"]
last_name_chars  = int(config["last_name_chars"])

# ── Helper: build pseudo-username ──────────────────────────────────────────────
def build_username(first: str, last: str) -> str:
    fn = first.lower()
    ln = last.lower()
    if str(first_name_chars).strip().lower() == "all":
        fn_part = fn
    else:
        fn_part = fn[:int(first_name_chars)]
    ln_part = ln[:last_name_chars]
    return f"{fn_part}_{ln_part}"

# ── Helper: parse LinkedIn title → (name, position) ───────────────────────────
TITLE_RE = re.compile(
    r"^(?P<name>[^|\-]+?)\s*[-–]\s*(?P<position>[^|\-][^|]+?)\s*(?:[-–|])",
    re.UNICODE
)

def parse_title(title: str):
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

# ── Collect Google results in batches to reach num_results ────────────────────
query = f'site:linkedin.com/in "{org_name}"'

print(f"[*] Starting search: {query}")
print(f"[*] Target results : {num_results}")

collected_urls   = {}   # url → SearchResult  (dedup by url)
BATCH_SIZE       = 100  # request this many per call (googlesearch max chunk)
collected_count  = 0
start            = 0
consecutive_dupes = 0
MAX_CONSEC_DUPES  = 3   # stop if N consecutive batches yield no new results

while collected_count < num_results:
    want = min(BATCH_SIZE, num_results - collected_count)
    print(f"[*] Fetching batch start={start}, want={want} "
          f"(have {collected_count}/{num_results}) …")
    try:
        batch = list(
            search(
                query,
                num_results=want,
                advanced=True,
                sleep_interval=random.uniform(3, 6),
                start=start,
            )
        )
    except Exception as e:
        print(f"[WARN] Search error at start={start}: {e}. "
              "Sleeping 30 s then retrying once …")
        time.sleep(30)
        try:
            batch = list(
                search(
                    query,
                    num_results=want,
                    advanced=True,
                    sleep_interval=random.uniform(5, 10),
                    start=start,
                )
            )
        except Exception as e2:
            print(f"[ERROR] Retry failed: {e2}. Stopping collection.")
            break

    if not batch:
        print("[*] Empty batch returned – Google has no more results.")
        break

    new_this_batch = 0
    for sr in batch:
        url = getattr(sr, "url", "") or ""
        if not url:
            continue
        if url not in collected_urls:
            collected_urls[url] = sr
            new_this_batch += 1

    print(f"    → {len(batch)} fetched, {new_this_batch} new unique")

    if new_this_batch == 0:
        consecutive_dupes += 1
        print(f"[WARN] No new results ({consecutive_dupes}/{MAX_CONSEC_DUPES} "
              "consecutive empty batches).")
        if consecutive_dupes >= MAX_CONSEC_DUPES:
            print("[*] Google appears exhausted for this query. Stopping.")
            break
    else:
        consecutive_dupes = 0

    collected_count = len(collected_urls)
    start += len(batch)

    # Polite delay between batches to avoid 429s
    sleep_secs = random.uniform(5, 12)
    print(f"    Sleeping {sleep_secs:.1f} s …")
    time.sleep(sleep_secs)

print(f"\n[*] Collection complete. Total unique URLs: {len(collected_urls)}")

# ── Parse every result → row ───────────────────────────────────────────────────
rows = []
for url, sr in collected_urls.items():
    title = getattr(sr, "title", "") or ""

    name, position = parse_title(title)

    pseudo_username = None
    if name:
        parts = name.split()
        if len(parts) >= 2:
            try:
                pseudo_username = build_username(parts[0], parts[-1])
            except Exception:
                pass
        elif len(parts) == 1:
            try:
                pseudo_username = build_username(parts[0], "")
            except Exception:
                pass

    rows.append({
        "Name":            name            or "",
        "Position":        position        or "",
        "Pseudo_Username": pseudo_username or "",
        "URL":             url,
        "Raw_Title":       title,          # kept for debugging
    })

# ── Build DataFrame ────────────────────────────────────────────────────────────
df = pd.DataFrame(rows, columns=["Name", "Position", "Pseudo_Username",
                                  "URL", "Raw_Title"])

# Split into parsed vs unparsed sheets
df_parsed   = df[df["Name"] != ""].drop(columns=["Raw_Title"]).reset_index(drop=True)
df_unparsed = df[df["Name"] == ""][["URL", "Raw_Title"]].reset_index(drop=True)

print(f"[*] Parsed rows   : {len(df_parsed)}")
print(f"[*] Unparsed rows : {len(df_unparsed)}")

# ── Export to Excel ────────────────────────────────────────────────────────────
output_file = f"OSINT_{org_name}.xlsx"

def auto_fit(ws):
    for col_cells in ws.columns:
        max_len = max(
            (len(str(c.value)) for c in col_cells if c.value is not None),
            default=10,
        )
        ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 4, 80)

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    df_parsed.to_excel(writer,   index=False, sheet_name="Parsed")
    df_unparsed.to_excel(writer, index=False, sheet_name="Unparsed_Raw")
    auto_fit(writer.sheets["Parsed"])
    auto_fit(writer.sheets["Unparsed_Raw"])

print(f"\n[OK] Saved → {output_file}")
print(f"     Sheet 'Parsed'       : {len(df_parsed)} records")
print(f"     Sheet 'Unparsed_Raw' : {len(df_unparsed)} records")
```
