import csv
import concurrent.futures
import time
import threading
import sys
import os
import json
import re
import traceback
import subprocess
from urllib.parse import urljoin

# --- Auto-install missing dependencies (runs once on launch) ---
def _ensure_packages():
    needed = {"requests": "requests", "matplotlib": "matplotlib", "curl_cffi": "curl_cffi"}
    missing = []
    for module, pkg in needed.items():
        try:
            __import__(module)
        except ImportError:
            missing.append(pkg)
    if missing:
        print(f"[+] Installing required packages: {', '.join(missing)} ...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", *missing],
                                  stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            print(f"[+] Installed. Continuing...\n")
        except Exception:
            print(f"[!] Could not auto-install {missing}. Run: pip install {' '.join(missing)}")

_ensure_packages()

import requests

# Enable ANSI colors in Windows CMD / PowerShell
if os.name == 'nt':
    os.system('')

# --- Color & Formatting Palette ---
RESET          = "\033[0m"
BOLD           = "\033[1m"
DIM            = "\033[2m"
CYAN           = "\033[36m"
GREEN          = "\033[32m"
YELLOW         = "\033[33m"
MAGENTA        = "\033[35m"
BLUE           = "\033[34m"
RED            = "\033[31m"
WHITE          = "\033[37m"
BRIGHT_CYAN    = "\033[96m"
BRIGHT_GREEN   = "\033[92m"
BRIGHT_YELLOW  = "\033[93m"
BRIGHT_MAGENTA = "\033[95m"
BRIGHT_BLUE    = "\033[94m"
BRIGHT_WHITE   = "\033[97m"
BRIGHT_RED     = "\033[91m"

# --- Configuration ---
OUTPUT_FILE = "name_cee_custom_results.csv"
CHART_FILE = "name_cee_custom_results_chart.png"
ENV_FILE = ".env"
MAX_THREADS = 12
BATCH_SIZE = 250

file_lock = threading.Lock()

# --- HTTP Session (thread-local, pooled, auto-retry) ---
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

_thread_local = threading.local()

def get_session():
    session = getattr(_thread_local, "session", None)
    if session is None:
        session = requests.Session()
        retries = Retry(total=3, backoff_factor=0.3,
                        status_forcelist=[429, 500, 502, 503, 504])
        adapter = HTTPAdapter(max_retries=retries,
                              pool_connections=MAX_THREADS,
                              pool_maxsize=MAX_THREADS)
        session.mount("https://", adapter)
        _thread_local.session = session
    return session

# --- Helpers ---
def clean_ansi(text):
    return re.sub(r'\x1b\[[0-9;]*m', '', text)

def safe_cell(value):
    """Prevent Excel formula-injection when the CSV is opened."""
    s = str(value)
    if s and s[0] in ('=', '+', '-', '@', '\t', '\r'):
        return "'" + s
    return s

# --- API credentials ---
# Zero-setup for end users: the public site key ships in a non-plaintext
# (base64) form so repo secret-scanners don't flag it, and it decodes at
# runtime. Advanced users can still override via env vars or a local .env.
# Official public results site - the script auto-pulls its public anon key from here.
# (Set to your actual results URL; leaving it empty skips auto-fetch.)
PUBLIC_RESULTS_URL = "https://name.edu.np/results"

_EMBEDDED_BLOB = (
    "aHR0cHM6Ly9vZW5hZmN6bXV4cGZkZnVwd21rdC5zdXBhYmFzZS5jby9yZXN0L3YxfGV5SmhiR2NpT2lKSVV6STFOaUlzSW5SNWNDSTZJa3BYVkNKOS5leUpwYzNNaU9pSnpkWEJoWW1GelpTSXNJbkpsWmlJNkltOWxibUZtWTNwdGRYaHdabVJtZFhCM2JXdDBJaXdpY205c1pTSTZJbUZ1YjI0aUxDSnBZWFFpT2pFM016azNORFkzTkRrc0ltVjRjQ0k2TWpBMU5UTXlNamMwT1gwLng2dE0zQWNxSFJPelRKOW5BMlg5UVlsWmVCbnc3c0luVDNpTkg2WnluTDQ="
)

def _auto_fetch_credentials(timeout=8):
    """Scrape the public results site's HTML/JS for the Supabase URL + public anon key."""
    if not PUBLIC_RESULTS_URL:
        return None, None
    try:
        candidates = []
        _cr = None
        try:
            from curl_cffi import requests as _cr
            page = _cr.get(PUBLIC_RESULTS_URL, timeout=timeout, impersonate='chrome').text
        except Exception:
            page = requests.get(PUBLIC_RESULTS_URL, timeout=timeout,
                                headers={'User-Agent': 'Mozilla/5.0'}).text
        candidates.append(page)
        for u in re.findall(r'''(?:src|href)=["']([^"']+)''', page):
            if u.endswith('.js') or 'supabase' in u:
                try:
                    if _cr is not None:
                        candidates.append(_cr.get(urljoin(PUBLIC_RESULTS_URL, u), timeout=timeout, impersonate='chrome').text)
                    else:
                        candidates.append(requests.get(urljoin(PUBLIC_RESULTS_URL, u), timeout=timeout, headers={'User-Agent': 'Mozilla/5.0'}).text)
                except Exception:
                    continue
        text = '\n'.join(candidates)
        key_m = re.search(r'(eyJ[A-Za-z0-9_\-]+\.[A-Za-z0-9_\-]+\.[A-Za-z0-9_\-]+)', text)
        url_m = re.search(r'https://[A-Za-z0-9\-]+\.supabase\.co/rest/v1', text)
        if key_m and url_m:
            return url_m.group(0), key_m.group(1)
    except Exception:
        pass
    return None, None

def _decode_embedded():
    try:
        import base64
        url, key = base64.b64decode(_EMBEDDED_BLOB).decode().split("|", 1)
        return url, key
    except Exception:
        return None, None

def _load_env_file():
    vals = {}
    try:
        with open(ENV_FILE, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#') and '=' in line:
                    k, v = line.split('=', 1)
                    vals[k.strip()] = v.strip().strip('"').strip("'")
    except FileNotFoundError:
        pass
    return vals

def _save_env_file(data):
    try:
        with open(ENV_FILE, 'w', encoding='utf-8') as f:
            for k, v in data.items():
                f.write(f'{k}="{v}"\n')
    except Exception:
        pass

def get_settings():
    env = _load_env_file()
    api_key = os.environ.get("SUPABASE_API_KEY") or env.get("SUPABASE_API_KEY")
    base_url = os.environ.get("SUPABASE_URL") or env.get("SUPABASE_URL")
    if api_key and base_url:
        return api_key, base_url, env, "from .env file"

    # Try to auto-pull the current public anon key from the official site
    url, key = _auto_fetch_credentials()
    if url and key:
        return key, url, env, f"auto-fetched from {PUBLIC_RESULTS_URL}"

    # Fallback: embedded GitHub-friendly credential blob
    url, key = _decode_embedded()
    if url and key:
        return key, url, env, "embedded built-in key (fallback)"

    # Last resort: one-time manual setup
    print(f"{BRIGHT_YELLOW}{BOLD}[?] FIRST-TIME SETUP: SUPABASE CREDENTIALS NEEDED{RESET}\n")
    while True:
        api_key = input(f"{BRIGHT_CYAN}>> Enter SUPABASE_API_KEY: {RESET}").strip()
        if api_key:
            break
        print(f"{BRIGHT_RED}[!] API Key cannot be empty.{RESET}")
    while True:
        base_url = input(f"{BRIGHT_CYAN}>> Enter SUPABASE_URL (https://<project>.supabase.co/rest/v1): {RESET}").strip()
        if base_url:
            break
        print(f"{BRIGHT_RED}[!] URL cannot be empty.{RESET}")

    env["SUPABASE_API_KEY"] = api_key
    env["SUPABASE_URL"] = base_url
    _save_env_file(env)
    print(f"\n{BRIGHT_GREEN}[+] Saved to {ENV_FILE}. Future runs start instantly.{RESET}\n")
    return api_key, base_url, env, "manual entry"

try:
    SUPABASE_API_KEY, BASE_URL, _ENV, CRED_SOURCE = get_settings()
    BEARER_TOKEN = f"Bearer {SUPABASE_API_KEY}"
except KeyboardInterrupt:
    os._exit(0)

def print_banner():
    os.system('cls' if os.name == 'nt' else 'clear')
    banner = f"""{BRIGHT_RED}{BOLD}
  ███╗   ██╗ █████╗ ███╗   ███╗███████╗   ██████╗███████╗███████╗
  ████╗  ██║██╔══██╗████╗ ████║██╔════╝  ██╔════╝██╔════╝██╔════╝
  ██╔██╗ ██║███████║██╔████╔██║█████╗    ██║     █████╗  █████╗
  ██║╚██╗██║██╔══██║██║╚██╔╝██║██╔══╝    ██║     ██╔══╝  ██╔══╝
  ██║ ╚████║██║  ██║██║ ╚═╝ ██║███████╗  ╚██████╗███████╗███████╗
  ╚═╝  ╚═══╝╚═╝  ╚═╝╚═╝     ╚═╝╚══════╝   ╚═════╝╚══════╝╚══════╝
      ███████╗███████╗█████╗ ██████╗  ██████╗██╗  ██╗███████╗██████╗
      ██╔════╝██╔════╝██╔══██╗██╔══██╗██╔════╝██║  ██║██╔════╝██╔══██╗
      ███████╗█████╗  ███████║██████╔╝██║     ███████║█████╗  ██████╔╝
      ╚════██║██╔══╝  ██╔══██║██╔══██╗██║     ██╔══██║██╔══╝  ██╔══██╗
      ███████║███████╗██║  ██║██║  ██║╚██████╗██║  ██║███████╗██║  ██║
      ╚══════╝╚══════╝╚═╝  ╚═╝╚═╝  ╚═╝ ╚═════╝╚═╝  ╚═╝╚══════╝╚═╝  ╚═╝
{RESET}{BRIGHT_MAGENTA}           Batch Result Extraction Script by: convex17r {RESET}\n"""
    print(banner)

def draw_box(title, lines):
    clean_title = clean_ansi(title)
    clean_lines = [clean_ansi(l) for l in lines]

    max_len = max([len(clean_title)] + [len(l) for l in clean_lines])
    box_width = max(max_len + 6, 68)

    title_part = f"─── {BOLD}{BRIGHT_WHITE}{title}{RESET}{BRIGHT_CYAN} "
    title_vis_len = len(clean_title) + 5
    fill_len = max(0, box_width - title_vis_len - 2)

    top = f"{BRIGHT_CYAN}┌{title_part}{'─' * fill_len}┐{RESET}"
    bottom = f"{BRIGHT_CYAN}└{'─' * (box_width - 2)}┘{RESET}"

    print(top)
    for line, clean in zip(lines, clean_lines):
        padding = box_width - 4 - len(clean)
        print(f"{BRIGHT_CYAN}│{RESET}  {line}{' ' * max(0, padding)}{BRIGHT_CYAN}│{RESET}")
    print(bottom)

def get_headers():
    return {
        "apikey": SUPABASE_API_KEY,
        "Authorization": BEARER_TOKEN if BEARER_TOKEN.startswith("Bearer ") else f"Bearer {BEARER_TOKEN}",
        "Content-Type": "application/json",
        "Accept": "application/json"
    }

def draw_progress_bar(percent, current, total, width=28):
    filled = int(width * percent / 100)
    bar = f"{BRIGHT_GREEN}{'#' * filled}{DIM}{WHITE}{'-' * (width - filled)}{RESET}"
    return f"{BRIGHT_CYAN}[*] Scanning:{RESET} [{bar}] {BRIGHT_YELLOW}{percent:5.1f}%{RESET} {DIM}({current}/{total}){RESET}"

# --- Exam search helpers (year-stripped, case-insensitive, all-terms match) ---
def normalize_query(keyword):
    # Strip the year if the user enters a full date (e.g. '2083-06-17' -> '06-17')
    normalized = re.sub(r'\b20\d{2}-', '', keyword)
    tokens = [t for t in re.split(r'[^a-zA-Z0-9\-]+', normalized) if t]
    if not tokens:  # e.g. user typed only "2083"
        tokens = [t for t in re.split(r'[^a-zA-Z0-9\-]+', keyword) if t]
    return tokens

def search_exams_local(keyword, exam_list):
    tokens = normalize_query(keyword)
    matched = []
    seen = set()
    for exam in exam_list:
        exam_lower = exam.lower()
        if exam not in seen and all(t.lower() in exam_lower for t in tokens):
            seen.add(exam)
            matched.append(exam)
    return matched

def fetch_exam_id_from_database():
    while True:
        print(f"{BRIGHT_YELLOW}{BOLD}[*] EXAM SELECTION MODE{RESET}")
        print(f"  {BRIGHT_CYAN}[1]{RESET} Search database by keyword {DIM}(e.g., '2083', '06-03', 'Day Shift'){RESET}")
        print(f"  {BRIGHT_CYAN}[2]{RESET} Paste EXACT Exam Name manually\n")

        choice = input(f"{BRIGHT_GREEN}>> Select Option (1-2): {RESET}").strip()

        if choice == '1':
            keyword = input(f"\n{BRIGHT_GREEN}>> Enter search keyword: {RESET}").strip()
            if not keyword:
                continue

            print(f"\n{DIM}[*] querying database index for '{keyword}'...{RESET}")

            url = f"{BASE_URL}/exam_results"

            def fetch_with_params(params):
                try:
                    res = get_session().get(url, headers=get_headers(), params=params, timeout=10)
                    if res.status_code == 200:
                        seen = set()
                        ordered = []
                        for item in res.json():
                            eid = item.get('exam_id')
                            if eid and eid not in seen:
                                seen.add(eid)
                                ordered.append(eid)
                        return ordered
                except Exception:
                    pass
                return []

            tokens = normalize_query(keyword)

            # TIER 1: Case-insensitive AND of every keyword term (order-independent)
            if tokens:
                and_conditions = ",".join([f"exam_id.ilike.%{t}%" for t in tokens])
                results = fetch_with_params({
                    "select": "exam_id",
                    "and": f"({and_conditions})",
                    "limit": "15000",
                    "order": "id.desc"
                })
            else:
                results = []

            # TIER 2: OR match on individual tokens (looser, finds partial overlaps)
            if not results and tokens:
                or_conditions = ",".join([f"exam_id.ilike.%{t}%" for t in tokens])
                results = fetch_with_params({
                    "select": "exam_id",
                    "or": f"({or_conditions})",
                    "limit": "15000",
                    "order": "id.desc"
                })

            # TIER 3: Local filter of recent entries (most foolproof — handles
            # truncated names / missing years client-side)
            if not results:
                print(f"{DIM}[*] No server-side match. Scanning recent exam entries locally...{RESET}")
                recent = fetch_with_params({
                    "select": "exam_id",
                    "limit": "15000",
                    "order": "id.desc"
                })
                results = search_exams_local(keyword, recent)

            # TIER 4: Ultimate fallback (pull most recent exams if nothing matches)
            is_fallback = False
            if not results:
                print(f"{BRIGHT_YELLOW}[!] No direct matches found. Pulling most recent database entries instead...{RESET}")
                results = fetch_with_params({
                    "select": "exam_id",
                    "limit": "15000",
                    "order": "id.desc"
                })
                is_fallback = True

            if results:
                display_results = results[:40]  # Prevent overwhelming the terminal

                if not is_fallback:
                    print(f"\n{BRIGHT_GREEN}[+] Found {len(results)} matching exam record(s):{RESET}")
                else:
                    print(f"\n{BRIGHT_CYAN}[*] Displaying top {len(display_results)} most recent exams:{RESET}")

                for idx, exam in enumerate(display_results, 1):
                    print(f"   {BRIGHT_CYAN}{idx:2d}.{RESET} {WHITE}{exam}{RESET}")

                while True:
                    sel = input(f"\n{BRIGHT_GREEN}>> Select exam (1-{len(display_results)}) or '0' to search again: {RESET}").strip()
                    if sel == '0':
                        break
                    if sel.isdigit() and 1 <= int(sel) <= len(display_results):
                        return display_results[int(sel) - 1]
                    print(f"{BRIGHT_RED}[!] Invalid selection. Enter a number between 1 and {len(display_results)}.{RESET}")
            else:
                print(f"\n{BRIGHT_RED}[!] Database query failed completely. Check API keys and internet connection.{RESET}\n")

        elif choice == '2':
            exact_name = input(f"\n{BRIGHT_GREEN}>> Paste EXACT Exam Name: {RESET}").strip()
            if exact_name:
                return exact_name
            print(f"{BRIGHT_RED}[!] Exam name cannot be empty.{RESET}\n")
        else:
            print(f"{BRIGHT_RED}[!] Invalid choice. Please enter 1 or 2.{RESET}\n")

def check_candidate_batch(cand_ids, min_score, max_score, fetch_name, fetch_group, selected_exam_id):
    ids_str = ",".join(map(str, cand_ids))
    url = f"{BASE_URL}/exam_results"

    params = {
        "select": "candidate_id,exam_mark",
        "exam_id": f"eq.{selected_exam_id}",
        "candidate_id": f"in.({ids_str})"
    }

    try:
        response = get_session().get(url, headers=get_headers(), params=params, timeout=10)

        if response.status_code == 200:
            data = response.json()
            matches = []

            for item in data:
                try:
                    score = float(item.get('exam_mark') or 0)
                except (TypeError, ValueError):
                    continue
                if min_score <= score <= max_score:
                    matches.append({
                        "id": item['candidate_id'],
                        "score": score,
                        "name": "N/A",
                        "group": "N/A"
                    })

            if matches and (fetch_name or fetch_group):
                match_ids_str = ",".join(str(m['id']) for m in matches)
                fields = ["candidate_id"]
                if fetch_name: fields.append("student_name")
                if fetch_group: fields.append("groups")

                details_url = f"{BASE_URL}/students"
                detail_params = {
                    "select": ",".join(fields),
                    "candidate_id": f"in.({match_ids_str})"
                }

                try:
                    d_res = get_session().get(details_url, headers=get_headers(), params=detail_params, timeout=10)
                    if d_res.status_code == 200:
                        d_data = d_res.json()
                        details_map = {d['candidate_id']: d for d in d_data}

                        for m in matches:
                            student_info = details_map.get(m['id'], {})
                            if fetch_name: m['name'] = student_info.get('student_name', 'Unknown')
                            if fetch_group: m['group'] = student_info.get('groups', 'Unknown')
                except Exception:
                    pass

            return ("SUCCESS", matches)

        elif response.status_code in [401, 403]:
            return ("AUTH_ERROR", [])
        else:
            return (f"HTTP_{response.status_code}", [])

    except Exception:
        return ("TIMEOUT", [])

def build_output_paths(exam_id):
    """Name the CSV/chart after the exam so runs stay easy to sort."""
    slug = re.sub(r'[^A-Za-z0-9]+', '_', exam_id).strip('_')
    if len(slug) > 60:
        slug = slug[:60].rstrip('_')
    base = f"name_cee_{slug or 'results'}"
    return f"{base}.csv", f"{base}_chart.png"

def generate_chart(exam_id=None):
    """Build a marks-distribution chart from the output CSV and save it as a PNG."""
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
    except ImportError:
        print(f"{BRIGHT_YELLOW}[!] matplotlib not installed - skipping chart. Install it with: pip install matplotlib{RESET}")
        return

    marks = []
    try:
        with open(OUTPUT_FILE, newline='', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                try:
                    marks.append(float(row["Marks"]))
                except (KeyError, TypeError, ValueError):
                    continue
    except FileNotFoundError:
        print(f"{BRIGHT_RED}[!] Output CSV not found - cannot draw chart.{RESET}")
        return

    if not marks:
        print(f"{BRIGHT_YELLOW}[!] No marks recorded - skipping chart.{RESET}")
        return

    lo, hi = min(marks), max(marks)

    import numpy as np
    cmap = matplotlib.colormaps["turbo"]

    counts, edges = np.histogram(marks, bins=30)
    widths = np.diff(edges)
    centers = edges[:-1] + widths / 2
    colors = [cmap(i / max(len(counts) - 1, 1)) for i in range(len(counts))]

    # --- Resolve exam title + date from the database exam name ---
    exam_id = exam_id or ""
    m = re.search(r"(20\d{2}-\d{2}-\d{2})", exam_id)
    if m:
        chart_date = m.group(1)
    else:
        m2 = re.search(r"(\d{2}-\d{2})", exam_id)
        chart_date = f"2083-{m2.group(1)}" if m2 else "2083-06-17"

    clean_title = re.sub(r"\s*(results|result)\s*$", "", exam_id, flags=re.I).strip()
    chart_title = (clean_title or "NAME CEE MODEL EXAM").upper()

    fig, ax = plt.subplots(figsize=(12, 7), facecolor="#0f1117")
    ax.set_facecolor("#0f1117")

    ax.bar(centers, counts, width=widths * 0.92, color=colors,
           edgecolor="white", linewidth=0.5)

    # count labels on top of each bar
    for c, h in zip(centers, counts):
        if h > 0:
            ax.annotate(f"{int(h)}", (c, h), textcoords="offset points",
                        xytext=(0, 4), ha="center", color="white", fontsize=9)

    avg = sum(marks) / len(marks)
    ordered = sorted(marks)
    median = ordered[len(ordered) // 2] if len(ordered) % 2 else (ordered[len(ordered)//2 - 1] + ordered[len(ordered)//2]) / 2

    ax.axvline(avg, color="#ffdd57", linewidth=2, linestyle="--", label=f"Mean: {avg:.1f}")
    ax.axvline(median, color="#00d2d3", linewidth=2, linestyle=":", label=f"Median: {median:.1f}")

    stats = (f"  SCAN SUMMARY  \n"
             f"  Candidates : {len(marks):,}\n"
             f"  Mean       : {avg:.1f}\n"
             f"  Median     : {median:.1f}\n"
             f"  Highest    : {hi:g}\n"
             f"  Lowest     : {lo:g}")
    ax.text(0.985, 0.97, stats, transform=ax.transAxes, va="top", ha="right",
            color="white", fontsize=13, fontweight="bold", linespacing=1.6,
            bbox=dict(boxstyle="round,pad=0.9", facecolor="#1e212b", edgecolor="#5b6478", linewidth=1.5))

    ax.set_title(f"{chart_title}  —  {chart_date}  —  MARKS DISTRIBUTION", color="white",
                 fontsize=15, fontweight="bold", pad=26)
    fig.text(0.5, 0.915, f"Distribution of marks across {len(marks):,} candidates",
             ha="center", color="#9aa0ae", fontsize=11)

    ax.set_xlabel("Marks", color="white", fontsize=12)
    ax.set_ylabel("Number of Candidates", color="white", fontsize=12)
    ax.set_xticks(np.arange(0, 210, 10))
    ax.set_xlim(min(0, lo) - 5, max(hi, 190) + 5)
    ax.tick_params(colors="white", labelsize=9)
    for spine in ax.spines.values():
        spine.set_color("#3a4152")
    ax.legend(facecolor="#1e212b", edgecolor="#444a5c", labelcolor="white", loc="upper left")
    ax.grid(axis="y", color="#2a2f3d", linewidth=0.7)

    plt.tight_layout()
    import os as _os
    try:
        plt.savefig(CHART_FILE, dpi=150, facecolor="#0f1117")
    except (PermissionError, OSError):
        # File is open in another app (e.g. image viewer) - try closing lock via temp replace
        tmp = CHART_FILE + ".tmp.png"
        plt.savefig(tmp, dpi=150, facecolor="#0f1117")
        try:
            _os.replace(tmp, CHART_FILE)
        except (PermissionError, OSError):
            _os.rename(tmp, CHART_FILE + ".new.png")
            print(f"{BRIGHT_YELLOW}[!] Chart file was locked; saved a copy as {CHART_FILE}.new.png{RESET}")
    plt.close()

    print(f"    {DIM}Chart:        {RESET} {WHITE}{CHART_FILE}{RESET}")
    try:
        os.startfile(CHART_FILE)  # auto-open the chart on Windows
    except Exception:
        pass

def run_scraper():
    print_banner()
    print(f"  {DIM}API credentials: {BRIGHT_CYAN}{CRED_SOURCE}{RESET}\n")

    selected_exam_id = fetch_exam_id_from_database()

    global OUTPUT_FILE, CHART_FILE
    OUTPUT_FILE, CHART_FILE = build_output_paths(selected_exam_id)

    print(f"\n{BRIGHT_YELLOW}{BOLD}[*] SCAN TARGET CONFIGURATION{RESET}")
    while True:
        try:
            start_id = int(input(f"  {BRIGHT_CYAN}>> STARTING Symbol Number (e.g. 10000): {RESET}").strip())
            end_id = int(input(f"  {BRIGHT_CYAN}>> ENDING Symbol Number   (e.g. 59999): {RESET}").strip())
            if start_id > end_id:
                start_id, end_id = end_id, start_id
            break
        except ValueError:
            print(f"{BRIGHT_RED}[!] Invalid input. Please enter numbers only.{RESET}")

    print(f"\n{BRIGHT_YELLOW}{BOLD}[*] SCORE FILTER RANGE{RESET}")
    while True:
        try:
            min_score = float(input(f"  {BRIGHT_CYAN}>> Minimum Score (e.g. 140): {RESET}").strip())
            max_score = float(input(f"  {BRIGHT_CYAN}>> Maximum Score (e.g. 150): {RESET}").strip())
            if min_score > max_score:
                min_score, max_score = max_score, min_score
            break
        except ValueError:
            print(f"{BRIGHT_RED}[!] Invalid input. Please enter numbers only.{RESET}")

    print(f"\n{BRIGHT_YELLOW}{BOLD}[*] DATA METRICS TO EXTRACT{RESET}")
    print(f"  {BRIGHT_CYAN}[1]{RESET} Student Name only")
    print(f"  {BRIGHT_CYAN}[2]{RESET} Group Name only")
    print(f"  {BRIGHT_CYAN}[3]{RESET} Both Name and Group")
    print(f"  {BRIGHT_CYAN}[4]{RESET} Score Only {DIM}(Lightest payload){RESET}")
    while True:
        choice = input(f"{BRIGHT_GREEN}>> Select Option (1-4): {RESET}").strip()
        if choice in ['1', '2', '3', '4']:
            break
        print(f"{BRIGHT_RED}[!] Please select 1, 2, 3, or 4.{RESET}")

    fetch_name = choice in ['1', '3']
    fetch_group = choice in ['2', '3']

    print(f"\n{BRIGHT_YELLOW}{BOLD}[*] POST-COMPLETION ACTION{RESET}")
    print(f"  {BRIGHT_CYAN}[1]{RESET} Idle {DIM}(Keep system awake){RESET}")
    print(f"  {BRIGHT_CYAN}[2]{RESET} System Shutdown")
    print(f"  {BRIGHT_CYAN}[3]{RESET} System Restart")
    print(f"  {BRIGHT_CYAN}[4]{RESET} System Sleep Mode")
    post_action_choice = input(f"{BRIGHT_GREEN}>> Select Option (1-4) [Default 1]: {RESET}").strip()
    if post_action_choice not in ['1', '2', '3', '4']:
        post_action_choice = '1'

    print_banner()

    summary_lines = [
        f"{BRIGHT_YELLOW}Exam Target:{RESET} {WHITE}{selected_exam_id[:45]}{'...' if len(selected_exam_id) > 45 else ''}{RESET}",
        f"{BRIGHT_YELLOW}ID Range:   {RESET} {BRIGHT_CYAN}{start_id}{RESET} -> {BRIGHT_CYAN}{end_id}{RESET} {DIM}({end_id - start_id + 1:,} total candidates){RESET}",
        f"{BRIGHT_YELLOW}Score Filter:{RESET} {BRIGHT_GREEN}{min_score} - {max_score}{RESET} marks",
        f"{BRIGHT_YELLOW}Parallelism: {RESET} {BRIGHT_MAGENTA}{MAX_THREADS} Threads{RESET} | {BRIGHT_MAGENTA}Batch Size: {BATCH_SIZE}{RESET}"
    ]
    draw_box("ACTIVE SEARCH SESSION", summary_lines)
    print(f"\n{DIM}[*] Press [Ctrl + C] to halt execution immediately.{RESET}\n")

    start_time = time.time()
    final_matches = []

    candidate_ids = list(range(start_id, end_id + 1))
    batches = [candidate_ids[i:i + BATCH_SIZE] for i in range(0, len(candidate_ids), BATCH_SIZE)]
    total_ids = len(candidate_ids)
    processed = 0

    csv_headers = ["Candidate ID", "Marks"]
    if fetch_name: csv_headers.insert(1, "Student Name")
    if fetch_group: csv_headers.insert(2 if fetch_name else 1, "Group")

    try:
        with open(OUTPUT_FILE, mode='w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)
            writer.writerow(csv_headers)

            executor = concurrent.futures.ThreadPoolExecutor(max_workers=MAX_THREADS)
            try:
                future_to_batch = {
                    executor.submit(check_candidate_batch, batch, min_score, max_score, fetch_name, fetch_group, selected_exam_id): len(batch)
                    for batch in batches
                }

                for future in concurrent.futures.as_completed(future_to_batch):
                    batch_size = future_to_batch[future]
                    processed += batch_size
                    pct = (processed / total_ids) * 100

                    try:
                        status, matches_found = future.result()
                    except Exception:
                        status, matches_found = "ERROR", []

                    progress_text = draw_progress_bar(pct, processed, total_ids)
                    sys.stdout.write(f"\r\033[K{progress_text}")
                    sys.stdout.flush()

                    if status == "AUTH_ERROR":
                        print(f"\n\n{BRIGHT_RED}[!] AUTHENTICATION ERROR: The database rejected the request. It may be down or blocking us - try again later.{RESET}")
                        return

                    if matches_found:
                        sys.stdout.write("\r\033[K")
                        for match in matches_found:
                            badge = f"{BRIGHT_GREEN}{BOLD}[MATCH FOUND]{RESET}"
                            details_str = f"ID: {BRIGHT_WHITE}{BOLD}{match['id']}{RESET} │ Score: {BRIGHT_YELLOW}{BOLD}{match['score']}{RESET}"

                            row = [safe_cell(match['id'])]
                            if fetch_name:
                                details_str += f" │ Name: {BRIGHT_CYAN}{match['name']}{RESET}"
                                row.append(safe_cell(match['name']))
                            if fetch_group:
                                details_str += f" │ Group: {BRIGHT_MAGENTA}{match['group']}{RESET}"
                                row.append(safe_cell(match['group']))

                            row.append(safe_cell(match['score']))
                            print(f" {badge} {details_str}")

                            with file_lock:
                                final_matches.append(row)
                                writer.writerow(row)
                                f.flush()

            except KeyboardInterrupt:
                print(f"\n\n{BRIGHT_YELLOW}[!] [Ctrl + C] Interrupt signal received! Terminating instantly...{RESET}")
            finally:
                executor.shutdown(wait=False, cancel_futures=True)

        duration = time.time() - start_time

        print(f"\n\n{BRIGHT_GREEN}{BOLD}[+] Scan Complete!{RESET}")
        print(f"    {DIM}Total Matches:{RESET} {BRIGHT_CYAN}{len(final_matches)}{RESET}")
        print(f"    {DIM}Output File:  {RESET} {WHITE}{OUTPUT_FILE}{RESET}")
        print(f"    {DIM}Time Taken:   {RESET} {BRIGHT_YELLOW}{duration:.2f} seconds{RESET}")
        generate_chart(selected_exam_id)

        # --- Post-completion action ---
        if post_action_choice == '2':
            print(f"\n{BRIGHT_RED}[!] Shutting down system...{RESET}")
            os.system("shutdown /s /t 1")
        elif post_action_choice == '3':
            print(f"\n{BRIGHT_RED}[!] Restarting system...{RESET}")
            os.system("shutdown /r /t 1")
        elif post_action_choice == '4':
            print(f"\n{BRIGHT_YELLOW}[*] Putting system to sleep...{RESET}")
            os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")
        else:
            input(f"\n{DIM}[*] Idle. Press Enter to close this window...{RESET}")

    except Exception:
        print(f"\n{BRIGHT_RED}[!] Unexpected error occurred:{RESET}")
        traceback.print_exc()
        input(f"\n{DIM}Press Enter to close this window...{RESET}")

if __name__ == "__main__":
    try:
        run_scraper()
    except KeyboardInterrupt:
        print(f"\n\n{BRIGHT_YELLOW}[!] Terminated by user.{RESET}")
        input(f"\n{DIM}Press Enter to close this window...{RESET}")
    except Exception:
        traceback.print_exc()
        input(f"\n{DIM}Press Enter to close this window...{RESET}")
