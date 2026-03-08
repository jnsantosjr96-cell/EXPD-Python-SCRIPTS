# -*- coding: utf-8 -*-
"""
HCPIF extractor R2.6

- Base: R2.4 (robust, Unicode-safe, anti-signature & anti-date, smart transliteration)
- R2.5: Integração com WEBADI preservando macros/estrutura; mantém "Preferred Language" e "BATCH_NAME".
- R2.6:
  * Desprotege workbook e todas as abas antes de escrever (opcional via CLI).
  * Detecção de cabeçalho robusta (normalização, sinônimos, tolerante a "HOTEL ID", NBSP, quebras de linha).
  * Busca de cabeçalho a partir da linha 5 (parametrizável).
  * CLI para nome da aba, linha inicial do cabeçalho e debug.

Ajustes solicitados:
- Quando --webadi_template for informado e --webadi_output não for, a saída será
  um novo arquivo na mesma pasta do template, com nome:
    OC WEBADI MMDDYYYY.xlsm  (ex.: OC WEBADI 03022026.xlsm).
- Somente a coluna BILL_TO deve vir com "BILL_TO" (localizada pelo cabeçalho; fallback na coluna Q);
  se existir cabeçalho "SITE_PURPOSE", também recebe "BILL_TO".
- Demais colunas são preenchidas a partir dos PDFs + mapeamento pelas colunas do WEBADI.
- Linha 3 da aba WebADI:
    B3 = BATCH_NAME        (mantido)
    D3 = * Text            (mantido)
    E3 = OC MMDDYYYY       (atualizado; remove "MS BRUNER" e cia)
"""

# -------------------- DEFAULT PATHS --------------------
DEFAULT_INPUT_DIR = r"C:\Users\josenjr\Downloads\HCPIFs"
DEFAULT_OUTPUT_FILE = r"C:\Users\josenjr\Downloads\HCPIFs\Output\HCPIF_extraction.xlsx"

# Template padrão (arquivo com macros boas)
DEFAULT_WEBADI_TEMPLATE = r"C:\Users\josenjr\Downloads\HCPIFs\OC WebADI SF 149540900 2-13-26 (002).xlsm"

# -------------------- ENSURE / INSTALL DEPS --------------------
import sys
import subprocess
import importlib
import argparse
import copy
import re
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Optional, Tuple

REQUIRED = [
    ("pdfplumber", "pdfplumber"),
    ("pandas", "pandas"),
    ("openpyxl", "openpyxl"),
    ("tqdm", "tqdm"),
    ("Unidecode", "unidecode"),
    ("pycountry", "pycountry"),
]

def ensure_package(pip_name: str, import_name: str):
    try:
        return importlib.import_module(import_name)
    except ImportError:
        print(f" Installing {pip_name} ...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", pip_name])
        return importlib.import_module(import_name)

pdfplumber = ensure_package("pdfplumber", "pdfplumber")
pd = ensure_package("pandas", "pandas")
openpyxl = ensure_package("openpyxl", "openpyxl")
tqdm = ensure_package("tqdm", "tqdm")
unidecode_mod = ensure_package("Unidecode", "unidecode")
from unidecode import unidecode  # noqa: E402

try:
    pycountry = ensure_package("pycountry", "pycountry")
except Exception:
    pycountry = None

# --- Hotfix: registrar .JPG/.JPEG no mimetypes interno do openpyxl ---
from openpyxl.packaging import manifest as _ox_manifest  # noqa: E402

for _ext in (".jpg", ".jpeg", ".JPG", ".JPEG"):
    try:
        _ox_manifest.mimetypes.add_type("image/jpeg", _ext)
    except Exception:
        pass
# --- fim do hotfix ---

from openpyxl import load_workbook  # noqa: E402
from openpyxl.utils import column_index_from_string  # noqa: E402

# -------------------- FIELD DEFINITIONS --------------------
FIELDS: List[Dict] = [
    {"col": "Today's date", "labels": [r"today[\']s\s*date"]},
    {"col": "Effective Date of Change", "labels": [r"effective\s*date\s*of\s*change"]},
    {"col": "Expedia ID", "labels": [r"expedia\s*id"]},
    {"col": "Country", "labels": [r"country"]},
    {"col": "Currency", "labels": [r"currency", r"currency\s*code", r"curr\.?"]},
    {"col": "Legal Name", "labels": [r"legal\s*name"]},
    {"col": "Hotel Name", "labels": [r"hotel\s*name"]},
    {"col": "Address Line 1", "labels": [r"address\s*line\s*1"]},
    {"col": "City", "labels": [r"city"]},
    {
        "col": "State/Province",
        "labels": [
            r"state\s*/\s*province",
            r"state\s*province",
            r"state",
            r"province",
        ],
    },
    {"col": "Postal Code", "labels": [r"postal\s*code", r"zip\s*code"]},
    {"col": "Tax Registration Number", "labels": [r"tax\s*registration\s*number"]},
    {"col": "Tax Registration Status", "labels": [r"tax\s*registration\s*status"]},
    {"col": "First Name", "labels": [r"first\s*name"]},
    {"col": "Last Name", "labels": [r"last\s*name"]},
    {"col": "Email Address", "labels": [r"email\s*address", r"email"]},
    {"col": "Preferred Language", "labels": [r"preferred\s*language", r"language\s*preference"]},
]

SAME_LINE_PATTERN_TMPL = r"(?im)^\s*{label}\s*:?\s*(?P<val>[^\n\r]+?)\s*$"
LABEL_ONLY_PATTERN_TMPL = r"(?im)^\s*{label}\s*:?\s*$"
EMAIL_FALLBACK = re.compile(r"(?i)[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}")

DATE_PATTERNS = [
    "%d/%b/%Y",
    "%d/%B/%Y",
    "%d/%m/%Y",
    "%Y-%m-%d",
    "%d-%b-%Y",
    "%d-%m-%Y",
]
DATE_TOKEN_REGEX = re.compile(
    r"\b(\d{1,2}/[A-Za-z]{3}/\d{4}|\d{1,2}/\d{1,2}/\d{4}|"
    r"\d{4}-\d{2}-\d{2}|\d{1,2}-[A-Za-z]{3}-\d{4}|\d{1,2}-\d{1,2}-\d{4})\b"
)
MONTH3 = r"(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)"

UNICODE_LETTERS = r"A-Za-z-\u0370-\u03FF\u1F00-\u1FFF\u3040-\u30FF\u4E00-\u9FFF"
SIGNATURE_NOISE = re.compile(r"(?i)^\s*(signature|signed|signee|signer)\b")
HEADER_NOISE = re.compile(
    r"(?i)\b(city|state|province|billing|contact|information|signature|signed|signee|signer)\b"
)
POSTAL_LIKE = re.compile(r"(?i)^([A-Z]?\d[\dA-Z\- ]{2,}|[A-Z]\d[A-Z]\s*\d[A-Z]\d)$")

ALL_LABELS = [lab for f in FIELDS for lab in f["labels"]]
BARRIER_LABELS = ALL_LABELS + [r"signature", r"signed", r"signee", r"signer"]
NEXT_LABEL_BARRIER = re.compile(r"(?i)\b(?:" + "|".join(BARRIER_LABELS) + r")\b\s*:?")
LABEL_LINE_RE = re.compile(r"(?im)^\s*(?:" + "|".join(BARRIER_LABELS) + r")\s*:?\b")

def is_valid_date_token(token: str) -> bool:
    token = token.strip()
    for fmt in DATE_PATTERNS:
        try:
            datetime.strptime(token, fmt)
            return True
        except ValueError:
            continue
    return False

def contains_date_like(s: str) -> bool:
    if not s:
        return False
    if DATE_TOKEN_REGEX.search(s):
        return True
    if re.search(rf"\b{MONTH3}\b\s+\d{{1,2}},\s*\d{{4}}", s, re.IGNORECASE):
        return True
    if re.search(r"GMT\s*[+-]\d+", s, re.IGNORECASE):
        return True
    return False

def normalize_block_text(s: str) -> str:
    return (s or "").replace("\u00A0", " ").replace("\t", " ")

def cut_at_next_label(s: str) -> str:
    if not s:
        return s
    m = NEXT_LABEL_BARRIER.search(s)
    return s[: m.start()].strip() if m else s.strip()

def strip_signature_prefix(s: Optional[str]) -> Optional[str]:
    if not s:
        return s
    return (
        re.sub(r"(?i)^\s*(signature|signed|signee|signer)\s*:\s*", "", s).strip() or None
    )

def strip_parenthesized_dates(s: Optional[str]) -> Optional[str]:
    if not s:
        return s
    s = re.sub(
        rf"\(\s*{MONTH3}\s+\d{{1,2}},\s*\d{{4}}[^()]*\)", "", s, flags=re.IGNORECASE
    )
    s = re.sub(r"\([^()]*\d[^()]*\)", "", s)
    return s.strip() or None

def clean_extracted_value(s: Optional[str]) -> Optional[str]:
    if s is None:
        return None
    v = normalize_block_text(s)
    v = cut_at_next_label(v)
    v = strip_signature_prefix(v) or v
    v = strip_parenthesized_dates(v) or v
    v = re.sub(r"\s{2,}", " ", v).strip()
    if not re.search(rf"[{UNICODE_LETTERS}0-9]", v):
        return None
    if contains_date_like(v):
        return None
    return v or None

def looks_like_postal(s: Optional[str]) -> bool:
    if not s:
        return False
    return bool(POSTAL_LIKE.match(s.strip()))

def sanitize_country(val: Optional[str], currency: Optional[str]) -> Optional[str]:
    if not val:
        return None
    v = clean_extracted_value(val)
    if not v:
        return None
    v_clean = v.strip()
    v_upper = v_clean.upper()
    if re.fullmatch(r"[A-Za-z]{2,4}", v_clean):
        if currency and v_upper == (currency or "").strip().upper():
            return None
        return None
    if (
        HEADER_NOISE.search(v_clean)
        or re.search(r"\d", v_clean)
        or SIGNATURE_NOISE.match(v_clean)
        or contains_date_like(v_clean)
    ):
        return None
    return v_clean

def sanitize_state_and_postal(
    state_val: Optional[str], postal_val: Optional[str]
) -> Tuple[Optional[str], Optional[str]]:
    s = clean_extracted_value(state_val) if state_val else None
    p = clean_extracted_value(postal_val) if postal_val else None

    if s and looks_like_postal(s):
        if not p:
            p = s
        s = None

    if s and not re.search(rf"[{UNICODE_LETTERS}]", s) and len(s) <= 2:
        s = None

    if s and (contains_date_like(s) or SIGNATURE_NOISE.match(s)):
        s = None

    return s, p

def is_alpha_line(s: str) -> bool:
    s = s.strip()
    if not s or "@" in s:
        return False
    if contains_date_like(s) or SIGNATURE_NOISE.match(s):
        return False
    return re.fullmatch(rf"[{UNICODE_LETTERS} .'\-()]{2,120}", s) is not None

def extract_two_dates_from_lines(full_text: str):
    for ln in full_text.splitlines():
        tokens = DATE_TOKEN_REGEX.findall(ln)
        if len(tokens) >= 2:
            d1, d2 = tokens[0], tokens[1]
            if is_valid_date_token(d1) and is_valid_date_token(d2):
                return d1, d2
    return None, None

def try_same_line_block(text: str, label_regex: str) -> Optional[str]:
    pattern = re.compile(
        SAME_LINE_PATTERN_TMPL.format(label=label_regex),
        flags=re.IGNORECASE | re.MULTILINE,
    )
    m = pattern.search(text)
    if m:
        cand = clean_extracted_value(m.group("val"))
        if cand and not contains_date_like(cand) and not SIGNATURE_NOISE.match(cand):
            return cand
    return None

def try_next_line_block(
    text: str, label_regex: str, all_label_variants: List[str]
) -> Optional[str]:
    label_only = re.compile(
        LABEL_ONLY_PATTERN_TMPL.format(label=label_regex),
        flags=re.IGNORECASE | re.MULTILINE,
    )
    m = label_only.search(text)
    if not m:
        return None
    tail = text[m.end() :]

    start_of_label = re.compile(
        r"(?im)^\s*(?:" + "|".join(all_label_variants) + r")\s*:\b"
    )
    for ln in tail.splitlines():
        cand = clean_extracted_value(ln)
        if not cand:
            continue
        if start_of_label.match(cand):
            return None
        if HEADER_NOISE.search(cand) or SIGNATURE_NOISE.match(cand) or contains_date_like(cand):
            continue
        return cand
    return None

def words_from_page(page) -> List[Dict]:
    return page.extract_words(
        x_tolerance=3,
        y_tolerance=3,
        keep_blank_chars=False,
        use_text_flow=True,
    ) or []

def find_label_anchors(words: List[Dict], label_tokens: List[str]) -> List[Dict]:
    def _norm(s: str) -> str:
        s = (s or "").lower().replace("’", "'").strip()
        return s[:-1] if s.endswith(":") else s

    anchors = []
    n, m = len(words), len(label_tokens)
    norm_labels = [_norm(x) for x in label_tokens]
    for i in range(n - m + 1):
        ok = True
        for j in range(m):
            if _norm(words[i + j]["text"]) != norm_labels[j]:
                ok = False
                break
        if ok:
            seq = words[i : i + m]
            anchors.append(
                {
                    "x0": min(w["x0"] for w in seq),
                    "x1": max(w["x1"] for w in seq),
                    "top": min(w["top"] for w in seq),
                    "bottom": max(w["bottom"] for w in seq),
                }
            )
    return anchors

def collect_tokens_on_same_line_right(
    words: List[Dict], anchor: Dict, y_tol: float = 6.0
):
    a_top, a_bot, a_x1 = anchor["top"], anchor["bottom"], anchor["x1"]
    line_tokens = []
    for w in words:
        if w["x0"] <= a_x1:
            continue
        if abs(w["top"] - a_top) <= y_tol or abs(w["bottom"] - a_bot) <= y_tol:
            line_tokens.append(w)
    line_tokens.sort(key=lambda k: (k["top"], k["x0"]))
    return line_tokens

def collect_tokens_next_line_down(
    words, anchor, x_window=360.0, v_range=(0.5, 80.0)
):
    a_bot, a_x1 = anchor["bottom"], anchor["x1"]
    v_min, v_max = a_bot + v_range[0], a_bot + v_range[1]
    h_min, h_max = a_x1, a_x1 + x_window
    cand = [
        w
        for w in words
        if (v_min <= w["top"] <= v_max and h_min <= w["x0"] <= h_max)
    ]
    cand.sort(key=lambda k: (k["top"], k["x0"]))
    return cand

def join_tokens(tokens: List[Dict]) -> str:
    return " ".join(t["text"] for t in tokens).strip()

CURRENCY3 = re.compile(r"^[A-Za-z]{3}$")

def extract_currency_positional(words: List[Dict], anchors: List[Dict]) -> Optional[str]:
    for a in anchors:
        same = collect_tokens_on_same_line_right(words, a)
        for t in same:
            if CURRENCY3.match(t["text"].strip()):
                return t["text"].upper()
        down = collect_tokens_next_line_down(words, a, x_window=360, v_range=(0.5, 100))
        for t in down:
            if CURRENCY3.match(t["text"].strip()):
                return t["text"].upper()
    return None

def extract_date_positional(words: List[Dict], anchors: List[Dict]) -> Optional[str]:
    for a in anchors:
        same = collect_tokens_on_same_line_right(words, a)
        if same:
            txt = join_tokens(same)
            m = DATE_TOKEN_REGEX.search(txt)
            if m and is_valid_date_token(m.group(1)):
                return m.group(1)
        down = collect_tokens_next_line_down(words, a, x_window=420, v_range=(0.5, 120))
        if down:
            txt = join_tokens(down)
            m = DATE_TOKEN_REGEX.search(txt)
            if m and is_valid_date_token(m.group(1)):
                return m.group(1)
    return None

def extract_country_positional(words: List[Dict], anchors: List[Dict]) -> Optional[str]:
    for a in anchors:
        same = collect_tokens_on_same_line_right(words, a)
        raw = (
            join_tokens(same)
            if same
            else join_tokens(
                collect_tokens_next_line_down(words, a, x_window=480, v_range=(0.5, 120))
            )
        )
        if raw:
            val = clean_extracted_value(raw)
            if (
                val
                and not HEADER_NOISE.search(val)
                and not re.search(r"\d", val)
                and not re.fullmatch(r"[A-Za-z]{2,4}", val)
            ):
                return val
    return None

TRN_BAD = re.compile(r"(?i)^(registered|select a status)$")

def sanitize_trn(val: Optional[str]) -> Optional[str]:
    v = clean_extracted_value(val)
    if not v:
        return None
    if TRN_BAD.match(v) or SIGNATURE_NOISE.match(v) or contains_date_like(v):
        return None
    if not re.search(r"\d", v):
        return None
    return v

def extract_trn_same_line(words: List[Dict], anchors: List[Dict]) -> Optional[str]:
    for a in anchors:
        same = collect_tokens_on_same_line_right(words, a, y_tol=6.0)
        if not same:
            continue
        raw = join_tokens(same)
        val = sanitize_trn(raw)
        if val:
            return val
    return None

def extract_fields_positional(pdf_path: Path) -> Dict[str, Optional[str]]:
    results = {f["col"]: None for f in FIELDS}
    with pdfplumber.open(str(pdf_path)) as pdf:
        full_text_pages = []
        pages_bundle = []
        for idx, page in enumerate(pdf.pages, start=1):
            words = words_from_page(page)
            text = page.extract_text() or ""
            full_text_pages.append(text)
            pages_bundle.append((idx, page, words, text))

        full_text = normalize_block_text("\n".join(full_text_pages))

        for idx, page, words, text in pages_bundle:
            if not words:
                continue

            anchors_currency = find_label_anchors(words, ["currency"])
            anchors_country = find_label_anchors(words, ["country"])
            anchors_edc = find_label_anchors(words, ["effective", "date", "of", "change"])
            anchors_today = find_label_anchors(words, ["today's", "date"])
            anchors_trn = find_label_anchors(words, ["tax", "registration", "number"])

            if not results["Currency"]:
                results["Currency"] = extract_currency_positional(words, anchors_currency)
            if not results["Effective Date of Change"]:
                results["Effective Date of Change"] = extract_date_positional(
                    words, anchors_edc
                )
            if not results["Today's date"]:
                results["Today's date"] = extract_date_positional(words, anchors_today)
            if not results["Country"]:
                results["Country"] = extract_country_positional(words, anchors_country)
            if not results["Tax Registration Number"]:
                results["Tax Registration Number"] = extract_trn_same_line(
                    words, anchors_trn
                )

        d1, d2 = extract_two_dates_from_lines(full_text)
        if d1 and d2:
            if not results["Today's date"]:
                results["Today's date"] = d1
            if not results["Effective Date of Change"]:
                results["Effective Date of Change"] = d2

        if not results["Currency"]:
            m = re.search(r"(?is)\bcurrency\s*:\s*([A-Za-z]{3})\b", full_text)
            if m:
                results["Currency"] = m.group(1).upper()

        if not results["Country"]:
            m_same = re.search(r"(?im)^\s*country\s*:\s*(.+)$", full_text)
            if m_same:
                val = clean_extracted_value(m_same.group(1))
                if val:
                    results["Country"] = val

        for f in FIELDS:
            col = f["col"]
            if results.get(col):
                continue
            found = None
            if col == "Tax Registration Number":
                for lab in f["labels"]:
                    found = try_same_line_block(full_text, lab)
                    if found:
                        found = sanitize_trn(found)
                    if found:
                        break
            else:
                for lab in f["labels"]:
                    found = try_same_line_block(full_text, lab)
                    if found:
                        break
                if not found:
                    for lab in f["labels"]:
                        found = try_next_line_block(full_text, lab, BARRIER_LABELS)
                        if found:
                            break
            if not found and col.lower().startswith("email"):
                m = EMAIL_FALLBACK.search(full_text)
                if m:
                    found = m.group(0)
            results[col] = found

        if results["Currency"]:
            cur = results["Currency"].upper().strip()
            results["Currency"] = cur if re.fullmatch(r"[A-Z]{3}", cur) else None

        full_text_safe = locals().get("full_text", "")

        results["State/Province"], results["Postal Code"] = sanitize_state_and_postal(
            results.get("State/Province"), results.get("Postal Code")
        )

        if not results.get("State/Province"):
            city_val = results.get("City")
            postal_val = results.get("Postal Code")
            if city_val and postal_val:
                patt = re.compile(
                    rf"(?i)\b{re.escape(city_val)}\b\s+(?P<state>.+?)\s+\b{re.escape(postal_val)}\b"
                )
                for ln in full_text_safe.splitlines():
                    if LABEL_LINE_RE.match(ln):
                        continue
                    m = patt.search(ln)
                    if m:
                        st_raw = m.group("state")
                        st = clean_extracted_value(st_raw)
                        if st and is_alpha_line(st):
                            results["State/Province"] = st
                            break

        for name_col in ("First Name", "Last Name"):
            if results.get(name_col):
                v = clean_extracted_value(results[name_col])
                v = strip_parenthesized_dates(strip_signature_prefix(v) or v) or v
                results[name_col] = None if (v and contains_date_like(v)) else v

        trn = results.get("Tax Registration Number")
        results["Tax Registration Number"] = sanitize_trn(trn) if trn else None

        cur = results.get("Currency")
        c_val = sanitize_country(results.get("Country"), cur)
        if c_val and (
            re.fullmatch(r"[A-Za-z]{2,4}", c_val)
            or (cur and c_val.upper() == cur.upper())
            or contains_date_like(c_val)
        ):
            c_val = None
        results["Country"] = c_val or results.get("Country")

        translit_fields = {
            "First Name",
            "Last Name",
            "City",
            "State/Province",
            "Country",
            "Legal Name",
            "Hotel Name",
            "Address Line 1",
            "Preferred Language",
            "Tax Registration Status",
        }
        for col in translit_fields:
            if results.get(col):
                results[col] = unidecode(results[col]).strip()

    return results

def find_pdfs(input_dir: Path) -> List[Path]:
    return sorted([p for p in input_dir.rglob("*.pdf") if p.is_file()])

WEBADI_SHEET_DEFAULT = "WebADI"

def to_iso2(country_name: Optional[str]) -> Optional[str]:
    if not country_name:
        return None
    v = str(country_name).strip()
    if re.fullmatch(r"[A-Za-z]{2}", v):
        return v.upper()
    if pycountry:
        try:
            c = pycountry.countries.lookup(v)
            return getattr(c, "alpha_2", None)
        except Exception:
            return None
    return None

def _norm_header_key(s: str) -> str:
    if s is None:
        return ""
    s = str(s).replace("\u00A0", " ")
    s = s.replace("\n", " ").replace("\r", " ")
    s = s.strip()
    s = re.sub(r"\s+", " ", s)
    s = s.upper()
    s = s.replace("/", "_").replace("-", "_").replace(" ", "_")
    s = re.sub(r"[^A-Z0-9_]", "", s)
    return s

HEADER_SYNONYMS = {
    "BATCH_NAME": {"BATCH", "BATCHNAME"},
    "HOTEL_ID": {"PROPERTY_ID", "PROPERTYID", "EXPEDIA_ID", "HOTEL", "HOTELID", "PROPERTY"},
    "CUSTOMER_NAME": {"CUSTOMER", "CUSTOMERNAME", "LEGAL_NAME"},
    "ADDRESS_LINE_1": {"ADDRESS1", "ADDR1", "ADDRESS_LINE1", "ADDRESS"},
    "CITY": {"TOWN"},
    "STATE": {"STATE_PROVINCE", "STATEORPROVINCE", "STATE_", "STATE__PROVINCE"},
    "PROVINCE": {"STATE_PROVINCE", "STATEORPROVINCE"},
    "POSTAL_CODE": {"ZIP", "ZIP_CODE", "POSTCODE", "POSTALCODE"},
    "COUNTRY": {"COUNTRY_CODE", "COUNTRYNAME"},
    "BILLING_CURRENCY": {"CURRENCY", "CURRENCY_CODE"},
    "FIRST_NAME": {"FIRSTNAME", "CONTACT_FIRST_NAME"},
    "LAST_NAME": {"LASTNAME", "CONTACT_LAST_NAME"},
    "EMAIL_ADDRESS": {"EMAIL", "CONTACT_EMAIL", "E_MAIL"},
    "PREFERRED_LANGUAGE": {"LANGUAGE", "LANG"},
    "TAX_REG_NUMBER": {
        "TAX_REGISTRATION_NUMBER",
        "TRN",
        "VAT",
        "VAT_NUMBER",
        "TAXID",
        "TAX_ID",
    },
    "SITE_PURPOSE": {"SITEPURPOSE", "SITE_PURP", "SITE_PURPOSE_"},
    "BILL_TO": {"BILLTO", "SITE_USE_CODE", "SITE_USE", "SITEUSECODE"},
}

def _apply_synonym(key: str) -> str:
    for target, aliases in HEADER_SYNONYMS.items():
        if key == target or key in aliases:
            return target
    return key

def _debug_dump_first_rows(ws, n=6):
    print("\n[DEBUG] Amostra das primeiras linhas (normalizadas):")
    for r in range(1, min(ws.max_row, n) + 1):
        vals = []
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            vals.append(_apply_synonym(_norm_header_key(v)) if isinstance(v, str) else "")
        print(f"  L{r:02d}: {vals}")

def _debug_dump_row(ws, r):
    vals = []
    for c in range(1, ws.max_column + 1):
        v = ws.cell(r, c).value
        vals.append((_apply_synonym(_norm_header_key(v)) if isinstance(v, str) else "", v))
    print(f"[DEBUG] Linha {r} -> {[x for x in vals if x[1]]}")

def find_header_row(
    ws,
    must_have_cols=("HOTEL_ID", "CUSTOMER_NAME"),
    start_at_row: int = 5,
    debug: bool = False,
) -> Tuple[int, Dict[str, int]]:
    must_have = {_apply_synonym(_norm_header_key(x)) for x in must_have_cols}

    if debug:
        _debug_dump_first_rows(ws, n=max(6, start_at_row + 2))
        print(
            f"[DEBUG] Procurando cabeçalho a partir da linha {start_at_row} "
            f"(must_have={sorted(list(must_have))})"
        )

    for r in range(max(1, start_at_row), ws.max_row + 1):
        header_map_norm: Dict[str, int] = {}

        for c in range(1, ws.max_column + 1):
            val = ws.cell(r, c).value
            if not isinstance(val, str):
                continue

            raw = (
                val.replace("\u00A0", " ")
                .replace("\n", " ")
                .replace("\r", " ")
                .strip()
            )
            if not raw:
                continue

            key = _apply_synonym(_norm_header_key(raw))
            if key and key not in header_map_norm:
                header_map_norm[key] = c

        if must_have.issubset(header_map_norm.keys()):
            if debug:
                print(
                    f"[DEBUG] Cabeçalho encontrado na linha {r}. "
                    f"Chaves: {sorted(header_map_norm.keys())}"
                )
                _debug_dump_row(ws, r)
            return r, header_map_norm

    raise RuntimeError(
        "Não encontrei a linha de cabeçalho na aba WebADI. Verifique o template."
    )

def last_data_row(ws, header_row: int, key_col_index: int) -> int:
    r = ws.max_row
    while r > header_row:
        if ws.cell(r, key_col_index).value not in (None, "", " "):
            return r
        r -= 1
    return header_row

def clone_row(ws, src_row: int, dest_row: int):
    for c in range(1, ws.max_column + 1):
        src = ws.cell(src_row, c)
        dst = ws.cell(dest_row, c)
        dst.value = src.value
        dst.font = copy.copy(src.font)
        dst.border = copy.copy(src.border)
        dst.fill = copy.copy(src.fill)
        dst.number_format = src.number_format
        dst.protection = copy.copy(src.protection)
        dst.alignment = copy.copy(src.alignment)

def unprotect_workbook_and_sheets(wb, debug: bool = False):
    try:
        if hasattr(wb, "security"):
            wb.security.lockStructure = False
            wb.security.workbookPassword = None
            if debug:
                print("[DEBUG] Workbook security: lockStructure=False, password cleared.")
    except Exception as e:
        if debug:
            print(f"[DEBUG] Falha ao alterar wb.security: {e}")

    for ws in wb.worksheets:
        try:
            ws.protection.sheet = False
            try:
                ws.protection.password = None
            except Exception:
                pass
            if debug:
                print(f"[DEBUG] Desprotegida aba: {ws.title}")
        except Exception as e:
            if debug:
                print(f"[DEBUG] Falha ao desproteger aba {ws.title}: {e}")

def inject_into_webadi(
    template_path: Path,
    out_path: Optional[Path],
    df: pd.DataFrame,
    mode: str = "replace",
    sheet_name: str = WEBADI_SHEET_DEFAULT,
    header_start_row: int = 5,
    debug: bool = False,
    unprotect: bool = True,
):
    if not template_path.exists():
        raise FileNotFoundError(f"WEBADI template não encontrado: {template_path}")

    wb = load_workbook(
        str(template_path),
        keep_vba=True,
        data_only=False,
        keep_links=True,
    )

    if unprotect:
        if debug:
            print("[DEBUG] webadi_unprotect=True: desprotegendo workbook/abas.")
        unprotect_workbook_and_sheets(wb, debug=debug)
    else:
        if debug:
            print("[DEBUG] webadi_unprotect=False: mantendo proteções do workbook/abas.")

    if sheet_name not in wb.sheetnames:
        raise RuntimeError(f"Aba '{sheet_name}' não encontrada no template.")

    ws = wb[sheet_name]

    # Atualiza a célula E3 (valor do BATCH_NAME) para "OC MMDDYYYY"
    try:
        d3 = ws["D3"].value
        if isinstance(d3, str) and d3.strip().lower().replace(" ", "") == "*text":
            today_str = datetime.now().strftime("%m%d%Y")  # MMDDYYYY
            ws["E3"].value = f"OC {today_str}"
            if debug:
                print(f"[DEBUG] E3 (BATCH_NAME) atualizado para 'OC {today_str}'")
    except Exception as e:
        if debug:
            print(f"[DEBUG] Falha ao atualizar BATCH_NAME em E3: {e}")

    header_row, header_map = find_header_row(
        ws, start_at_row=header_start_row, debug=debug
    )

    wanted = {
        "HOTEL_ID": "Expedia ID",
        "CUSTOMER_NAME": "Legal Name",
        "ADDRESS_LINE_1": "Address Line 1",
        "CITY": "City",
        "POSTAL_CODE": "Postal Code",
        "COUNTRY": "Country",
        "BILLING_CURRENCY": "Currency",
        "FIRST_NAME": "First Name",
        "LAST_NAME": "Last Name",
        "EMAIL_ADDRESS": "Email Address",
        "TAX_REG_NUMBER": "Tax Registration Number",
    }

    col_state = header_map.get("STATE")
    col_province = header_map.get("PROVINCE")

    col_bill_to = header_map.get("BILL_TO")
    if not col_bill_to:
        col_bill_to = column_index_from_string("Q")
        if debug:
            print(
                "[DEBUG] Coluna BILL_TO não encontrada por cabeçalho; "
                "usando fallback na coluna Q."
            )

    col_site_purpose = header_map.get("SITE_PURPOSE")

    key_col = header_map.get("HOTEL_ID") or header_map.get("CUSTOMER_NAME")
    if not key_col:
        raise RuntimeError(
            "Não encontrei colunas-chave (HOTEL_ID/CUSTOMER_NAME) no cabeçalho WebADI."
        )

    if mode.lower() == "replace":
        last = ws.max_row
        for _ in range(header_row + 1, last + 1):
            ws.delete_rows(header_row + 1)
        base_row = header_row
        if debug:
            print(
                f"[DEBUG] Modo replace: limpei linhas após o cabeçalho (linha {header_row})."
            )
    else:
        base_row = last_data_row(ws, header_row, key_col)
        if debug:
            print(
                f"[DEBUG] Modo append: última linha de dados detectada = {base_row}."
            )

    current = base_row

    for _, rec in df.iterrows():
        iso2 = to_iso2(rec.get("Country"))
        stateprov = rec.get("State/Province")

        current += 1
        src_row = base_row if base_row > header_row else header_row + 1
        clone_row(ws, src_row, current)

        if col_bill_to:
            ws.cell(current, col_bill_to).value = "BILL_TO"
        if col_site_purpose:
            ws.cell(current, col_site_purpose).value = "BILL_TO"

        for webadi_col, script_col in wanted.items():
            col_idx = header_map.get(webadi_col)
            if not col_idx:
                continue
            val = rec.get(script_col)

            if webadi_col == "COUNTRY":
                v2 = to_iso2(val) or val
                if v2 not in (None, "", " "):
                    ws.cell(current, col_idx).value = v2
            elif webadi_col == "BILLING_CURRENCY":
                if isinstance(val, str) and re.fullmatch(r"[A-Za-z]{3}", val.strip()):
                    ws.cell(current, col_idx).value = val.strip().upper()
            else:
                if val not in (None, "", " "):
                    ws.cell(current, col_idx).value = val

        if stateprov and isinstance(stateprov, str):
            if iso2 == "CA" and col_province:
                ws.cell(current, col_province).value = stateprov
            elif iso2 != "CA" and col_state:
                ws.cell(current, col_state).value = stateprov

    if not out_path:
        out_path = template_path

    # NÃO removemos imagens/ícones para manter a formatação o mais fiel possível
    wb.save(str(out_path))
    print(f"[OK] WEBADI preenchido: {out_path}")

def main():
    ap = argparse.ArgumentParser(
        description="Extract HCPIF fields into Excel (R2.6 + WEBADI injection robusta)."
    )
    ap.add_argument("--input_dir", type=str, default=str(DEFAULT_INPUT_DIR))
    ap.add_argument("--output", type=str, default=str(DEFAULT_OUTPUT_FILE))
    ap.add_argument(
        "--webadi_template",
        type=str,
        default=None,
        help="Caminho do .xlsm/.xlsx do WEBADI",
    )
    ap.add_argument(
        "--webadi_output",
        type=str,
        default=None,
        help="Caminho p/ salvar cópia preenchida. Se omitido, gera 'OC WEBADI MMDDYYYY.xlsm' na mesma pasta do template.",
    )
    ap.add_argument(
        "--webadi_mode",
        type=str,
        choices=["append", "replace"],
        default="replace",
        help="append (clona última linha); replace (limpa região de dados antes).",
    )
    ap.add_argument(
        "--webadi_sheet",
        type=str,
        default=WEBADI_SHEET_DEFAULT,
        help="Nome da aba do WEBADI (default: WebADI).",
    )
    ap.add_argument(
        "--webadi_header_start_row",
        type=int,
        default=5,
        help="Linha inicial para procurar o cabeçalho (default: 5).",
    )
    ap.add_argument(
        "--webadi_debug",
        action="store_true",
        help="Imprime diagnósticos de cabeçalho/aba.",
    )
    ap.add_argument(
        "--webadi_unprotect",
        dest="webadi_unprotect",
        action="store_true",
        help="Desprotege workbook/abas antes de escrever (padrão: True).",
    )
    ap.add_argument(
        "--webadi_keep_protection",
        dest="webadi_unprotect",
        action="store_false",
        help="Não desproteger workbook/abas antes de escrever.",
    )
    ap.set_defaults(webadi_unprotect=True)

    args = ap.parse_args()

    input_dir = Path(args.input_dir)
    raw_out = Path(args.output)

    output_path = (
        raw_out / "HCPIF_extraction.xlsx"
        if (raw_out.exists() and raw_out.is_dir())
        else (raw_out if raw_out.suffix.lower() == ".xlsx" else raw_out.with_suffix(".xlsx"))
    )
    output_path.parent.mkdir(parents=True, exist_ok=True)

    if not input_dir.exists():
        print(f"[ERROR] Input folder not found: {input_dir}")
        sys.exit(1)

    pdf_files = find_pdfs(input_dir)
    if not pdf_files:
        print(f"[ERROR] No PDFs found in: {input_dir}")
        sys.exit(2)

    rows = []
    print(f"\nReading {len(pdf_files)} PDF(s) from: {input_dir}\n")
    for pdf in tqdm.tqdm(pdf_files, desc="Processing PDFs"):
        try:
            data = extract_fields_positional(pdf)
            data["file_name"] = pdf.name
            data["file_path"] = str(pdf.resolve())
            rows.append(data)
        except Exception as e:
            rows.append(
                {
                    "file_name": pdf.name,
                    "file_path": str(pdf.resolve()),
                    "status": f"error: {e}",
                }
            )

    df = pd.DataFrame(rows)
    ordered_cols = [f["col"] for f in FIELDS]
    aux_cols = [c for c in df.columns if c not in ordered_cols]
    df = df[ordered_cols + aux_cols]

    try:
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Extraction", index=False)
        print("\n[OK] Done!")
        print(f"- PDFs read from: {input_dir}")
        print(f"- Excel saved to: {output_path}\n")
    except PermissionError:
        print("\n[ERROR] Could not write the Excel file.")
        print(f"Destination: {output_path}")
        print("Common causes:")
        print("  1) File is currently OPEN in Excel.")
        print("  2) You lack permission.")
        raise

    template_path = None
    if args.webadi_template:
        template_path = Path(args.webadi_template)
    else:
        default_tpl = Path(DEFAULT_WEBADI_TEMPLATE)
        if default_tpl.exists():
            template_path = default_tpl

    if template_path:
        if args.webadi_output:
            out_path = Path(args.webadi_output)
            out_path.parent.mkdir(parents=True, exist_ok=True)
        else:
            today_str = datetime.now().strftime("%m%d%Y")  # MMDDYYYY
            out_name = f"OC WEBADI {today_str}.xlsm"
            out_path = template_path.parent / out_name

        print(f"[INFO] WEBADI template: {template_path}")
        print(f"[INFO] WEBADI output: {out_path}")
        print(
            f"[INFO] WEBADI sheet: {args.webadi_sheet} | header_start_row={args.webadi_header_start_row}"
        )

        need_cols = {
            "Expedia ID",
            "Legal Name",
            "Address Line 1",
            "City",
            "State/Province",
            "Postal Code",
            "Country",
            "Currency",
            "First Name",
            "Last Name",
            "Email Address",
            "Preferred Language",
            "Tax Registration Number",
        }
        present = [c for c in need_cols if c in df.columns]

        if not present:
            print("[WARN] Nenhuma coluna necessária para o WEBADI está disponível.")
        else:
            df2 = df.loc[
                (df["Expedia ID"].notna()) | (df["Legal Name"].notna()), present
            ]
            if not len(df2):
                print("[WARN] Nada para inserir no WEBADI (sem Expedia ID / Legal Name).")
            else:
                print(f"[INFO] WEBADI mode: {args.webadi_mode}")
                try:
                    inject_into_webadi(
                        template_path,
                        out_path,
                        df2,
                        mode=args.webadi_mode,
                        sheet_name=args.webadi_sheet,
                        header_start_row=args.webadi_header_start_row,
                        debug=args.webadi_debug,
                        unprotect=args.webadi_unprotect,
                    )
                except Exception as ex:
                    print(f"[ERROR] Falha ao preencher WEBADI: {ex}")
                    print(
                        "Sugestões: confira o nome da aba, desproteja a planilha (se necessário), e valide a linha do cabeçalho."
                    )
    else:
        print(
            "[INFO] Passo WEBADI não executado (sem --webadi_template e sem template padrão encontrado)."
        )

if __name__ == "__main__":
    main()
