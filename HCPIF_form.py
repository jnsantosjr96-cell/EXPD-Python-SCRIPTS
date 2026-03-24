# -*- coding: utf-8 -*-
"""
HCPIF extractor R2.6 + OCR fallback
- Base: R2.4 (robust, Unicode-safe, anti-signature & anti-date, smart transliteration)
- R2.5: WEBADI integration preserving macros/structure; keeps "Preferred Language" and "BATCH_NAME".
- R2.6:
  * Optionally unprotects workbook and all sheets before writing (CLI flag).
  * Robust header detection (normalization, synonyms, tolerant to "HOTEL ID", NBSP, line breaks).
  * Header search starting at row 5 (configurable).
  * CLI for sheet name, header start row, and debug.
Extras:
- OCR fallback via Tesseract + pdf2image for scrambled-text PDFs.
- Cleaning of junk characters at the beginning of fields (dates, currency, emails etc.).
- Normalization of Today's date, Effective Date of Change and Currency, with fallbacks.

R2.7:
- Oracle (TCA) connection by Expedia ID (ATTRIBUTE1) with dynamic IN list.
- Match HCPIF x Oracle by Expedia ID + country (ISO2).
- New columns in HCPIF Extraction:
    * Country ISO2 (right after Country)
    * Found SLE OID (ORACLE_ID from query)
    * Found SLE Name (SLE_NAME from query)
    * Oracle Currency
    * Currency Matches Oracle (YES/NO)
    * Comments ("Updated Ownership per HCPIF - <file_name>")
- Forces Hotel Name to be filled (uses Oracle HOTEL_NAME if empty from extraction).
- Removes special characters from Legal Name (keeps only letters, numbers and space) in Output and WEBADI.
- WEBADI: CUSTOMER_NAME now comes from Hotel Name (not Legal Name / SLE Name).
- Rows where RECEIPT_METHOD_NAME contains 'DIRECT DEBIT' or TAI_BM contains 'GROUP'
  are copied to a separate review sheet (but kept in main output and WEBADI).

R2.7.1:
- Sends Effective Date of Change from the form to WEBADI column EFFECTIVE_DATE.
- Comments: "Updated Ownership per HCPIF - <Case ... - EID ...>" (Case first, then EID).
- Expands currency handling for Greece/GR (EUR), Hong Kong/HK (HKD), Canada/CA (CAD).

R2.7.3:
- OCR improved (DPI 350, Tesseract --oem 3 --psm 6).
- Postal Code:
  * Map typical OCR confusions: O/o -> 0, G/g -> 8.
  * Force Postal Code to contain only digits; if nothing left, set to blank.
- State/Province:
  * Value comes only from the labeled State/Province field (positional extraction),
    not inferred from the "City ... State ... Postal" line or generic text search.
"""

# -------------------- DEFAULT PATHS --------------------
from pathlib import Path

DEFAULT_INPUT_DIR = r"C:\Users\josenjr\Downloads\HCPIFs"
DEFAULT_OUTPUT_FILE = r"C:\Users\josenjr\Downloads\HCPIFs\Output\HCPIF_extraction.xlsx"

DEFAULT_WEBADI_TEMPLATE = r"C:\Users\josenjr\Downloads\HCPIFs\OC WebADI SF 149540900 2-13-26 (002).xlsm"

# -------------------- POPPLER / TESSERACT CONFIG --------------------

POPPLER_PATH = r"C:\Users\josenjr\OneDrive - Expedia Group\Desktop\Poppler\poppler-25.12.0\Library\bin"

TESSERACT_EXE = r"C:\Users\josenjr\OneDrive - Expedia Group\Desktop\Tessdata2\tesseract.exe"
TESSDATA_PREFIX = r"C:\Users\josenjr\OneDrive - Expedia Group\Desktop\Tessdata2\tessdata"

# -------------------- ORACLE CONFIG --------------------
ORACLE_CLIENT_DIR = r"C:\Users\josenjr\OneDrive - Expedia Group\Desktop\Oracle Instant Client\instantclient_23_0"
ORACLE_USERNAME = "josenjr"
ORACLE_PASSWORD = "qyuxYQZFs13"  # Replace locally
ORACLE_DSN = "ashworaebsdb02-vip.datawarehouse.expecn.com:1526/ORAPRD_UI"

ORACLE_TCA_SQL_TEMPLATE = """
SELECT
    HCA.ACCOUNT_NUMBER AS ORACLE_ID,
    COUNT(HCA.ACCOUNT_NUMBER) OVER (PARTITION BY HCA.ACCOUNT_NUMBER) AS DUPLICATE_COUNT,
    HZP.PARTY_NUMBER,
    HCA.ATTRIBUTE1 AS EXPEDIA_ID,
    HZP.PARTY_NAME AS HOTEL_NAME,
    TAI.Attribute1 AS TAI_PC,
    TAI.Business_Model AS TAI_BM,
    SITEOU.NAME AS SITE_OU,
    HCA.CUSTOMER_CLASS_CODE,
    HCA.ATTRIBUTE3 AS SOURCE_SYSTEM,
    HCA.CREATION_DATE AS ACCOUNT_CREATION_DATE,
    SLE.PARTY_NAME AS SLE_NAME,
    SLE.SLE_OID,
    REL.START_DATE AS RELATIONSHIP_DATE,
    HZP.ADDRESS1,
    HZP.CITY,
    HZP.POSTAL_CODE,
    HZP.STATE,
    HZP.PROVINCE,
    HZP.COUNTRY AS COUNTRY_ABBR,
    TER.TERRITORY_SHORT_NAME AS COUNTRY_FULL_NAME,
    CONTACT.USER_NAME AS LAST_UPDATED_BY,
    CONTACT.FIRST_NAME,
    CONTACT.LAST_NAME,
    CONTACT.EMAIL_ADDRESS,
    CONTACT.PRIMARY_FLAG AS CONTACT_PRIMARY_FLAG,
    TAX.REGISTRATION_NUMBER,
    TAX.TAX,
    TAX.TAX_REGIME_CODE,
    TAX.REGISTRATION_STATUS_CODE,
    TAX.EFFECTIVE_FROM AS TAX_START_DATE,
    CURR.CURRENCY_CODE AS ORACLE_CURRENCY,
    RM.NAME AS RECEIPT_METHOD_NAME,
    RM.CURRENT_RM,
    RM.RM_START_DATE
FROM AR.HZ_CUST_ACCOUNTS HCA
LEFT JOIN AR.HZ_PARTIES HZP
    ON HZP.PARTY_ID = HCA.PARTY_ID
LEFT JOIN 
    (SELECT HR.OBJECT_ID,HR.SUBJECT_ID,HR.RELATIONSHIP_CODE,HR.START_DATE
        FROM AR.HZ_RELATIONSHIPS HR
            WHERE HR.RELATIONSHIP_CODE = 'LEGAL ENTITY/OWNER OF'
                AND HR.END_DATE LIKE '31-DEC-12') REL
                    ON REL.OBJECT_ID = HCA.PARTY_ID
LEFT JOIN
    (SELECT HZP2.PARTY_NAME,HZP2.PARTY_ID,HCA2.ACCOUNT_NUMBER AS SLE_OID
        FROM AR.HZ_PARTIES HZP2
            LEFT JOIN AR.HZ_CUST_ACCOUNTS HCA2
                ON HCA2.PARTY_ID = HZP2.PARTY_ID) SLE
                    ON SLE.PARTY_ID = REL.SUBJECT_ID
LEFT JOIN 
    (SELECT PTP.PARTY_ID,ZR.REGISTRATION_NUMBER,ZR.REGISTRATION_STATUS_CODE,ZR.EFFECTIVE_FROM,ZR.EFFECTIVE_TO,ZR.TAX_REGIME_CODE,ZR.TAX
        FROM APPS.ZX_PARTY_TAX_PROFILE PTP
            LEFT JOIN APPS.ZX_REGISTRATIONS ZR
                ON PTP.PARTY_TAX_PROFILE_ID = ZR.PARTY_TAX_PROFILE_ID
                    WHERE ZR.REGISTRATION_NUMBER IS NOT NULL
                        AND ZR.EFFECTIVE_TO IS NULL) TAX
                            ON TAX.PARTY_ID = HCA.PARTY_ID
LEFT JOIN
    (SELECT HCPA.CURRENCY_CODE,HCPA.CUST_ACCOUNT_ID
        FROM AR.HZ_CUST_PROFILE_AMTS HCPA
            WHERE HCPA.ATTRIBUTE1 IS NULL) CURR
                ON CURR.CUST_ACCOUNT_ID = HCA.CUST_ACCOUNT_ID
LEFT JOIN
    (SELECT CRM.CUSTOMER_ID,CRM.ATTRIBUTE1 AS CURRENT_RM,CRM.START_DATE AS RM_START_DATE,ARM.NAME
        FROM AR.RA_CUST_RECEIPT_METHODS CRM
            LEFT JOIN AR.AR_RECEIPT_METHODS ARM
                ON CRM.RECEIPT_METHOD_ID = ARM.RECEIPT_METHOD_ID
                    WHERE CRM.PRIMARY_FLAG = 'Y'
                        AND CRM.END_DATE IS NULL) RM
                            ON RM.CUSTOMER_ID = HCA.CUST_ACCOUNT_ID
LEFT JOIN
    (SELECT ACV.CUSTOMER_ID,ACV.FIRST_NAME,ACV.LAST_NAME,ACV.EMAIL_ADDRESS,CRV.PRIMARY_FLAG,FNDU.USER_NAME
        FROM APPS.AR_CONTACTS_V ACV
            LEFT JOIN APPS.AR_CONTACT_ROLES_V CRV
                ON ACV.CONTACT_ID = CRV.CONTACT_ID
                    LEFT JOIN APPS.FND_USER FNDU
                        ON FNDU.USER_ID = ACV.LAST_UPDATED_BY
                    WHERE STATUS = 'A'
                        AND CRV.PRIMARY_FLAG = 'Y') CONTACT
                            ON CONTACT.CUSTOMER_ID = HCA.CUST_ACCOUNT_ID
LEFT JOIN
    (SELECT TERRITORY_CODE,TERRITORY_SHORT_NAME
        FROM APPLSYS.FND_TERRITORIES_TL) TER
            ON TER.TERRITORY_CODE = HZP.COUNTRY
LEFT JOIN XXEXPD.XXEXPD_TCA_ADDITIONAL_INFO TAI
    ON HCA.Cust_Account_ID = TAI.partner_id
LEFT JOIN HZ_PARTY_SITES PS
    ON HZP.party_id = PS.party_id
LEFT JOIN HR_OPERATING_UNITS SITEOU  
    ON HCA.ORG_ID = SITEOU.ORGANIZATION_ID
WHERE HCA.CUSTOMER_CLASS_CODE = 'HOTEL'
  AND HCA.Attribute1 IN ({expedia_ids})
"""

# -------------------- ENSURE / INSTALL DEPS --------------------
import sys
import subprocess
import importlib
import argparse
import copy
import re
import os
import time
from datetime import datetime
from typing import Dict, List, Optional, Tuple

import oracledb

REQUIRED = [
    ("pdfplumber", "pdfplumber"),
    ("pandas", "pandas"),
    ("openpyxl", "openpyxl"),
    ("tqdm", "tqdm"),
    ("Unidecode", "unidecode"),
    ("pycountry", "pycountry"),
    ("pytesseract", "pytesseract"),
    ("pdf2image", "pdf2image"),
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
pytesseract = ensure_package("pytesseract", "pytesseract")
pdf2image = ensure_package("pdf2image", "pdf2image")

from unidecode import unidecode
from pdf2image import convert_from_path

pytesseract.pytesseract.tesseract_cmd = TESSERACT_EXE
os.environ["TESSDATA_PREFIX"] = TESSDATA_PREFIX

try:
    pycountry = ensure_package("pycountry", "pycountry")
except Exception:
    pycountry = None

from openpyxl.packaging import manifest as _ox_manifest

for _ext in (".jpg", ".jpeg", ".JPG", ".JPEG"):
    try:
        _ox_manifest.mimetypes.add_type("image/jpeg", _ext)
    except Exception:
        pass

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

# -------------------- FIELD DEFINITIONS --------------------
FIELDS: List[Dict] = [
    {"col": "Today's date", "labels": [r"today[\\']s\s*date"]},
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

KNOWN_CURRENCIES = {"AED", "AFN", "ALL", "AMD", "ANG", "AOA", "ARS", "AUD", "AWG", "AZN",
    "BAM", "BBD", "BDT", "BGN", "BHD", "BIF", "BMD", "BND", "BOB", "BRL",
    "BSD", "BTN", "BWP", "BYN", "BZD",
    "CAD", "CDF", "CHF", "CLP", "CNY", "COP", "CRC", "CUP", "CVE", "CZK",
    "DJF", "DKK", "DOP", "DZD",
    "EGP", "ERN", "ETB", "EUR",
    "FJD", "FKP",
    "GBP", "GEL", "GHS", "GIP", "GMD", "GNF", "GTQ", "GYD",
    "HKD", "HNL", "HTG", "HUF",
    "IDR", "ILS", "INR", "IQD", "IRR", "ISK",
    "JMD", "JOD", "JPY",
    "KES", "KGS", "KHR", "KMF", "KPW", "KRW", "KWD", "KYD", "KZT",
    "LAK", "LBP", "LKR", "LRD", "LSL", "LYD",
    "MAD", "MDL", "MGA", "MKD", "MMK", "MNT", "MOP", "MRU", "MUR", "MVR",
    "MWK", "MXN", "MYR", "MZN",
    "NAD", "NGN", "NIO", "NOK", "NPR", "NZD",
    "OMR",
    "PAB", "PEN", "PGK", "PHP", "PKR", "PLN", "PYG",
    "QAR",
    "RON", "RSD", "RUB", "RWF",
    "SAR", "SBD", "SCR", "SDG", "SEK", "SGD", "SHP", "SLE", "SLL", "SOS",
    "SRD", "SSP", "STN", "SVC", "SYP", "SZL",
    "THB", "TJS", "TMT", "TND", "TOP", "TRY", "TTD", "TWD", "TZS",
    "UAH", "UGX", "USD", "UYU", "UZS",
    "VED", "VES", "VND", "VUV",
    "WST",
    "XAF", "XCD", "XDR", "XOF", "XPF",
    "YER",
    "ZAR", "ZMW", "ZWL",}

DEFAULT_CURRENCY_BY_COUNTRY = {
    "GERMANY": "EUR",
    "DE": "EUR",
    "SPAIN": "EUR",
    "ES": "EUR",
    "ITALY": "EUR",
    "IT": "EUR",
    "MONTENEGRO": "EUR",
    "ME": "EUR",
    "MEXICO": "MXN",
    "MX": "MXN",
    "UNITED STATES": "USD",
    "USA": "USD",
    "US": "USD",
    "SWITZERLAND": "CHF",
    "CH": "CHF",
    "THAILAND": "THB",
    "TH": "THB",
    "AUSTRALIA": "AUD",
    "AU": "AUD",
    "GREECE": "EUR",
    "GR": "EUR",
    "HONG KONG": "HKD",
    "HK": "HKD",
    "CANADA": "CAD",
    "CA": "CAD",
}

# -------------------- DATE / VALUE HELPERS --------------------
def is_valid_date_token(token: str) -> bool:
    token = token.strip()
    for fmt in DATE_PATTERNS:
        try:
            datetime.strptime(token, fmt)
            return True
        except ValueError:
            continue
    return False

def extract_first_date_token(s: Optional[str]) -> Optional[str]:
    if not s:
        return None
    for m in DATE_TOKEN_REGEX.finditer(s):
        token = m.group(0)
        if is_valid_date_token(token):
            return token
    return None

def extract_two_dates_from_lines(full_text: str):
    tokens = DATE_TOKEN_REGEX.findall(full_text or "")
    valids = [t for t in tokens if is_valid_date_token(t)]
    if len(valids) >= 2:
        return valids[0], valids[1]
    if len(valids) == 1:
        return valids[0], None
    return None, None

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
    s = re.sub(r"\([^0-9()]*\d[^()]*\)", "", s)
    return s.strip() or None

LEADING_JUNK_RE = re.compile(r"^[^0-9A-Za-z]+")

def strip_leading_junk(s: Optional[str]) -> Optional[str]:
    if s is None:
        return None
    s = str(s).lstrip()
    s = LEADING_JUNK_RE.sub("", s)
    return s

def clean_extracted_value(s: Optional[str]) -> Optional[str]:
    if s is None:
        return None
    v = normalize_block_text(s)
    v = cut_at_next_label(v)
    v = strip_signature_prefix(v) or v
    v = strip_parenthesized_dates(v) or v
    v = re.sub(r"\s{2,}", " ", v).strip()
    v = strip_leading_junk(v) or ""
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
        if HEADER_NOISE.search(cand) or SIGNATURE_NOISE.match(cand) or contains_date_like(
            cand
        ):
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
        if all(_norm(words[i + j]["text"]) == norm_labels[j] for j in range(m)):
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

def is_garbage_text(text: str) -> bool:
    """
    Simple heuristic to detect scrambled text.
    If we don't find any typical HCPIF keywords, treat as "garbage" and use OCR.
    """
    if not text:
        return True
    s = text.strip()
    if len(s) < 40:
        return True
    lowered = s.lower()
    keywords = [
        "hotel collect payment information form",
        "legal name",
        "expedia id",
        "country",
        "currency",
        "first name",
        "last name",
        "email address",
        "preferred language",
    ]
    if any(kw in lowered for kw in keywords):
        return False
    return True

def ocr_extract_full_text(pdf_path: Path) -> str:
    """
    Convert PDF to images and use OCR (pytesseract) to extract readable text.
    Uses DPI 350 and LSTM engine (oem 3, psm 6) for better quality.
    """
    texts = []
    images = convert_from_path(str(pdf_path), dpi=350, poppler_path=POPPLER_PATH)
    for img in images:
        try:
            t = pytesseract.image_to_string(
                img,
                lang="eng",
                config="--oem 3 --psm 6",
            )
            if t:
                texts.append(t)
        except Exception:
            continue
    return "\n".join(texts)

def strip_state_if_person_name(results: Dict[str, Optional[str]]) -> None:
    """
    Se State/Province for igual ao First Name, Last Name ou "First Last", zera o campo.
    Evita casos em que state vazio puxa nome de contato.
    """
    s = results.get("State/Province")
    if not s:
        return
    fn = results.get("First Name")
    ln = results.get("Last Name")
    if not fn and not ln:
        return

    def norm(x):
        return unidecode(str(x)).strip().lower()

    s_n = norm(s)
    candidates = set()
    if fn:
        candidates.add(norm(fn))
    if ln:
        candidates.add(norm(ln))
    if fn and ln:
        candidates.add((norm(fn) + " " + norm(ln)).strip())

    if s_n in candidates:
        results["State/Province"] = None

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

        full_text_raw = normalize_block_text("\n".join(full_text_pages))
        use_ocr = is_garbage_text(full_text_raw)
        if use_ocr:
            ocr_text = ocr_extract_full_text(pdf_path)
            full_text = normalize_block_text(ocr_text)
            pages_bundle = []
        else:
            full_text = full_text_raw

        for idx, page, words, text in pages_bundle:
            if not words:
                continue

            anchors_currency = find_label_anchors(words, ["currency"])
            anchors_country = find_label_anchors(words, ["country"])
            anchors_edc = find_label_anchors(words, ["effective", "date", "of", "change"])
            anchors_today = find_label_anchors(words, ["today's", "date"])
            anchors_trn = find_label_anchors(words, ["tax", "registration", "number"])
            anchors_state = find_label_anchors(words, ["state", "province"]) + \
                            find_label_anchors(words, ["state"]) + \
                            find_label_anchors(words, ["province"])

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

            # State/Province ONLY from labeled field (positional)
            if not results["State/Province"] and anchors_state:
                for a in anchors_state:
                    same = collect_tokens_on_same_line_right(words, a, y_tol=6.0)
                    if same:
                        raw = join_tokens(same)
                        val = clean_extracted_value(raw)
                        if val:
                            results["State/Province"] = val
                            break

        d1, d2 = extract_two_dates_from_lines(full_text)
        if d1 and not results["Today's date"]:
            results["Today's date"] = d1
        if d2 and not results["Effective Date of Change"]:
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

        # ------- Fallback por regex para campos restantes -------
        for f in FIELDS:
            col = f["col"]
            if results.get(col):
                continue

            # Fallback controlado para State/Province:
            if col == "State/Province":
                found = None
                specific_labels = [r"state\s*/\s*province", r"state\s*province"]
                for lab in specific_labels:
                    found = try_same_line_block(full_text, lab)
                    if not found:
                        found = try_next_line_block(full_text, lab, BARRIER_LABELS)
                    if found:
                        break
                results[col] = found
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
        # ------- fim fallback FIELDS -------

        if results["Currency"]:
            cur0 = results["Currency"].upper().strip()
            results["Currency"] = cur0 if re.fullmatch(r"[A-Z]{3}", cur0) else None

        # Limpa state se for igual a nome de contato
        strip_state_if_person_name(results)

        # Only use sanitize_state_and_postal to handle obvious swaps;
        results["State/Province"], results["Postal Code"] = sanitize_state_and_postal(
            results.get("State/Province"), results.get("Postal Code")
        )

        for name_col in ("First Name", "Last Name"):
            if results.get(name_col):
                v = clean_extracted_value(results[name_col])
                v = strip_parenthesized_dates(strip_signature_prefix(v) or v) or v
                results[name_col] = None if (v and contains_date_like(v)) else v

        trn = results.get("Tax Registration Number")
        results["Tax Registration Number"] = sanitize_trn(trn) if trn else None

        # --- Extra sanity for Postal Code (only digits, OCR fix O->0, G->8) ---
        p_val = results.get("Postal Code")
        if p_val:
            s = str(p_val)
            trans = str.maketrans({"O": "0", "o": "0", "G": "8", "g": "8"})
            s2 = s.translate(trans)
            digits = "".join(ch for ch in s2 if ch.isdigit())
            results["Postal Code"] = digits or None
            p_val = results["Postal Code"]

        if p_val and not re.search(r"\d", str(p_val)):
            results["Postal Code"] = None

        # Special case: Hong Kong should not have short numeric "CEP" like 852
        country_raw = (results.get("Country") or "").strip()
        iso2_country = None
        try:
            iso2_country = to_iso2(country_raw) or ""
        except Exception:
            iso2_country = ""
        if iso2_country.upper() == "HK":
            p_val = results.get("Postal Code")
            if p_val and re.fullmatch(r"\d{1,4}", str(p_val).strip()):
                results["Postal Code"] = None

        cur = results.get("Currency")
        c_val = sanitize_country(results.get("Country"), cur)
        if c_val and (
            re.fullmatch(r"[A-Za-z]{2,4}", c_val)
            or (cur and c_val.upper() == cur.upper())
            or contains_date_like(c_val)
        ):
            c_val = None
        results["Country"] = c_val or results.get("Country")

        for date_col in ("Today's date", "Effective Date of Change"):
            raw = results.get(date_col)
            token = extract_first_date_token(raw)
            results[date_col] = token

        if results.get("Today's date") and not results.get("Effective Date of Change"):
            d1, d2 = extract_two_dates_from_lines(full_text)
            if not d2 and d1 == results["Today's date"]:
                results["Effective Date of Change"] = d1

        cur = results.get("Currency")
        if isinstance(cur, str):
            c = cur.strip().upper()
            if not re.fullmatch(r"[A-Z]{3}", c):
                m = re.search(r"\b([A-Z]{3})\b", c)
                c = m.group(1).upper() if m else None
            if c and c not in KNOWN_CURRENCIES:
                c = None
            results["Currency"] = c

        if not results.get("Currency"):
            country_raw = (results.get("Country") or "").strip()
            country_norm = country_raw.upper()
            iso2_country = None
            try:
                iso2_country = to_iso2(country_raw) or ""
            except Exception:
                iso2_country = ""
            cur_final = (
                DEFAULT_CURRENCY_BY_COUNTRY.get(country_norm)
                or DEFAULT_CURRENCY_BY_COUNTRY.get(iso2_country.upper(), None)
            )
            results["Currency"] = cur_final

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

    # Normalização forte para tratar casos como Türkiye/Turkiye/Turkey
    v_norm = unidecode(v).strip().upper()

    MANUAL_ISO2 = {
        "TURKEY": "TR",
        "TURKIYE": "TR",
    }
    if v_norm in MANUAL_ISO2:
        return MANUAL_ISO2[v_norm]

    # Se já for código de 2 letras (ex.: "TR")
    if re.fullmatch(r"[A-Za-z]{2}", v):
        return v.upper()

    if pycountry:
        # Tenta primeiro com a forma normalizada, depois com o valor original
        for cand in (v_norm, v):
            try:
                c = pycountry.countries.lookup(cand)
                return getattr(c, "alpha_2", None)
            except Exception:
                continue
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
        "TAX_REG_NUM",
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
    print("\n[DEBUG] Sample of first rows (normalized):")
    for r in range(1, min(ws.max_row, n) + 1):
        vals = []
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            vals.append(_apply_synonym(_norm_header_key(v)) if isinstance(v, str) else "")
        print(f"  R{r:02d}: {vals}")

def _debug_dump_row(ws, r):
    vals = []
    for c in range(1, ws.max_column + 1):
        v = ws.cell(r, c).value
        vals.append((_apply_synonym(_norm_header_key(v)) if isinstance(v, str) else "", v))
    print(f"[DEBUG] Row {r} -> {[x for x in vals if x[1]]}")

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
            f"[DEBUG] Searching for header starting at row {start_at_row} "
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
                    f"[DEBUG] Header found at row {r}. "
                    f"Keys: {sorted(header_map_norm.keys())}"
                )
                _debug_dump_row(ws, r)
            return r, header_map_norm
    raise RuntimeError(
        "Could not find header row in WebADI sheet. "
        "Check the template."
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
            print(f"[DEBUG] Failed to modify wb.security: {e}")

    for ws in wb.worksheets:
        try:
            ws.protection.sheet = False
            try:
                ws.protection.password = None
            except Exception:
                pass
            if debug:
                print(f"[DEBUG] Unprotected sheet: {ws.title}")
        except Exception as e:
            if debug:
                print(f"[DEBUG] Failed to unprotect sheet {ws.title}: {e}")

def _is_blank(val) -> bool:
    """True para None, string vazia ou NaN."""
    if val is None:
        return True
    if isinstance(val, str) and val.strip() == "":
        return True
    try:
        import pandas as _pd
        return bool(_pd.isna(val))
    except Exception:
        return False

def normalize_language_to_code(val: Optional[str]) -> Optional[str]:
    """
    Converte a linguagem (string) em código ISO 639-1 de 2 letras (EN, PT, ES, ...).
    - Usa pycountry.languages como fonte principal (padrão internacional).
    - Aceita nomes ('English', 'Portuguese (Brazil)'), códigos com região ('en-US', 'pt_BR').
    - Se não reconhecer com segurança, retorna None (fica em branco no WEBADI).
    """
    if val is None:
        return None

    s_raw = str(val).strip()
    if not s_raw:
        return None

    # Normaliza acentos e caixa
    s = unidecode(s_raw).strip()

    # 1) Já vier um código tipo "en", "EN", "pt", etc.
    if len(s) == 2 and s.isalpha():
        return s.upper()

    # 2) Pega prefixo de 2 letras antes de "-" ou "_" (ex: "en-US", "pt_BR", "en (US)")
    m = re.match(r"\s*([A-Za-z]{2})\b", s)
    if m:
        return m.group(1).upper()

    # 3) Tenta lookup pelo pycountry (nome, alpha_2, alpha_3, bibliographic, etc.)
    if pycountry is not None:
        try:
            lang = pycountry.languages.lookup(s.lower())
            alpha2 = getattr(lang, "alpha_2", None)
            if alpha2 and len(alpha2) == 2:
                return alpha2.upper()
        except Exception:
            pass

    # 4) Fallback manual para nomes comuns (em inglês/português/espanhol)
    s_low = s.lower()
    if "english" in s_low:
        return "EN"
    if "portugu" in s_low:
        return "PT"
    if "spanish" in s_low or "espanol" in s_low or "espanhol" in s_low:
        return "ES"
    if "french" in s_low or "frances" in s_low:
        return "FR"
    if "german" in s_low or "alemao" in s_low or "aleman" in s_low:
        return "DE"
    if "italian" in s_low or "italiano" in s_low:
        return "IT"
    if "dutch" in s_low or "holandes" in s_low:
        return "NL"
    if "japanese" in s_low or "japones" in s_low:
        return "JA"
    if "chinese" in s_low or "mandarin" in s_low:
        return "ZH"
    if "korean" in s_low:
        return "KO"
    if "russian" in s_low:
        return "RU"

    # Se não conseguiu mapear com segurança, deixa em branco no WEBADI
    return None

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
        raise FileNotFoundError(f"WEBADI template not found: {template_path}")

    wb = load_workbook(
        str(template_path),
        keep_vba=True,
        data_only=False,
        keep_links=True,
    )

    if unprotect:
        if debug:
            print("[DEBUG] webadi_unprotect=True: unprotecting workbook/sheets.")
        unprotect_workbook_and_sheets(wb, debug=debug)
    else:
        if debug:
            print("[DEBUG] webadi_unprotect=False: keeping protections.")

    if sheet_name not in wb.sheetnames:
        raise RuntimeError(f"Sheet '{sheet_name}' not found in template.")
    ws = wb[sheet_name]

    try:
        d3 = ws["D3"].value
        if isinstance(d3, str) and d3.strip().lower().replace(" ", "") == "*text":
            today_str = datetime.now().strftime("%m%d%Y")
            ws["E3"].value = f"OC {today_str}"
            if debug:
                print(f"[DEBUG] E3 (BATCH_NAME) set to 'OC {today_str}'")
    except Exception as e:
        if debug:
            print(f"[DEBUG] Failed to update BATCH_NAME in E3: {e}")

    header_row, header_map = find_header_row(
        ws, start_at_row=header_start_row, debug=debug
    )

    wanted = {
        "HOTEL_ID": "Expedia ID",
        "CUSTOMER_NAME": "Hotel Name",
        "ADDRESS_LINE_1": "Address Line 1",
        "CITY": "City",
        "POSTAL_CODE": "Postal Code",
        "COUNTRY": "Country",
        "BILLING_CURRENCY": "Currency",
        "FIRST_NAME": "First Name",
        "LAST_NAME": "Last Name",
        "EMAIL_ADDRESS": "Email Address",
        "TAX_REG_NUMBER": "Tax Registration Number",
        "EFFECTIVE_DATE": "Effective Date of Change",
        "PREFERRED_LANGUAGE": "Preferred Language",
    }

    col_state = header_map.get("STATE")
    col_province = header_map.get("PROVINCE")
    col_bill_to = header_map.get("BILL_TO")

    if not col_bill_to:
        col_bill_to = column_index_from_string("Q")
        if debug:
            print(
                "[DEBUG] BILL_TO column not found by header; "
                "using fallback column Q."
            )

    col_site_purpose = header_map.get("SITE_PURPOSE")

    key_col = header_map.get("HOTEL_ID") or header_map.get("CUSTOMER_NAME")
    if not key_col:
        raise RuntimeError(
            "Could not find key columns (HOTEL_ID/CUSTOMER_NAME) in WebADI header."
        )

    if mode.lower() == "replace":
        last = ws.max_row
        for _ in range(header_row + 1, last + 1):
            ws.delete_rows(header_row + 1)
        base_row = header_row
        if debug:
            print(
                f"[DEBUG] Replace mode: cleared rows after header row {header_row}."
            )
    else:
        base_row = last_data_row(ws, header_row, key_col)
        if debug:
            print(
                f"[DEBUG] Append mode: last data row detected = {base_row}."
            )

    current = base_row
    for _, rec in df.iterrows():
        iso2 = to_iso2(rec.get("Country"))
        stateprov = rec.get("State/Province")

        current += 1
        src_row = base_row if base_row > header_row else header_row + 1
        clone_row(ws, src_row, current)

        # Sempre sobrescreve BILL_TO e SITE_PURPOSE
        if col_bill_to:
            ws.cell(current, col_bill_to).value = "BILL_TO"
        if col_site_purpose:
            ws.cell(current, col_site_purpose).value = "BILL_TO"

        # Zera sempre STATE/PROVINCE nas linhas novas para não herdar lixo
        if col_state:
            ws.cell(current, col_state).value = None
        if col_province:
            ws.cell(current, col_province).value = None

        for webadi_col, script_col in wanted.items():
            col_idx = header_map.get(webadi_col)
            if not col_idx:
                continue
            val = rec.get(script_col)

            # País (ISO2 ou nome). Se não houver valor, limpa a célula.
            if webadi_col == "COUNTRY":
                if _is_blank(val):
                    ws.cell(current, col_idx).value = None
                else:
                    v2 = to_iso2(val) or val
                    ws.cell(current, col_idx).value = v2

            # Moeda: 3 letras ou, se não tiver valor, limpa.
            elif webadi_col == "BILLING_CURRENCY":
                if _is_blank(val):
                    ws.cell(current, col_idx).value = None
                else:
                    if isinstance(val, str) and re.fullmatch(r"[A-Za-z]{3}", val.strip()):
                        ws.cell(current, col_idx).value = val.strip().upper()
                    else:
                        ws.cell(current, col_idx).value = val

            # Preferred Language: converte para código de 2 letras.
            elif webadi_col == "PREFERRED_LANGUAGE":
                if _is_blank(val):
                    ws.cell(current, col_idx).value = None
                else:
                    code = normalize_language_to_code(val)
                    ws.cell(current, col_idx).value = code if code else None

            # Demais campos (inclui TAX_REG_NUMBER, ADDRESS_LINE_1, etc.)
            else:
                if _is_blank(val):
                    ws.cell(current, col_idx).value = None
                else:
                    ws.cell(current, col_idx).value = val

        # Preencher State/Province no WebADI, após limpar as colunas acima
        if not _is_blank(stateprov) and isinstance(stateprov, str):
            if iso2 == "CA" and col_province:
                ws.cell(current, col_province).value = stateprov
            elif iso2 != "CA" and col_state:
                ws.cell(current, col_state).value = stateprov

    if not out_path:
        out_path = template_path

    wb.save(str(out_path))
    print(f"[OK] WEBADI filled: {out_path}")

# -------------------- ORACLE ENRICH / HELPER FUNCS --------------------

LEGAL_NAME_CLEAN_RE = re.compile(r"[^A-Za-z0-9 ]+")

def add_country_iso2_column(df: pd.DataFrame) -> pd.DataFrame:
    if "Country" not in df.columns:
        return df
    iso2_values = []
    for v in df["Country"]:
        try:
            iso2_values.append(to_iso2(v))
        except Exception:
            iso2_values.append(None)
    insert_pos = list(df.columns).index("Country") + 1
    df.insert(loc=insert_pos, column="Country ISO2", value=iso2_values)
    return df

def sanitize_legal_name_column(df: pd.DataFrame) -> pd.DataFrame:
    if "Legal Name" not in df.columns:
        return df

    def _clean_ln(x):
        if x in (None, "", " "):
            return x
        x = unidecode(str(x))
        x = LEGAL_NAME_CLEAN_RE.sub("", x)
        return x.strip()

    df["Legal Name"] = df["Legal Name"].astype(str).apply(_clean_ln)
    return df

def format_comment_from_filename(fn: str) -> str:
    if not fn:
        return ""
    base = os.path.splitext(str(fn))[0]

    m = re.search(r"(case[^-_]*)[-_ ]*(eid[^-_]*)", base, flags=re.IGNORECASE)
    if m:
        case_part = m.group(1).strip()
        eid_part = m.group(2).strip()
        return f"{case_part} - {eid_part}"

    m = re.search(r"(eid[^-_]*)[-_ ]*(case[^-_]*)", base, flags=re.IGNORECASE)
    if m:
        eid_part = m.group(1).strip()
        case_part = m.group(2).strip()
        return f"{case_part} - {eid_part}"

    return base.strip()

def fetch_oracle_tca(expedia_ids: List[str]) -> pd.DataFrame:
    expedia_ids = sorted({str(x).strip() for x in expedia_ids if pd.notna(x) and str(x).strip()})
    if not expedia_ids:
        return pd.DataFrame()

    ids_literal = ", ".join(f"'{eid}'" for eid in expedia_ids)
    sql = ORACLE_TCA_SQL_TEMPLATE.format(expedia_ids=ids_literal)

    print(f"\n[INFO] Running Oracle query for {len(expedia_ids)} Expedia ID(s)...")

    oracledb.init_oracle_client(lib_dir=ORACLE_CLIENT_DIR)

    with oracledb.connect(user=ORACLE_USERNAME, password=ORACLE_PASSWORD, dsn=ORACLE_DSN) as con:
        with con.cursor() as cur:
            cur.execute("ALTER SESSION SET CURRENT_SCHEMA = APPS")

        t0 = time.time()
        with con.cursor() as cur:
            cur.execute(sql)
            cols = [d[0] for d in cur.description]
            rows = cur.fetchall()
        elapsed = time.time() - t0

    print(f"[INFO] Oracle query completed. Rows: {len(rows)}. Time: {elapsed:.2f}s")

    df_oracle = pd.DataFrame(rows, columns=cols)

    if not df_oracle.empty:
        df_oracle["EXPEDIA_ID"] = df_oracle["EXPEDIA_ID"].astype(str).str.strip()
        df_oracle["COUNTRY_ABBR"] = df_oracle["COUNTRY_ABBR"].astype(str).str.strip()
        df_oracle = df_oracle.sort_values(
            ["EXPEDIA_ID", "COUNTRY_ABBR", "DUPLICATE_COUNT"],
            ascending=[True, True, True],
        )
        df_oracle = df_oracle.drop_duplicates(subset=["EXPEDIA_ID", "COUNTRY_ABBR"], keep="first")

    return df_oracle

def enrich_hcpif_with_oracle(
    df: pd.DataFrame, base_columns: List[str]
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    if "Expedia ID" not in df.columns:
        print("[WARN] Column 'Expedia ID' not found; skipping Oracle enrichment.")
        return df, pd.DataFrame(), pd.DataFrame()

    expedia_ids = df["Expedia ID"].tolist()
    df_oracle = fetch_oracle_tca(expedia_ids)
    if df_oracle.empty:
        print("[INFO] No Oracle data returned; HCPIF remains without enrichment.")
        return df, pd.DataFrame(), pd.DataFrame()

    df["Expedia ID"] = df["Expedia ID"].astype(str).str.strip()
    if "Country ISO2" not in df.columns:
        df["Country ISO2"] = df["Country"].map(lambda x: to_iso2(x) or "")

    df_oracle["EXPEDIA_ID"] = df_oracle["EXPEDIA_ID"].astype(str).str.strip()
    df_oracle["COUNTRY_ABBR"] = df_oracle["COUNTRY_ABBR"].astype(str).str.strip()

    merged = df.merge(
        df_oracle,
        how="left",
        left_on=["Expedia ID", "Country ISO2"],
        right_on=["EXPEDIA_ID", "COUNTRY_ABBR"],
        suffixes=("", "_ORACLE"),
    )

    merged["Found SLE OID"] = merged["ORACLE_ID"]
    merged["Found SLE Name"] = merged["SLE_NAME"]

    if "Hotel Name" in merged.columns:
        merged["Hotel Name"] = merged["Hotel Name"].where(
            merged["Hotel Name"].notna()
            & (merged["Hotel Name"].astype(str).str.strip() != ""),
            merged["HOTEL_NAME"],
        )
    else:
        merged["Hotel Name"] = merged["HOTEL_NAME"]

    merged["Oracle Currency"] = merged["ORACLE_CURRENCY"]
    if "Currency" in merged.columns:
        merged["Currency"] = merged["Currency"].astype(str).str.upper().str.strip()
        merged["Oracle Currency"] = merged["Oracle Currency"].astype(str).str.upper().str.strip()

        def _match(cur, oc):
            if not cur or not oc:
                return "NO"
            return "YES" if cur == oc else "NO"

        merged["Currency Matches Oracle"] = [
            _match(c, o) for c, o in zip(merged["Currency"], merged["Oracle Currency"])
        ]
    else:
        merged["Currency Matches Oracle"] = pd.NA

    comment_col = "Comments"
    new_comment = merged.get("file_name", pd.Series([""] * len(merged))).apply(
        lambda fn: f"Updated Ownership per HCPIF - {format_comment_from_filename(fn)}"
    )
    if comment_col in merged.columns:
        merged[comment_col] = merged[comment_col].fillna("").astype(str)
        merged[comment_col] = merged[comment_col].where(
            merged[comment_col].str.strip() != "",
            new_comment,
        )
    else:
        merged[comment_col] = new_comment

    oracle_key_cols = [c for c in ["Expedia ID", "Country ISO2", "file_name", "file_path"] if c in merged.columns]
    exclude_for_details = set(base_columns) | {
        "Found SLE OID",
        "Found SLE Name",
        "Oracle Currency",
        "Currency Matches Oracle",
        "Comments",
    }
    oracle_value_cols = [c for c in merged.columns if c not in exclude_for_details]
    mask_has_oracle = merged["ORACLE_ID"].notna() if "ORACLE_ID" in merged.columns else merged["Found SLE OID"].notna()
    df_oracle_details = merged.loc[mask_has_oracle, oracle_key_cols + oracle_value_cols].copy()

    series_rm = merged.get("RECEIPT_METHOD_NAME", pd.Series([""] * len(merged))).astype(str)
    mask_direct_debit = series_rm.str.contains("DIRECT DEBIT", case=False, na=False)

    series_tai = merged.get("TAI_BM", pd.Series([""] * len(merged))).astype(str)
    mask_tai_group = series_tai.str.contains("GROUP", case=False, na=False)

    mask_excluded = mask_direct_debit | mask_tai_group

    df_main_raw = merged.copy()
    df_excluded_raw = merged[mask_excluded].copy()

    if not df_excluded_raw.empty:
        print(
            f"[INFO] {len(df_excluded_raw)} row(s) flagged for review due to "
            "RECEIPT_METHOD_NAME DIRECT DEBIT and/or TAI_BM GROUP. "
            "They remain in the main Extraction sheet and WEBADI."
        )

    extra_cols = ["Found SLE OID", "Found SLE Name", "Oracle Currency", "Currency Matches Oracle", "Comments"]
    final_main_cols = list(base_columns)
    for col in extra_cols:
        if col in df_main_raw.columns and col not in final_main_cols:
            final_main_cols.append(col)

    df_main = df_main_raw.loc[:, final_main_cols].copy()
    df_excluded = df_excluded_raw.loc[:, final_main_cols].copy()

    return df_main, df_excluded, df_oracle_details

# -------------------- MAIN --------------------
def main():
    ap = argparse.ArgumentParser(
        description="Extract HCPIF fields into Excel (R2.6 + WEBADI injection + OCR fallback + Oracle enrich)."
    )
    ap.add_argument("--input_dir", type=str, default=str(DEFAULT_INPUT_DIR))
    ap.add_argument("--output", type=str, default=str(DEFAULT_OUTPUT_FILE))
    ap.add_argument(
        "--webadi_template",
        type=str,
        default=None,
        help="Path to WEBADI .xlsm/.xlsx template",
    )
    ap.add_argument(
        "--webadi_output",
        type=str,
        default=None,
        help="Path to save filled copy. "
        "If omitted, generates 'OC WEBADI MMDDYYYY.xlsm' in the same folder as the template.",
    )
    ap.add_argument(
        "--webadi_mode",
        type=str,
        choices=["append", "replace"],
        default="replace",
        help="append (clone last row); replace (clear data region before writing).",
    )
    ap.add_argument(
        "--webadi_sheet",
        type=str,
        default=WEBADI_SHEET_DEFAULT,
        help="WEBADI sheet name (default: WebADI).",
    )
    ap.add_argument(
        "--webadi_header_start_row",
        type=int,
        default=5,
        help="Row to start searching for header (default: 5).",
    )
    ap.add_argument(
        "--webadi_debug",
        action="store_true",
        help="Prints header/sheet diagnostics.",
    )
    ap.add_argument(
        "--webadi_unprotect",
        dest="webadi_unprotect",
        action="store_true",
        help="Unprotect workbook/sheets before writing (default: True).",
    )
    ap.add_argument(
        "--webadi_keep_protection",
        dest="webadi_unprotect",
        action="store_false",
        help="Do NOT unprotect workbook/sheets before writing.",
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
            print(f"[ERROR] Failed to process {pdf.name}: {e}")
            rows.append(
                {
                    "file_name": pdf.name,
                    "file_path": str(pdf.resolve()),
                    "status": f"error: {e}",
                }
            )

    df = pd.DataFrame(rows)

    ordered_cols = [f["col"] for f in FIELDS]
    for col in ordered_cols:
        if col not in df.columns:
            df[col] = pd.NA

    aux_cols = [c for c in df.columns if c not in ordered_cols]
    df = df[ordered_cols + aux_cols]

    def _strip_cell(v):
        if isinstance(v, str):
            return strip_leading_junk(v)
        return v

    df = df.applymap(_strip_cell)

    df = add_country_iso2_column(df)

    base_columns = list(df.columns)

    df, df_excluded, df_oracle_details = enrich_hcpif_with_oracle(df, base_columns)

    df = sanitize_legal_name_column(df)
    if not df_excluded.empty:
        df_excluded = sanitize_legal_name_column(df_excluded)

    try:
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Extraction", index=False)
            if not df_excluded.empty:
                df_excluded.to_excel(
                    writer, sheet_name="Excluded_DIRECTDEBIT_TAIBM", index=False
                )
            if not df_oracle_details.empty:
                df_oracle_details.to_excel(
                    writer, sheet_name="Oracle_Details", index=False
                )

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
            today_str = datetime.now().strftime("%m%d%Y")
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
            "Hotel Name",
            "Effective Date of Change",
        }

        present = [c for c in need_cols if c in df.columns]
        if not present:
            print("[WARN] No required columns for WEBADI are available.")
        else:
            df2 = df.loc[
                (df["Expedia ID"].notna()) | (df["Legal Name"].notna()),
                present,
            ]
            if not len(df2):
                print("[WARN] Nothing to insert into WEBADI (no Expedia ID / Legal Name).")
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
                    print(f"[ERROR] Failed to fill WEBADI: {ex}")
                    print(
                        "Suggestions: check sheet name, unprotect workbook (if needed), "
                        "and validate the header row."
                    )
    else:
        print(
            "[INFO] WEBADI step not executed (no --webadi_template and no default template found)."
        )

if __name__ == "__main__":
    main()
