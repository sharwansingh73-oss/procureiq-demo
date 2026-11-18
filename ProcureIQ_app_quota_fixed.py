import os

import re

import json

import io

import math

import hashlib
import time

from datetime import datetime, date, timedelta

import pandas as pd

import streamlit as st

# Optional document readers

try:

    import fitz  # PyMuPDF

except Exception:

    fitz = None

try:

    from docx import Document

except Exception:

    Document = None

try:

    from PIL import Image

except Exception:

    Image = None

try:

    import pytesseract

except Exception:

    pytesseract = None

try:

    from google import genai

except Exception:

    genai = None

# ---------------------------------------------------------------------

# ProcureIQ — self-contained Streamlit procurement workflow prototype

# ---------------------------------------------------------------------

st.set_page_config(

    page_title="ProcureIQ | Procurement workspace",

    page_icon="◈",

    layout="wide",

    initial_sidebar_state="expanded",

)

CSS = """

<style>

@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@500;600;700;800&display=swap');

:root { --ink:#17243b; --muted:#718096; --line:#e8edf4; --blue:#315efb; --navy:#111d33; --bg:#f5f7fb; }

html, body, [class*="css"] { font-family:'DM Sans',sans-serif; color:var(--ink); }

.stApp { background:var(--bg); }

.block-container { padding-top:1.5rem; padding-bottom:3rem; max-width:1500px; }

section[data-testid="stSidebar"] { background:var(--navy); border-right:1px solid #26344d; }

section[data-testid="stSidebar"] * { color:#e9efff !important; }

section[data-testid="stSidebar"] .stRadio label { padding:9px 12px; border-radius:9px; }

section[data-testid="stSidebar"] .stRadio label:hover { background:#22314c; }

section[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p { color:#9fb0cb !important; }

h1,h2,h3 { font-family:'Manrope',sans-serif !important; letter-spacing:-.035em; color:#17243b; }

h1 { font-size:2rem !important; font-weight:800 !important; }

h2 { font-size:1.35rem !important; font-weight:750 !important; }

h3 { font-size:1.02rem !important; font-weight:700 !important; }

p, li { color:#52627a; }

[data-testid="stMetric"] { background:#fff; border:1px solid var(--line); padding:17px 18px; border-radius:14px; box-shadow:0 3px 12px rgba(25,42,70,.035); }

[data-testid="stMetricLabel"] { color:#718096 !important; font-size:.82rem; }

[data-testid="stMetricValue"] { color:#17243b !important; font-family:'Manrope',sans-serif; font-weight:800; }

div[data-testid="stVerticalBlockBorderWrapper"] { background:#fff; border:1px solid var(--line); border-radius:15px; }

.stButton button, .stDownloadButton button { border-radius:9px; font-weight:650; min-height:2.55rem; }

.stButton button[kind="primary"] { background:var(--blue); border-color:var(--blue); }

.stTextInput input, .stTextArea textarea, .stSelectbox div[data-baseweb="select"], .stNumberInput input { border-radius:9px; }

[data-testid="stDataFrame"] { border:1px solid var(--line); border-radius:12px; overflow:hidden; }

.piq-brand { font-family:'Manrope',sans-serif; font-weight:800; font-size:1.25rem; letter-spacing:-.04em; color:white; }

.piq-brand span { color:#8ea7ff; }

.piq-eyebrow { text-transform:uppercase; font-size:.69rem; letter-spacing:.13em; color:#91a2be; font-weight:700; margin-top:20px; }

.piq-muted { color:#718096; font-size:.88rem; }

.piq-card { background:white; border:1px solid #e8edf4; border-radius:15px; padding:18px 20px; }

.piq-pill { display:inline-block; padding:4px 9px; border-radius:99px; background:#edf2ff; color:#315efb; font-size:.72rem; font-weight:700; }

.piq-good { color:#16835d; font-weight:700; }

.piq-warn { color:#b7791f; font-weight:700; }

hr { border-color:#e8edf4; }

</style>

"""

st.markdown(CSS, unsafe_allow_html=True)

# ---------- Configuration and state ----------

def get_secret(name, default=""):

    try:

        value = st.secrets.get(name, "")

        if value:

            return str(value)

    except Exception:

        pass

    return os.getenv(name, default)

def init_state():

    defaults = {

        "page": "Overview",

        "rfx": {

            "title": "Corrugated Packaging — Annual Sourcing",

            "category": "Corrugated packaging",

            "description": "Source corrugated cartons for multi-location fulfilment operations.",

            "currency": "INR",

            "deadline": date.today() + timedelta(days=14),

            "items": [

                {"item": f"Packaging item {i:02d}", "description": desc, "quantity": qty, "unit": unit}

                for i, desc, qty, unit in [

                    (1,"3-ply corrugated carton, 12 x 10 x 8 in",5000,"pieces"),

                    (2,"5-ply corrugated carton, 16 x 12 x 10 in",3500,"pieces"),

                    (3,"Printed shipping carton, 18 x 14 x 12 in",2500,"pieces"),

                    (4,"Corrugated roll, 1 m width",800,"rolls"),

                    (5,"Die-cut insert for fragile product",4000,"pieces"),

                ]

            ],

            "questions": ["Do you hold valid ISO 9001 certification?", "What is your standard lead time?", "Can you support monthly volume variation of ±20%?"],

            "terms": "Quote validity: 60 days. Delivery: to buyer-designated locations. Payment: net 30 days from accepted invoice.",

        },

        "vendors": [],

        "vendor_data": {},
        "extraction_cache": {},

        "chat_history": [],

        "fx": {"INR": 1.0, "USD": 84.0, "EUR": 91.0, "GBP": 108.0},

    }

    for k, v in defaults.items():

        if k not in st.session_state:

            st.session_state[k] = v

init_state()

# ---------- Utility functions ----------

def safe_float(value, default=0.0):

    try:

        if pd.isna(value):

            return default

        s = str(value).replace(",", "").replace("₹", "").replace("$", "").strip()

        m = re.search(r"-?\d+(?:\.\d+)?", s)

        return float(m.group()) if m else default

    except Exception:

        return default

def extract_text(upload):

    name = upload.name.lower()

    raw = upload.getvalue()

    if name.endswith(".pdf"):

        if not fitz:

            return "", "Install pymupdf to read PDFs."

        doc = fitz.open(stream=raw, filetype="pdf")

        return "\n".join(p.get_text() for p in doc), "PDF text extraction"

    if name.endswith(".docx"):

        if not Document:

            return "", "Install python-docx to read Word files."

        doc = Document(io.BytesIO(raw))

        return "\n".join(p.text for p in doc.paragraphs), "Word text extraction"

    if name.endswith((".xlsx", ".xls")):

        xls = pd.ExcelFile(io.BytesIO(raw))

        chunks = []

        for sheet in xls.sheet_names:

            df = pd.read_excel(xls, sheet_name=sheet)

            chunks.append(f"SHEET: {sheet}\n" + df.to_csv(index=False))

        return "\n\n".join(chunks), "Excel extraction"

    if name.endswith(".csv"):

        return pd.read_csv(io.BytesIO(raw)).to_csv(index=False), "CSV extraction"

    if name.endswith((".txt", ".eml")):

        return raw.decode("utf-8", errors="ignore"), "Text extraction"

    if name.endswith((".png", ".jpg", ".jpeg", ".webp")):

        if not Image or not pytesseract:

            return "", "Image OCR needs Pillow, pytesseract and the Tesseract application installed."

        image = Image.open(io.BytesIO(raw))

        return pytesseract.image_to_string(image), "Image OCR"

    return "", "Unsupported file type. Use PDF, DOCX, XLSX, CSV, TXT, EML, PNG or JPG."

def gemini_client():

    key = get_secret("GEMINI_API_KEY")

    if not key:

        st.error("Debug: Gemini API key is not being loaded.")

        return None

    if genai is None:

        st.error("Debug: Gemini library is not loaded.")

        return None

    try:

        return genai.Client(api_key=key)

    except Exception as e:

        st.error(f"Gemini client error: {e}")

        return None

class GeminiQuotaError(RuntimeError):
    """Raised when Gemini free-tier or paid quota is exhausted."""


def call_gemini(prompt, model="gemini-3.8-flash"):
    client = gemini_client()
    if not client:
        raise RuntimeError("Gemini is not configured. Check GEMINI_API_KEY in .streamlit/secrets.toml.")

    try:
        response = client.models.generate_content(model=model, contents=prompt)
        return response.text or ""
    except Exception as exc:
        message = str(exc)
        if "429" in message or "RESOURCE_EXHAUSTED" in message or "quota" in message.lower():
            raise GeminiQuotaError(
                "Gemini API quota is exhausted. You can continue using Demo Analyst mode "
                "without Gemini. The API quota will reset according to your Google AI Studio limits."
            ) from exc
        raise

def parse_json_response(text):

    cleaned = re.sub(r"^```(?:json)?|```$", "", text.strip(), flags=re.I | re.M).strip()

    start, end = cleaned.find("{"), cleaned.rfind("}")

    if start >= 0 and end > start:

        cleaned = cleaned[start:end+1]

    return json.loads(cleaned)

def parse_quote_with_ai(text, vendor_name, filename):

    prompt = f"""

You are a careful procurement quote extraction engine. Extract only facts explicitly present.

Return valid JSON only with this schema:

{{

 "vendor_name":"{vendor_name}",

 "currency":"INR",

 "quote_validity":"",

 "lead_time":"",

 "payment_terms":"",

 "questionnaire":{{}},

 "items":[{{"item_name":"","description":"","quantity":0,"unit":"pieces","unit_price":0,"line_total":0,"currency":"INR","quoted_unit_basis":"per piece","source_evidence":"","confidence":"high"}}],

 "notes":[]

}}

Rules: preserve original quoted unit basis (e.g. per box, per 100 pieces); do not assume a missing item or price. If quantity/unit price are absent, use null. Include source_evidence short text. Document: {filename}

Vendor label: {vendor_name}

QUOTE TEXT:

{text[:28000]}

"""

    return parse_json_response(call_gemini(prompt))

def basic_parse_quote(text, vendor_name):

    """Fallback parser: detects simple CSV/Excel-like rows; mark confidence low."""

    lines = [x.strip() for x in text.splitlines() if x.strip()]

    items = []

    for line in lines:

        parts = re.split(r"\t|,|;|\|", line)

        if len(parts) < 3:

            continue

        nums = [safe_float(p, None) for p in parts]

        numeric = [n for n in nums if n is not None]

        if len(numeric) >= 1:

            items.append({

                "item_name": parts[0][:100],

                "description": parts[1] if len(parts) > 1 else "",

                "quantity": numeric[0] if len(numeric) > 1 else None,

                "unit": "pieces",

                "unit_price": numeric[-1],

                "line_total": None,

                "currency": "INR",

                "quoted_unit_basis": "per piece (assumed)",

                "source_evidence": line[:240],

                "confidence": "low",

            })

    return {"vendor_name": vendor_name, "currency": "INR", "items": items, "questionnaire": {}, "notes": ["Fallback text parser used; validate all fields."]}

def normalise_unit_price(price, basis, currency):

    """Return INR price per individual piece where the basis is understood."""

    if price is None:

        return None, "Price missing"

    fx = st.session_state.fx.get(str(currency).upper(), None)

    if fx is None:

        return None, f"Unknown currency conversion: {currency}"

    b = (basis or "per piece").lower().replace(" ", "")

    divisor = 1

    if "per100" in b or "/100" in b or "100pieces" in b or "perhundred" in b:

        divisor = 100

    elif "perbox" in b or "/box" in b:

        return None, "Box quantity required to normalise"

    elif "per1000" in b or "/1000" in b:

        divisor = 1000

    elif "perpiece" not in b and "each" not in b and "piece" not in b:

        return None, f"Unrecognised price basis: {basis}"

    return float(price) * float(fx) / divisor, "Normalised to INR / piece"

def get_comparison_rows():

    rows = []

    for vendor, data in st.session_state.vendor_data.items():

        for item in data.get("items", []):

            price = item.get("unit_price")

            basis = item.get("quoted_unit_basis") or item.get("unit") or "per piece"

            currency = item.get("currency") or data.get("currency") or "INR"

            norm, status = normalise_unit_price(safe_float(price, None), basis, currency)

            rows.append({

                "Vendor": vendor,

                "Quoted item": item.get("item_name") or item.get("description") or "Unlabelled item",

                "Description": item.get("description", ""),

                "Qty": item.get("quantity"),

                "Unit": item.get("unit", ""),

                "Quoted price": price,

                "Currency": currency,

                "Price basis": basis,

                "INR / piece": round(norm, 4) if norm is not None else None,

                "Normalisation": status,

                "Confidence": item.get("confidence", "unknown"),

                "Evidence": item.get("source_evidence", ""),

            })

    return pd.DataFrame(rows)

def vendor_summary():

    result = []

    for name, data in st.session_state.vendor_data.items():

        items = data.get("items", [])

        result.append({

            "Vendor": name,

            "Lines extracted": len(items),

            "Questionnaire answers": len(data.get("questionnaire", {})),

            "Currency": data.get("currency", "—"),

            "Lead time": data.get("lead_time", "Not provided"),

            "File": data.get("_filename", ""),

        })

    return pd.DataFrame(result)

def download_csv(df, filename):

    return df.to_csv(index=False).encode("utf-8-sig")

def demo_analyst_answer(question):
    """Answer common procurement questions from extracted data without calling Gemini."""
    df = get_comparison_rows()
    q = (question or "").strip().lower()
    vendors = st.session_state.vendor_data
    expected = len(st.session_state.rfx.get("items", []))

    if df.empty:
        return "No extracted quote lines are available yet. Upload vendor responses first."

    lines = [
        "**Demo Analyst — local calculation mode**",
        "",
        "This answer uses only the quote data currently stored in this app. It does not call Gemini.",
        "",
    ]

    if any(word in q for word in ["missing", "coverage", "27/30", "line items"]):
        coverage = []
        for name, data in vendors.items():
            count = len(data.get("items", []))
            coverage.append({"Vendor": name, "Extracted lines": count,
                             "Expected RFx lines": expected, "Count-based gap": max(0, expected-count)})
        lines += [
            "### Vendor quote coverage",
            pd.DataFrame(coverage).to_string(index=False) if coverage else "No vendors found.",
            "",
            "Coverage is calculated by counting extracted rows. It does **not** prove that each RFx item ID has been matched; verify line-item identity before award.",
        ]
        return "\n".join(lines)

    if any(word in q for word in ["exception", "normalis", "unit basis", "currency"]):
        exceptions = df[df["INR / piece"].isna()]
        lines.append("### Normalisation exceptions")
        if exceptions.empty:
            lines.append("No exceptions were detected in the extracted rows.")
        else:
            cols = ["Vendor", "Quoted item", "Quoted price", "Currency", "Price basis", "Normalisation", "Evidence"]
            lines.append(exceptions[cols].to_string(index=False))
        lines += ["", "Per-box pricing is not converted unless pack size is known. FX rates are the editable assumptions shown on the Bid Comparison page."]
        return "\n".join(lines)

    if any(word in q for word in ["lead time", "leadtime", "payment term", "payment terms", "delivery"]):
        summary = []
        for name, data in vendors.items():
            summary.append({
                "Vendor": name,
                "Lead time": data.get("lead_time") or "Not provided",
                "Payment terms": data.get("payment_terms") or "Not provided",
                "Quote validity": data.get("quote_validity") or "Not provided",
            })
        lines += ["### Vendor commercial terms", pd.DataFrame(summary).to_string(index=False)]
        return "\n".join(lines)

    normalised = df[df["INR / piece"].notna()].copy()
    # If the question names an item, narrow the analysis to matching extracted labels/descriptions.
    item_match = re.search(r"\b(?:CP[- ]?\d{1,2}|item\s*\d{1,2})\b", question, flags=re.I)
    if item_match:
        token = item_match.group(0).lower().replace(" ", "")
        candidate = normalised[
            normalised["Quoted item"].astype(str).str.lower().str.replace(" ", "", regex=False).str.contains(token, na=False)
            | normalised["Description"].astype(str).str.lower().str.replace(" ", "", regex=False).str.contains(token, na=False)
        ]
        if not candidate.empty:
            normalised = candidate

    if any(word in q for word in ["lowest", "cheapest", "price", "compare", "best value", "supplier"]):
        if normalised.empty:
            lines.append("There are no rows with a recognised, normalised INR-per-piece price.")
        else:
            lowest = normalised.sort_values("INR / piece").groupby("Quoted item", as_index=False).first()
            lines += [
                "### Lowest extracted comparable price by quoted item",
                lowest[["Quoted item", "Vendor", "INR / piece", "Confidence", "Price basis"]].to_string(index=False),
                "",
                "These are price-only signals based on rows that could be normalised to INR per piece. This is **not an award recommendation**. Check item matching, evidence, quality, delivery, taxes, freight and contractual terms.",
            ]
        return "\n".join(lines)

    summary = []
    for name, data in vendors.items():
        summary.append({
            "Vendor": name,
            "Extracted lines": len(data.get("items", [])),
            "Currency": data.get("currency", "Not provided"),
            "Lead time": data.get("lead_time") or "Not provided",
            "Payment terms": data.get("payment_terms") or "Not provided",
        })
    lines += [
        "### Supplier response summary",
        pd.DataFrame(summary).to_string(index=False),
        "",
        "Try asking about lowest comparable prices, missing lines, normalisation exceptions, lead times or payment terms.",
    ]
    return "\n".join(lines)


# ---------- Sidebar ----------

with st.sidebar:

    st.markdown('<div class="piq-brand">◈ Procure<span>IQ</span></div>', unsafe_allow_html=True)

    st.markdown('<div style="color:#91a2be;font-size:.75rem;margin-top:4px">INTELLIGENT SOURCING WORKSPACE</div>', unsafe_allow_html=True)

    st.markdown('<div class="piq-eyebrow">WORKSPACE</div>', unsafe_allow_html=True)

    pages = ["Overview", "RFx Builder", "Vendor Responses", "Bid Comparison", "AI Analyst"]

    selected = st.radio("Navigation", pages, label_visibility="collapsed",

                        index=pages.index(st.session_state.page) if st.session_state.page in pages else 0)

    st.session_state.page = selected

    st.markdown("---")

    st.markdown('<div class="piq-eyebrow">CURRENT EVENT</div>', unsafe_allow_html=True)

    st.markdown(f"**{st.session_state.rfx['title']}**")

    st.caption(f"{len(st.session_state.rfx['items'])} line items · {len(st.session_state.vendor_data)} vendor responses")

    st.markdown("---")

    st.caption("Prototype workspace · Demo data should be validated before any sourcing decision.")

# ---------- Header helpers ----------

def page_header(title, subtitle, tag="PROCUREMENT WORKSPACE"):

    c1, c2 = st.columns([5, 1])

    with c1:

        st.caption(tag)

        st.title(title)

        st.markdown(f'<div class="piq-muted">{subtitle}</div>', unsafe_allow_html=True)

    with c2:

        st.markdown("<div style='text-align:right;padding-top:12px;color:#718096;font-size:.8rem'>● Workspace active</div>", unsafe_allow_html=True)

    st.write("")

def section_title(title, hint=None):

    a, b = st.columns([4, 2])

    with a:

        st.subheader(title)

    if hint:

        with b:

            st.markdown(f"<div style='text-align:right;color:#8793a7;font-size:.8rem;padding-top:9px'>{hint}</div>", unsafe_allow_html=True)

# ---------- Overview ----------

if st.session_state.page == "Overview":

    page_header("Good afternoon, procurement team", "Your sourcing activity, supplier responses and commercial signals at a glance.")

    vendors = len(st.session_state.vendor_data)

    all_items = sum(len(x.get("items", [])) for x in st.session_state.vendor_data.values())

    cols = st.columns(4)

    cols[0].metric("Active RFx", "01", "Corrugated packaging")

    cols[1].metric("Invited vendors", "05", "Target supplier panel")

    cols[2].metric("Responses received", vendors, f"{max(0,5-vendors)} pending")

    cols[3].metric("Quote lines extracted", all_items, "Across uploaded responses")

    st.write("")

    left, right = st.columns([1.55, 1])

    with left:

        with st.container(border=True):

            section_title("Sourcing event", "CURRENT EVENT")

            st.markdown(f"### {st.session_state.rfx['title']}")

            st.markdown(f"<span class='piq-pill'>{st.session_state.rfx['category']}</span>", unsafe_allow_html=True)

            st.write(st.session_state.rfx["description"])

            m1, m2, m3 = st.columns(3)

            m1.metric("Line items", len(st.session_state.rfx["items"]))

            m2.metric("Buyer currency", st.session_state.rfx["currency"])

            m3.metric("Submission deadline", st.session_state.rfx["deadline"].strftime("%d %b %Y"))

            st.progress(min(vendors / 5, 1.0), text=f"{vendors} of 5 vendor responses uploaded")

            if st.button("Open RFx builder", type="primary"):

                st.session_state.page = "RFx Builder"

                st.rerun()

    with right:

        with st.container(border=True):

            section_title("Workflow health", "LIVE STATUS")

            stages = [

                ("RFx configuration", True),

                ("Vendor invitations", True),

                ("Quote intake", vendors > 0),

                ("Normalisation", vendors > 0),

                ("Award analysis", False),

            ]

            for label, done in stages:

                st.markdown(f"{'🟢' if done else '⚪'}  **{label}**")

            st.caption("Stages update as you add and process vendor responses.")

    st.write("")

    section_title("Recent vendor activity", "UPLOADS & EXTRACTION")

    if vendors:

        st.dataframe(vendor_summary(), use_container_width=True, hide_index=True)

    else:

        st.info("No vendor responses uploaded yet. Go to Vendor Responses to add Excel, PDF, Word, CSV, email text or image quotes.")

    with st.container(border=True):

        st.markdown("**Recommended next action**")

        st.write("Upload vendor quotations, review extracted evidence and resolve any unit or currency warnings before comparing commercial offers.")

        if st.button("Go to vendor responses", type="primary"):

            st.session_state.page = "Vendor Responses"

            st.rerun()

# ---------- RFx Builder ----------

elif st.session_state.page == "RFx Builder":

    page_header("RFx builder", "Define the sourcing event, line items, supplier questionnaire and commercial terms.", "SOURCE-TO-CONTRACT · RFx")

    rfx = st.session_state.rfx

    with st.container(border=True):

        st.subheader("Event details")

        c1, c2 = st.columns(2)

        rfx["title"] = c1.text_input("RFx title", rfx["title"])

        rfx["category"] = c2.text_input("Category", rfx["category"])

        rfx["description"] = st.text_area("Scope and business context", rfx["description"], height=90)

        c3, c4 = st.columns(2)

        rfx["currency"] = c3.selectbox("Target currency", ["INR", "USD", "EUR", "GBP"], index=["INR","USD","EUR","GBP"].index(rfx["currency"]))

        rfx["deadline"] = c4.date_input("Submission deadline", rfx["deadline"])

    with st.container(border=True):

        section_title("Line items", f"{len(rfx['items'])} configured")

        items_df = pd.DataFrame(rfx["items"])

        edited = st.data_editor(items_df, num_rows="dynamic", use_container_width=True, key="rfx_items_editor",

                                column_config={"quantity": st.column_config.NumberColumn(min_value=0, step=1)})

        if st.button("Save line items"):

            st.session_state.rfx["items"] = edited.fillna("").to_dict("records")

            st.success("Line items saved.")

        with st.expander("Add 30-line corrugated packaging starter list"):

            if st.button("Load 30 sample line items"):

                st.session_state.rfx["items"] = [

                    {"item": f"CP-{i:02d}", "description": f"Corrugated packaging specification {i:02d}", "quantity": 1000 + i*100, "unit": "pieces"}

                    for i in range(1, 31)

                ]

                st.rerun()

    with st.container(border=True):

        st.subheader("Supplier questionnaire")

        q_text = st.text_area("One question per line", "\n".join(rfx["questions"]), height=120)

        rfx["questions"] = [q.strip() for q in q_text.splitlines() if q.strip()]

        st.subheader("Commercial terms")

        rfx["terms"] = st.text_area("Terms and conditions", rfx["terms"], height=110)

    with st.container(border=True):

        st.subheader("AI-assisted RFx drafting")

        st.caption("Use Gemini to improve the scope, questionnaire and terms. Review the generated content before sharing.")

        if st.button("Draft RFx with AI", type="primary"):

            prompt = f"""Act as a senior procurement manager. Improve this RFx for {rfx['category']}.

Title: {rfx['title']}

Scope: {rfx['description']}

Line items: {json.dumps(rfx['items'][:30], ensure_ascii=False)}

Existing questions: {rfx['questions']}

Terms: {rfx['terms']}

Return concise sections: improved_scope, recommended_questionnaire (list), recommended_terms, evaluation_criteria (list). Do not invent legal commitments."""

            try:

                answer = parse_json_response(call_gemini(prompt))

                st.session_state["rfx_ai_draft"] = answer

            except Exception as e:

                st.error(str(e))

        if "rfx_ai_draft" in st.session_state:

            st.json(st.session_state["rfx_ai_draft"])

    st.download_button("Download RFx configuration (JSON)", json.dumps(rfx, indent=2, default=str), "procureiq_rfx.json", "application/json")

# ---------- Vendor Responses ----------

elif st.session_state.page == "Vendor Responses":

    page_header("Vendor responses", "Bring supplier quotations into one workspace. Extracted values remain traceable to source evidence.", "SUPPLIER COLLABORATION · INTAKE")

    with st.container(border=True):

        st.subheader("Upload supplier quote")

        st.caption("Supported formats: XLSX, CSV, PDF, DOCX, TXT, EML, PNG and JPG. For images, install the Tesseract OCR application separately.")

        uploads = st.file_uploader("Choose vendor response files", type=["xlsx","xls","csv","pdf","docx","txt","eml","png","jpg","jpeg","webp"], accept_multiple_files=True)

        vendor_names = st.text_input("Vendor names (optional; separate in same order with commas)", placeholder="NorthStar, Apex, GreenFold, MetroBox, Reliable")

        if st.button("Extract uploaded responses", type="primary", disabled=not uploads):

            names = [n.strip() for n in vendor_names.split(",")] if vendor_names.strip() else []

            progress = st.progress(0)

            for i, upload in enumerate(uploads):
                fallback_name = os.path.splitext(upload.name)[0].replace("_", " ").replace("-", " ")
                vendor = names[i] if i < len(names) and names[i] else fallback_name
                raw_bytes = upload.getvalue()
                cache_key = hashlib.sha256(
                    raw_bytes + vendor.encode("utf-8", errors="ignore")
                ).hexdigest()

                cached = st.session_state.extraction_cache.get(cache_key)
                if cached:
                    parsed = json.loads(json.dumps(cached["data"], default=str))
                    st.session_state.vendor_data[vendor] = parsed
                    st.info(f"{vendor}: reused cached extraction — no Gemini request made.")
                    progress.progress((i + 1) / len(uploads))
                    continue

                text, method = extract_text(upload)
                if not text.strip():
                    parsed = {
                        "vendor_name": vendor, "items": [], "questionnaire": {},
                        "notes": [method], "_filename": upload.name, "_method": method
                    }
                    st.session_state.vendor_data[vendor] = parsed
                    st.warning(f"{upload.name}: {method}")
                    st.session_state.extraction_cache[cache_key] = {"data": parsed}
                    progress.progress((i + 1) / len(uploads))
                    continue

                try:
                    parsed = parse_quote_with_ai(text, vendor, upload.name)
                    parsed["_filename"] = upload.name
                    parsed["_method"] = method + " + Gemini structuring"
                    st.session_state.vendor_data[vendor] = parsed
                    st.session_state.extraction_cache[cache_key] = {"data": parsed}
                    st.success(f"{vendor}: extracted {len(parsed.get('items', []))} quote lines.")
                except Exception as e:
                    parsed = basic_parse_quote(text, vendor)
                    parsed["_filename"] = upload.name
                    parsed["_method"] = method + " + fallback parser"
                    parsed.setdefault("notes", []).append(str(e))
                    st.session_state.vendor_data[vendor] = parsed
                    st.session_state.extraction_cache[cache_key] = {"data": parsed}
                    st.warning(
                        f"{vendor}: AI extraction unavailable; fallback extraction used. "
                        f"Review every field. Details: {e}"
                    )
                progress.progress((i + 1) / len(uploads))
            st.rerun()

    if st.session_state.vendor_data:

        section_title("Received responses", f"{len(st.session_state.vendor_data)} supplier(s)")
        if st.button("Clear extraction cache", help="Removes saved extraction results from this browser session. Uploading the files again will run extraction again."):
            st.session_state.extraction_cache = {}
            st.success("Extraction cache cleared.")

        st.dataframe(vendor_summary(), use_container_width=True, hide_index=True)

        for vendor, data in st.session_state.vendor_data.items():

            with st.expander(f"{vendor} · {len(data.get('items', []))} quote lines"):

                st.caption(f"Source: {data.get('_filename','')} · Method: {data.get('_method','')}")

                st.json(data)

                if st.button(f"Remove {vendor}", key=f"remove_{vendor}"):

                    del st.session_state.vendor_data[vendor]

                    st.rerun()

        st.download_button("Download all extracted responses (JSON)",

                           json.dumps(st.session_state.vendor_data, indent=2, ensure_ascii=False, default=str),

                           "vendor_responses.json", "application/json")

    else:

        st.info("Your supplier response register is empty. Upload the five sample vendor quote workbooks to test the workflow.")

# ---------- Bid Comparison ----------

elif st.session_state.page == "Bid Comparison":

    page_header("Bid comparison", "Compare supplier offers on a like-for-like basis and investigate every normalisation exception.", "COMMERCIAL EVALUATION")

    st.info("Currency rates below are editable prototype assumptions. Confirm live FX rates and contractual pricing rules before making a real award.")

    with st.container(border=True):

        st.subheader("Currency conversion assumptions")

        fx_cols = st.columns(4)

        for col, cur in zip(fx_cols, ["USD", "EUR", "GBP", "INR"]):

            st.session_state.fx[cur] = col.number_input(f"{cur} → INR", min_value=0.0001, value=float(st.session_state.fx[cur]), step=0.5, format="%.4f")

    df = get_comparison_rows()

    if df.empty:

        st.warning("No extracted vendor quote lines yet. Upload supplier responses first.")

    else:

        normalised = df[df["INR / piece"].notna()].copy()

        exceptions = df[df["INR / piece"].isna()].copy()

        c1, c2, c3 = st.columns(3)

        c1.metric("Vendors compared", df["Vendor"].nunique())

        c2.metric("Extracted quote lines", len(df))

        c3.metric("Needs review", len(exceptions))

        with st.container(border=True):

            section_title("Normalised quote lines", "INR PER PIECE · WHERE BASIS IS UNDERSTOOD")

            st.dataframe(normalised, use_container_width=True, hide_index=True)

            st.download_button("Export normalised comparison CSV", download_csv(normalised, "comparison.csv"), "procureiq_comparison.csv", "text/csv")

        with st.container(border=True):

            section_title("Exceptions requiring buyer review", "DO NOT AUTO-COMPARE")

            if exceptions.empty:

                st.success("No unit or currency normalisation exceptions detected.")

            else:

                st.dataframe(exceptions[["Vendor","Quoted item","Quoted price","Currency","Price basis","Normalisation","Evidence"]], use_container_width=True, hide_index=True)

                st.warning("Per-box pricing needs the pack size. Unknown currencies and ambiguous bases must be resolved before comparing.")

        with st.container(border=True):

            section_title("Supplier coverage", "MISSING-LINE CHECK")

            coverage = []

            expected = st.session_state.rfx["items"]

            for vendor, data in st.session_state.vendor_data.items():

                got = len(data.get("items", []))

                coverage.append({"Vendor": vendor, "Lines found": got, "Expected RFx lines": len(expected), "Gap": max(0, len(expected)-got)})

            if coverage:

                st.dataframe(pd.DataFrame(coverage), use_container_width=True, hide_index=True)

        with st.container(border=True):

            section_title("Indicative lowest comparable price", "NOT AN AWARD RECOMMENDATION")

            if not normalised.empty:

                cheapest = normalised.sort_values("INR / piece").groupby("Quoted item", as_index=False).first()

                st.dataframe(cheapest[["Quoted item","Vendor","INR / piece","Confidence","Evidence"]], use_container_width=True, hide_index=True)

                st.caption("This is a price-only signal, not a recommendation. Quality, delivery, questionnaire compliance, taxes, freight and contractual terms are not yet scored.")

# ---------- AI Analyst ----------

elif st.session_state.page == "AI Analyst":

    page_header("AI procurement analyst", "Ask questions about supplier responses. Demo mode works without Gemini API calls.", "DECISION SUPPORT · ANALYST")

    if not st.session_state.vendor_data:
        st.warning("Upload and extract at least one vendor response before asking data questions.")
    else:
        examples = [
            "Which supplier has the lowest comparable unit price?",
            "Which vendors have missing quote lines?",
            "Show all unit-basis or currency normalisation exceptions.",
            "Summarise lead time and payment terms by vendor.",
        ]

        with st.container(border=True):
            st.markdown("**Suggested questions**")
            ec = st.columns(2)
            for i, q in enumerate(examples):
                if ec[i % 2].button(q, key=f"suggest_{i}", use_container_width=True):
                    st.session_state["analyst_question"] = q

            analyst_mode = st.radio(
                "Analysis mode",
                ["Demo Analyst (no Gemini API)", "Gemini AI (uses API quota)"],
                horizontal=True,
                help="Demo mode calculates answers locally from extracted data. Gemini mode sends the data and question to Gemini."
            )
            question = st.text_area(
                "Ask your procurement question",
                key="analyst_question",
                height=90,
                placeholder="e.g. Which vendor offers the lowest comparable price for CP-01?"
            )

            if st.button("Analyse responses", type="primary", disabled=not question.strip()):
                if analyst_mode == "Demo Analyst (no Gemini API)":
                    answer = demo_analyst_answer(question)
                    st.session_state["last_analyst_answer"] = answer
                    st.session_state.chat_history.append({"question": question, "answer": answer})
                else:
                    df = get_comparison_rows()
                    context = {
                        "rfx": st.session_state.rfx,
                        "vendors": st.session_state.vendor_data,
                        "comparison_rows": df.where(pd.notna(df), None).to_dict("records"),
                        "currency_rates_to_INR": st.session_state.fx,
                    }
                    prompt = f"""You are ProcureIQ's procurement analyst. Answer ONLY from the supplied procurement data.
Do not invent missing facts. Clearly distinguish extracted facts, calculations, assumptions and unknowns.
For any price comparison, mention currency/unit basis and whether it was normalised.
Do not recommend an award based only on price. Mention missing evidence and confidence.
Use concise headings and a table when useful.

QUESTION: {question}

DATA:
{json.dumps(context, ensure_ascii=False, default=str)[:45000]}
"""
                    try:
                        answer = call_gemini(prompt)
                        st.session_state["last_analyst_answer"] = answer
                        st.session_state.chat_history.append({"question": question, "answer": answer})
                    except GeminiQuotaError as e:
                        st.warning(f"{e} Showing a local Demo Analyst answer instead.")
                        answer = demo_analyst_answer(question)
                        st.session_state["last_analyst_answer"] = answer
                        st.session_state.chat_history.append({"question": question, "answer": answer})
                    except Exception as e:
                        st.warning(f"Gemini could not run ({e}). Showing a local Demo Analyst answer instead.")
                        answer = demo_analyst_answer(question)
                        st.session_state["last_analyst_answer"] = answer
                        st.session_state.chat_history.append({"question": question, "answer": answer})

        if st.session_state.get("last_analyst_answer"):
            with st.container(border=True):
                st.markdown("### Analysis")
                st.markdown(st.session_state["last_analyst_answer"])

        if st.session_state.chat_history:
            with st.expander("Previous analysis"):
                for entry in reversed(st.session_state.chat_history[-10:]):
                    st.markdown(f"**You:** {entry['question']}")
                    st.markdown(entry["answer"])
                    st.divider()

# ---------- Footer ----------

st.markdown("---")

st.markdown(

    "<div style='display:flex;justify-content:space-between;color:#91a0b5;font-size:.75rem'>"

    "<span>◈ ProcureIQ · Intelligent sourcing workspace</span>"

    "<span>Prototype only · Validate extracted data before commercial decisions</span></div>",

    unsafe_allow_html=True,

)
