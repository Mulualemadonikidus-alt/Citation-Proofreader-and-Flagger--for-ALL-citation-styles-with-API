import streamlit as st
import re
import requests
import pandas as pd
import json
from docx import Document
from pypdf import PdfReader

# Page configurations
st.set_page_config(page_title="CiteShield AI Pro", page_icon="🛡️", layout="wide")

st.title("🛡️ CiteShield AI Pro: Multi-Style Citation Auditor")
st.markdown("Verify reference authenticity and context across Vancouver, APA, and AMA styles using Crossref and Google Gemini.")

# ─── SIDEBAR CONFIGURATION ───
st.sidebar.header("Configuration")
gemini_api_key = st.sidebar.text_input("Gemini API Key (Required for Context Check)", type="password")
if not gemini_api_key:
    st.sidebar.warning("⚠️ Provide your Gemini API key from Google AI Studio to unlock Section 4 (Contextual Fit Analysis).")

# ─── TEXT EXTRACTION LAYER ───
@st.cache_data(ttl=3600, max_entries=100)
def extract_text_and_paragraphs(file):
    """Extracts raw text blocks and a structured array of clean paragraphs."""
    paragraphs = []
    full_text = ""
    
    if file.name.endswith('.docx'):
        doc = Document(file)
        for p in doc.paragraphs:
            text = p.text.strip()
            if text:
                # SANITIZATION: Clean AI line markers, source tags, and page annotations
                text = re.sub(r"^\\s*", "", text, flags=re.IGNORECASE)
                text = re.sub(r"^---\s*PAGE\s*\d+\s*---", "", text, flags=re.IGNORECASE)
                if text.strip():
                    paragraphs.append(text.strip())
        full_text = "\n".join(paragraphs)
        
    elif file.name.endswith('.pdf'):
        reader = PdfReader(file)
        for page in reader.pages:
            text = page.extract_text()
            if text:
                lines = [line.strip() for line in text.split('\n') if line.strip()]
                clean_lines = []
                for line in lines:
                    # SANITIZATION: Clean AI line markers, source tags, and page annotations
                    line = re.sub(r"^\\s*", "", line, flags=re.IGNORECASE)
                    line = re.sub(r"^---\s*PAGE\s*\d+\s*---", "", line, flags=re.IGNORECASE)
                    if line.strip():
                        clean_lines.append(line.strip())
                if clean_lines:
                    paragraphs.extend(clean_lines)
                    
        full_text = "\n".join(paragraphs)
        
    return full_text, paragraphs

# ─── SEQUENCE-VALIDATED BIBLIOGRAPHY PARSER ───
def parse_bibliography(full_text):
    """Isolates the reference chunk and groups items using an active sequence verification filter."""
    split_keywords = [r"\nreferences\n", r"\nbibliography\n", r"\nliterature cited\n", r"\nreference list\n"]
    bib_text = ""
    body_text = full_text
    
    for kw in split_keywords:
        match = re.search(kw, full_text, re.IGNORECASE)
        if match:
            split_idx = match.start()
            body_text = full_text[:split_idx]
            bib_text = full_text[split_idx:]
            break
            
    if not bib_text:
        split_idx = int(len(full_text) * 0.75)
        body_text = full_text[:split_idx]
        bib_text = full_text[split_idx:]

    lines = [line.strip() for line in bib_text.split('\n') if line.strip()]
    
    # Pre-scan layout type
    numbered_line_count = 0
    total_lines_checked = 0
    for line in lines[:40]:
        if line.lower() in ['references', 'bibliography', 'literature cited', 'reference list']:
            continue
        total_lines_checked += 1
        if re.match(r"^\[?(\d+)\]?[\s\.\-\)]+", line):
            numbered_line_count += 1
            
    is_numbered_system = (numbered_line_count / max(1, total_lines_checked)) > 0.20
    
    bib_mapping = {}
    current_id = None
    next_expected_id = 1
    generated_id = 1
    doi_regex = r"(10\.\d{4,9}/[-._;()/:A-Z0-9]+)"
    
    for line in lines:
        # Filter layout/journal noise tracking signatures
        if line.lower() in ['references', 'bibliography', 'literature cited', 'reference list']:
            continue
        if "plos one" in line.lower() or "doi.org/10.1371" in line.lower():
            continue
            
        num_match = re.match(r"^\[?(\d+)\]?(?:[\s\-\)]+|\.(?!\d))", line)
        
        if is_numbered_system:
            if num_match:
                potential_id = int(num_match.group(1))
                
                # STRICT SEQUENCE TRACKING: Must follow the established chronological flow
                if potential_id == next_expected_id or (potential_id == next_expected_id + 1 and next_expected_id > 1):
                    current_id = potential_id
                    next_expected_id = current_id + 1
                    clean_citation = re.sub(r"^\[?(\d+)\]?[\s\.\-\)]+", "", line).strip()
                    bib_mapping[current_id] = {
                        "raw_lines": [line],
                        "clean_query": clean_citation,
                        "marker": str(current_id)
                    }
                else:
                    # Out of sequence numbers (years/volumes) are merged back to the active tracking parent
                    if current_id in bib_mapping:
                        bib_mapping[current_id]["raw_lines"].append(line)
                        bib_mapping[current_id]["clean_query"] += " " + line
            else:
                if current_id in bib_mapping:
                    bib_mapping[current_id]["raw_lines"].append(line)
                    bib_mapping[current_id]["clean_query"] += " " + line
                elif not bib_mapping:
                    current_id = generated_id
                    bib_mapping[current_id] = {"raw_lines": [line], "clean_query": line, "marker": str(current_id)}
                    generated_id += 1
        else:
            # Author-Date/Alphabetical system rules
            is_new_author_block = re.match(r"^[A-Z][A-Za-z\-]+,?\s[A-Z]", line) or re.match(r"^[A-Z][a-z]+", line)
            if not bib_mapping or is_new_author_block:
                current_id = generated_id
                author_match = re.match(r"^([A-Za-z\-]+)", line)
                author_token = author_match.group(1) if author_match else ""
                in_text_marker = author_token if author_token else line[:15]
                
                bib_mapping[current_id] = {
                    "raw_lines": [line],
                    "clean_query": line,
                    "marker": in_text_marker
                }
                generated_id += 1
            else:
                if current_id in bib_mapping:
                    bib_mapping[current_id]["raw_lines"].append(line)
                    bib_mapping[current_id]["clean_query"] += " " + line

    final_mapping = {}
    for idx, data in bib_mapping.items():
        combined_text = " ".join(data["raw_lines"])
        doi_match = re.search(doi_regex, combined_text, re.IGNORECASE)
        explicit_doi = doi_match.group(1).rstrip('.,);]') if doi_match else None
        
        final_mapping[idx] = {
            "raw_line": combined_text,
            "clean_query": re.sub(r"\s+", " ", data["clean_query"]).strip(),
            "explicit_doi": explicit_doi,
            "marker": data["marker"]
        }
        
    return body_text, final_mapping

# ─── REGISTRY RESOLUTION ENGINE ───
def resolve_and_verify_citation(bib_info):
    """Resolves citation strings against official Crossref metadata indices."""
    headers = {"User-Agent": "CiteShieldPro/2.0 (mailto:analytics@citeshield.pro)"}
    
    if bib_info["explicit_doi"]:
        url = f"https://api.crossref.org/works/{bib_info['explicit_doi']}"
        try:
            res = requests.get(url, headers=headers, timeout=5)
            if res.status_code == 200:
                msg = res.json().get("message", {})
                return package_crossref_metadata(msg, "Valid (Direct DOI Check)")
        except Exception:
            pass

    search_url = "https://api.crossref.org/works"
    params = {"query": bib_info["clean_query"], "rows": 1}
    
    try:
        res = requests.get(search_url, params=params, headers=headers, timeout=6)
        if res.status_code == 200:
            items = res.json().get("message", {}).get("items", [])
            if not items:
                return {"status": "Hallucinated/Fake Reference", "title": "N/A", "journal": "N/A", "main_author": "N/A", "doi": "N/A", "year": "N/A"}
                
            top_match = items[0]
            resolved_meta = package_crossref_metadata(top_match, "Valid")
            
            query_lower = bib_info["clean_query"].lower()
            title_words = set(re.findall(r'\w+', resolved_meta["title"].lower())) - {'the', 'a', 'of', 'in', 'and', 'for', 'with', 'on', 'at', 'by', 'to'}
            query_words = set(re.findall(r'\w+', query_lower))
            
            word_overlap = title_words.intersection(query_words)
            author_surname = resolved_meta["main_author"].split(',')[0].lower()
            author_match = author_surname in query_lower if author_surname != "unknown author" else True
            
            if len(word_overlap) < 2 and not author_match:
                return {
                    "status": "Hallucinated/Fake Reference", 
                    "title": f"Mismatched Match: {resolved_meta['title']}", 
                    "journal": resolved_meta["journal"], 
                    "main_author": resolved_meta["main_author"], 
                    "doi": resolved_meta["doi"], 
                    "year": resolved_meta["year"]
                }
                
            return resolved_meta
    except Exception:
        pass
        
    return {"status": "Unverified API Failure", "title": "N/A", "journal": "N/A", "main_author": "N/A", "doi": "N/A", "year": "N/A"}

def package_crossref_metadata(item, base_status):
    """Packages dictionary metadata cleanly."""
    title = item.get("title", ["Unknown Title"])[0]
    journal = item.get("container-title", ["Unknown Journal"])[0]
    doi = item.get("DOI", "N/A")
    
    authors = item.get("author", [])
    main_author = "Unknown Author"
    if authors:
        main_author = f"{authors[0].get('family', '')}, {authors[0].get('given', '')}".strip()
        if main_author == ",":
            main_author = authors[0].get("name", "Unknown Author")
            
    year = "N/A"
    for date_field in ["published-print", "published-online", "created"]:
        if date_field in item and item[date_field].get("date-parts"):
            year = item[date_field]["date-parts"][0][0]
            break
            
    return {"status": base_status, "title": title, "journal": journal, "main_author": main_author, "doi": doi, "year": year}

@st.cache_data(show_spinner=False)
def fetch_abstract_via_semantic_scholar(doi):
    """Queries Semantic Scholar database to fetch paper abstracts."""
    if not doi or doi == "N/A":
        return ""
    url = f"https://api.semanticscholar.org/graph/v1/paper/DOI:{doi}?fields=abstract"
    try:
        res = requests.get(url, timeout=5)
        if res.status_code == 200:
            return res.json().get("abstract", "")
    except Exception:
        pass
    return ""

def get_contextual_fit_analysis(context_paragraph, reference_content):
    """
    Performs a forensic audit of the citation context using a 6-point 
    verification schema to ensure factual alignment.
    """
    prompt = f"""
    You are an academic forensic auditor. Compare the provided Body Text excerpt and the cited Reference for factual consistency.

    1. Body Text: "{context_paragraph}"
    2. Reference: "{reference_content}"

    Audit the following 6 core data points:
    1. Study Sites: Do the hospitals match?
    2. Methodology: Does the study design/type match?
    3. Sample Size: Does the 'n' match?
    4. Study Year: Do the timeframes match?
    5. Objectives: Do the study goals align?
    6. Abstract Context: Does the reference abstract support the body text's claim?

    Output the result in the following JSON format:
    {{
        "Status": "ALIGNED" or "MISMATCH",
        "Discrepancies": ["list any specific failures here or 'None'"],
        "Reasoning": "One sentence explanation."
    }}
    """
    
    # Generate content using the Gemini model
    model = genai.GenerativeModel('gemini-1.5-flash')
    response = model.generate_content(prompt)
    
    # Parse the JSON response
    try:
        json_text = response.text.replace("```json", "").replace("```", "").strip()
        data = json.loads(json_text)
        
        return {
            "status": data.get("Status", "MISMATCH"),
            "discrepancies": data.get("Discrepancies", []),
            "reasoning": data.get("Reasoning", "No explanation provided.")
        }
    except Exception as e:
        return {
            "status": "MISMATCH",
            "discrepancies": ["Audit parsing error"],
            "reasoning": "Could not parse audit output."
        }

# ─── USER INTERFACE RUNTIME ───
uploaded_file = st.file_uploader("Upload Manuscript File", type=["docx", "pdf"])

if uploaded_file:
    with st.spinner("Decoding document structural layout matrices..."):
        full_text, paragraphs = extract_text_and_paragraphs(uploaded_file)
        body_text, bib_mapping = parse_bibliography(full_text)
        
    if not bib_mapping:
        st.error("❌ Critical Error: Could not parse any recognizable reference list entries.")
    else:
        st.success(f"Parsed {len(bib_mapping)} bibliography elements successfully!")
        
        audit_records = []
        progress_bar = st.progress(0)
        
        for idx, (serial_num, info) in enumerate(bib_mapping.items()):
            verification = resolve_and_verify_citation(info)
            
            v_status = verification.get("status", "Unverified")
            v_doi = verification.get("doi", "N/A")
            v_title = verification.get("title", "N/A")
            v_journal = verification.get("journal", "N/A")
            v_author = verification.get("main_author", "N/A")
            
            marker = info["marker"]
            context_paragraphs = [p for p in paragraphs if marker in p]
            combined_context = " | ".join(context_paragraphs) if context_paragraphs else "Not explicitly mapped inside body text layout."
            
            abstract_text = ""
            if "Valid" in v_status:
                abstract_text = fetch_abstract_via_semantic_scholar(v_doi)
                
            record = {
                "Serial No": serial_num,
                "DOI": v_doi,
                "Title": v_title,
                "Journal": v_journal,
                "Main Author": v_author,
                "Status": v_status,
                "Context Fit": "Not Verified",
                "Context Paragraph": combined_context
            }
            
            if "Valid" in v_status and gemini_api_key and abstract_text:
                result = get_contextual_fit_analysis(combined_context, abstract_text)
                fit = (result["status"] == "ALIGNED")
                reason = result["reasoning"]
                if not fit:
                    record["Status"] = "Contextual Mismatch"
                    record["Context Fit"] = f"🚨 Misalignment: {reason}"
                else:
                    record["Context Fit"] = "✅ Aligned"
            elif "Valid" in v_status and not gemini_api_key:
                record["Context Fit"] = "Provide Gemini Key to Verify"
            elif "Hallucinated" in v_status:
                record["Context Fit"] = "Skipped (Fake Reference)"
                
            audit_records.append(record)
            progress_bar.progress((idx + 1) / len(bib_mapping))
            
        df_audit = pd.DataFrame(audit_records)
        
        # ─── SUMMARY TABLE ───
        st.markdown("## 📊 Reference Audit Summary Table")
        flagged_df = df_audit[df_audit["Status"].isin(["Hallucinated/Fake Reference", "Contextual Mismatch", "Unverified API Failure"])]
        
        if not flagged_df.empty:
            st.warning(f"Attention Required: {len(flagged_df)} anomalies found.")
            st.dataframe(
                flagged_df[["Serial No", "DOI", "Status", "Main Author", "Journal", "Context Fit"]], 
                use_container_width=True,
                hide_index=True
            )
        else:
            st.success("Clean Audit! Every reference matched its live registry parameters seamlessly.")

        # ─── DETAILED LOGS ───
        st.markdown("## 🔍 Detailed Reference Audit Log")
        
        for rec in audit_records:
            with st.container():
                col_status, col_content = st.columns([1, 4])
                
                with col_status:
                    st.markdown(f"### `Ref #{rec['Serial No']}`")
                    if "Valid" in rec["Status"]:
                        st.markdown("🟢 **AUTHENTIC**")
                    elif rec["Status"] == "Contextual Mismatch":
                        st.markdown("🟡 **MISMATCH**")
                    else:
                        st.markdown("🔴 **FAKE/UNVERIFIED**")
                        
                with col_content:
                    st.markdown(f"**Title:** *{rec['Title']}*")
                    st.markdown(f"**Journal:** {rec['Journal']} | **Main Author:** {rec['Main Author']} | **DOI:** `{rec['DOI']}`")
                    
                    if rec["Status"] == "Contextual Mismatch":
                        st.error(f"**Context Fit Flag:** {rec['Context Fit']}")
                        st.markdown("#### Highlighted Paragraph with Inaccurate Context:")
                        st.info(f"👉 \"{rec['Context Paragraph']}\"")
                    elif rec["Context Paragraph"] != "Not explicitly mapped inside body text layout." and gemini_api_key and "Valid" in rec["Status"]:
                        st.success(f"**Context Fit Check:** {rec['Context Fit']}")
                        
                st.markdown("---")
