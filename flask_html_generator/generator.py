
from __future__ import annotations

import json
import re
from collections import OrderedDict, defaultdict
from dataclasses import dataclass
from pathlib import Path

import pandas as pd


# ---------- Generic helpers ----------

def clean(value) -> str:
    if pd.isna(value):
        return ""
    s = str(value).strip()
    return "" if s.lower() == "nan" else s


def norm_issn(value: str) -> str:
    return "".join(ch for ch in clean(value) if ch.isdigit() or ch.upper() == "X").upper()


def dashed_issn(value: str) -> str:
    normalized = norm_issn(value)
    return f"{normalized[:4]}-{normalized[4:]}" if len(normalized) == 8 else clean(value)


def dedupe_issns_prefer_dash(values) -> list[str]:
    chosen: OrderedDict[str, str] = OrderedDict()
    for raw in values:
        raw = clean(raw)
        if not raw:
            continue
        key = norm_issn(raw)
        if not key:
            continue
        candidate = dashed_issn(raw) if len(key) == 8 else raw
        if key not in chosen or ("-" not in chosen[key] and "-" in candidate):
            chosen[key] = candidate
    return list(chosen.values())

def valid_norm_issn(value: str) -> str:
    key = norm_issn(value)
    return key if len(key) == 8 else ""


def dedupe_names(values) -> list[str]:
    seen = set()
    out = []
    for value in values:
        value = clean(value)
        if value and value not in seen:
            seen.add(value)
            out.append(value)
    return out


def is_all_caps_name(name: str) -> bool:
    letters = [ch for ch in name if ch.isalpha()]
    return bool(letters) and all(ch.upper() == ch for ch in letters)


def choose_canonical_name(candidates: list[tuple[int, str]]) -> str:
    cleaned = []
    for order, name in candidates:
        n = clean(name)
        if n:
            cleaned.append((order, n))
    if not cleaned:
        return ""

    non_caps = [item for item in cleaned if not is_all_caps_name(item[1])]
    if non_caps:
        non_caps.sort(key=lambda x: x[0])
        return non_caps[0][1]

    cleaned.sort(key=lambda x: x[0])
    return cleaned[0][1]

def normalize_name_key(name: str) -> str:
    s = clean(name).lower()
    s = re.sub(r"[\(\)\[\]\{\}:;,\.\-_/&]", " ", s)
    s = re.sub(r"\bthe\b", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def normalized_display_name(name: str) -> str:
    s = clean(name)
    if not s:
        return ""
    # remove trailing artifact digits when attached to the final token
    s = re.sub(r"(?<=[A-Za-z\)])\d+$", "", s).strip()
    s = re.sub(r"^The\s+", "", s, flags=re.I)
    s = re.sub(r"\s+", " ", s).strip()
    if is_all_caps_name(s):
        s = s.title()
    return s


def row_has_any_content(row: pd.Series) -> bool:
    return any(clean(v) != "" for v in row.tolist())


class DSU:
    def __init__(self):
        self.parent: dict[str, str] = {}

    def find(self, x: str) -> str:
        self.parent.setdefault(x, x)
        if self.parent[x] != x:
            self.parent[x] = self.find(self.parent[x])
        return self.parent[x]

    def union(self, a: str, b: str) -> None:
        ra, rb = self.find(a), self.find(b)
        if ra != rb:
            self.parent[rb] = ra


# ---------- Cleaning: agreement ----------

AGREEMENT_NAME_OVERRIDES = {
    "Lyell": "Lyell Collection Journals",
    "ACM": "Title",
    "Sage": "Journal",
}

def agreement_issn_like_columns(cols) -> list:
    out = []
    for c in cols:
        s = str(c).strip().lower()
        if any(tok in s for tok in ["issn", "identifier"]):
            out.append(c)
    return out


def detect_agreement_name_column(sheet_name: str, cols) -> str | None:
    if sheet_name in AGREEMENT_NAME_OVERRIDES and AGREEMENT_NAME_OVERRIDES[sheet_name] in cols:
        return AGREEMENT_NAME_OVERRIDES[sheet_name]

    priority = [
        "Journal", "Journal Name", "Journal name", "Journal Title", "Title",
        "title", "Name", "name", "short name"
    ]
    lower_map = {str(c).strip().lower(): c for c in cols}
    for p in priority:
        if p.lower() in lower_map:
            return lower_map[p.lower()]

    for c in cols:
        s = str(c).strip().lower()
        if not any(tok in s for tok in ["issn", "identifier", "percentile", ">90"]):
            return c
    return None


def clean_agreement_workbook(input_path: str, output_path: str) -> dict[str, pd.DataFrame]:
    xls = pd.ExcelFile(input_path)
    cleaned: dict[str, pd.DataFrame] = {}

    for sheet in xls.sheet_names:
        df = pd.read_excel(input_path, sheet_name=sheet)
        cols = list(df.columns)

        name_col = detect_agreement_name_column(sheet, cols)
        issn_cols = agreement_issn_like_columns(cols)

        keep_cols = []
        if name_col:
            keep_cols.append(name_col)
        keep_cols.extend([c for c in issn_cols if c not in keep_cols])

        # Sage special keep/rename
        if sheet == "Sage":
            lower_map = {str(c).strip().lower(): c for c in cols}
            for candidate in ["jif percentile", ">90"]:
                if candidate in lower_map:
                    c = lower_map[candidate]
                    if c not in keep_cols:
                        keep_cols.append(c)
                    break

        if not keep_cols:
            continue

        out = df[keep_cols].copy()
        out = out[out.apply(row_has_any_content, axis=1)].copy()

        if sheet == "Lyell":
            def lyell_keep(row):
                name = clean(row.get(name_col)).lower()
                issns = [norm_issn(row.get(c)) for c in issn_cols if norm_issn(row.get(c))]
                if "book series" in name:
                    return False
                if not issns and any(tok in name for tok in ["series", "collection"]):
                    return False
                return True
            out = out[out.apply(lyell_keep, axis=1)].copy()

        rename_map = {}
        if name_col and name_col != "Journal Name":
            rename_map[name_col] = "Journal Name"
        if sheet == "Sage":
            for c in list(out.columns):
                s = str(c).strip().lower()
                if s in {"jif percentile", ">90"}:
                    rename_map[c] = "Open Access Fund Better"
        out = out.rename(columns=rename_map)
        cleaned[sheet] = out.reset_index(drop=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet, df in cleaned.items():
            df.to_excel(writer, sheet_name=sheet[:31], index=False)

    return cleaned


# ---------- Cleaning: keren ----------

def keren_find_header_row(df: pd.DataFrame):
    for i in range(len(df)):
        row = df.iloc[i]
        if any("issn" in clean(x).lower() for x in row.tolist()):
            return i
    return None


def find_books_header_row(df: pd.DataFrame):
    for i in range(len(df)):
        vals = [clean(x).lower() for x in df.iloc[i].tolist()]
        if any(v in {"publisher", "press", "book publisher"} or "publisher" in v for v in vals):
            return i
    return None

def keren_trim_source(df: pd.DataFrame):
    cut = None
    for i in range(len(df) - 1, -1, -1):
        vals = [clean(x).lower() for x in df.iloc[i].tolist()]
        if any(re.match(r"^source:?$", v) for v in vals):
            cut = i
            break
    if cut is None:
        return df.reset_index(drop=True)
    start = cut
    j = cut - 1
    while j >= 0:
        vals = [clean(x) for x in df.iloc[j].tolist()]
        if all(v == "" for v in vals):
            start = j
            j -= 1
        else:
            break
    return df.iloc[:start].copy().reset_index(drop=True)


def keren_detect_name_issn(cols):
    lower = {str(c).strip().lower(): c for c in cols}
    name = None
    for p in ["journal name", "journal title", "book title", "title", "name", "press", "book publisher", "publisher", "journal"]:
        if p in lower:
            name = lower[p]
            break
    if name is None:
        for c in cols:
            s = str(c).strip().lower()
            if not any(tok in s for tok in ["issn", "identifier", "rank", "publisher", "field", "year", "source", "edition", "category", "citations", "jif", "jci", "oa", "percentile", "abbreviation", "acronym", "id#", "eraid", "#"]):
                name = c
                break
    issns = [c for c in cols if any(tok in str(c).strip().lower() for tok in ["issn", "identifier"])]
    return name, issns


def publisher_key(name):
    s = clean(name).lower()
    s = re.sub(r"[^\w\s]", " ", s)
    tokens = [t for t in s.split() if t not in {"of", "the", "and"}]
    tokens = [t for t in tokens if t not in {"press", "publisher", "publishers", "publishing", "university"}]
    return tuple(sorted(tokens))


def choose_preferred_publisher(names: list[str]) -> str:
    uniq = []
    seen = set()
    for n in names:
        n = clean(n)
        if n and n not in seen:
            uniq.append(n)
            seen.add(n)

    def pref_key(x: str):
        lx = x.lower()
        return (
            lx.startswith("university of "),
            0 if " university " in f" {lx} " else 1,
            0 if " press" in lx else 1,
            len(lx),
            lx,
        )

    uniq.sort(key=pref_key)
    return uniq[0]


def clean_keren_workbook(input_path: str, output_path: str) -> dict[str, pd.DataFrame]:
    xls = pd.ExcelFile(input_path)
    cleaned: dict[str, pd.DataFrame] = {}
    book_groups: dict[tuple[str, ...], list[str]] = defaultdict(list)

    for sheet in xls.sheet_names:
        if sheet == "Menu":
            continue

        if sheet.endswith("_Books"):
            raw = pd.read_excel(input_path, sheet_name=sheet, header=None)
            raw = keren_trim_source(raw)

            header_idx = find_books_header_row(raw)
            if header_idx is not None:
                headers = [clean(x) for x in raw.iloc[header_idx].tolist()]
                body = raw.iloc[header_idx + 1:].copy().reset_index(drop=True)
                body.columns = headers
            else:
                body = raw.copy().reset_index(drop=True)

            if len(body.columns) == 0:
                continue

            publisher_col = None
            for c in body.columns:
                s = str(c).strip().lower()
                if s in {"publisher", "press", "book publisher"} or "publisher" in s:
                    publisher_col = c
                    break
            if publisher_col is None:
                publisher_col = body.columns[0]

            for _, row in body.iterrows():
                name = clean(row.get(publisher_col))
                if not name:
                    continue
                if name.isdigit():
                    continue
                if name.lower() in {"press", "publisher", "book publisher", "journal name", "title"}:
                    continue
                key = publisher_key(name)
                if key:
                    book_groups[key].append(name)
            continue

        raw = pd.read_excel(input_path, sheet_name=sheet)
        header_idx = keren_find_header_row(raw)

        if header_idx is not None:
            trimmed = raw.iloc[header_idx:].copy().reset_index(drop=True)
            headers = [clean(x) for x in trimmed.iloc[0].tolist()]
            body = trimmed.iloc[1:].copy().reset_index(drop=True)
            body.columns = headers
        else:
            body = raw.copy()

        body = keren_trim_source(body)
        cols = list(body.columns)
        name_col, issn_cols = keren_detect_name_issn(cols)

        # explicit fixes
        if sheet == "Law" and "Journal Name" in body.columns:
            name_col = "Journal Name"
        elif sheet == "Humanities" and "Title" in body.columns:
            name_col = "Title"
        elif sheet == "Business_Admin (AJG)" and "Journal Title" in body.columns:
            name_col = "Journal Title"
        elif sheet == "Math_ERA" and "Title" in body.columns:
            name_col = "Title"
        elif sheet == "CS_Conferences" and "Title" in body.columns:
            name_col = "Title"
            issn_cols = []

        keep_cols = []
        if name_col and name_col in body.columns:
            keep_cols.append(name_col)
        keep_cols.extend([c for c in issn_cols if c in body.columns and c not in keep_cols])

        if not keep_cols:
            continue

        out = body[keep_cols].copy()
        if name_col != "Journal Name" and name_col in out.columns:
            out = out.rename(columns={name_col: "Journal Name"})

        keep_mask = []
        for _, row in out.iterrows():
            name = clean(row.get("Journal Name"))
            row_issns = [norm_issn(row.get(c)) for c in out.columns if c != "Journal Name" and norm_issn(row.get(c))]
            if name == "" and not row_issns:
                keep_mask.append(False)
                continue
            lname = name.lower()
            if "book series" in lname:
                keep_mask.append(False)
                continue
            if not row_issns and any(tok in lname for tok in ["series", "collection"]):
                keep_mask.append(False)
                continue
            keep_mask.append(True)

        out = out[pd.Series(keep_mask, index=out.index)].copy().reset_index(drop=True)
        cleaned[sheet] = out

    # Create unified Books tab
    book_values = [choose_preferred_publisher(names) for names in book_groups.values()]
    book_values = sorted(dict.fromkeys(book_values), key=lambda x: x.lower())
    cleaned["Books"] = pd.DataFrame({"Book Publisher": book_values})

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet, df in cleaned.items():
            df.to_excel(writer, sheet_name=sheet[:31], index=False)

    return cleaned


# ---------- Rebuild myName / name_changed ----------

@dataclass
class RowRef:
    workbook: str
    sheet: str
    idx: int
    row_num: int
    orig_name: str
    issns: list[str]
    is_book: bool


def detect_processed_name_column(sheet: str, cols) -> str | None:
    lower = {str(c).strip().lower(): c for c in cols}
    for p in ["journal name", "book publisher", "publisher", "press", "title", "name"]:
        if p in lower:
            return lower[p]
    return cols[0] if cols else None


def rebuild_myname_on_cleaned_workbooks(keren_dfs: dict[str, pd.DataFrame], agreement_dfs: dict[str, pd.DataFrame]):
    refs: list[RowRef] = []
    dsu = DSU()
    candidates: dict[str, list[tuple[int, str]]] = defaultdict(list)
    appearance = 0

    # First pass over both workbooks
    for workbook, dfs in [("keren", keren_dfs), ("agreement", agreement_dfs)]:
        for sheet, df in dfs.items():
            cols = list(df.columns)
            name_col = detect_processed_name_column(sheet, cols)
            issn_cols = [c for c in cols if any(tok in str(c).strip().lower() for tok in ["issn", "identifier"])]
            is_book = sheet.strip().lower() == "books"

            for idx, row in df.iterrows():
                if not row_has_any_content(row):
                    continue
                orig_name = clean(row.get(name_col))
                raw_issns = dedupe_issns_prefer_dash([row.get(c) for c in issn_cols])
                refs.append(RowRef(workbook, sheet, idx, idx + 2, orig_name, raw_issns, is_book))

                if is_book:
                    continue
                keys = [norm_issn(x) for x in raw_issns if norm_issn(x)]
                for i in range(1, len(keys)):
                    dsu.union(keys[0], keys[i])

    # candidate names from all non-book rows with ISSN
    for ref in refs:
        if ref.is_book:
            continue
        keys = [norm_issn(x) for x in ref.issns if norm_issn(x)]
        if not keys:
            continue
        appearance += 1
        cluster = dsu.find(keys[0])
        candidates[cluster].append((appearance, ref.orig_name))

    # Assign myName + name_changed
    for workbook, dfs in [("keren", keren_dfs), ("agreement", agreement_dfs)]:
        for sheet, df in list(dfs.items()):
            cols = list(df.columns)
            name_col = detect_processed_name_column(sheet, cols)
            issn_cols = [c for c in cols if any(tok in str(c).strip().lower() for tok in ["issn", "identifier"])]
            is_book = sheet.strip().lower() == "books"

            my_names = []
            name_changed = []

            for idx, row in df.iterrows():
                if not row_has_any_content(row):
                    my_names.append("")
                    name_changed.append("")
                    continue

                orig_name = clean(row.get(name_col))
                raw_issns = dedupe_issns_prefer_dash([row.get(c) for c in issn_cols])

                if is_book:
                    myname = orig_name
                else:
                    keys = [norm_issn(x) for x in raw_issns if norm_issn(x)]
                    if keys:
                        myname = choose_canonical_name(candidates.get(dsu.find(keys[0]), []))
                    else:
                        myname = orig_name

                my_names.append(myname)
                name_changed.append("X" if clean(orig_name) != clean(myname) else "")

            df_out = df.copy()
            if "myName" in df_out.columns:
                df_out = df_out.drop(columns=["myName"])
            if "name_changed" in df_out.columns:
                df_out = df_out.drop(columns=["name_changed"])

            df_out["myName"] = my_names
            df_out["name_changed"] = name_changed

            base_cols = list(df_out.columns)
            name_col_actual = detect_processed_name_column(sheet, base_cols)
            issn_cols_actual = [c for c in base_cols if any(tok in str(c).strip().lower() for tok in ["issn", "identifier"])]
            lead_cols = []
            if name_col_actual and name_col_actual in df_out.columns:
                lead_cols.append(name_col_actual)
            lead_cols.extend([c for c in issn_cols_actual if c in df_out.columns and c not in lead_cols])

            tail_priority = [c for c in ["myName", "name_changed"] if c in df_out.columns]
            middle_cols = [c for c in df_out.columns if c not in lead_cols and c not in tail_priority]
            ordered_cols = lead_cols + tail_priority + middle_cols
            df_out = df_out[ordered_cols]

            dfs[sheet] = df_out

    # Sage Open Access Fund Better based on JCR JIF >90
    sage_sheet = next((s for s in agreement_dfs.keys() if s.strip().lower() == "sage"), None)
    if sage_sheet and "JCR JIF >90" in keren_dfs:
        jcr_names = {clean(v).lower() for v in keren_dfs["JCR JIF >90"]["myName"].tolist() if clean(v)}
        sage_df = agreement_dfs[sage_sheet].copy()
        values = []
        for _, row in sage_df.iterrows():
            name = clean(row.get("myName")).lower()
            values.append("X" if name and name in jcr_names else "")
        if "Open Access Fund Better" in sage_df.columns:
            sage_df = sage_df.drop(columns=["Open Access Fund Better"])
        sage_df["Open Access Fund Better"] = values
        agreement_dfs[sage_sheet] = sage_df

    return keren_dfs, agreement_dfs


# ---------- Build HTML dataset from processed workbooks ----------

def merge_html_items(raw_items, include_issns: bool = True) -> list[dict]:
    if not raw_items:
        return []

    dsu = DSU()
    issn_to_ids: dict[str, list[str]] = defaultdict(list)
    no_issn_name_to_ids: dict[str, list[str]] = defaultdict(list)

    for idx, item in enumerate(raw_items):
        item_id = f"item_{idx}"
        issns = [norm_issn(x) for x in item["searchIssns"] if norm_issn(x)]
        if issns:
            for key in issns:
                for prev in issn_to_ids[key]:
                    dsu.union(item_id, prev)
                issn_to_ids[key].append(item_id)
        else:
            key = item["displayName"].strip().lower()
            for prev in no_issn_name_to_ids[key]:
                dsu.union(item_id, prev)
            no_issn_name_to_ids[key].append(item_id)

    merged: OrderedDict[str, dict] = OrderedDict()
    for idx, item in enumerate(raw_items):
        root = dsu.find(f"item_{idx}")
        if root not in merged:
            merged[root] = {
                "displayName": item["displayName"],
                "names": [],
                "displayIssns": [],
                "searchIssns": [],
                "kerenTabs": [],
                "agreementTabs": [],
                "sageHasOpenFund": False,
            }
        m = merged[root]
        m["displayName"] = m["displayName"] or item["displayName"]
        m["names"].extend(item["names"])
        m["displayIssns"].extend(item["displayIssns"])
        m["searchIssns"].extend(item["searchIssns"])
        m["kerenTabs"].extend(item["kerenTabs"])
        m["agreementTabs"].extend(item["agreementTabs"])
        m["sageHasOpenFund"] = m["sageHasOpenFund"] or item["sageHasOpenFund"]

    out = []
    for m in merged.values():
        names = dedupe_names(m["names"])
        display_issns = dedupe_issns_prefer_dash(m["displayIssns"]) if include_issns else []
        search_issns = dedupe_issns_prefer_dash(m["searchIssns"]) if include_issns else []
        agreement_tabs = dedupe_names(m["agreementTabs"])
        if "ACS" in agreement_tabs:
            agreement_publisher = "ACS"
        elif "Sage" in agreement_tabs or "SAGE" in agreement_tabs:
            agreement_publisher = "Sage"
        else:
            agreement_publisher = agreement_tabs[0] if agreement_tabs else ""
        unique_label = f"{m['displayName']} ({','.join(display_issns)})" if display_issns else m["displayName"]
        out.append({
            "displayName": m["displayName"],
            "uniqueLabel": unique_label,
            "names": names,
            "normSearchIssns": sorted({norm_issn(x) for x in search_issns if norm_issn(x)}),
            "isInKeren": len(dedupe_names(m["kerenTabs"])) > 0,
            "agreementTabs": agreement_tabs,
            "agreementPublisher": agreement_publisher,
            "sageHasOpenFund": bool(m["sageHasOpenFund"]),
        })
    return sorted(out, key=lambda x: x["displayName"].lower())


def build_html_from_processed(keren_dfs: dict[str, pd.DataFrame], agreement_dfs: dict[str, pd.DataFrame], output_path: str, template_path: str | None = None) -> str:
    books_raw = []
    journals_raw = []

    # Keren
    for sheet, df in keren_dfs.items():
        cols = list(df.columns)
        name_col = detect_processed_name_column(sheet, cols)
        issn_cols = [c for c in cols if any(tok in str(c).strip().lower() for tok in ["issn", "identifier"])]

        for _, row in df.iterrows():
            if not row_has_any_content(row):
                continue
            myname = clean(row.get("myName"))
            orig_name = clean(row.get(name_col))
            display_name = myname or orig_name
            if not display_name:
                continue
            row_issns = dedupe_issns_prefer_dash([row.get(c) for c in issn_cols if valid_norm_issn(row.get(c))])

            record = {
                "displayName": display_name,
                "names": [display_name],
                "displayIssns": row_issns,
                "searchIssns": row_issns,
                "kerenTabs": [sheet] if sheet.strip().lower() != "books" else [],
                "agreementTabs": [],
                "sageHasOpenFund": False,
            }
            if sheet.strip().lower() == "books":
                record["displayIssns"] = []
                record["searchIssns"] = []
                books_raw.append(record)
            elif sheet != "CS_Conferences":
                journals_raw.append(record)

    # Agreement
    sage_sheet = next((s for s in agreement_dfs.keys() if s.strip().lower() == "sage"), None)
    sage_map = {}
    if sage_sheet:
        for _, row in agreement_dfs[sage_sheet].iterrows():
            name = clean(row.get("myName")).lower()
            if name:
                sage_map[name] = clean(row.get("Open Access Fund Better")) == "X"

    for sheet, df in agreement_dfs.items():
        cols = list(df.columns)
        name_col = detect_processed_name_column(sheet, cols)
        issn_cols = [c for c in cols if any(tok in str(c).strip().lower() for tok in ["issn", "identifier"])]

        for _, row in df.iterrows():
            if not row_has_any_content(row):
                continue
            myname = clean(row.get("myName"))
            orig_name = clean(row.get(name_col))
            display_name = myname or orig_name
            if not display_name:
                continue
            row_issns = dedupe_issns_prefer_dash([row.get(c) for c in issn_cols if valid_norm_issn(row.get(c))])

            journals_raw.append({
                "displayName": display_name,
                "names": [display_name],
                "displayIssns": row_issns,
                "searchIssns": row_issns,
                "kerenTabs": [],
                "agreementTabs": [sheet],
                "sageHasOpenFund": sage_map.get(display_name.lower(), False) if sheet == sage_sheet else False,
            })

    books_items = merge_html_items(books_raw, include_issns=False)
    journals_items = merge_html_items(journals_raw, include_issns=True)

    books_json = json.dumps(books_items, ensure_ascii=False)
    journals_json = json.dumps(journals_items, ensure_ascii=False)

    if template_path is None:
        template_file = Path(__file__).resolve().parent / "templates" / "app_template.html"
    else:
        template_file = Path(template_path)

    template_html = template_file.read_text(encoding="utf-8")
    final_html = template_html.replace("__BOOKS__", books_json).replace("__JOURNALS__", journals_json)
    Path(output_path).write_text(final_html, encoding="utf-8")
    return output_path


# ---------- Public API ----------

def generate_outputs(raw_keren_path: str, raw_agreement_path: str, output_dir: str, include_cleaned_excels: bool = False) -> dict[str, str]:
    output_root = Path(output_dir)
    output_root.mkdir(parents=True, exist_ok=True)

    ts = pd.Timestamp.now().strftime("%Y%m%d_%H%M%S")
    cleaned_keren_path = output_root / f"keren_clean__{ts}.xlsx"
    cleaned_agreement_path = output_root / f"agreement_clean__{ts}.xlsx"
    processed_keren_path = output_root / f"keren_processed__{ts}.xlsx"
    processed_agreement_path = output_root / f"agreement_processed__{ts}.xlsx"
    html_path = output_root / f"open_access_search__{ts}.html"

    cleaned_keren = clean_keren_workbook(raw_keren_path, str(cleaned_keren_path))
    cleaned_agreement = clean_agreement_workbook(raw_agreement_path, str(cleaned_agreement_path))

    processed_keren, processed_agreement = rebuild_myname_on_cleaned_workbooks(cleaned_keren, cleaned_agreement)

    with pd.ExcelWriter(processed_keren_path, engine="openpyxl") as writer:
        for sheet, df in processed_keren.items():
            df.to_excel(writer, sheet_name=sheet[:31], index=False)
    with pd.ExcelWriter(processed_agreement_path, engine="openpyxl") as writer:
        for sheet, df in processed_agreement.items():
            df.to_excel(writer, sheet_name=sheet[:31], index=False)

    build_html_from_processed(processed_keren, processed_agreement, str(html_path))

    result = {"html_path": str(html_path)}
    if include_cleaned_excels:
        result["cleaned_keren_xlsx_path"] = str(cleaned_keren_path)
        result["cleaned_agreement_xlsx_path"] = str(cleaned_agreement_path)
        result["processed_keren_xlsx_path"] = str(processed_keren_path)
        result["processed_agreement_xlsx_path"] = str(processed_agreement_path)
    return result
