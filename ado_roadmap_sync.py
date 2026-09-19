# ============================================================
# ADO Roadmap Sync Script v7.1
# Pulls Product Backlog Items + orphan Features/Epics
# from HMIS (TFS on-prem) and generates a roadmap Excel
# matching the original format, with a Dashboard sheet.
#
# New in v7: ticket system analysis. Ticket numbers linked to
# Features/PBIs (Custom.Ticketnumber — the lowest level PBI
# overrides the Feature value) are matched against the ticket
# system export Excel. A new 'Tickets' sheet lists every
# ticket with name, severity, number, status and a
# 'Covered by Roadmap' flag; the Dashboard shows tickets per
# severity and covered/not-covered counts.
#
# New in v7.1: only 'Product Backlog Item' is fetched (Product Non
# Backlog Item excluded). A 'Ticket Priority' column (from the ticket
# system export) sits next to Ticket Number on the Roadmap and New
# Stories sheets — when several tickets are linked, the highest-rated
# ticket's priority wins (Critical > High > Medium > Normal). The
# Completed Stories sheet also shows Ticket Number + Ticket Priority,
# and the Tickets sheet now includes the ticket's Module.
#
# New in v6: no carry-over — the roadmap is built completely
# fresh from TFS on every run.
#
# New in v5: TFS fields Ticketnumber / BusinessImpactValue /
# BusinessValueCategory / Impactlevel are pulled from work
# items. Feature-level values inherit down to every child
# story; story-level values override the feature's values.
#
# Run: python ado_roadmap_sync.py
# Output: roadmap_<YYYY-MM-DD_HH-MM>.xlsx in the same folder
#
# At startup the script asks for the two cutoff dates (ENTER keeps
# the configured defaults). Pass --defaults to skip the wizard.
# ============================================================

import requests
import base64
import re
import os
import sys
import time
import fnmatch
import warnings
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from collections import defaultdict
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo, TableColumn
from openpyxl.worksheet.filters import AutoFilter
from urllib.parse import quote

# ============================================================
# CONFIGURATION
# ============================================================

TFS_URL = "http://ahq-tfs-azure/DefaultCollection"
PAT = "gym5zjd2luca3exar5mc2izkjsu7hudv4bxbaxz6bixundvpskjq"
# Folder where the script/exe lives — output Excel and the ticket export
# are expected here. PyInstaller-aware: when frozen, use the exe's folder
if getattr(sys, "frozen", False):
    OUTPUT_FOLDER = os.path.dirname(os.path.abspath(sys.executable))
else:
    OUTPUT_FOLDER = os.path.dirname(os.path.abspath(__file__))
API_VERSION = "5.1"

PROJECT = "HMIS"

# All projects to query
PROJECTS = ["HMIS", "HR System", "Mobile Applications", "Websites"]

# Only real Product Backlog Items are pulled ('Product Non Backlog Item'
# is intentionally excluded)
PBI_TYPES = ["Product Backlog Item"]
FEATURE_TYPES = ["Feature", "Epic"]

ALLOWED_OWNERS = [
    "ahmed nasr younis abdelwahed",
    "mohamed sharshira",
    "mohamed ahmed mohamed aly",
    "ibrahim abdelfattah mohamed ghanem",
    "nada adel khamis",
    "mohamed adel khalifa",
    "mohamed moataz",
    "elzohery",
]

# New TFS fields (roadmap columns). Values live on Features and stories;
# a Feature's value applies to every child story, unless the story has
# its own value for that field (story level overrides feature level).
NEW_FIELD_REFS = {
    "Ticket Number": "Custom.Ticketnumber",
    "Notes": "Custom.BusinessImpactValue",
    "Category": "Custom.BusinessValueCategory",
    "Impact": "Custom.Impactlevel",
}
# Some of these fields are HTML-formatted in TFS and need cleaning
HTML_NEW_FIELDS = {"Custom.Ticketnumber", "Custom.BusinessImpactValue"}

# Standard Weight (note: the field is spelled 'Standared weight' in TFS).
# Falls back to the item's Effort when Standard Weight is empty.
STANDARD_WEIGHT_REF = "Custom.Standaredweight"
EFFORT_REF = "Microsoft.VSTS.Scheduling.Effort"

# Cutoff date for dashboard "new items" counting
CUTOFF_DATE = "2026-09-01"

# Roadmap cutoff date — stories with a Done/Delivery date BEFORE this are excluded.
# Stories not done yet, or done on/after this date, are included.
ROADMAP_CUTOFF_DATE = "2026-08-01"

# Projects where unparented PBIs are EXCLUDED (they must have a Feature/Epic parent).
# Empty list = unparented PBIs are included everywhere; each becomes its own
# feature group (story title = Business Area).
UNPARENTED_EXCLUDE_PROJECTS = []

# Ticket system analysis: the script reads the ticket export Excel generated
# by the ticket system. TICKET_EXPORT_FOLDER may be a FOLDER (the newest file
# matching TICKET_EXPORT_PATTERN inside it is used) or a direct FILE path.
# By default the ticket export is expected in the SAME FOLDER as this script.
# Set TICKET_EXPORT_FOLDER = "" to disable ticket analysis.
TICKET_EXPORT_FOLDER = OUTPUT_FOLDER
TICKET_EXPORT_PATTERN = "Active All Requests*"

# Display order for ticket severities (matched by keyword, so emoji prefixes
# like 🟥 / 🟦 in the export values are handled automatically).
SEVERITY_ORDER = ["Critical", "High", "Medium", "Normal"]

# Column names in the ticket export (header row, matched case-insensitively)
TICKET_EXPORT_COLUMNS = {
    "number": ["ID"],
    "name": ["Subject", "Title"],
    "severity": ["Priority", "Severity"],
    "status": ["Status"],
    "request_type": ["Request Type", "Type"],
    "module": ["Subcategory", "Module"],
    # Dynamics record GUID — used to build the ticket deep link
    "guid": ["(Do Not Modify) All Requests ID", "All Requests ID"],
}

# Deep link into Dynamics 365 for a ticket record ('{guid}' is replaced
# with the ticket's '(Do Not Modify) All Requests ID' value)
TICKET_URL_TEMPLATE = (
    "https://org2f45e702.crm4.dynamics.com/main.aspx"
    "?appid=1176d8e9-1da0-ef11-8a6a-6045bd94a6e7"
    "&forceUCI=1&pagetype=entityrecord&etn=cr603_allrequests&id={guid}"
)

# Only tickets matching these are tracked in the analysis:
# - Request Type must be one of TICKET_REQUEST_TYPES (exact, case-insensitive)
# - Status must NOT contain any TICKET_EXCLUDED_STATUS_KEYWORDS (Rejected and
#   On hold are dropped; Open / New / Committed / Resolved / Closed all stay —
#   resolved tickets count as Done; resolved tickets without any roadmap
#   reference are filtered out later)
TICKET_REQUEST_TYPES = ["CR"]
TICKET_EXCLUDED_STATUS_KEYWORDS = ["rejected", "on hold"]

# Ticket numbers referenced by the roadmap but missing from the ticket export
# are listed on the dashboard up to this many
TICKET_MISSING_LIST_MAX = 20

DETAILS_BATCH = 200
PARENTS_BATCH = 200
MAX_WORKERS = 6  # parallel batch downloads (network round-trips are the bottleneck)

grandparent_assignee_lookup = {}

# Status ordering for sorting (lower = higher priority in display)
STATUS_ORDER = {
    "Done": 0,
    "Testing": 1,
    "Development": 2,
    "Backlog": 3,
}


# ============================================================
# AUTH
# ============================================================

def get_auth_header():
    token = f":{PAT}"
    encoded = base64.b64encode(token.encode("utf-8")).decode("utf-8")
    return {"Authorization": f"Basic {encoded}", "Content-Type": "application/json"}


# Shared HTTP session with keep-alive + automatic retries. Transient DNS
# failures ('Failed to resolve' / getaddrinfo errors, common when several
# batches resolve at once) and blips are retried with backoff instead of
# falling through to the per-item path.
_HTTP = requests.Session()
_RETRY = Retry(total=3, connect=3, read=3, backoff_factor=0.5,
               status_forcelist=(502, 503, 504))
_HTTP.mount("http://", HTTPAdapter(max_retries=_RETRY))
_HTTP.mount("https://", HTTPAdapter(max_retries=_RETRY))


# ============================================================
# TFS API
# ============================================================

def run_wiql(query, top=5000):
    url = f"{TFS_URL}/_apis/wit/wiql?api-version={API_VERSION}"
    payload = {"query": query, "top": top}
    resp = _HTTP.post(url, json=payload, headers=get_auth_header(), timeout=60)
    resp.raise_for_status()
    return resp.json().get("workItems", [])


def fetch_work_items_with_relations(ids):
    if not ids:
        return []
    batches = [ids[i:i + DETAILS_BATCH] for i in range(0, len(ids), DETAILS_BATCH)]
    total = len(batches)
    print(f"      Fetching {len(ids)} items in {total} batches "
          f"(batch size {DETAILS_BATCH}, up to {min(MAX_WORKERS, total)} in parallel)...")

    def fetch_batch(batch_num, batch):
        ids_param = ",".join(str(x) for x in batch)
        url = (
            f"{TFS_URL}/_apis/wit/workitems"
            f"?ids={ids_param}&$expand=relations&api-version={API_VERSION}"
        )
        try:
            resp = _HTTP.get(url, headers=get_auth_header(), timeout=120)
            if resp.status_code == 404:
                return []
            resp.raise_for_status()
            return resp.json().get("value", [])
        except Exception as e:
            print(f"      ERROR in batch {batch_num}: {e}")
            print(f"      Retrying individually...")
            out = []
            for wid in batch:
                try:
                    url2 = f"{TFS_URL}/_apis/wit/workItems/{wid}?$expand=relations&api-version={API_VERSION}"
                    r2 = _HTTP.get(url2, headers=get_auth_header(), timeout=30)
                    if r2.status_code == 200:
                        out.append(r2.json())
                except Exception:
                    print(f"      Skipped ID {wid}")
            return out

    all_items = []
    with ThreadPoolExecutor(max_workers=min(MAX_WORKERS, total)) as executor:
        futures = [executor.submit(fetch_batch, n, b) for n, b in enumerate(batches, 1)]
        for fut in as_completed(futures):
            all_items.extend(fut.result())
    return all_items


def fetch_work_items_basic(ids):
    if not ids:
        return {}
    batches = [ids[i:i + PARENTS_BATCH] for i in range(0, len(ids), PARENTS_BATCH)]

    def fetch_batch(batch):
        ids_param = ",".join(str(x) for x in batch)
        # No 'fields' restriction: fetches all populated fields so the
        # Custom.* roadmap fields come through on any work item type.
        url = (
            f"{TFS_URL}/_apis/wit/workitems"
            f"?ids={ids_param}&api-version={API_VERSION}"
        )
        try:
            resp = _HTTP.get(url, headers=get_auth_header(), timeout=120)
            if resp.status_code == 404:
                return []
            resp.raise_for_status()
            return resp.json().get("value", [])
        except Exception as e:
            print(f"      ERROR fetching parents batch: {e}")
            return []

    lookup = {}
    with ThreadPoolExecutor(max_workers=min(MAX_WORKERS, len(batches))) as executor:
        for value in executor.map(fetch_batch, batches):
            for wi in value:
                flds = wi.get("fields", {})
                lookup[wi["id"]] = {
                    "title": flds.get("System.Title", ""),
                    "type": flds.get("System.WorkItemType", ""),
                    "assigned_to": extract_display_name(flds.get("System.AssignedTo")),
                    "new_values": extract_new_values(flds),
                }
    return lookup


# ============================================================
# RELATIONS PARSING
# ============================================================

def extract_parent_id(relations):
    if not relations:
        return None
    for rel in relations:
        if rel.get("rel") == "System.LinkTypes.Hierarchy-Reverse":
            url = rel.get("url", "")
            parts = url.rstrip("/").split("/")
            try:
                return int(parts[-1])
            except (ValueError, IndexError):
                return None
    return None


def extract_child_ids(relations):
    if not relations:
        return []
    children = []
    for rel in relations:
        if rel.get("rel") == "System.LinkTypes.Hierarchy-Forward":
            url = rel.get("url", "")
            parts = url.rstrip("/").split("/")
            try:
                children.append(int(parts[-1]))
            except (ValueError, IndexError):
                pass
    return children


# ============================================================
# FIELD MAPPING HELPERS
# ============================================================

def strip_html(html_text):
    if not html_text:
        return ""
    clean = re.sub(r"<[^>]+>", " ", html_text)
    clean = clean.replace("&nbsp;", " ").replace("&amp;", "&")
    clean = clean.replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"')
    clean = clean.replace("&#39;", "'")
    return re.sub(r"\s+", " ", clean).strip()


def extract_display_name(identity_field):
    if not identity_field:
        return ""
    if isinstance(identity_field, dict):
        return identity_field.get("displayName", "")
    return str(identity_field)


def format_date(date_string):
    if not date_string:
        return ""
    try:
        dt = datetime.fromisoformat(date_string.replace("Z", "+00:00"))
        return dt.strftime("%Y-%m-%d")
    except:
        return date_string


def to_number(value):
    """Coerce a cell value to float for weight sums (None when not numeric)."""
    if value in ("", None):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def resolve_standard_weight(fields):
    """Standard Weight — falls back to the item's Effort when the
    Standard Weight field is empty; empty string when neither is set."""
    std_weight = fields.get(STANDARD_WEIGHT_REF)
    if std_weight in (None, ""):
        std_weight = fields.get(EFFORT_REF)
    return std_weight if std_weight not in (None, "") else ""


def parse_date_for_cutoff(date_string):
    if not date_string:
        return None
    try:
        dt = datetime.fromisoformat(date_string.replace("Z", "+00:00"))
        return dt.replace(tzinfo=None)
    except:
        return None


def is_allowed_owner(display_name):
    if not display_name:
        return False
    name_lower = display_name.lower()
    for allowed in ALLOWED_OWNERS:
        if allowed in name_lower:
            return True
    return False


def map_status(state, working_status):
    """Map TFS state + Custom.WorkingStatus to roadmap status.
    Only 4 statuses: Backlog, Development, Testing, Done.
    - New/Approved -> Backlog
    - Committed + NOT testing -> Development
    - Committed + Testing -> Testing
    - Done -> Done
    """
    state_lower = (state or "").lower().strip()
    ws_lower = (working_status or "").lower().strip()

    if state_lower in ("new", "proposed", "approved"):
        return "Backlog"

    if state_lower in ("done", "closed"):
        return "Done"

    if state_lower in ("committed", "active", "in progress"):
        if "test" in ws_lower:
            return "Testing"
        return "Development"

    return "Backlog"


def build_business_area_with_grandparent(parent_id, parent_lookup, grandparent_lookup):
    if not parent_id or parent_id not in parent_lookup:
        return ""
    parent_info = parent_lookup[parent_id]
    parent_title = parent_info["title"]
    parent_type = parent_info["type"]
    if parent_type == "Epic":
        return parent_title
    epic_title = grandparent_lookup.get(parent_id)
    if epic_title:
        return f"{epic_title} - {parent_title}"
    return parent_title


def extract_new_values(fields):
    """Pull the new TFS roadmap fields (ticket number, business impact
    value, category, impact level) from a work item's fields dict."""
    vals = {}
    for col, ref in NEW_FIELD_REFS.items():
        raw = fields.get(ref) or ""
        if isinstance(raw, dict):
            raw = raw.get("displayName", "")
        if ref in HTML_NEW_FIELDS:
            raw = strip_html(str(raw))
        vals[col] = str(raw).strip()
    return vals


def resolve_new_values(fields, parent_id, parent_lookup):
    """Story-level values override feature-level values. A story with no
    value of its own inherits the parent Feature's value for each field."""
    own = extract_new_values(fields)
    parent = parent_lookup.get(parent_id, {}).get("new_values", {}) if parent_id else {}
    return {
        col: (own.get(col, "") or parent.get(col, ""))
        for col in NEW_FIELD_REFS
    }


# ============================================================
# FETCH ALL WORK ITEMS
# ============================================================

def fetch_all_pbis():
    types_filter = ", ".join(f"'{t}'" for t in PBI_TYPES)
    projects_filter = ", ".join(f"'{p}'" for p in PROJECTS)
    wiql = (
        f"SELECT [System.Id] FROM WorkItems "
        f"WHERE [System.TeamProject] IN ({projects_filter}) "
        f"AND [System.WorkItemType] IN ({types_filter}) "
        f"AND [System.State] <> 'Removed' "
        # Server-side pre-filter: items Done before the roadmap cutoff are
        # discarded during mapping anyway — never download them at all.
        # (Delivery Date falls back to StateChangeDate for edge cases, so
        # those items are still fetched and filtered locally.)
        f"AND NOT ([System.State] IN ('Done', 'Closed') "
        f"AND [Microsoft.VSTS.Common.ClosedDate] < '{ROADMAP_CUTOFF_DATE}') "
        f"ORDER BY [System.Id] DESC"
    )
    print(f"\n[1/5] Querying {PROJECTS} for {PBI_TYPES} (excluding Removed, "
          f"and Done before {ROADMAP_CUTOFF_DATE})...")
    refs = run_wiql(wiql)
    print(f"      Found {len(refs)} work items")
    if not refs:
        return []
    all_ids = [r["id"] for r in refs]
    print(f"\n      Fetching PBIs with relations...")
    items = fetch_work_items_with_relations(all_ids)
    print(f"      Fetched {len(items)} PBIs with relations")
    return items


def fetch_all_features_and_epics():
    types_filter = ", ".join(f"'{t}'" for t in FEATURE_TYPES)
    projects_filter = ", ".join(f"'{p}'" for p in PROJECTS)
    wiql = (
        f"SELECT [System.Id] FROM WorkItems "
        f"WHERE [System.TeamProject] IN ({projects_filter}) "
        f"AND [System.WorkItemType] IN ({types_filter}) "
        f"AND [System.State] <> 'Removed' "
        f"ORDER BY [System.Id] DESC"
    )
    print(f"\n[2/5] Querying {PROJECTS} for Features and Epics...")
    refs = run_wiql(wiql)
    print(f"      Found {len(refs)} Features/Epics")
    if not refs:
        return []
    all_ids = [r["id"] for r in refs]
    print(f"\n      Fetching Features/Epics with relations...")
    items = fetch_work_items_with_relations(all_ids)
    print(f"      Fetched {len(items)} items")
    orphans = []
    for wi in items:
        relations = wi.get("relations", [])
        children = extract_child_ids(relations)
        if len(children) == 0:
            orphans.append(wi)
    print(f"      {len(orphans)} orphan Features/Epics (no children)")
    return orphans


# ============================================================
# PARENT + GRANDPARENT LOOKUP
# ============================================================

def build_parent_and_grandparent_lookup(items, orphan_features):
    parent_ids = set()
    for wi in items:
        relations = wi.get("relations", [])
        pid = extract_parent_id(relations)
        if pid:
            parent_ids.add(pid)
    orphan_parent_ids = set()
    for wi in orphan_features:
        relations = wi.get("relations", [])
        pid = extract_parent_id(relations)
        if pid:
            orphan_parent_ids.add(pid)
    all_parent_ids = parent_ids | orphan_parent_ids
    print(f"      Fetching {len(all_parent_ids)} parent work items...")
    parent_lookup = fetch_work_items_basic(list(all_parent_ids))
    print(f"      {len(parent_lookup)} parents loaded")
    grandparent_lookup = {}
    gp_assignee = {}
    feature_ids_to_check = []
    for pid, info in parent_lookup.items():
        if info["type"] == "Feature":
            feature_ids_to_check.append(pid)
    if feature_ids_to_check:
        print(f"      Fetching {len(feature_ids_to_check)} Feature parents with relations...")
        feature_items = fetch_work_items_with_relations(feature_ids_to_check)
        grandparent_ids = set()
        feature_to_grandparent_id = {}
        for fi in feature_items:
            fid = fi.get("id")
            relations = fi.get("relations", [])
            gpid = extract_parent_id(relations)
            if gpid:
                feature_to_grandparent_id[fid] = gpid
                grandparent_ids.add(gpid)
            else:
                feature_to_grandparent_id[fid] = None
        if grandparent_ids:
            print(f"      Fetching {len(grandparent_ids)} Epic grandparents...")
            grandparent_basic = fetch_work_items_basic(list(grandparent_ids))
            for fid, gpid in feature_to_grandparent_id.items():
                if gpid and gpid in grandparent_basic:
                    grandparent_lookup[fid] = grandparent_basic[gpid]["title"]
                    gp_assignee[fid] = grandparent_basic[gpid].get("assigned_to", "")
                else:
                    grandparent_lookup[fid] = None
                    gp_assignee[fid] = None
        else:
            for fid in feature_ids_to_check:
                grandparent_lookup[fid] = None
                gp_assignee[fid] = None
    return parent_lookup, grandparent_lookup, gp_assignee


# ============================================================
# ROW MAPPING
# ============================================================

def map_pbi_to_row(wi, parent_lookup, grandparent_lookup):
    fields = wi.get("fields", {})
    relations = wi.get("relations", [])
    pbi_assignee = extract_display_name(fields.get("System.AssignedTo"))
    module = fields.get("Custom.Module", "")
    parent_id = extract_parent_id(relations)
    feature = build_business_area_with_grandparent(parent_id, parent_lookup, grandparent_lookup)

    # Unparented PBI — becomes its own feature group (same treatment as an orphan feature)
    own_group = not feature
    if own_group:
        feature = fields.get("System.Title", "")

    # Owner: check PBI -> parent Feature -> grandparent Epic
    owner = ""
    if pbi_assignee and is_allowed_owner(pbi_assignee):
        owner = pbi_assignee
    elif parent_id and parent_id in parent_lookup:
        parent_assignee = parent_lookup[parent_id].get("assigned_to", "")
        if parent_assignee and is_allowed_owner(parent_assignee):
            owner = parent_assignee
        else:
            gp_assignee = grandparent_assignee_lookup.get(parent_id, "")
            if gp_assignee and is_allowed_owner(gp_assignee):
                owner = gp_assignee

    requirement = fields.get("System.Title", "")
    ref_id = str(wi.get("id", ""))
    sync_id = str(wi.get("id", ""))
    state = fields.get("System.State", "")
    working_status = fields.get("Custom.WorkingStatus", "")
    status = map_status(state, working_status)

    # Added on = PBI created date
    created_date = fields.get("System.CreatedDate", "")
    added_on = format_date(created_date)

    # Start date = ActivatedDate
    start_date = ""
    activated = fields.get("Microsoft.VSTS.Common.ActivatedDate", "")
    if activated:
        start_date = format_date(activated)

    # Delivery date
    delivery_date = ""
    closed = fields.get("Microsoft.VSTS.Common.ClosedDate", "")
    if closed:
        delivery_date = format_date(closed)
    elif state_lower_check(state):
        state_change = fields.get("Microsoft.VSTS.Common.StateChangeDate", "")
        if state_change:
            delivery_date = format_date(state_change)

    # Parent ID for grouping
    parent_id_str = str(parent_id) if parent_id else ""

    # New fields: story's own values, else inherited from parent Feature
    new_vals = resolve_new_values(fields, parent_id, parent_lookup)

    return {
        "Owner": owner,
        "Module": module,
        "Business Area / Feature": feature,
        "Requirement": requirement,
        "Azure ID": ref_id,
        "Standard Weight": resolve_standard_weight(fields),
        "Ticket Number": new_vals.get("Ticket Number", ""),
        "Impact": new_vals.get("Impact", ""),
        "Status": status,
        "Added on": added_on,
        "Start Date": start_date,
        "Done Date": delivery_date,
        "Category": new_vals.get("Category", ""),
        "Notes": new_vals.get("Notes", ""),
        "_id": sync_id,
        "_item_type": "story",
        "_created_date": created_date,
        "_parent_id": parent_id_str,
        "_project": fields.get("System.TeamProject", ""),
        "_own_group": own_group,
    }


def map_orphan_to_row(wi, parent_lookup, grandparent_lookup):
    fields = wi.get("fields", {})
    relations = wi.get("relations", [])
    title = fields.get("System.Title", "")
    owner = extract_display_name(fields.get("System.AssignedTo"))
    module = fields.get("Custom.Module", "")
    parent_id = extract_parent_id(relations)

    if parent_id and parent_id in parent_lookup:
        parent_info = parent_lookup[parent_id]
        parent_title = parent_info["title"]
        parent_type = parent_info["type"]
        if parent_type == "Epic":
            business_area = parent_title
        elif parent_type == "Feature":
            epic_title = grandparent_lookup.get(parent_id)
            if epic_title:
                business_area = f"{epic_title} - {parent_title}"
            else:
                business_area = parent_title
        else:
            business_area = parent_title
    else:
        business_area = title

    created_date = fields.get("System.CreatedDate", "")

    # New fields: the feature's own values (inherited from its parent, if any)
    new_vals = resolve_new_values(fields, parent_id, parent_lookup)

    return {
        "Owner": owner,
        "Module": module,
        "Business Area / Feature": business_area,
        "Requirement": title,
        "Azure ID": "",
        "Standard Weight": resolve_standard_weight(fields),
        "Ticket Number": new_vals.get("Ticket Number", ""),
        "Impact": new_vals.get("Impact", ""),
        "Status": "Backlog",
        "Added on": format_date(created_date),
        "Start Date": "",
        "Done Date": "",
        "Category": new_vals.get("Category", ""),
        "Notes": new_vals.get("Notes", ""),
        "_id": str(wi.get("id", "")),
        "_item_type": "feature",
        "_created_date": created_date,
        "_parent_id": "",
        "_project": fields.get("System.TeamProject", ""),
        "_own_group": False,
    }


def state_lower_check(state):
    return (state or "").lower().strip() in ("done", "closed")


# ============================================================
# FEATURE GROUP STATUS
# ============================================================

def compute_feature_group_status(rows):
    """For each Business Area, compute the overall feature status.
    Done = all stories are Done.
    Testing = at least one story in Testing, rest Done/Testing.
    Development = at least one in Development, rest Done/Testing/Development.
    Backlog = at least one in Backlog.
    Returns: {business_area: status}
    """
    area_statuses = defaultdict(list)
    for r in rows:
        ba = r.get("Business Area / Feature", "")
        if ba:
            area_statuses[ba].append(r.get("Status", "Backlog"))

    result = {}
    for ba, statuses in area_statuses.items():
        if all(s == "Done" for s in statuses):
            result[ba] = "Done"
        elif any(s == "Testing" for s in statuses):
            result[ba] = "Testing"
        elif any(s == "Development" for s in statuses):
            result[ba] = "Development"
        else:
            result[ba] = "Backlog"
    return result


# ============================================================
# TICKET SYSTEM ANALYSIS
# ============================================================

def normalize_ticket_id(value):
    """Normalize a ticket number to a plain string ('5487', whether it
    comes as int, float or str)."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    if isinstance(value, int):
        return str(value)
    s = str(value).strip()
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    if s.isdigit():
        return str(int(s))
    return s


def parse_ticket_numbers(value):
    """Extract ticket numbers from a Ticket Number field value.
    Handles HTML remnants, commas, spaces and multiple numbers
    (e.g. '4537, 3004'). Returns unique normalized numbers, in order."""
    if value is None:
        return []
    text = strip_html(str(value))
    if not text:
        return []
    result = []
    for num in re.findall(r"\d+", text):
        n = normalize_ticket_id(num)
        if n and n not in result:
            result.append(n)
    return result


def severity_sort_key(severity):
    """Sort key for severities: SEVERITY_ORDER keywords first (by name),
    unknown values last (alphabetically)."""
    s = str(severity or "").lower()
    for i, name in enumerate(SEVERITY_ORDER):
        if name.lower() in s:
            return (0, i, s)
    return (1, 1, s)


def ticket_number_sort_key(number):
    """Numeric sort for ticket numbers ('3004' < '10000' < 'abc')."""
    n = str(number or "")
    return (0, int(n), "") if n.isdigit() else (1, 0, n)


def resolve_ticket_priority(ticket_number_value, number_to_priority):
    """Priority of the ticket(s) referenced by a Ticket Number value.
    When several tickets are linked, the HIGHEST-RATED ticket's priority
    wins (Critical > High > Medium > Normal, unknown values last).
    Returns '' when no referenced ticket is found in the export."""
    priorities = [number_to_priority.get(n) for n in parse_ticket_numbers(ticket_number_value)]
    priorities = [p for p in priorities if p]
    if not priorities:
        return ""
    return sorted(priorities, key=severity_sort_key)[0]


def find_latest_ticket_export():
    """Locate the ticket system export Excel.
    TICKET_EXPORT_FOLDER can be a direct file path, or a folder in which
    the newest file matching TICKET_EXPORT_PATTERN is used.
    Returns the path or None if not found/disabled."""
    if not TICKET_EXPORT_FOLDER:
        return None
    path = os.path.expanduser(TICKET_EXPORT_FOLDER)
    if os.path.isfile(path):
        return path
    if not os.path.isdir(path):
        return None
    try:
        files = [
            f for f in os.listdir(path)
            if fnmatch.fnmatch(f, TICKET_EXPORT_PATTERN)
            and f.lower().endswith(".xlsx")
            and not f.startswith("~$")
        ]
    except OSError:
        return None
    if not files:
        return None
    files.sort(key=lambda f: os.path.getmtime(os.path.join(path, f)), reverse=True)
    return os.path.join(path, files[0])


def load_ticket_export(path):
    """Read the ticket system export and return a list of ticket dicts:
    {number, name, severity, status}. Columns are located by header name.
    Returns None if no sheet with the expected headers is found."""
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            twb = load_workbook(path, read_only=True, data_only=True)
    except Exception as e:
        print(f"      ERROR reading ticket export: {e}")
        return None

    try:
        for sheet_name in twb.sheetnames:
            if "hidden" in sheet_name.lower():
                continue
            sheet = twb[sheet_name]
            header = [
                (str(c.value).strip() if c.value is not None else "")
                for c in next(sheet.iter_rows(max_row=1))
            ]
            if not header:
                continue
            col_map = {}
            for field, aliases in TICKET_EXPORT_COLUMNS.items():
                for idx, h in enumerate(header):
                    if h.lower() in [a.lower() for a in aliases]:
                        col_map[field] = idx
                        break
            if "number" in col_map and "name" in col_map:
                tickets = []
                get = lambda row, key: (row[col_map[key]] if key in col_map and col_map[key] < len(row) else None)
                for row in sheet.iter_rows(min_row=2, values_only=True):
                    number = normalize_ticket_id(get(row, "number"))
                    if not number:
                        continue
                    tickets.append({
                        "number": number,
                        "name": str(get(row, "name") or ""),
                        "severity": str(get(row, "severity") or "") or "Unknown",
                        "status": str(get(row, "status") or "") or "Unknown",
                        "request_type": str(get(row, "request_type") or "").strip(),
                        "module": str(get(row, "module") or "").strip() or "Unknown",
                        "guid": str(get(row, "guid") or "").strip(),
                    })
                return tickets
        return None
    finally:
        twb.close()


def is_tracked_ticket(ticket):
    """A ticket is tracked if its Request Type is one of TICKET_REQUEST_TYPES
    and its Status does NOT contain any TICKET_EXCLUDED_STATUS_KEYWORDS.
    (Emoji prefixes are ignored — '🚩Open' counts as 'Open'. Resolved and
    Closed tickets stay tracked: they count as Done when linked to the
    roadmap and are dropped entirely when not linked.)"""
    allowed_types = [t.lower() for t in TICKET_REQUEST_TYPES]
    rt = str(ticket.get("request_type", "")).strip().lower()
    if allowed_types and rt not in allowed_types:
        return False
    st = str(ticket.get("status", "")).strip().lower()
    if any(k in st for k in [s.lower() for s in TICKET_EXCLUDED_STATUS_KEYWORDS]):
        return False
    return True


def build_ticket_analysis(rows):
    """Match the roadmap's ticket numbers (Custom.Ticketnumber — the
    lowest-level PBI value overrides the Feature value, empty story
    inherits the Feature's value) against the ticket system export.

    Tracked tickets are CR tickets that are not Rejected / On hold.
    Per ticket, Progress = share of linked roadmap stories that are Done.
    Resolved/closed tickets count as Done when linked to the roadmap and
    are dropped entirely when they have no roadmap reference.

    Returns a dict with the export path, per-ticket covered flags and
    progress, per-severity/per-module progress stats and tickets
    referenced by the roadmap but missing from the export — or None if
    the export is unavailable."""
    export_path = find_latest_ticket_export()
    if not export_path:
        print("      No ticket export found — ticket analysis skipped")
        return None
    print(f"      Ticket export: {os.path.basename(export_path)}")

    tickets = load_ticket_export(export_path)
    if tickets is None:
        print("      WARNING: ticket export has no recognizable sheet — analysis skipped")
        return None

    # Filter to tracked tickets (CR, not Rejected / On hold)
    tracked = [t for t in tickets if is_tracked_ticket(t)]
    excluded = len(tickets) - len(tracked)

    # Ticket number -> priority lookup, built from ALL tickets in the
    # export (not just tracked ones) so that resolved/inactive tickets
    # referenced by the roadmap still resolve to their priority
    number_to_priority = {}
    number_to_guid = {}
    for t in tickets:
        if t["number"] not in number_to_priority:
            number_to_priority[t["number"]] = t.get("severity") or ""
        if t["number"] not in number_to_guid and t.get("guid"):
            number_to_guid[t["number"]] = t["guid"]

    if excluded:
        print(f"      {excluded} tickets excluded (Request Type not in "
              f"{TICKET_REQUEST_TYPES} or Rejected / On hold)")
    if not tracked:
        print("      WARNING: no tracked tickets in the export — analysis skipped")
        return None

    # Ticket numbers covered by the roadmap = numbers referenced by any
    # roadmap row's resolved Ticket Number (story override / feature inherit)
    covered_set = set()
    for r in rows:
        for num in parse_ticket_numbers(r.get("Ticket Number", "")):
            covered_set.add(num)

    # PBI progress per ticket number: how many of the linked stories are Done
    ticket_pbis = defaultdict(lambda: {"total": 0, "done": 0})
    for r in rows:
        if r.get("_item_type") != "story":
            continue
        for num in parse_ticket_numbers(r.get("Ticket Number", "")):
            stats = ticket_pbis[num]
            stats["total"] += 1
            if r.get("Status") == "Done":
                stats["done"] += 1

    # Progress per ticket; resolved/closed tickets without any roadmap
    # reference are dropped entirely
    kept = []
    dropped_resolved = 0
    for t in tracked:
        st = str(t.get("status", "")).strip().lower()
        t["resolved"] = "resolved" in st or "closed" in st
        if t["resolved"] and t["number"] not in covered_set:
            dropped_resolved += 1
            continue
        if t["resolved"]:
            t["progress"] = 1.0  # resolved/closed counts as fully completed
        else:
            stats = ticket_pbis.get(t["number"], {"total": 0, "done": 0})
            t["progress"] = (stats["done"] / stats["total"]) if stats["total"] else None
        kept.append(t)
    tracked = kept
    excluded += dropped_resolved
    if dropped_resolved:
        print(f"      {dropped_resolved} resolved/closed tickets without roadmap references dropped")

    per_severity = {}
    per_module = {}
    covered_total = 0
    cat_counts = {"Done": 0, "In Progress": 0, "Not Started": 0, "Not Covered": 0}

    def progress_category(covered, progress):
        if not covered:
            return "Not Covered"
        if progress is not None and progress >= 1:
            return "Done"
        if progress is not None and progress > 0:
            return "In Progress"
        return "Not Started"

    for t in tracked:
        covered = t["covered"] = t["number"] in covered_set
        cat = t["category"] = progress_category(covered, t["progress"])
        cat_counts[cat] += 1
        if covered:
            covered_total += 1

        sev = t.get("severity") or "Unknown"
        stats = per_severity.setdefault(sev, {"total": 0, "Done": 0, "In Progress": 0,
                                              "Not Started": 0, "Not Covered": 0})
        stats["total"] += 1
        stats[cat] += 1

        mod = t.get("module") or "Unknown"
        stats = per_module.setdefault(mod, {"total": 0, "covered": 0})
        stats["total"] += 1
        if covered:
            stats["covered"] += 1

    export_ids = {t["number"] for t in tickets}
    missing_from_export = sorted(covered_set - export_ids, key=ticket_number_sort_key)

    print(f"      {len(tracked)} tracked tickets — Done: {cat_counts['Done']}, "
          f"In Progress: {cat_counts['In Progress']}, Not Started: {cat_counts['Not Started']}, "
          f"Not Covered: {cat_counts['Not Covered']}")
    if missing_from_export:
        print(f"      {len(missing_from_export)} ticket numbers referenced in the roadmap are not in the export")

    return {
        "export_path": export_path,
        "tickets": tracked,
        "tracked_total": len(tracked),
        "excluded_total": excluded,
        "per_severity": per_severity,
        "per_module": per_module,
        "total": len(tracked),
        "covered_total": covered_total,
        "not_covered_total": len(tracked) - covered_total,
        "cat_counts": cat_counts,
        "missing_from_export": missing_from_export,
        "roadmap_ticket_count": len(covered_set),
        "number_to_priority": number_to_priority,
        "number_to_guid": number_to_guid,
    }


# ============================================================
# EXCEL GENERATION
# ============================================================

def generate_excel(rows):
    print(f"\n      Generating Excel with {len(rows)} rows...")

    # Compute feature group status for coloring only (Done groups stay green)
    feature_status = compute_feature_group_status(rows)

    # Sort: by Business Area, then by status within the group, then title.
    # Feature groups are NOT pulled to the top anymore — colors only.
    rows.sort(key=lambda r: (
        r.get("Business Area / Feature", "") or "~",
        STATUS_ORDER.get(r.get("Status", "Backlog"), 99),
        r.get("Requirement", "") or "",
    ))

    # --- New items computation (shared by Dashboard + New Stories sheet) ---
    cutoff_dt = parse_date_for_cutoff(CUTOFF_DATE + "T00:00:00+00:00")
    new_detail_rows = []
    new_stories = 0
    new_orphan_features = 0
    features_with_new_stories = {}  # area -> {project, owner} of its first new story

    for r in rows:
        created = r.get("_created_date", "")
        created_dt = parse_date_for_cutoff(created)
        if not (created_dt and cutoff_dt and created_dt > cutoff_dt):
            continue

        if r.get("_item_type") == "story":
            new_stories += 1
            ba = r.get("Business Area / Feature", "")
            if ba and not r.get("_own_group") and ba not in features_with_new_stories:
                features_with_new_stories[ba] = {
                    "project": r.get("_project", ""),
                    "owner": r.get("Owner", ""),
                }
        else:
            new_orphan_features += 1

        new_detail_rows.append({
            "Owner": r.get("Owner", ""),
            "Project": r.get("_project", ""),
            "Module": r.get("Module", ""),
            "Feature": r.get("Business Area / Feature", ""),
            "Story Title": r.get("Requirement", ""),
            "Story ID": r.get("Azure ID", ""),
            "Ticket Number": r.get("Ticket Number", ""),
            "Status": r.get("Status", ""),
            "Added on": r.get("Added on", ""),
            "_id": r.get("_id", ""),
        })

    # Order: Owner > Project > Module > Feature > Story Title > Story number
    new_detail_rows.sort(key=lambda d: (
        d["Owner"], d["Project"], d["Module"], d["Feature"],
        d["Story Title"], d["_id"],
    ))

    # --- Completed since the new-items cutoff (stories set as Done) ---
    completed_detail_rows = []
    completed_counts = {}  # (project, owner) -> count
    for r in rows:
        if r.get("Status") != "Done" or r.get("_item_type") != "story":
            continue
        delivery = r.get("Done Date", "")
        delivery_dt = parse_date_for_cutoff(delivery + "T00:00:00+00:00") if delivery else None
        if not (delivery_dt and cutoff_dt and delivery_dt > cutoff_dt):
            continue
        k = (r.get("_project", ""), r.get("Owner", ""))
        completed_counts[k] = completed_counts.get(k, 0) + 1
        completed_detail_rows.append({
            "Owner": r.get("Owner", ""),
            "Project": r.get("_project", ""),
            "Module": r.get("Module", ""),
            "Feature": r.get("Business Area / Feature", ""),
            "Story Title": r.get("Requirement", ""),
            "Story ID": r.get("Azure ID", ""),
            "Standard Weight": r.get("Standard Weight", ""),
            "Ticket Number": r.get("Ticket Number", ""),
            "Done Date": delivery,
            "_id": r.get("_id", ""),
        })

    completed_detail_rows.sort(key=lambda d: (
        d["Owner"], d["Project"], d["Module"], d["Feature"],
        d["Story Title"], d["_id"],
    ))

    new_data = {
        "new_stories": new_stories,
        "new_orphan_features": new_orphan_features,
        "features_with_new_stories": features_with_new_stories,
        "new_features_total": new_orphan_features + len(features_with_new_stories),
        "completed_counts": completed_counts,
        "completed_total": len(completed_detail_rows),
        "completed_detail_rows": completed_detail_rows,
    }

    wb = Workbook()

    # --- TICKET SYSTEM ANALYSIS (before sheets are built) ---
    print("\n      Ticket system analysis...")
    ticket_data = build_ticket_analysis(rows)

    # Ticket Priority for every row / detail row — from the ticket system
    # export; the highest-rated linked ticket wins (Critical > High > ...)
    ticket_priority_lookup = (ticket_data or {}).get("number_to_priority", {})
    for r in rows:
        r["Ticket Priority"] = resolve_ticket_priority(
            r.get("Ticket Number", ""), ticket_priority_lookup
        )
    for d in new_detail_rows:
        d["Ticket Priority"] = resolve_ticket_priority(
            d.get("Ticket Number", ""), ticket_priority_lookup
        )
    for d in completed_detail_rows:
        d["Ticket Priority"] = resolve_ticket_priority(
            d.get("Ticket Number", ""), ticket_priority_lookup
        )

    # Ticket deep links into Dynamics — first ticket listed in the cell
    number_to_guid = (ticket_data or {}).get("number_to_guid", {})

    def first_ticket_url(ticket_number_value):
        nums = parse_ticket_numbers(ticket_number_value)
        guid = number_to_guid.get(nums[0]) if nums else None
        return TICKET_URL_TEMPLATE.format(guid=guid) if guid else ""

    for r in rows:
        r["Ticket URL"] = first_ticket_url(r.get("Ticket Number", ""))
    for d in new_detail_rows:
        d["Ticket URL"] = first_ticket_url(d.get("Ticket Number", ""))
    for d in completed_detail_rows:
        d["Ticket URL"] = first_ticket_url(d.get("Ticket Number", ""))

    # --- DASHBOARD SHEET (created first = opens first) ---
    dash_ws = wb.active
    dash_ws.title = "Dashboard"
    build_dashboard(dash_ws, rows, feature_status, new_data, ticket_data)

    # --- ROADMAP SHEET ---
    ws = wb.create_sheet("Roadmap", 1)
    ws.sheet_properties.tabColor = "2F5496"

    # Column definitions
    columns = [
        ("Owner", 28),
        ("Module", 18),
        ("Business Area / Feature", 40),
        ("Requirement", 50),
        ("Azure ID", 12),
        ("Standard Weight", 14),
        ("Ticket Number", 14),
        ("Ticket Priority", 14),
        ("Impact", 12),
        ("Status", 14),
        ("Added on", 14),
        ("Start Date", 14),
        ("Done Date", 14),
        ("Category", 18),
        ("Notes", 35),
    ]

    header_font = Font(name="Segoe UI", bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    data_font = Font(name="Segoe UI", size=10)
    link_font = Font(name="Segoe UI", size=10, color="0563C1", underline="single")
    data_align = Alignment(vertical="top", wrap_text=True)
    weight_align = Alignment(vertical="top", horizontal="left")
    feature_font = Font(name="Segoe UI", size=10, bold=True)
    feature_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    # Feature group fully done = Done status green (feature column only)

    status_colors = {
        "Backlog": PatternFill(start_color="E3F2FD", end_color="E3F2FD", fill_type="solid"),
        "Development": PatternFill(start_color="FFF3E0", end_color="FFF3E0", fill_type="solid"),
        "Testing": PatternFill(start_color="FCE4EC", end_color="FCE4EC", fill_type="solid"),
        "Done": PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid"),
    }

    # Write header
    for col_idx, (col_name, col_width) in enumerate(columns, 1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border
        ws.column_dimensions[get_column_letter(col_idx)].width = col_width

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}1"

    # Write data rows
    col_names = [c[0] for c in columns]
    feature_col_idx = col_names.index("Business Area / Feature") + 1
    status_col_idx = col_names.index("Status") + 1
    weight_col_idx = col_names.index("Standard Weight") + 1
    num_cols = len(columns)

    for row_idx, row_data in enumerate(rows, 2):
        status = row_data.get("Status", "")
        status_fill = status_colors.get(status)
        ba = row_data.get("Business Area / Feature", "")
        group_status = feature_status.get(ba, "Backlog")
        is_done_group = (group_status == "Done")

        for col_idx, (col_name, _) in enumerate(columns, 1):
            value = row_data.get(col_name, "")
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font = data_font
            cell.alignment = weight_align if col_idx == weight_col_idx else data_align
            cell.border = thin_border

            # Azure ID links straight to the work item in TFS
            if col_name == "Azure ID" and row_data.get("Azure ID"):
                cell.hyperlink = (
                    f"{TFS_URL}/{quote(row_data.get('_project') or PROJECT)}/"
                    f"_workitems/edit/{row_data['Azure ID']}"
                )
                cell.font = link_font

            # Ticket Number links to the Dynamics ticket record (first ticket)
            if col_name == "Ticket Number" and row_data.get("Ticket URL"):
                cell.hyperlink = row_data["Ticket URL"]
                cell.font = link_font

            # Determine fill
            if col_idx == feature_col_idx:
                cell.font = feature_font
                cell.fill = status_colors["Done"] if is_done_group else feature_fill
            elif status_fill and col_idx == status_col_idx:
                cell.fill = status_fill

    # Merge feature cells
    print("      Merging feature cells...")
    start_row = 2
    current_feature = rows[0].get("Business Area / Feature", "") if rows else ""
    for i in range(1, len(rows)):
        feat = rows[i].get("Business Area / Feature", "")
        if feat != current_feature:
            end_row = i + 1
            if end_row > start_row:
                ws.merge_cells(
                    start_row=start_row, start_column=feature_col_idx,
                    end_row=end_row, end_column=feature_col_idx,
                )
            start_row = i + 2
            current_feature = feat
    if len(rows) > 0:
        end_row = len(rows) + 1
        if end_row > start_row:
            ws.merge_cells(
                start_row=start_row, start_column=feature_col_idx,
                end_row=end_row, end_column=feature_col_idx,
            )

    # --- NEW STORIES SHEET (filterable detail table) ---
    build_new_stories_sheet(wb, new_detail_rows)

    # --- COMPLETED STORIES SHEET (delivered to the PO since the cutoff) ---
    build_completed_stories_sheet(wb, new_data.get("completed_detail_rows", []))

    # --- TICKETS SHEET (ticket system analysis detail table) ---
    build_tickets_sheet(wb, ticket_data)

    # --- SUMMARY SHEET ---
    summary_ws = wb.create_sheet("Summary")
    summary_ws.sheet_properties.tabColor = "808080"
    summary_ws.column_dimensions["A"].width = 30
    summary_ws.column_dimensions["B"].width = 15
    summary_ws.column_dimensions["C"].width = 15
    summary_ws.column_dimensions["D"].width = 15
    summary_ws.column_dimensions["E"].width = 15
    summary_font = Font(name="Segoe UI", bold=True, size=14, color="2F5496")
    summary_ws.cell(row=1, column=1, value="Sync Summary").font = summary_font

    status_counts_summary = defaultdict(int)
    for r in rows:
        status_counts_summary[r.get("Status", "Unknown")] += 1
    owner_counts = defaultdict(int)
    for r in rows:
        owner_counts[r.get("Owner", "Unassigned")] += 1

    summary_data = [
        ("", ""),
        ("Generated At:", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        ("TFS Server:", TFS_URL),
        ("Projects:", ", ".join(PROJECTS)),
        ("Roadmap Start Date:", ROADMAP_CUTOFF_DATE),
        ("New Items Cutoff:", CUTOFF_DATE),
        ("Total Items:", len(rows)),
        ("", ""),
        ("--- By Status ---", ""),
    ]
    for s, c in sorted(status_counts_summary.items(), key=lambda x: -x[1]):
        summary_data.append((s, c))
    summary_data.append(("", ""))
    summary_data.append(("--- By Owner ---", ""))
    for o, c in sorted(owner_counts.items(), key=lambda x: -x[1]):
        summary_data.append((o, c))

    # --- Standard Weight by Owner (Done / Development / Testing) ---
    summary_data.append(("", ""))
    summary_data.append(("--- Standard Weight by Owner ---", "", "", ""))
    summary_data.append(("Owner", "Done", "Development", "Testing", "Total"))

    weight_by_owner = defaultdict(lambda: {"Done": 0.0, "Development": 0.0, "Testing": 0.0})
    for r in rows:
        w = to_number(r.get("Standard Weight"))
        if w is None:
            continue
        st = r.get("Status")
        if st in ("Done", "Development", "Testing"):
            weight_by_owner[r.get("Owner") or "Unassigned"][st] += w

    def fmt_weight(v):
        return int(v) if float(v).is_integer() else round(float(v), 1)

    weight_total = {"Done": 0.0, "Development": 0.0, "Testing": 0.0}
    for owner in sorted(weight_by_owner):
        sums = weight_by_owner[owner]
        for k in weight_total:
            weight_total[k] += sums[k]
        summary_data.append((owner, fmt_weight(sums["Done"]),
                             fmt_weight(sums["Development"]),
                             fmt_weight(sums["Testing"]),
                             fmt_weight(sums["Done"] + sums["Development"] + sums["Testing"])))
    summary_data.append(("Total", fmt_weight(weight_total["Done"]),
                         fmt_weight(weight_total["Development"]),
                         fmt_weight(weight_total["Testing"]),
                         fmt_weight(weight_total["Done"] + weight_total["Development"]
                                    + weight_total["Testing"])))

    for row_idx, vals in enumerate(summary_data, 3):
        for col_idx, val in enumerate(vals, 1):
            c = summary_ws.cell(row=row_idx, column=col_idx, value=val)
            c.font = Font(name="Segoe UI", bold=(col_idx == 1), size=11)

    # Output file name includes the generation date & time
    output_path = os.path.join(
        OUTPUT_FOLDER,
        f"roadmap_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.xlsx",
    )
    print(f"\n      Saving to: {output_path}")
    try:
        wb.save(output_path)
        print(f"      Done!")
        return output_path
    except PermissionError:
        # Same file name already open in Excel (re-run within the same minute) — add seconds
        alt = output_path.replace(".xlsx", f"_{datetime.now().strftime('%S')}.xlsx")
        wb.save(alt)
        print(f"      WARNING: '{output_path}' is locked (close it in Excel).")
        print(f"      Saved to fallback file instead: {alt}")
        return alt


# ============================================================
# DASHBOARD
# ============================================================

def build_dashboard(dash_ws, rows, feature_status, new_data, ticket_data=None):
    """Professional dashboard: overall + per-project status breakdown
    (all tables reconcile with each other), and new items broken down
    by project & owner. The filterable detail list lives on the
    'New Stories' sheet. Ticket system analysis (per-severity counts +
    covered/not covered) is added as the last section when available."""

    # --- Column widths ---
    widths = {"A": 30, "B": 26, "C": 14, "D": 14, "E": 14, "F": 14, "G": 24}
    for col, w in widths.items():
        dash_ws.column_dimensions[col].width = w

    # --- Style definitions ---
    thin_b = Border(
        left=Side(style="thin", color="D6D6D6"),
        right=Side(style="thin", color="D6D6D6"),
        top=Side(style="thin", color="D6D6D6"),
        bottom=Side(style="thin", color="D6D6D6"),
    )
    title_font = Font(name="Segoe UI", bold=True, size=20, color="1F3864")
    cutoff_font = Font(name="Segoe UI", bold=True, size=11, color="C00000")
    meta_font = Font(name="Segoe UI", size=9, color="808080")
    section_font = Font(name="Segoe UI", bold=True, size=11, color="FFFFFF")
    section_fill = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
    hdr_font = Font(name="Segoe UI", bold=True, size=10)
    hdr_fill = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
    hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    label_font = Font(name="Segoe UI", size=10)
    num_font = Font(name="Segoe UI", size=10)
    bold_font = Font(name="Segoe UI", size=10, bold=True)
    total_font = Font(name="Segoe UI", size=10, bold=True, color="1F3864")
    total_fill = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
    data_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")

    status_fills = {
        "Done": PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),
        "Testing": PatternFill(start_color="FFCCC7", end_color="FFCCC7", fill_type="solid"),
        "Development": PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid"),
        "Backlog": PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid"),
    }

    dash_ws.sheet_properties.tabColor = "1F3864"

    # --- Title block (roadmap cutoff shown prominently) ---
    c = dash_ws.cell(row=1, column=1, value="Roadmap Dashboard")
    c.font = title_font
    dash_ws.merge_cells("A1:G1")
    dash_ws.row_dimensions[1].height = 32

    c = dash_ws.cell(row=2, column=1, value=f"Roadmap cutoff: {ROADMAP_CUTOFF_DATE} — stories completed before this date are excluded")
    c.font = cutoff_font
    dash_ws.merge_cells("A2:G2")

    c = dash_ws.cell(row=3, column=1, value=f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}  |  Projects: {', '.join(PROJECTS)}  |  New items cutoff: {CUTOFF_DATE}")
    c.font = meta_font
    dash_ws.merge_cells("A3:G3")

    status_order = ["Done", "Testing", "Development", "Backlog"]
    cutoff_dt = parse_date_for_cutoff(CUTOFF_DATE + "T00:00:00+00:00")

    # --- Shared counts (single source of truth — every table reconciles) ---
    proj_stories = {p: defaultdict(int) for p in PROJECTS}
    for r in rows:
        if r.get("_item_type") == "story":
            p = r.get("_project", "")
            if p in proj_stories:
                proj_stories[p][r.get("Status", "Other")] += 1

    proj_areas = {p: set() for p in PROJECTS}
    for r in rows:
        if r.get("_own_group"):
            continue  # standalone unparented story — its group is not a feature
        p = r.get("_project", "")
        ba = r.get("Business Area / Feature", "")
        if p in proj_areas and ba:
            proj_areas[p].add(ba)

    story_status_totals = defaultdict(int)
    for p in PROJECTS:
        for s, cnt in proj_stories[p].items():
            story_status_totals[s] += cnt
    total_stories = sum(story_status_totals.values())
    total_feature_groups = sum(len(proj_areas[p]) for p in PROJECTS)

    # --- Small table helpers ---
    def write_section(r, title, width):
        c = dash_ws.cell(row=r, column=1, value=title)
        c.font = section_font
        c.fill = section_fill
        dash_ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=width)
        return r + 1

    def write_table_header(r, headers):
        for col_idx, h in enumerate(headers, 1):
            c = dash_ws.cell(row=r, column=col_idx, value=h)
            c.font = hdr_font
            c.fill = hdr_fill
            c.alignment = hdr_align
            c.border = thin_b
        return r + 1

    def write_data_row(r, vals, fills=None, total=False):
        for col_idx, val in enumerate(vals, 1):
            c = dash_ws.cell(row=r, column=col_idx, value=val)
            if total:
                c.font = total_font
                c.fill = total_fill
            elif col_idx == 1:
                c.font = bold_font
            else:
                c.font = num_font
            c.alignment = left_align if col_idx == 1 else data_align
            c.border = thin_b
            if fills and not total:
                f = fills.get(col_idx)
                if f:
                    c.fill = f
        return r + 1

    # --- Section 1: Overall Status Breakdown ---
    row = 5
    row = write_section(row, "Overall Status Breakdown", 3)
    row = write_table_header(row, ["Status", "Stories", "Feature Groups"])

    for status in status_order:
        s = story_status_totals.get(status, 0)
        f = 0
        for p in PROJECTS:
            for area in proj_areas[p]:
                if feature_status.get(area, "Backlog") == status:
                    f += 1
        fill = status_fills.get(status)
        row = write_data_row(row, [status, s, f],
                             fills={1: fill, 2: fill, 3: fill} if fill else None)
    row = write_data_row(row, ["Total", total_stories, total_feature_groups], total=True)

    # --- Section 2: Status Breakdown by Project (reconciles with Section 1) ---
    row += 2
    row = write_section(row, "Status Breakdown by Project", 7)
    row = write_table_header(row, ["Project", "Done", "Testing", "Development", "Backlog", "Stories", "Feature Groups"])

    for proj in PROJECTS:
        counts = proj_stories.get(proj, defaultdict(int))
        vals = [proj] + [counts.get(s, 0) for s in status_order] \
               + [sum(counts.values()), len(proj_areas.get(proj, set()))]
        row = write_data_row(row, vals)
    row = write_data_row(
        row,
        ["Total"] + [story_status_totals.get(s, 0) for s in status_order]
        + [total_stories, total_feature_groups],
        total=True,
    )

    # --- Section 3: New Items broken down by Project & Owner ---
    row += 2
    row = write_section(row, f"New Items After {CUTOFF_DATE} — by Project & Owner", 4)
    row = write_table_header(row, ["Project", "Owner", "New Stories", "New Features"])

    # Per (project, owner) counts — sums reconcile with the totals below
    po = {}  # (project, owner) -> [new stories, new features]
    for r in rows:
        created = r.get("_created_date", "")
        created_dt = parse_date_for_cutoff(created)
        if not (created_dt and cutoff_dt and created_dt > cutoff_dt):
            continue
        k = (r.get("_project", ""), r.get("Owner", ""))
        if k not in po:
            po[k] = [0, 0]
        if r.get("_item_type") == "story":
            po[k][0] += 1
        else:
            po[k][1] += 1  # new orphan feature

    # Features that gained new stories — attributed to the first new story's project/owner
    for info in new_data["features_with_new_stories"].values():
        k = (info.get("project", ""), info.get("owner", ""))
        if k not in po:
            po[k] = [0, 0]
        po[k][1] += 1

    proj_order = {p: i for i, p in enumerate(PROJECTS)}
    for k in sorted(po.keys(), key=lambda x: (proj_order.get(x[0], 99), x[1])):
        row = write_data_row(row, [k[0], k[1], po[k][0], po[k][1]])

    row = write_data_row(
        row,
        ["Total", "", new_data["new_stories"], new_data["new_features_total"]],
        total=True,
    )

    # Pointer to the filterable detail sheet
    row += 1
    c = dash_ws.cell(row=row, column=1,
                     value=f"Full filterable list: see the 'New Stories' sheet (items created after {CUTOFF_DATE})")
    c.font = Font(name="Segoe UI", size=9, italic=True, color="808080")
    dash_ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)

    # --- Section 4: Completed After Cutoff (counts by project & owner) ---
    row += 2
    row = write_section(row, f"Completed After {CUTOFF_DATE} — Stories Set as Done", 3)
    row = write_table_header(row, ["Project", "Owner", "Completed Stories"])

    completed_counts = new_data.get("completed_counts", {})
    for k in sorted(completed_counts.keys(), key=lambda x: (proj_order.get(x[0], 99), x[1])):
        row = write_data_row(row, [k[0], k[1], completed_counts[k]])

    row = write_data_row(row, ["Total", "", new_data.get("completed_total", 0)], total=True)

    # Pointer to the dedicated detail sheet
    completed_list = new_data.get("completed_detail_rows", [])
    row += 1
    if completed_list:
        c = dash_ws.cell(row=row, column=1,
                         value=f"Full list: see the 'Completed Stories' sheet "
                               f"({len(completed_list)} user stories delivered to the PO since {CUTOFF_DATE})")
    else:
        c = dash_ws.cell(row=row, column=1, value=f"No stories completed after {CUTOFF_DATE}")
    c.font = Font(name="Segoe UI", size=9, italic=True, color="808080" if completed_list else "999999")
    dash_ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)

    # --- Section 5: Ticket System Analysis & Progress ---
    row += 2
    row = write_section(row, "Ticket System Analysis & Progress", 6)

    if ticket_data:
        row = write_table_header(row, ["Severity", "Tickets", "Done", "In Progress", "Not Started", "Not Covered"])

        sevs = sorted(ticket_data["per_severity"].keys(), key=severity_sort_key)
        for sev in sevs:
            s = ticket_data["per_severity"][sev]
            row = write_data_row(row, [sev, s["total"], s["Done"], s["In Progress"],
                                       s["Not Started"], s["Not Covered"]])

        cc = ticket_data["cat_counts"]
        row = write_data_row(
            row,
            ["Total", ticket_data["total"], cc["Done"], cc["In Progress"],
             cc["Not Started"], cc["Not Covered"]],
            total=True,
        )

        # Tickets per module (under the severity table)
        row += 1
        c = dash_ws.cell(row=row, column=1, value="Tickets per Module")
        c.font = Font(name="Segoe UI", bold=True, size=10, color="1F3864")
        row += 1
        row = write_table_header(row, ["Module", "Tickets", "Covered by Roadmap", "Not Covered"])

        per_module = ticket_data.get("per_module") or {}
        modules = sorted(per_module.keys(), key=lambda m: (-per_module[m]["total"], m.lower()))
        for mod in modules:
            m = per_module[mod]
            row = write_data_row(row, [mod, m["total"], m["covered"], m["total"] - m["covered"]])

        row = write_data_row(
            row,
            ["Total", ticket_data["total"], ticket_data["covered_total"],
             ticket_data["not_covered_total"]],
            total=True,
        )

        # Progress explanation
        row += 1
        c = dash_ws.cell(
            row=row, column=1,
            value=("Progress = share of linked roadmap items (PBIs) that are Done. 'Done' = all linked PBIs "
                   "completed or the ticket is resolved/closed in the ticket system. Yellow rows on the Tickets "
                   "sheet are 100% complete in ADO but still open there. Resolved tickets without any roadmap "
                   "link are excluded."),
        )
        c.font = Font(name="Segoe UI", size=9, italic=True, color="808080")
        dash_ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)

        # Ticket numbers referenced by the roadmap but missing from the export
        missing = ticket_data.get("missing_from_export") or []
        if missing:
            row += 1
            shown = ", ".join(missing[:TICKET_MISSING_LIST_MAX])
            if len(missing) > TICKET_MISSING_LIST_MAX:
                shown += ", ..."
            c = dash_ws.cell(
                row=row, column=1,
                value=(f"Ticket numbers linked to roadmap items but not found in the ticket "
                       f"export (likely resolved or inactive tickets): {shown}"),
            )
            c.font = Font(name="Segoe UI", size=9, italic=True, color="808080")
            dash_ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)

        # Pointer to the detail sheet
        row += 1
        c = dash_ws.cell(
            row=row, column=1,
            value=(f"Full ticket list: see the 'Tickets' sheet "
                   f"(source: {os.path.basename(ticket_data['export_path'])})"),
        )
        c.font = Font(name="Segoe UI", size=9, italic=True, color="808080")
        dash_ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
    else:
        row += 1
        c = dash_ws.cell(
            row=row, column=1,
            value=(f"Ticket analysis skipped — no export found (pattern '{TICKET_EXPORT_PATTERN}' in "
                   f"'{TICKET_EXPORT_FOLDER}') or no tracked tickets "
                   f"(Request Type '{'/'.join(TICKET_REQUEST_TYPES)}', excluding Rejected / On hold)."),
        )
        c.font = Font(name="Segoe UI", size=10, italic=True, color="999999")
        dash_ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)

    dash_ws.freeze_panes = "A4"


def build_new_stories_sheet(wb, detail_rows):
    """Dedicated sheet holding the new stories/features as a real Excel
    Table (filter dropdowns + banded rows), ordered:
    Owner > Project > Module > Feature > Story Title > Story number."""
    ws = wb.create_sheet("New Stories", 2)
    ws.sheet_properties.tabColor = "70AD47"

    headers = ["Owner", "Project", "Module", "Feature", "Story Title", "Story ID",
               "Ticket Number", "Ticket Priority", "Status", "Added on"]
    widths = [30, 20, 20, 42, 55, 12, 14, 14, 13, 13]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    thin_b = Border(
        left=Side(style="thin", color="D6D6D6"),
        right=Side(style="thin", color="D6D6D6"),
        top=Side(style="thin", color="D6D6D6"),
        bottom=Side(style="thin", color="D6D6D6"),
    )
    title_font = Font(name="Segoe UI", bold=True, size=14, color="1F3864")
    hdr_font = Font(name="Segoe UI", bold=True, size=10, color="FFFFFF")
    hdr_fill = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
    hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    c = ws.cell(row=1, column=1, value=f"New Stories & Features — Created After {CUTOFF_DATE}")
    c.font = title_font
    ws.merge_cells("A1:J1")

    header_row = 3

    if detail_rows:
        for col_idx, h in enumerate(headers, 1):
            c = ws.cell(row=header_row, column=col_idx, value=h)
            c.font = hdr_font
            c.fill = hdr_fill
            c.alignment = hdr_align
            c.border = thin_b

        data_font = Font(name="Segoe UI", size=10)
        link_font = Font(name="Segoe UI", size=10, color="0563C1", underline="single")
        data_align = Alignment(vertical="top", wrap_text=True)

        r = header_row + 1
        for d in detail_rows:
            vals = [d["Owner"], d["Project"], d["Module"], d["Feature"],
                    d["Story Title"], d["Story ID"], d["Ticket Number"],
                    d["Ticket Priority"], d["Status"], d["Added on"]]
            for col_idx, val in enumerate(vals, 1):
                c = ws.cell(row=r, column=col_idx, value=val)
                c.font = data_font
                c.alignment = data_align
                c.border = thin_b
                if col_idx == 7 and d.get("Ticket URL"):
                    c.hyperlink = d["Ticket URL"]
                    c.font = link_font
            r += 1
        last_row = r - 1

        # Real Excel Table -> filter dropdowns + banded styling
        table_ref = f"A{header_row}:J{last_row}"
        tab = Table(displayName="NewStoriesTable", ref=table_ref,
                    autoFilter=AutoFilter(ref=table_ref))
        tab.tableColumns = [TableColumn(id=i + 1, name=h) for i, h in enumerate(headers)]
        tab.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showRowStripes=True,
            showColumnStripes=False,
            showFirstColumn=False,
            showLastColumn=False,
        )
        ws.add_table(tab)
        ws.freeze_panes = f"A{header_row + 1}"
    else:
        c = ws.cell(row=header_row, column=1,
                    value=f"No new stories or features created after {CUTOFF_DATE}")
        c.font = Font(name="Segoe UI", size=10, italic=True, color="999999")


def build_completed_stories_sheet(wb, detail_rows):
    """Dedicated 'Completed Stories' sheet: user stories delivered to the
    PO since the cutoff, as a real Excel Table (filter dropdowns + banded
    rows), ordered: Owner > Project > Module > Feature > Story Title >
    Story number."""
    ws = wb.create_sheet("Completed Stories", 3)
    ws.sheet_properties.tabColor = "548235"

    headers = ["Owner", "Project", "Module", "Feature", "Story Title", "Story ID",
               "Standard Weight", "Ticket Number", "Ticket Priority", "Done Date"]
    widths = [30, 20, 20, 42, 55, 12, 14, 14, 14, 13]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    thin_b = Border(
        left=Side(style="thin", color="D6D6D6"),
        right=Side(style="thin", color="D6D6D6"),
        top=Side(style="thin", color="D6D6D6"),
        bottom=Side(style="thin", color="D6D6D6"),
    )
    title_font = Font(name="Segoe UI", bold=True, size=14, color="1F3864")
    meta_font = Font(name="Segoe UI", size=9, color="808080")
    hdr_font = Font(name="Segoe UI", bold=True, size=10, color="FFFFFF")
    hdr_fill = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
    hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    c = ws.cell(row=1, column=1,
                value="List of user stories delivered to the PO since the cut off")
    c.font = title_font
    ws.merge_cells("A1:J1")

    c = ws.cell(row=2, column=1,
                value=f"Stories set as Done after {CUTOFF_DATE}  |  {len(detail_rows)} user stories")
    c.font = meta_font
    ws.merge_cells("A2:J2")

    header_row = 4

    if detail_rows:
        for col_idx, h in enumerate(headers, 1):
            c = ws.cell(row=header_row, column=col_idx, value=h)
            c.font = hdr_font
            c.fill = hdr_fill
            c.alignment = hdr_align
            c.border = thin_b

        data_font = Font(name="Segoe UI", size=10)
        link_font = Font(name="Segoe UI", size=10, color="0563C1", underline="single")
        weight_align = Alignment(vertical="top", horizontal="left")
        data_align = Alignment(vertical="top", wrap_text=True)

        r = header_row + 1
        for d in detail_rows:
            vals = [d["Owner"], d["Project"], d["Module"], d["Feature"],
                    d["Story Title"], d["Story ID"], d["Standard Weight"],
                    d["Ticket Number"], d["Ticket Priority"], d["Done Date"]]
            for col_idx, val in enumerate(vals, 1):
                c = ws.cell(row=r, column=col_idx, value=val)
                c.font = data_font
                c.alignment = weight_align if col_idx == 7 else data_align
                c.border = thin_b
                if col_idx == 8 and d.get("Ticket URL"):
                    c.hyperlink = d["Ticket URL"]
                    c.font = link_font
            r += 1
        last_row = r - 1

        # Real Excel Table -> filter dropdowns + banded styling
        table_ref = f"A{header_row}:J{last_row}"
        tab = Table(displayName="CompletedStoriesTable", ref=table_ref,
                    autoFilter=AutoFilter(ref=table_ref))
        tab.tableColumns = [TableColumn(id=i + 1, name=h) for i, h in enumerate(headers)]
        tab.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showRowStripes=True,
            showColumnStripes=False,
            showFirstColumn=False,
            showLastColumn=False,
        )
        ws.add_table(tab)
        ws.freeze_panes = f"A{header_row + 1}"
    else:
        c = ws.cell(row=header_row, column=1,
                    value=f"No user stories delivered to the PO since {CUTOFF_DATE}")
        c.font = Font(name="Segoe UI", size=10, italic=True, color="999999")


def build_tickets_sheet(wb, ticket_data):
    """Dedicated 'Tickets' sheet: every tracked ticket from the ticket
    system export with name, module, severity, number, Progress, status
    and a 'Covered by Roadmap' flag — as a real Excel Table (filter
    dropdowns + banded rows). Progress = share of linked roadmap stories
    that are Done ('Done' at 100%, empty at 0% / not covered); rows that
    are 100% complete in ADO but still open in the ticket system are
    highlighted faint yellow. Sorted by severity, then not-covered
    first, then ticket number. Skipped gracefully when the export is
    unavailable."""
    ws = wb.create_sheet("Tickets", 4)
    ws.sheet_properties.tabColor = "C55A11"

    headers = ["Ticket Name", "Module", "Severity", "Ticket Number", "Progress", "Status", "Covered by Roadmap"]
    widths = [60, 22, 16, 16, 12, 16, 20]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    thin_b = Border(
        left=Side(style="thin", color="D6D6D6"),
        right=Side(style="thin", color="D6D6D6"),
        top=Side(style="thin", color="D6D6D6"),
        bottom=Side(style="thin", color="D6D6D6"),
    )
    title_font = Font(name="Segoe UI", bold=True, size=14, color="1F3864")
    note_font = Font(name="Segoe UI", size=9, italic=True, color="C00000")
    meta_font = Font(name="Segoe UI", size=9, color="808080")
    hdr_font = Font(name="Segoe UI", bold=True, size=10, color="FFFFFF")
    hdr_fill = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
    hdr_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    covered_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    covered_font = Font(name="Segoe UI", size=10, color="006100")
    not_covered_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    not_covered_font = Font(name="Segoe UI", size=10, color="9C0006")
    done_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    done_font = Font(name="Segoe UI", size=10, bold=True, color="006100")
    highlight_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    # Title + note
    if ticket_data:
        c = ws.cell(row=1, column=1,
                    value=f"Ticket System Analysis — {os.path.basename(ticket_data['export_path'])}")
        c.font = title_font
        ws.merge_cells("A1:G1")

        c = ws.cell(row=2, column=1,
                    value=("Progress = share of linked roadmap items (PBIs) that are Done. 'Done' = all linked "
                           "PBIs completed or the ticket is resolved/closed. Yellow rows are 100% complete in "
                           "ADO but still open in the ticket system."))
        c.font = note_font
        ws.merge_cells("A2:G2")

        excluded = ticket_data.get("excluded_total", 0)
        filter_note = (f"Tracking Request Type '{'/'.join(TICKET_REQUEST_TYPES)}' — excluding Rejected / On hold "
                       f"and resolved tickets without roadmap references")
        if excluded:
            filter_note += f" — {excluded} tickets excluded"
        c = ws.cell(row=3, column=1, value=filter_note)
        c.font = meta_font
        ws.merge_cells("A3:G3")

        header_row = 4
    else:
        c = ws.cell(row=1, column=1, value="Ticket System Analysis")
        c.font = title_font
        ws.merge_cells("A1:G1")

        c = ws.cell(row=2, column=1,
                    value=(f"No ticket export found (pattern '{TICKET_EXPORT_PATTERN}' in "
                           f"'{TICKET_EXPORT_FOLDER}') — ticket analysis skipped."))
        c.font = meta_font
        ws.merge_cells("A2:G2")

        header_row = 4

    if ticket_data and ticket_data["tickets"]:
        for col_idx, h in enumerate(headers, 1):
            c = ws.cell(row=header_row, column=col_idx, value=h)
            c.font = hdr_font
            c.fill = hdr_fill
            c.alignment = hdr_align
            c.border = thin_b

        data_font = Font(name="Segoe UI", size=10)
        link_font = Font(name="Segoe UI", size=10, color="0563C1", underline="single")
        data_align = Alignment(vertical="top", wrap_text=True)

        # Sort: severity order, not-covered first, then ticket number
        tickets = sorted(
            ticket_data["tickets"],
            key=lambda t: (
                severity_sort_key(t.get("severity")),
                0 if not t.get("covered") else 1,
                ticket_number_sort_key(t.get("number")),
            ),
        )

        def progress_text(progress):
            if progress is None or progress <= 0:
                return ""  # not covered or 0% -> empty
            if progress >= 1:
                return "Done"
            return f"{round(progress * 100)}%"

        r = header_row + 1
        for t in tickets:
            covered = t.get("covered")
            progress = t.get("progress")
            row_highlight = (progress is not None and progress >= 1 and not t.get("resolved"))
            vals = [t.get("name", ""), t.get("module", ""), t.get("severity", ""),
                    t.get("number", ""), progress_text(progress), t.get("status", ""),
                    "Yes" if covered else "No"]
            for col_idx, val in enumerate(vals, 1):
                c = ws.cell(row=r, column=col_idx, value=val)
                c.border = thin_b
                c.font = data_font
                c.alignment = data_align
                if row_highlight:
                    c.fill = highlight_fill
                if col_idx == 5 and progress is not None and progress >= 1:
                    c.fill = done_fill
                    c.font = done_font
                    c.alignment = Alignment(horizontal="center", vertical="center")
                elif col_idx == 7:
                    c.font = covered_font if covered else not_covered_font
                    c.fill = covered_fill if covered else not_covered_fill
                    c.alignment = Alignment(horizontal="center", vertical="center")
                if col_idx == 4 and t.get("guid"):
                    c.hyperlink = TICKET_URL_TEMPLATE.format(guid=t["guid"])
                    c.font = link_font
            r += 1
        last_row = r - 1

        # Real Excel Table -> filter dropdowns + banded styling
        table_ref = f"A{header_row}:G{last_row}"
        tab = Table(displayName="TicketsTable", ref=table_ref,
                    autoFilter=AutoFilter(ref=table_ref))
        tab.tableColumns = [TableColumn(id=i + 1, name=h) for i, h in enumerate(headers)]
        tab.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showRowStripes=True,
            showColumnStripes=False,
            showFirstColumn=False,
            showLastColumn=False,
        )
        ws.add_table(tab)
        ws.freeze_panes = f"A{header_row + 1}"
    elif ticket_data:
        c = ws.cell(row=header_row, column=1,
                    value=f"No tracked tickets in the export "
                          f"(Request Type '{'/'.join(TICKET_REQUEST_TYPES)}', excluding Rejected / On hold)")
        c.font = Font(name="Segoe UI", size=10, italic=True, color="999999")


# ============================================================
# DATE WIZARD
# ============================================================

def prompt_cutoff_dates():
    """Ask the user for the two cutoff dates before running.
    ENTER keeps the configured default. Input must be YYYY-MM-DD."""
    global CUTOFF_DATE, ROADMAP_CUTOFF_DATE

    def ask(question, current):
        print(f"\n  {question}")
        while True:
            try:
                raw = input(f"  Date YYYY-MM-DD (ENTER = keep {current}): ").strip()
            except EOFError:
                return current
            if not raw:
                return current
            try:
                return datetime.strptime(raw, "%Y-%m-%d").strftime("%Y-%m-%d")
            except ValueError:
                print(f"      Invalid date '{raw}' — expected YYYY-MM-DD, try again.")

    print("\n" + "-" * 60)
    print("  Roadmap dates — press ENTER to keep the current value")
    print("-" * 60)
    ROADMAP_CUTOFF_DATE = ask(
        "1) When should the roadmap start? (stories completed before this date are excluded)",
        ROADMAP_CUTOFF_DATE,
    )
    CUTOFF_DATE = ask(
        "2) When to count new items added to the roadmap? (dashboard 'new items' threshold)",
        CUTOFF_DATE,
    )
    print(f"\n  Roadmap start date:     {ROADMAP_CUTOFF_DATE}")
    print(f"  New items counted from: {CUTOFF_DATE}")
    print("-" * 60)


# ============================================================
# MAIN
# ============================================================

def main():
    # UTF-8 console (ticket severities come with emoji prefixes from the export)
    for stream in (sys.stdout, sys.stderr):
        try:
            if stream and hasattr(stream, "reconfigure"):
                stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    print("=" * 60)
    print("  ADO Roadmap Sync v7.1")
    print("  TFS: " + TFS_URL)
    print(f"  Projects: {', '.join(PROJECTS)}")
    print("=" * 60)

    # Date wizard (skip with --defaults)
    if "--defaults" in sys.argv:
        print(f"\n  --defaults: skipping the date wizard.")
        print(f"  Roadmap start date:     {ROADMAP_CUTOFF_DATE}")
        print(f"  New items counted from: {CUTOFF_DATE}")
    else:
        prompt_cutoff_dates()

    if PAT == "YOUR_PAT_HERE" or not PAT:
        print("\nERROR: No PAT configured!")
        sys.exit(1)

    try:
        t_start = time.time()

        # 1. Fetch PBIs
        t0 = time.time()
        pbi_items = fetch_all_pbis()
        print(f"      [1/5] took {time.time() - t0:.1f}s")

        # 2. Fetch orphan Features/Epics
        t0 = time.time()
        orphan_items = fetch_all_features_and_epics()
        print(f"      [2/5] took {time.time() - t0:.1f}s")

        # 3. Build parent + grandparent lookup
        print(f"\n[3/5] Building parent and grandparent (Epic) lookup...")
        t0 = time.time()
        parent_lookup, grandparent_lookup, gp_assignee = build_parent_and_grandparent_lookup(
            pbi_items, orphan_items
        )
        global grandparent_assignee_lookup
        grandparent_assignee_lookup = gp_assignee
        print(f"      Parent lookup: {len(parent_lookup)} items")
        print(f"      [3/5] took {time.time() - t0:.1f}s")

        # 4. Map + filter
        print(f"\n[4/5] Mapping work items and filtering...")
        t0 = time.time()
        print(f"      Roadmap cutoff: {ROADMAP_CUTOFF_DATE} (stories done before this are excluded)")
        all_rows = []
        skipped_no_parent = 0
        skipped_no_owner = 0
        skipped_done_before_cutoff = 0

        roadmap_cutoff_dt = parse_date_for_cutoff(ROADMAP_CUTOFF_DATE + "T00:00:00+00:00")

        for wi in pbi_items:
            # Exclude unparented PBIs — but only in configured projects
            # (HMIS stories live under Features; HR/Mobile/Websites PBIs are
            #  mostly top-level, so they are kept and become their own group)
            relations = wi.get("relations", [])
            parent_id = extract_parent_id(relations)
            wi_project = wi.get("fields", {}).get("System.TeamProject", "")
            if not parent_id and wi_project in UNPARENTED_EXCLUDE_PROJECTS:
                skipped_no_parent += 1
                continue

            row = map_pbi_to_row(wi, parent_lookup, grandparent_lookup)
            if not row["Owner"]:
                skipped_no_owner += 1
                continue

            # Exclude stories done before roadmap cutoff
            if row["Status"] == "Done" and row["Done Date"]:
                delivery_dt = parse_date_for_cutoff(row["Done Date"] + "T00:00:00+00:00")
                if delivery_dt and roadmap_cutoff_dt and delivery_dt < roadmap_cutoff_dt:
                    skipped_done_before_cutoff += 1
                    continue

            all_rows.append(row)

        for wi in orphan_items:
            row = map_orphan_to_row(wi, parent_lookup, grandparent_lookup)
            if not (row["Owner"] and is_allowed_owner(row["Owner"])):
                skipped_no_owner += 1
                continue

            # Orphan features are Backlog status — never excluded by cutoff
            all_rows.append(row)

        print(f"      {len(all_rows)} items included")
        print(f"      {skipped_no_parent} PBIs skipped (unparented)")
        print(f"      {skipped_no_owner} items skipped (no allowed owner in hierarchy)")
        print(f"      {skipped_done_before_cutoff} stories skipped (done before {ROADMAP_CUTOFF_DATE})")

        # Diagnostics: how many rows have the new TFS fields populated
        for col in NEW_FIELD_REFS:
            n = sum(1 for r in all_rows if str(r.get(col, "") or "").strip())
            print(f"      '{col}' populated on {n} rows")
        n_sw = sum(1 for r in all_rows if str(r.get("Standard Weight", "") or "").strip())
        print(f"      'Standard Weight' populated on {n_sw} rows")
        print(f"      [4/5] took {time.time() - t0:.1f}s")

        if not all_rows:
            print("\nNo work items match criteria. Exiting.")
            sys.exit(0)

        # 5. Generate Excel — fresh build from TFS every run, no carry-over
        print(f"\n[5/5] Generating Excel...")
        t0 = time.time()
        output_file = generate_excel(all_rows)
        print(f"      [5/5] took {time.time() - t0:.1f}s")
        print(f"      Total run: {time.time() - t_start:.1f}s")

        print(f"\n{'=' * 60}")
        print(f"  SYNC COMPLETE")
        print(f"  Total items in roadmap: {len(all_rows)}")
        print(f"  Output: {output_file}")
        print(f"{'=' * 60}")

    except requests.exceptions.ConnectionError:
        print(f"\nERROR: Cannot connect to TFS at {TFS_URL}")
        sys.exit(1)
    except requests.exceptions.HTTPError as e:
        print(f"\nERROR: HTTP {e.response.status_code}: {e}")
        if e.response.status_code == 401:
            print("   Authentication failed. Check your PAT.")
        sys.exit(1)
    except Exception as e:
        print(f"\nERROR: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
