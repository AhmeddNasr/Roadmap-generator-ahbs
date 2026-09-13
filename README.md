# ADO Roadmap Sync v6

Builds `roadmap_<date>_<time>.xlsx` from scratch on every run: pulls **Product Backlog Items + Product Non Backlog Items + orphan Features/Epics** from 4 TFS projects and generates a local Excel roadmap with a **Dashboard**. Nothing is carried over between runs — TFS is the single source of truth.

## Quick Start

### 1. Install packages
```
pip install -r requirements.txt
```

### 2. Get a TFS PAT
1. Open http://ahq-tfs-azure/DefaultCollection
2. Profile (top right) -> Security -> Personal access tokens -> New Token
3. Scopes: **Work Items (Read)**
4. Copy the token

### 3. Configure
Open `ado_roadmap_sync.py`, edit:
```python
PAT = "YOUR_PAT_HERE"           # Your TFS token
```

### 4. Run
```
python ado_roadmap_sync.py
```
Or double-click `run_sync.bat`.

A wizard asks for two dates before the sync starts (press **ENTER** to keep each default, or type a new date as `YYYY-MM-DD`):
1. **When should the roadmap start?** — stories completed before this date are excluded (`ROADMAP_CUTOFF_DATE`, default 2026-08-01)
2. **When to count new items added to the roadmap?** — dashboard "new items" threshold (`CUTOFF_DATE`, default 2026-09-01)

To run without the wizard (e.g. scheduled), pass `--defaults`:
```
python ado_roadmap_sync.py --defaults
```

### 5. Upload
Upload the generated `roadmap_<date>_<time>.xlsx` to SharePoint manually.

## What it does

### Projects synced

HMIS, HR System, Mobile Applications, Websites

### Data sources (2 kinds of rows)

| Source | Work Item Type | Condition |
|--------|---------------|-----------|
| User Stories | Product Backlog Item, Product Non Backlog Item | Not Removed, allowed owner (own or via parent Feature / Epic) |
| Orphan Features/Epics | Feature, Epic | No children, not Removed, directly assigned to an allowed owner |

### Owner filter (8 people)
- Ahmed Nasr Younis AbdElWahed
- Mohamed Sharshira
- Mohamed Ahmed Mohamed Aly
- Ibrahim AbdElFattah Mohamed Ghanem
- Nada Adel Khamis
- Mohamed Adel Khalifa
- Mohamed Moataz
- Elzohery

Owner resolution for stories: own assignee → parent Feature assignee → grandparent Epic assignee (first allowed match wins). Unassigned stories under an owned feature inherit the feature's owner.

### Business Area / Feature column

| Item Type | Parent | Grandparent | Business Area |
|-----------|--------|-------------|---------------|
| PBI | Feature | Epic | `{Epic} - {Feature}` |
| PBI | Feature | (none) | `{Feature}` |
| PBI | Epic | — | `{Epic}` |
| Unparented PBI | — | — | Own title (own feature group) |
| Orphan Feature | Epic | — | `{Epic} - {Feature}` |
| Orphan Feature | (none) | — | `{Feature}` |

### Status mapping

| TFS State | Working Status (Custom.WorkingStatus) | Roadmap Status |
|-----------|----------------------------------------|----------------|
| New / Approved | — | Backlog |
| Committed | contains "test" | Testing |
| Committed | (other / empty) | Development |
| Done | — | Done |

### New TFS fields (v5+)

| TFS Field | Roadmap Column | Notes |
|-----------|----------------|-------|
| Custom.Ticketnumber | Ticket Number | HTML-cleaned |
| Custom.Impactlevel | Impact | |
| Custom.BusinessValueCategory | Category | |
| Custom.BusinessImpactValue | Business Value | HTML-cleaned |

**Inheritance:** a Feature's values apply to every child story. Story-level values override the Feature's values **per field** (a story can set its own Category and inherit the rest).

### Dates

- **Added on** = CreatedDate
- **Start Date** = ActivatedDate
- **Delivery Date** = ClosedDate (or StateChangeDate if Done)
- Stories Done **before `ROADMAP_CUTOFF_DATE` (2026-08-01)** are excluded from the roadmap
- Dashboard "new items" use `CUTOFF_DATE` (2026-09-01)

### Dashboard sheet

1. **Overall Status Breakdown** — stories + feature groups per status
2. **Status Breakdown by Project** — reconciles with table 1
3. **New Items After CUTOFF — by Project & Owner** (+ filterable "New Stories" sheet)
4. **Completed After CUTOFF — by Project & Owner** (+ full list below)

## Output columns (16)

| # | Column | Source |
|---|--------|--------|
| 1 | Owner | ADO: AssignedTo (story, or inherited from Feature/Epic) |
| 2 | Module | ADO: Custom.Module |
| 3 | Business Area / Feature | ADO: `{Epic} - {Feature}` or `{Feature}` (merged cells) |
| 4 | Requirement | ADO: Work item title |
| 5 | Reference ID | ADO: ID (blank for orphan features) |
| 6 | Ticket Number | ADO: Custom.Ticketnumber (story or inherited from Feature) |
| 7 | Priority | ADO: Microsoft.VSTS.Common.Priority |
| 8 | Impact | ADO: Custom.Impactlevel (story or inherited from Feature) |
| 9 | Status | Mapped from State + Custom.WorkingStatus |
| 10 | Added on | ADO: CreatedDate |
| 11 | Start Date | ADO: ActivatedDate |
| 12 | Delivery Date | ADO: ClosedDate |
| 13 | Stakeholder | *(reserved — currently empty)* |
| 14 | Category | ADO: Custom.BusinessValueCategory (story or inherited from Feature) |
| 15 | Business Value | ADO: Custom.BusinessImpactValue (story or inherited from Feature) |
| 16 | Reviewed | *(reserved — currently empty)* |

## Sheets

1. **Dashboard** — status breakdown, new items, completed items
2. **Roadmap** — main data with merged feature cells, color-coded status
3. **New Stories** — filterable Excel table of items created after CUTOFF_DATE
4. **Summary** — counts by status, owner, module

## Configuration

Edit these at the top of `ado_roadmap_sync.py`:

```python
TFS_URL = "http://ahq-tfs-azure/DefaultCollection"
PAT = "YOUR_PAT_HERE"
PROJECTS = [...]                   # 4 projects to sync
ALLOWED_OWNERS = [...]              # 8 owner name patterns
CUTOFF_DATE = "2026-09-01"          # Default: new-items threshold (wizard question 2)
ROADMAP_CUTOFF_DATE = "2026-08-01"  # Default: roadmap start date (wizard question 1)
```

## Troubleshooting

| Error | Fix |
|-------|-----|
| Cannot connect to TFS | Check VPN/network |
| 401 Authentication failed | PAT expired or wrong scope |
| No work items found | Check work item type / owner filters |
| Ticket/Category columns empty | Populate Custom.Ticketnumber etc. on the Features/stories in TFS — the team fills these |
| Folder filling with old files | Every run creates a new `roadmap_<date>_<time>.xlsx` — delete old ones you no longer need |
