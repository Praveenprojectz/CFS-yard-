# CFS Yard Utilisation Dashboard

An interactive dashboard for daily CFS yard utilisation and availability. It reads
`CFS_Yard_Data.xlsx` from this repository, so updating that one Excel file updates
the dashboard for everyone who has the link.

## Files

| File | Purpose |
|---|---|
| `index.html` | The complete dashboard — page, styling, code and built-in fallback data in one file |
| `style.css` | Source copy of the styling (already included inside index.html — optional to upload) |
| `script.js` | Source copy of the logic (already included inside index.html — optional to upload) |
| `CFS_Yard_Data.xlsx` | The data — one row per day, entered by the operations team |
| `README.md` | This guide |

`index.html` also opens directly on any phone or computer (open it with Chrome or
Safari). On its own it shows the built-in data up to 28 Jul 2026; hosted on GitHub
Pages next to `CFS_Yard_Data.xlsx`, it always loads the latest data automatically.

The dashboard has a **Stacking band** selector: Total yard (670 slots), Warehouse
container stacking (221 slots, export open yard cargo) or Yard container stacking
(449 slots: export load, empty, import, special). Band views apply from
06 Jun 2026, when the 670-slot layout took effect.

## Publish the dashboard (one-time, about 5 minutes)

1. Sign in at github.com and click **New repository**. Name it `cfs-yard-dashboard`, set it to **Public**, and click **Create repository**.
2. Click the **uploading an existing file** link (or **Add file › Upload files**), drag the files above in (index.html and CFS_Yard_Data.xlsx are the two that matter), and click **Commit changes**.
3. Open **Settings › Pages**. Under *Build and deployment*, set **Source** to *Deploy from a branch*, choose branch **main** and folder **/ (root)**, and click **Save**.
4. Wait one to two minutes, then refresh the Pages screen. Your link appears at the top: `https://<your-username>.github.io/cfs-yard-dashboard/`
5. Share that link. Anyone can open it — no sign-in needed.

> A public repository means the data file is publicly downloadable. Free GitHub
> Pages does not work from private repositories, so keep that in mind for
> commercially sensitive figures.

## Update the data (daily, about 1 minute)

1. Add the day's row in `CFS_Yard_Data.xlsx` on your computer (the *Instructions* sheet inside the file explains each column).
2. In the repository, click **Add file › Upload files**, drag the saved file in, and click **Commit changes**. Because the name matches, it replaces the old file.
3. The dashboard shows the new data within a minute or two. Refresh the page (Ctrl+F5 if it looks unchanged).

The **Upload Excel** button on the page itself is for previewing a file on your own
device before publishing — it does not change what other people see.

## Data format rules

- Keep the file name `CFS_Yard_Data.xlsx`, the sheet name `Daily Data`, and the row-1 headings exactly as they are.
- Enter box **counts** (not TEU) for the 20' and 40' columns. The dashboard applies 1 TEU per 20' and 2 TEU per 40'.
- Special containers are entered directly in TEU. Warehouse figures are in sq ft.
- One row per date; if a date appears twice, the lower (later) row wins.
- If the yard is re-based, enter the new slot count from that day onward — the trend chart marks the change automatically.

## Calculations

Occupancy = export load + empty + open yard + import + special (TEU) ·
Utilisation = occupancy ÷ (ground slots × tiers) ·
Availability = capacity − occupancy ·
Warehouse % = area used ÷ total area.

## Troubleshooting

- **Opened locally, the page shows "built-in data"** — expected; the built-in data ends 28 Jul 2026. Use *Upload Excel* to view a newer file, or open the live GitHub Pages link, which always loads the latest.
- **The page opens as plain text or code on a phone** — use "Open with" and choose Chrome (Android) or Safari (iPhone).
- **New data not showing** — hard-refresh (Ctrl+F5); GitHub Pages can take a minute or two after a commit.
- **"No daily rows were found"** — the sheet name or row-1 headings were changed; restore them from the template.
