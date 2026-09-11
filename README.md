# product_performance

Merges GA4 item-level e-commerce data with Google Ads and Meta Ads spend
(joined by product/item ID) into a Brand/Category performance xlsx report.
Built and verified as a Claude Code skill — see `SKILL.md` for the full
workflow and `references/mcp_pull_reference.md` for exact data-pull specs.

## Getting this from a git clone onto a new machine

Cloning the repo gets you the three files (`SKILL.md`, `scripts/build_report.py`,
`references/mcp_pull_reference.md`) — that part is just plain text, no secrets,
no machine-specific paths, verified portable in an isolated test. Three things
are **not** included in the clone and need to be done separately on the new
machine:

1. **Put it where Claude Code looks for skills.** A clone in some random
   folder is not auto-discovered. Copy (or symlink) this folder to either:
   - `~/.claude/skills/product_performance/` (available in every project for
     that user), or
   - `<project>/.claude/skills/product_performance/` (scoped to one project)

   Skills are per-user/per-project directories, not something git or GitHub
   syncs on their own — every machine that should have this skill needs its
   own copy in one of those two locations.

2. **Install the Python dependencies** the script needs:
   ```bash
   pip install -r requirements.txt
   ```
   (just pandas + openpyxl — nothing else, verified by grepping the script's
   imports). `scripts/build_report.py` itself is a fully standalone script at
   that point — it can be run directly with `python3`, from any directory,
   with no dependency on Claude Code, this repo's other files, or any prior
   session context. Only the *data-pulling* half of the workflow (Step 5 in
   SKILL.md) needs an agent to drive it.

3. **Have your own access to GA4 / Google Ads / Meta Ads.** The skill doesn't
   bundle any credentials (there aren't any to bundle) — whoever runs it
   needs their own working connection to pull the three data sources
   (MCP tools, or any other means of producing the 3-4 CSVs the script
   expects — see the "Input CSV contracts" docstring at the top of
   `build_report.py`).

## Running it from Codex

The intended Codex users have the same `ga4-mcp` / `gads-mcp` / `meta-mcp`
connections as the Claude Code side (confirmed, not just same-named
equivalents) — so `references/mcp_pull_reference.md`'s tool names and call
shapes apply as-is, no adaptation needed. That part isn't the uncertain
piece.

**What's genuinely untested**: whether Codex auto-discovers a skill folder
the way Claude Code does, or needs to be told explicitly. Two ways to
handle either case:

- **If Codex has no auto-discovery**: tell it directly, e.g. *"Follow the
  steps in SKILL.md and references/mcp_pull_reference.md in this folder to
  build a product performance report"* — the files are plain instructions,
  nothing Claude-Code-specific in their content, so this should work as a
  manual playbook regardless of how Codex is wired.
- **Skip the agent step entirely once you have the CSVs**: the script itself
  needs no agent, Codex or otherwise -
  ```bash
  python3 scripts/build_report.py \
    --ga4-csv ga4_items.csv \
    --google-ads-csv google_ads_spend.csv \
    --meta-csv meta_spend.csv \
    --category-levels 2 \
    --include-brand-analysis \
    --output report.xlsx
  ```
