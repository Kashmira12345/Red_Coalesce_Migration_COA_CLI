"""Generate the United Rentals WhereScape RED -> Coalesce migration summary and runbook (.docx)."""

import sys
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Inches

NAVY = RGBColor(0x00, 0x20, 0x60)
SLATE = RGBColor(0x59, 0x59, 0x59)
CHARCOAL = RGBColor(0x26, 0x26, 0x26)
FONT = "Calibri"
MONO = "Consolas"

REPO_URL = "https://github.com/Kashmira12345/Red_Coalesce_Migration_COA_CLI.git"
WORKSPACE = r"I:\DataEntrega\Red_Coalesce_Migration\Red_Coalesce_Migration_COA_CLI"
DISCLAIMER = ("This is for informational purposes only. For medical advice or diagnosis, "
              "consult a professional. AI responses may include mistakes.")


# ---------------------------------------------------------------- styling helpers

def set_font(run_or_style, name=FONT, size=None, bold=None, color=None, italic=None):
    f = run_or_style.font
    f.name = name
    rpr = run_or_style.element.get_or_add_rPr() if hasattr(run_or_style, "element") else None
    if rpr is not None:
        fonts = rpr.find(qn("w:rFonts"))
        if fonts is None:
            fonts = OxmlElement("w:rFonts")
            rpr.append(fonts)
        for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
            fonts.set(qn(attr), name)
    if size is not None:
        f.size = Pt(size)
    if bold is not None:
        f.bold = bold
    if italic is not None:
        f.italic = italic
    if color is not None:
        f.color.rgb = color


def configure_styles(doc):
    styles = doc.styles
    normal = styles["Normal"]
    set_font(normal, size=11, color=CHARCOAL)
    pf = normal.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = 1.15
    pf.space_after = Pt(6)
    pf.space_before = Pt(0)

    for name, size, color, before, after in (
        ("Title", 24, NAVY, 0, 6),
        ("Subtitle", 13, SLATE, 0, 12),
        ("Heading 1", 16, NAVY, 18, 6),
        ("Heading 2", 13, NAVY, 12, 6),
        ("Heading 3", 11, SLATE, 10, 4),
    ):
        st = styles[name]
        set_font(st, size=size, bold=name != "Subtitle", color=color, italic=False)
        st.paragraph_format.space_before = Pt(before)
        st.paragraph_format.space_after = Pt(after)
        st.paragraph_format.line_spacing = 1.15
        st.paragraph_format.keep_with_next = True
        # Remove theme borders/underline that some Title styles carry
        ppr = st.element.get_or_add_pPr()
        for b in ppr.findall(qn("w:pBdr")):
            ppr.remove(b)

    for name in ("List Bullet", "List Bullet 2", "List Number"):
        st = styles[name]
        set_font(st, size=11, color=CHARCOAL)
        st.paragraph_format.space_after = Pt(4)
        st.paragraph_format.line_spacing = 1.15


def shade(cell, hex_fill):
    tcpr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tcpr.append(shd)


def set_cell_margins(table, top=80, bottom=80, left=120, right=120):
    tblpr = table._tbl.tblPr
    mar = OxmlElement("w:tblCellMar")
    for side, val in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), str(val))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tblpr.append(mar)


def set_table_borders(table, color="BFBFBF", size=4):
    tblpr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), str(size))
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    tblpr.append(borders)


def mark_header_row(row):
    trpr = row._tr.get_or_add_trPr()
    hdr = OxmlElement("w:tblHeader")
    hdr.set(qn("w:val"), "true")
    trpr.append(hdr)


# ---------------------------------------------------------------- content helpers

def para(doc, text="", bold_lead=None, style=None, align=None):
    p = doc.add_paragraph(style=style)
    if bold_lead:
        r = p.add_run(bold_lead)
        set_font(r, bold=True, color=CHARCOAL)
    if text:
        add_inline(p, text)
    if align is not None:
        p.alignment = align
    return p


def add_inline(p, text):
    """Render text, treating `backtick` spans as monospace inline code."""
    parts = text.split("`")
    for i, chunk in enumerate(parts):
        if not chunk:
            continue
        r = p.add_run(chunk)
        if i % 2 == 1:
            set_font(r, name=MONO, size=10, color=NAVY)
        else:
            set_font(r, color=CHARCOAL)


def bullet(doc, text, bold_lead=None, level=1):
    return para(doc, text, bold_lead=bold_lead, style="List Bullet" if level == 1 else "List Bullet 2")


def numbered(doc, text, bold_lead=None):
    return para(doc, text, bold_lead=bold_lead, style="List Number")


def code_block(doc, lines):
    """Render a shaded, bordered single-cell table holding monospace command/SQL text."""
    t = doc.add_table(rows=1, cols=1)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t, color="D9D9D9")
    set_cell_margins(t, 100, 100, 160, 160)
    cell = t.cell(0, 0)
    shade(cell, "F2F4F8")
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    for i, line in enumerate(lines):
        r = p.add_run(line)
        set_font(r, name=MONO, size=9.5, color=CHARCOAL)
        if i < len(lines) - 1:
            r.add_break()
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def table(doc, headers, rows, widths=None):
    t = doc.add_table(rows=1, cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(t)
    set_cell_margins(t)
    hdr = t.rows[0]
    mark_header_row(hdr)
    for i, h in enumerate(headers):
        c = hdr.cells[i]
        shade(c, "002060")
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(h)
        set_font(r, size=10, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
    for ri, row in enumerate(rows):
        cells = t.add_row().cells
        for i, val in enumerate(row):
            c = cells[i]
            if ri % 2 == 1:
                shade(c, "F2F4F8")
            p = c.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_after = Pt(0)
            parts = str(val).split("`")
            for j, chunk in enumerate(parts):
                if not chunk:
                    continue
                r = p.add_run(chunk)
                if j % 2 == 1:
                    set_font(r, name=MONO, size=9, color=NAVY)
                else:
                    set_font(r, size=10, color=CHARCOAL)
    if widths:
        for row in t.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Inches(w)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


def add_field(run, instr):
    for kind, text in (("begin", None), (None, instr), ("separate", None), ("end", None)):
        if kind:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), kind)
        else:
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = text
        run._r.append(el)


def configure_section(doc):
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Inches(8.5), Inches(11)
    for side in ("left_margin", "right_margin"):
        setattr(sec, side, Inches(1))
    sec.top_margin, sec.bottom_margin = Inches(0.9), Inches(0.9)
    sec.different_first_page_header_footer = False

    hp = sec.header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = hp.add_run("United Rentals  |  WhereScape RED to Coalesce Migration Summary & Runbook")
    set_font(r, size=8.5, color=SLATE)

    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fp.paragraph_format.space_after = Pt(2)
    r = fp.add_run(DISCLAIMER)
    set_font(r, size=8, italic=True, color=SLATE)

    pp = sec.footer.add_paragraph()
    pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = pp.add_run("Page ")
    set_font(r, size=8, color=SLATE)
    r = pp.add_run()
    set_font(r, size=8, color=SLATE)
    add_field(r, "PAGE")


# ---------------------------------------------------------------- document body

def title_page(doc):
    for _ in range(5):
        doc.add_paragraph()
    p = doc.add_paragraph("United Rentals", style="Title")
    p = doc.add_paragraph("Data Warehouse Modernization: WhereScape RED to Coalesce.io on Snowflake",
                          style="Subtitle")
    p = doc.add_paragraph("Migration Summary and Operational Runbook", style="Subtitle")
    for r in p.runs:
        r.font.bold = True

    doc.add_paragraph()
    table(doc, ["Attribute", "Value"], [
        ["Project Name", "United Rentals"],
        ["Source System", "WhereScape RED v8.4.2.0 (targeting automated cloud migration)"],
        ["Target System", "Coalesce Desktop (Coalesce.io) on Snowflake"],
        ["Code Repository", REPO_URL],
        ["Working Branch", "feature/new_migration (merged to main via pull requests)"],
        ["Local Workspace", f"`{WORKSPACE}`"],
        ["Architecture", "Data Vault 2.0 (Stage, Hub, Link, Satellite) with EDW 3NF relational layer"],
        ["Document Type", "Migration Summary, Technical Runbook and Maintenance Guide"],
        ["Document Date", "30 October 2026"],
    ], widths=[1.8, 4.7])
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def toc(doc):
    doc.add_paragraph("Contents", style="Heading 1")
    for line in (
        "1. Executive Summary and Local Workspace Setup",
        "2. Data Vault 2.0 and EDW Migration Methodologies",
        "3. Challenges Faced, Technical Discoveries and Resolutions",
        "4. Architectural Limitations and Proactive Maintenance Operations",
        "5. Migration Prompt Library",
        "Appendix A. Migrated Object Inventory",
        "Appendix B. Repository Commit Timeline",
    ):
        p = doc.add_paragraph(line)
        p.paragraph_format.left_indent = Inches(0.25)
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def section_1(doc):
    doc.add_paragraph("1. Executive Summary and Local Workspace Setup", style="Heading 1")

    doc.add_paragraph("1.1 Executive Summary", style="Heading 2")
    para(doc, "United Rentals operated its enterprise data warehouse on WhereScape RED v8.4.2.0, a "
              "procedural, repository-driven automation tool in which load, stage and Data Vault objects "
              "are generated as database-resident stored procedures and scheduled through the RED "
              "scheduler. While RED delivered a functional Data Vault, its object definitions lived inside "
              "a metadata repository database rather than in source control, change review depended on "
              "manual inspection of generated procedures, and promotion between environments required "
              "repository exports. These characteristics limited peer review, slowed delivery and coupled "
              "the warehouse to a single on-premises automation server.")
    para(doc, "The modernization programme re-platforms the warehouse onto Coalesce.io running against "
              "Snowflake. Coalesce is a metadata-driven, column-aware ELT platform in which every node "
              "(source, stage, hub, link, satellite or dimension) is expressed as a declarative YAML "
              "definition, rendered to Snowflake SQL by reusable node-type templates, and versioned in Git. "
              "The migration was executed with an AI-assisted engineering workflow: Claude Code, connected "
              "through the Model Context Protocol (MCP) to both the WhereScape RED metadata repository and "
              "the Coalesce Transform platform, read the legacy object definitions and authored equivalent "
              "Coalesce nodes through the Coalesce CLI (`coa`).")
    para(doc, "Key outcomes of the migration:")
    bullet(doc, " 23 source (load) nodes registered for the Traders North and Traders South operational "
                "systems and the country-code file source.", bold_lead="Source layer:")
    bullet(doc, " 29 Stage nodes rebuilt from WhereScape RED metadata, preserving the regional "
                "north/south split and the DSS audit columns.", bold_lead="Stage layer:")
    bullet(doc, " 9 Hubs, 5 Links and 38 Satellites implemented on the Coalesce Data Vault "
                "package, with hash keys and hash-diff change detection.", bold_lead="Raw Vault:")
    bullet(doc, " The R_COUNTRY reference table migrated as a 3NF storage node, plus an initial "
                "fact node used to validate the end-to-end pipeline.", bold_lead="EDW layer:")
    bullet(doc, " All 106 node definitions are under Git control on the feature/new_migration branch "
                "and promoted to main through reviewed pull requests.", bold_lead="Governance:")

    doc.add_paragraph("1.2 Local Workstation Initialization and Tool Setup", style="Heading 2")
    para(doc, "The following procedure establishes a migration workstation from a clean Windows 10/11 "
              "machine. Execute the steps in order from an elevated PowerShell or Git Bash terminal. "
              "Prerequisites are Node.js LTS with npm, Git for Windows, network access to "
              "app.coalescesoftware.io and to the SQL Server instance hosting the WhereScape RED "
              "repository, and an active Coalesce user with an access token.")

    doc.add_paragraph("Step 1. Create the local project directory structure", style="Heading 3")
    para(doc, "Create the root folder that will hold the MCP bridge and the Coalesce workspace clone.")
    code_block(doc, [r"mkdir I:\DataEntrega\Red_Coalesce_Migration",
                     r"cd I:\DataEntrega\Red_Coalesce_Migration"])

    doc.add_paragraph("Step 2. Initialize the Node.js workspace", style="Heading 3")
    para(doc, "Generate a default package.json so that the MCP server dependencies are tracked locally.")
    code_block(doc, ["npm init -y"])

    doc.add_paragraph("Step 3. Install the MCP and database packages", style="Heading 3")
    para(doc, "`@modelcontextprotocol/sdk` provides the MCP server runtime used to expose the WhereScape "
              "RED repository to Claude; `knex` supplies the SQL query builder and connection pooling; "
              "`mssql` is the SQL Server driver for the RED metadata database.")
    code_block(doc, ["npm install @modelcontextprotocol/sdk knex mssql"])

    doc.add_paragraph("Step 4. Enable Windows Authentication", style="Heading 3")
    para(doc, "The WhereScape RED repository is secured with integrated Windows security. The native "
              "ODBC-backed driver `msnodesqlv8` allows the MCP server to authenticate as the logged-on "
              "domain user without storing a SQL password.")
    code_block(doc, ["npm install msnodesqlv8"])

    doc.add_paragraph("Step 5. Configure the MCP server matrix", style="Heading 3")
    para(doc, "Register the `wherescape-red-mcp` server (stdio transport, launched by Node) and the "
              "`coalesce-transform` server (HTTP transport to app.coalescesoftware.io) in the MCP "
              "configuration file. The RED server receives the SQL Server host, repository database and "
              "the trusted-connection flag through environment variables; the Coalesce server reads its "
              "token from `.env.coalesce`. See Section 4.2 regarding the required file name.")

    doc.add_paragraph("Step 6. Install the Coalesce CLI globally", style="Heading 3")
    code_block(doc, ["npm install -g @coalescesoftware/coa"])

    doc.add_paragraph("Step 7. Verify the tool chain", style="Heading 3")
    para(doc, "Confirm that the CLI is resolvable on PATH and report its version. A version string "
              "confirms a successful installation; a 'command not found' error indicates the global npm "
              "prefix is not on PATH.")
    code_block(doc, ["coa --version"])

    doc.add_paragraph("Step 8. Synchronize with the remote Git repository", style="Heading 3")
    para(doc, "Clone the Coalesce workspace repository. Git registers the source as the remote named "
              "`origin` automatically; check out the migration branch afterwards.")
    code_block(doc, [f"git clone {REPO_URL}",
                     "cd Red_Coalesce_Migration_COA_CLI",
                     "git checkout feature/new_migration",
                     "git remote -v"])

    doc.add_paragraph("Step 9. Validate the MCP link inside Claude", style="Heading 3")
    para(doc, "Within the Claude Code session opened at the workspace root, run the platform skill below. "
              "It installs or verifies the Coalesce Transform MCP connection and confirms that the "
              "coalesce-transform tools are visible to the agent.")
    code_block(doc, ["/coalesce-platform-skills:coalesce-transform-mcp-install"])

    doc.add_paragraph("1.3 Workspace Layout", style="Heading 2")
    table(doc, ["Path", "Purpose"], [
        ["`workspace.yml`", "Workspace-level configuration for the Coalesce project."],
        ["`locations.yml`", "Storage location mapping (SRC and TARGET) to Snowflake databases/schemas."],
        ["`nodes/`", "One YAML file per node, named `<LOCATION>-<NODE_NAME>.yml` (106 files)."],
        ["`nodeTypes/`", "Node-type templates (create and run SQL) used to render nodes."],
        ["`packages/`", "`base-node-types.yml` and `DV.yml` (Data Vault package registration)."],
        ["`environments/`", "Environment definitions used for deployment targets."],
        ["`AGENTS.md`, `CLAUDE.md`", "Agent instructions for running `coa` and authoring nodes."],
    ], widths=[2.0, 4.5])


def section_2(doc):
    doc.add_paragraph("2. Data Vault 2.0 and EDW Migration Methodologies", style="Heading 1")
    para(doc, "Each architectural layer was migrated with a dedicated prompt blueprint: a structured "
              "instruction set given to the AI agent that defines how to read the RED object, which "
              "Coalesce node type to emit, how columns and transforms must be mapped, and which "
              "validations to run before committing. The blueprints below are the governing standards; "
              "any future object must conform to them.")

    # Stage
    doc.add_paragraph("2.1 Stage Node Layer", style="Heading 2")
    para(doc, "WhereScape RED loaded identical entity structures from two operational systems, Traders "
              "North and Traders South, producing paired load tables such as "
              "`load_customers_traders_north` and `load_customers_traders_south`. The stage blueprint "
              "applies the following rules:")
    bullet(doc, " Enumerate RED stage objects and group them by business domain (Customer, Employee, "
                "Order, Order Details, Product, Category, Supplier, Shipper, Region, Territory). The "
                "regional suffix `_traders_north` / `_traders_south` is parsed as a source qualifier, not "
                "as part of the entity name.", bold_lead="Domain grouping.")
    bullet(doc, " Each regional variant becomes its own Coalesce Stage node "
                "(for example `STAGE_CUSTOMER_TRADERS_NORTH`) sourced from its matching SRC node. "
                "Keeping regions separate preserves record source lineage and allows independent reloads.",
           bold_lead="Isolation.")
    bullet(doc, " Multi-entity stages that pre-join keys for links "
                "(for example `STAGE_ORDERS_SHIPPERS_EMPLOYEES_CUSTOMERS_TRADERS_NORTH`) are migrated "
                "with the RED join conditions translated into the node's join clause.",
           bold_lead="Composite stages.")
    bullet(doc, " Every business column in the RED stage definition must appear in the Coalesce node with "
                "the same name, data type, nullability and transform.", bold_lead="Column fidelity.")
    para(doc, "Audit column mappings applied to every Stage node:")
    table(doc, ["RED Column", "Coalesce Transform", "Purpose"], [
        ["`dss_load_date`", "`CURRENT_TIMESTAMP()::TIMESTAMP_NTZ`", "Batch load timestamp"],
        ["`dss_record_source`", "`'TRADERS_NORTH'` / `'TRADERS_SOUTH'`", "Originating system"],
        ["`dss_create_time`", "`CAST(CURRENT_TIMESTAMP AS TIMESTAMP)`", "Row creation time"],
        ["`dss_update_time`", "`CAST(CURRENT_TIMESTAMP AS TIMESTAMP)`", "Row update time"],
    ], widths=[1.6, 2.9, 2.0])

    # Hub
    doc.add_paragraph("2.2 Hub Node Layer", style="Heading 2")
    para(doc, "Hubs carry the unique list of business keys for a core business concept. Nine hubs were "
              "created: H_CATEGORY, H_CUSTOMER, H_EMPLOYEE, H_ORDER, H_PRODUCT, H_REGION, H_SHIPPER, "
              "H_SUPPLIER and H_TERRITORY.")
    bullet(doc, " Business keys are cast to text and normalised with `UPPER(TRIM(...))`; null keys "
                "are replaced with a sentinel value so that the hash remains deterministic.",
           bold_lead="Business key cleansing.")
    bullet(doc, " The hash key (`HK_H_<ENTITY>`) is computed as `MD5_BINARY` (16-byte BINARY) for "
                "storage efficiency and join performance in Snowflake; `SHA2_BINARY(..., 256)` is the "
                "approved alternative where collision tolerance must be stricter. Composite keys are "
                "concatenated with a `'||'` delimiter before hashing.", bold_lead="Hash key compilation.")
    bullet(doc, " Each hub carries `dss_load_date` and `dss_record_source` from the first stage "
                "that introduced the key, giving a lineage trail back to the Traders North or South "
                "system.", bold_lead="Lineage columns.")
    bullet(doc, " Hubs are insert-only. The Data Vault hub node type inserts only keys not already "
                "present in the target.", bold_lead="Load pattern.")
    code_block(doc, [
        "MD5_BINARY(UPPER(TRIM(COALESCE(CAST(customer_id AS VARCHAR), '-1')))) AS HK_H_CUSTOMER",
    ])

    # Link
    doc.add_paragraph("2.3 Link Node Layer", style="Heading 2")
    para(doc, "Links record relationships between hubs. Five links were migrated: "
              "L_EMPLOYEE_TERRITORIES, L_ORDERS_SHIPPERS_EMPLOYEES_CUSTOMERS, L_ORDER_DETAILS, "
              "L_PRODUCTS_CATEGORIES_SUPPLIERS and L_TERRITORY_REGIONS.")
    bullet(doc, " Each participating hub contributes its hash key as a foreign key column, computed "
                "with exactly the same cleansing and hashing expression as the hub itself so the keys "
                "join without translation.", bold_lead="Foreign hash keys.")
    bullet(doc, " The link hash key (`HK_L_<RELATIONSHIP>`) is the hash of the concatenated, "
                "cleansed business keys of all participating hubs in a fixed, documented order.",
           bold_lead="Link hash key.")
    bullet(doc, " Links are insert-only and unique on the link hash key; they carry "
                "`dss_load_date` and `dss_record_source`.", bold_lead="Load pattern.")
    code_block(doc, [
        "MD5_BINARY(",
        "    UPPER(TRIM(COALESCE(CAST(employee_id AS VARCHAR), '-1'))) || '||' ||",
        "    UPPER(TRIM(COALESCE(CAST(territory_id AS VARCHAR), '-1')))",
        ") AS HK_L_EMPLOYEE_TERRITORIES",
    ])

    # Satellite
    doc.add_paragraph("2.4 Satellite Node Layer", style="Heading 2")
    para(doc, "Satellites hold descriptive attributes and their history. Thirty-eight satellites were "
              "migrated, split by region and by rate of change in line with the RED design: LROC "
              "(low rate of change), MROC (medium), HROC (high), TRANSACT (transactional) and GDPR "
              "(personally identifiable attributes isolated for retention control).")
    bullet(doc, " A hash diff (`HD_S_<NAME>`) is computed over the cleansed, delimiter-joined "
                "descriptive attributes of the satellite.", bold_lead="Hash diff.")
    bullet(doc, " A new row is inserted only when no row for the same parent hash key carries the "
                "same hash diff as the latest record. This blocks redundant versions and keeps "
                "satellites append-only.", bold_lead="Delta detection.")
    code_block(doc, [
        "INSERT INTO S_CUSTOMER_LROC_TRADERS_NORTH",
        "SELECT stg.*",
        "FROM   STAGE_CUSTOMER_TRADERS_NORTH stg",
        "WHERE  NOT EXISTS (",
        "    SELECT 1",
        "    FROM   S_CUSTOMER_LROC_TRADERS_NORTH sat",
        "    WHERE  sat.HK_H_CUSTOMER   = stg.HK_H_CUSTOMER",
        "    AND    sat.HD_S_CUSTOMER_LROC = stg.HD_S_CUSTOMER_LROC",
        ");",
    ])
    para(doc, "The equivalent `LEFT JOIN ... WHERE sat.HK_H_CUSTOMER IS NULL` form against the latest "
              "satellite row per key is used where the Snowflake query profile shows a cheaper plan.")

    # EDW
    doc.add_paragraph("2.5 EDW 3NF Relational Layer", style="Heading 2")
    para(doc, "RED implemented EDW 3NF tables with generated T-SQL/PL-SQL procedures that ran cursor or "
              "row-by-row UPDATE and INSERT statements. In Coalesce these become storage nodes whose run "
              "templates issue set-based Snowflake SQL. R_COUNTRY was the first object converted.")
    table(doc, ["Pattern", "Legacy RED Procedure", "Coalesce / Snowflake Equivalent"], [
        ["SCD Type 1", "UPDATE existing rows, then INSERT missing keys in separate statements",
         "Single `MERGE INTO ... WHEN MATCHED THEN UPDATE WHEN NOT MATCHED THEN INSERT`"],
        ["SCD Type 2", "Close current row with UPDATE of end date, INSERT new version, manage "
                       "current flag in procedure logic",
         "Dimension node type with change-tracking columns; `MERGE` closes the prior version and "
         "inserts the new one in one transaction"],
        ["Truncate/reload", "TRUNCATE plus INSERT ... SELECT in procedure",
         "Table node with truncate-before-insert option"],
    ], widths=[1.2, 2.5, 2.8])


def section_3(doc):
    doc.add_paragraph("3. Challenges Faced, Technical Discoveries and Resolutions", style="Heading 1")

    doc.add_paragraph("3.1 Network Security and Interceptor Barriers", style="Heading 2")
    para(doc, bold_lead="Problem. ", text=
               "The Claude MCP bridge repeatedly failed to connect to the coalesce-transform server "
               "(v7.48.2). Tool discovery timed out or returned invalid HTTP responses even though the "
               "Coalesce web application was reachable in a browser.")
    para(doc, bold_lead="Root cause. ", text=
               "Norton Web/Mail Shield was performing active HTTPS scanning (SSL inspection) on the "
               "workstation. The shield re-signed TLS sessions with its own certificate and intercepted "
               "local HTTP traffic, which the Node-based MCP client neither trusted nor tolerated, so the "
               "loopback and outbound connections were dropped.")
    para(doc, bold_lead="Resolution. ", text=
               "An explicit exclusion for `app.coalescesoftware.io` was added in Norton Web/Mail Shield "
               "(HTTPS scanning was temporarily disabled during diagnosis to confirm the cause). With "
               "interception removed, the Coalesce server was reconfigured as a direct HTTP transport "
               "in `.mcp.json` and connected successfully.")
    para(doc, bold_lead="Preventive guidance. ", text=
               "Any endpoint security product with TLS inspection must have Coalesce endpoints "
               "excluded before onboarding a new workstation. If the MCP server reports "
               "'InvalidHTTPResponse', check the security software first.")

    doc.add_paragraph("3.2 Git Credential and Terminal Push Failures", style="Heading 2")
    para(doc, bold_lead="Problem. ", text=
               "Commit aa242cd ('Stage nodes are added') contained the 26 initial Stage nodes together "
               "with `.gitignore`, `AGENTS.md`, `CLAUDE.md` and `packages/base-node-types.yml`. Pushing "
               "it from the agent terminal failed with: `could not read Username for "
               "'https://github.com'`. The agent's non-interactive shell had no TTY on which Git could "
               "prompt for credentials and no credential helper was configured.")
    para(doc, bold_lead="Resolution. ", text=
               "Git Credential Manager was registered globally so that credentials are acquired once "
               "through the browser and cached in the Windows Credential Manager. The push was then "
               "issued with an explicit working-tree path, decoupled from the shell's current "
               "directory:")
    code_block(doc, [
        "git config --global credential.helper manager",
        f'git -C "{WORKSPACE}" push origin feature/new_migration',
    ])
    para(doc, "The pull request was reconciled and merged in the GitHub web interface, after "
                          "which the branch was refreshed in the Coalesce UI so that the workspace "
                          "reflected the committed nodes.")

    doc.add_paragraph("3.3 Data Vault Object Structural Anomalies", style="Heading 2")
    para(doc, bold_lead="Problem. ", text=
               "Three related defects appeared in the first migration pass:")
    bullet(doc, " Generated Hub nodes were created with the standard Stage node type because the "
                "Data Vault node-type package was not yet installed in the target workspace, so no hub "
                "type existed to assign.", bold_lead="Misclassified hubs.")
    bullet(doc, " Several business columns present in the RED stage definitions were absent from the "
                "generated Stage nodes.", bold_lead="Missing columns.")
    bullet(doc, " The transform for `dss_create_time` compiled incorrectly against Snowflake.",
           bold_lead="Audit timestamp error.")
    para(doc, bold_lead="Resolution. ", text=
               "The native Data Vault node-type package was provisioned and registered in the Coalesce "
               "interface (`packages/DV.yml`), and the hubs were regenerated on the hub node type "
               "(commit 8648d07). Stage nodes were rebuilt directly from WhereScape RED metadata with a "
               "column-by-column audit to restore every missing attribute (commit ac2ebe0). The "
               "`dss_create_time` transform was corrected to `CAST(CURRENT_TIMESTAMP AS TIMESTAMP)` so "
               "that it aligns with the target warehouse timestamp type (commit f65cf53).")

    doc.add_paragraph("3.4 Issue Summary", style="Heading 2")
    table(doc, ["Issue", "Root Cause", "Resolution", "Status"], [
        ["MCP connection failure", "Norton HTTPS scanning / SSL inspection",
         "Exclusion for app.coalescesoftware.io; direct HTTP in `.mcp.json`", "Resolved"],
        ["Git push authentication", "No credential helper; non-interactive shell",
         "`credential.helper manager`; `git -C` targeted push", "Resolved"],
        ["Hubs created as Stage nodes", "Data Vault package not installed",
         "Registered DV package; regenerated hubs", "Resolved"],
        ["Missing Stage columns", "Incomplete metadata extraction",
         "Rebuilt stages from RED metadata with column audit", "Resolved"],
        ["`dss_create_time` compile error", "Incorrect transform expression",
         "`CAST(CURRENT_TIMESTAMP AS TIMESTAMP)`", "Resolved"],
    ], widths=[1.5, 1.7, 2.5, 0.8])


def section_4(doc):
    doc.add_paragraph("4. Architectural Limitations and Proactive Maintenance Operations", style="Heading 1")

    doc.add_paragraph("4.1 Refreshing Environment State and Platform Tokens", style="Heading 2")
    para(doc, "The coalesce-transform MCP server and the `coa` CLI authenticate with a Coalesce access "
              "token held in `.env.coalesce`. Tokens are time-limited and are not refreshed "
              "automatically. When a token expires, MCP tool calls and `coa` commands fail with "
              "authentication errors.")
    numbered(doc, "Sign in to app.coalescesoftware.io and generate a new access token from the user "
                  "settings page.")
    numbered(doc, "Open `.env.coalesce` in the workspace and replace the value on the token line with the "
                  "new token. Do not commit this file; it must remain listed in `.gitignore`.")
    numbered(doc, "Restart the Claude Code session (or reconnect the MCP server) so that the new value "
                  "is loaded.")
    numbered(doc, "Confirm connectivity by listing workspaces through the MCP tools or by running a "
                  "read-only `coa` command.")

    doc.add_paragraph("4.2 MCP Configuration File Boundary", style="Heading 2")
    para(doc, "The `wherescape-red-mcp` server is currently defined only in `mcp.json` (no leading dot), "
              "the file created during the initial bridge setup. Claude Code discovers project-scoped "
              "servers exclusively from `.mcp.json` (with a leading dot). As a result, the agent sees "
              "the coalesce-transform server but not the WhereScape RED server, and RED metadata cannot "
              "be queried in the same session as Coalesce authoring.")
    para(doc, "Required maintenance task:")
    numbered(doc, "Open both `mcp.json` and `.mcp.json`.")
    numbered(doc, "Copy the complete `wherescape-red-mcp` entry (command, args and env block) from "
                  "`mcp.json` into the `mcpServers` object of `.mcp.json`, alongside "
                  "`coalesce-transform`.")
    numbered(doc, "Validate the JSON, restart Claude Code and approve the project server when prompted.")
    numbered(doc, "Confirm both servers are listed as connected, then retire `mcp.json` (or keep it "
                  "only as a documented backup) to remove the duplicate source of truth.")

    doc.add_paragraph("4.3 Other Operational Considerations", style="Heading 2")
    bullet(doc, " Disabling or upgrading endpoint security can reintroduce TLS interception; "
                "re-verify the exclusions after any security product update.", bold_lead="Endpoint security.")
    bullet(doc, " `msnodesqlv8` is a native module tied to the installed Node.js version; reinstall "
                "it after a Node major-version upgrade.", bold_lead="Native driver.")
    bullet(doc, " After merging to main, refresh the branch in the Coalesce UI before deploying so "
                "that the platform builds from the merged commit.", bold_lead="Branch sync.")
    bullet(doc, " Hash expressions in hubs, links and satellites must stay identical across layers; any "
                "change requires a coordinated update and a historical reload.",
           bold_lead="Hash consistency.")


PROMPT_LAYERS = (
    ("5.2 Stage Node Prompt", "stage.txt",
     "Supplied with a screenshot of the Stage table list from the WhereScape RED object tree. "
     "Produced the 29 Stage nodes, the north/south grouping and the audit column blueprint "
     "(commits aa242cd and ac2ebe0)."),
    ("5.3 Hub Node Prompt", "hub.txt",
     "Supplied with a screenshot of the nine Hub objects in the RED repository. Produced the "
     "hash key standard and the nine Hub nodes on the Data Vault package (commit 8648d07)."),
    ("5.4 Link Node Prompt", "link.txt",
     "Used as a template: the bracketed line is replaced with the Link metadata, DDL or generated "
     "SQL exported from RED. Produced the five Link nodes (commit 887503e)."),
    ("5.5 Satellite Node Prompt", "satellite.txt",
     "Used as a template: the bracketed line is replaced with the Satellite metadata, DDL or "
     "generated SQL exported from RED. Produced the 38 Satellite nodes (commit 887503e)."),
    ("5.6 EDW 3NF Node Prompt", "edw_3nf.txt",
     "Used as a template: the bracketed line is replaced with the 3NF table metadata or update "
     "procedure from RED. Produced the R_COUNTRY storage node (commit 1412d32)."),
)


def section_5(doc):
    doc.add_paragraph("5. Migration Prompt Library", style="Heading 1")
    para(doc, "This section records the prompts given to Claude to migrate each layer, so that the "
              "migration can be repeated for new subject areas or audited later. The prompt text is "
              "reproduced as it was used; the Markdown markers inside it (headings, bold and code "
              "quotes) are part of the prompt and should be kept when it is reused.")

    doc.add_paragraph("5.1 Source Node Layer", style="Heading 2")
    para(doc, "Source nodes were not generated from a prompt. The source tables were created in "
              "Snowflake by generating their DDL from WhereScape RED, and the resulting tables were "
              "then imported into Coalesce as Source nodes in the SRC location (commit 0a2ded8). The "
              "Stage prompt below assumes this step is complete.")

    prompt_dir = Path(__file__).with_name("prompts")
    for heading, filename, usage in PROMPT_LAYERS:
        doc.add_paragraph(heading, style="Heading 2")
        para(doc, usage, bold_lead="Usage. ")
        lines = (prompt_dir / filename).read_text(encoding="utf-8").splitlines()
        code_block(doc, [ln.rstrip() for ln in lines])


def appendices(doc):
    doc.add_paragraph("Appendix A. Migrated Object Inventory", style="Heading 1")
    table(doc, ["Layer", "Location", "Count", "Examples"], [
        ["Source", "SRC", "23", "LOAD_CUSTOMERS_TRADERS_NORTH, LOAD_COUNTRYCODES_FILE_SOURCES"],
        ["Stage", "TARGET", "29", "STAGE_CUSTOMER_TRADERS_SOUTH, STAGE_COUNTRY"],
        ["Hub", "TARGET", "9", "H_CUSTOMER, H_ORDER, H_PRODUCT"],
        ["Link", "TARGET", "5", "L_ORDER_DETAILS, L_TERRITORY_REGIONS"],
        ["Satellite", "TARGET", "38", "S_CUSTOMER_LROC_TRADERS_NORTH, S_EMPLOYEE_GDPR_TRADERS_SOUTH"],
        ["EDW 3NF", "TARGET", "1", "R_COUNTRY"],
        ["Fact", "TARGET", "1", "FCT_NODE"],
        ["Total", "", "106", ""],
    ], widths=[1.1, 0.9, 0.7, 3.8])

    doc.add_paragraph("Appendix B. Repository Commit Timeline", style="Heading 1")
    table(doc, ["Commit", "Description"], [
        ["099944d", "Init commit"],
        ["6a04e91 / 49e2cba", "Fact node added"],
        ["0a2ded8", "Added source nodes"],
        ["aa242cd", "Stage nodes are added (26 Stage nodes plus workspace files)"],
        ["ac2ebe0", "Rebuild Stage nodes from WhereScape RED metadata"],
        ["f65cf53", "Use CAST(CURRENT_TIMESTAMP AS TIMESTAMP) for DSS_CREATE_TIME"],
        ["8648d07", "Add 9 Data Vault hubs migrated from WhereScape RED"],
        ["887503e", "Add Data Vault links and satellites migrated from WhereScape RED"],
        ["1412d32", "Add EDW 3NF table R_COUNTRY migrated from WhereScape RED"],
    ], widths=[1.5, 5.0])


def main(out_path):
    doc = Document()
    configure_styles(doc)
    configure_section(doc)
    doc.core_properties.title = "United Rentals Migration Summary and Runbook"
    doc.core_properties.subject = "WhereScape RED to Coalesce.io on Snowflake"
    title_page(doc)
    toc(doc)
    section_1(doc)
    section_2(doc)
    section_3(doc)
    section_4(doc)
    section_5(doc)
    appendices(doc)
    doc.save(out_path)
    print(f"Saved {out_path}")


if __name__ == "__main__":
    default = Path(__file__).with_name("United_Rentals_Migration_Summary_and_Runbook.docx")
    main(Path(sys.argv[1]) if len(sys.argv) > 1 else default)
