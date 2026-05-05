---
name: Definian FBDI Compliance Report
description: Oracle FBDI field-level change report for internal Definian Oracle integration consultants
colors:
  blue-authority: "#0D2C71"
  green-signal: "#00AB63"
  midnight: "#02072D"
  slate-annotation: "#3C405B"
  mist-divider: "#D8D7EE"
  cloud-surface: "#F7F7FB"
  canvas: "#FFFFFF"
  caution-amber: "#B8860B"
  caution-amber-bg: "#FFFBF0"
  remove-red: "#C0392B"
typography:
  display:
    fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif"
    fontSize: "36px"
    fontWeight: 800
    lineHeight: 1.1
    letterSpacing: "-0.5px"
  headline:
    fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif"
    fontSize: "22px"
    fontWeight: 700
    lineHeight: 1.2
  title:
    fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif"
    fontSize: "18px"
    fontWeight: 700
    lineHeight: 1.3
  body:
    fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif"
    fontSize: "13px"
    fontWeight: 400
    lineHeight: 1.5
  label:
    fontFamily: "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif"
    fontSize: "10.5px"
    fontWeight: 600
    lineHeight: 1
    letterSpacing: "0.5px"
  mono:
    fontFamily: "ui-monospace, SFMono-Regular, Consolas, monospace"
    fontSize: "11.5px"
    fontWeight: 400
    lineHeight: 1.4
rounded:
  xs: "2px"
  sm: "4px"
  md: "6px"
  pill: "12px"
spacing:
  xs: "4px"
  sm: "8px"
  md: "14px"
  lg: "22px"
  xl: "36px"
components:
  file-head:
    backgroundColor: "{colors.blue-authority}"
    textColor: "{colors.canvas}"
    rounded: "0"
    padding: "14px 18px"
  badge-added:
    backgroundColor: "{colors.green-signal}"
    textColor: "{colors.canvas}"
    rounded: "{rounded.pill}"
    padding: "1px 8px"
  badge-removed:
    backgroundColor: "{colors.remove-red}"
    textColor: "{colors.canvas}"
    rounded: "{rounded.pill}"
    padding: "1px 8px"
  badge-modified:
    backgroundColor: "{colors.caution-amber}"
    textColor: "{colors.canvas}"
    rounded: "{rounded.pill}"
    padding: "1px 8px"
  badge-shifted:
    backgroundColor: "{colors.slate-annotation}"
    textColor: "{colors.canvas}"
    rounded: "{rounded.pill}"
    padding: "1px 8px"
  module-tag:
    backgroundColor: "{colors.mist-divider}"
    textColor: "{colors.midnight}"
    rounded: "{rounded.pill}"
    padding: "1px 8px"
  module-tag-financials:
    backgroundColor: "{colors.green-signal}"
    textColor: "{colors.canvas}"
    rounded: "{rounded.pill}"
    padding: "2px 10px"
---

# Design System: Definian FBDI Compliance Report

## 1. Overview

**Creative North Star: "The Signal Extract"**

Each Oracle Cloud quarterly release buries hundreds of FBDI field changes across dozens of template files. This report's entire purpose is to extract the signal — exactly what changed, in exactly which field, at exactly which position — and deliver it with the authority of Definian's brand. The design system exists to serve that extraction: nothing decorative, nothing ambient, nothing that slows a consultant's read-through. Every visual decision is justified by whether it helps a consultant orient faster, act with more confidence, or communicate the finding more clearly to their client team.

The system is built on the Definian canonical palette — Authority Blue, Signal Green, and their four supporting neutrals — supplemented by two semantic additions: Caution Amber for modified fields and Remove Red for removals. Color is the primary semantic layer; shape and spacing are structure. The system is deliberately flat (no shadows), tightly typeset (system sans for prose, monospace for all technical identifiers), and unambiguous about hierarchy. HTML and PDF are treated as separate media with separate design constraints; a consultant running the HTML report in a browser gets interaction and space; a consultant printing the PDF gets density and legibility. These are different deliverables that happen to share a source template.

The one thing that must be constant across both media: an unmistakable Definian visual signature. Deep navy headers, green signal accents, clean label typography — a consultant who has seen one Definian report should recognize the next one instantly.

**Key Characteristics:**
- Flat tonal depth — no shadows, surfaces differentiated by background color alone
- Semantic color system — every color has one job and never doubles as decoration
- Mono for machines — all technical identifiers (field names, table prefixes, filenames) in monospace
- Definian brand through structure — blue headers, green accents, the palette does the branding work
- Two-medium design — HTML and PDF optimized independently for their respective audiences and constraints

## 2. Colors: The Definian Signal Palette

Six canonical Definian brand colors plus two approved semantic supplementals. Every color has a single defined role. Introducing off-palette colors is prohibited except for functional semantic extension (a new change type requiring a new semantic color); decorative expansion is never justified.

### Primary
- **Authority Blue** (`#0D2C71`): The structural backbone. Section headers, file header bars, table column headers, summary table fills, and the cover hero background. This is Definian's identity color; it appears on every screen and anchors every major section boundary.
- **Signal Green** (`#00AB63`): Positive signal. Used for: field additions (ADDED), Financials module identification, brand wordmark on the cover, and the count badges of Added change blocks. Its meaning is "this is new, this is good, this is Definian." Never used decoratively.

### Secondary
- **Caution Amber** (`#B8860B`): Modified fields only. The change-block header color and badge fill for MODIFIED and MULTI change types. Paired with its tinted background (`#FFFBF0`) for in-base notes and warning callouts.
- **Remove Red** (`#C0392B`): Removed fields only. The change-block header color and badge fill for REMOVED change types. Also the delete-state for action checkboxes. Used only when something is gone.

### Neutral
- **Near-Black Midnight** (`#02072D`): Primary text across all document surfaces. Also the deep end of the cover gradient. Never pure black; this dark navy-tinted value keeps all body text within the Definian palette.
- **Slate Annotation** (`#3C405B`): Supporting text on light surfaces — lede paragraphs, metadata rows, table annotations, renamed/shifted change headers. Also used for Dark Gray on Midnight backgrounds per the brand guide.
- **Mist Divider** (`#D8D7EE`): All 1px structural borders: file section containers, table row dividers, detail element borders, section rule lines. Also the default module tag background.
- **Cloud Surface** (`#F7F7FB`): Card and section backgrounds — module cards, shift summary boxes, collapsible content areas. One step above the canvas white to create tonal depth without a border.
- **Canvas** (`#FFFFFF`): Document ground. The base page background and the file-body interior. This is the Definian brand white; use the token name, not the raw hex.

### Named Rules

**The Signal Rule.** Color carries semantic meaning, never decoration. Green = added/positive, Red = removed/critical, Amber = modified/caution, Blue = structure/authority. Using Signal Green on a non-addition element is a violation. The palette's power comes from predictability.

**The Six Plus Two Rule.** The six canonical Definian colors are primary. Caution Amber and Remove Red are the only approved supplementals, existing solely because the report requires two additional semantic signals. Any future semantic extension (a new change classification) may introduce a new supplemental with a documented role — but only if no existing color can carry the meaning without ambiguity.

**The No Off-Palette Rule.** Introducing a color outside the eight documented tokens — even a tint or shade of an existing color — is prohibited without explicit justification and documentation here.

## 3. Typography: Precision Sans

**Display/Body Font:** System sans — `-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif`
**Mono Font:** `ui-monospace, SFMono-Regular, Consolas, monospace` — for all Oracle field names, table prefixes, and technical identifiers

**Character:** The system font stack is an explicit choice, not a default. These are the fonts consultants see in their terminals, IDEs, and OS interfaces every day. The report reads as a professional, technical document at home in the same environment where the work happens. Monospace for all field names draws a clear boundary between "text you read" and "text you type"; a consultant should never have to wonder whether `PERSON_ID` is a prose label or a real identifier.

### Hierarchy

- **Display** (800 weight, 36px, line-height 1.1, letter-spacing −0.5px): Cover title only. One instance per report. Tight spacing reinforces the compressed, authoritative character.
- **Headline** (700 weight, 22px, line-height 1.2, color: Authority Blue): Section `h2` headers. Report-level structure anchors. Authority Blue mandatory.
- **Title** (700 weight, 18px, line-height 1.3, color: white on blue): File section names in the blue header bar. Per-FBDI-file identity marker.
- **Body** (400 weight, 13px, line-height 1.5, color: Midnight): Main prose, table data cells, lede paragraphs, metadata rows.
- **Label** (600–700 weight, 10–11px, letter-spacing 0.5px, uppercase): All table column headers, module names, badge text, category sub-labels, change-type h4 headings. Uppercase + spacing is mandatory; this is the visual signature of all structural annotation in the report.
- **Mono** (400 weight, 11.5px, line-height 1.4, color: Midnight or Slate Annotation): All Oracle/Applaud field names, table prefix codes, filename references. No exceptions.

### Named Rules

**The Mono Rule.** Every technical identifier — Oracle field name, Applaud column name, table prefix, filename — must render in monospace. A field name in regular sans is a defect, not a style choice.

**The Label Rule.** All column headers and structural sub-labels are uppercase with letter-spacing ≥0.5px. Consistency in label treatment is what creates reading rhythm across the report's dense tables. Breaking this rule — even once — disrupts the visual contract with the reader.

## 4. Elevation: Flat by Signal

This system has no shadows, no blurs, no elevation effects. Depth is communicated entirely through tonal layering and borders:

- **Canvas** (`#FFFFFF`) is the document ground — the base surface.
- **Cloud Surface** (`#F7F7FB`) is the card layer — module cards, shift summaries, collapsible interiors.
- **Mist Divider** (`#D8D7EE`) is the structure layer — dividers, 1px borders, table row rules.
- **Authority Blue** and **Midnight** are the authority layer — file headers, cover, summary table fills.

The three light surfaces (Canvas → Cloud Surface → Mist Divider) provide all the tonal differentiation the layout needs. The PDF rendering path through weasyprint further enforces this: shadows are unreliable in PDF generation, flat tonal surfaces are not. Flatness is a constraint made into a design principle.

### Named Rules

**The Flat Signal Rule.** No `box-shadow` anywhere. If you need to distinguish a surface, use a background color shift (Canvas → Cloud Surface) or a 1px Mist Divider border. If you're reaching for a shadow, you are solving a hierarchy problem — solve it with color and border instead.

## 5. Components

### Cover Hero
The only fully branded surface in the report. A navy-to-midnight gradient carries the Definian identity on the first visible surface. The brand wordmark in Signal Green locks the Definian palette before the consultant reads a word of data.

- **Background:** `linear-gradient(135deg, #0D2C71 0%, #02072D 100%)`
- **Brand wordmark:** 11px, weight 700, uppercase, letter-spacing 3px, color: Signal Green
- **Report title:** Display scale, white, centered
- **Subtitle:** 16px, weight 300, white, opacity 0.85
- **Meta strip:** `rgba(255,255,255,0.08)` background, `4px` radius, body scale — inline-flex with 24px gaps

### File Section
The primary structural unit. Each FBDI file in the report occupies one file section. The dark blue header bar makes each file's boundary immediately visible when scanning the page.

- **Container:** `1px solid #D8D7EE` border, `6px` radius, white interior, overflow hidden
- **File Head:** Authority Blue fill, white type; file name at Title scale; metadata row at 12px body
- **Module Pill (in head):** Signal Green fill, white text, `12px` radius, inside the metadata row
- **File Body:** `18px 22px` interior padding

### Change Table
The workhorse. Every field-level change lives in a change table. Column headers establish the data contract; data rows are the deliverable. Numeric columns right-align with tabular-nums. Field names mono. Action columns signal work items.

- **Header row:** Cloud Surface or Mist Divider background; Label scale uppercase text; Midnight
- **Data cells:** Body scale, 8px/9px padding, `1px solid #f0f0f0` bottom border
- **Field cells:** Mono font, Midnight
- **Numeric cells:** Right-aligned, tabular-nums, whitespace nowrap, Slate Annotation
- **Action columns:** 56px fixed width, centered, subtle Authority Blue tint background (`rgba(13,44,113,0.03)`), 1px left border

### Change Block Header
The semantic color system expressed as a section label. Each change type (Added, Removed, Modified, Renamed, Shifted) has its own color-coded h4 with a count badge. The color is the entire semantic signal.

| Change Type | Header Color | Badge Color |
|---|---|---|
| Added | Authority Blue text | Signal Green fill |
| Removed | Remove Red text | Remove Red fill |
| Modified | Caution Amber text | Caution Amber fill |
| Renamed | Slate Annotation text | Slate Annotation fill |
| Shifted | Slate Annotation text | Slate Annotation fill |

- **Typography:** Label scale, uppercase, letter-spacing 1px, weight 700
- **Layout:** flex row, align-items center, 8px gap between text and badge

### Count Badge
- **Shape:** Pill (`12px` radius)
- **Typography:** Label scale (11px), weight 600, white text
- **Padding:** `1px 8px`
- **Color:** Inherits from parent change block (Green / Red / Amber / Slate)

### Module Tag
Small category pill used in the summary table and pending-base list.

- **Default:** Mist Divider background, Midnight text, `10px` radius, `10px` label scale
- **Financials variant:** Signal Green background, white text — the one place Signal Green appears outside change data
- **Padding:** `1px 8px`

### Shift Details (Collapsible, HTML only)
The one interactive component. In HTML mode, SHIFTED fields collapse under a disclosure summary; in PDF mode (print media query), they expand inline as a two-column grid. The same Jinja2 template serves both; behavior diverges via the `print_mode` flag.

- **Summary bar:** Cloud Surface background, Authority Blue text, 12px body, triangle marker (inline SVG or CSS `content: '\25B6'`), `8px 14px` padding
- **Open state:** triangle rotates 90°, `transition: transform 0.15s`
- **Content area:** `1px solid Mist Divider` top border, `10px 14px` padding
- **Shift grid:** Two-column CSS grid, `gap: 4px 24px`; each row flex with space-between; old position in Slate Annotation, arrow in Slate Annotation, new position in Signal Green weight 600

### Action Checkbox
Visual placeholder for consultant action tracking. Not a real input — a styled `<span>` indicating an action column cell.

- **Default (open):** `16px × 16px`, `1.5px` border, `2px` radius, Authority Blue border
- **Dashed (informational only):** Same dimensions, dashed border, 0.4 opacity
- **Delete state:** Remove Red border
- **Warning state:** Caution Amber border

## 6. Do's and Don'ts

### Do:
- **Do** use Authority Blue (`#0D2C71`) for all structural headers, file head bars, table column fills, and section h2 headings. It is the document's identity and hierarchy color.
- **Do** use Signal Green (`#00AB63`) only for: ADDED field counts and labels, Financials module identification, and the cover brand wordmark. Its meaning is "new, positive, Definian."
- **Do** render every Oracle field name, Applaud column name, table prefix, and filename in monospace. No exceptions.
- **Do** use uppercase + letter-spacing ≥0.5px for all table column headers and structural sub-labels. This is the visual contract for the report's dense tabular content.
- **Do** use `1px solid #D8D7EE` for all surface containment borders. Thin, structural, consistent.
- **Do** use Cloud Surface (`#F7F7FB`) for all card-like backgrounds (module cards, shift summaries, collapsible interiors). It is one tonal step above Canvas and provides sufficient differentiation without a border.
- **Do** treat HTML and PDF as separate design targets. The HTML may be richer, more interactive, and more spatially generous. The PDF may be denser and more compact. Optimize each for its medium.
- **Do** use Midnight (`#02072D`) — not `#000` — for all body text. All Definian text uses the palette.

### Don't:
- **Don't** use `border-left` greater than `1px` as a colored accent stripe on cards, panels, or callouts. Side-stripe borders are prohibited. Use a full `1px` border with a background tint instead.
- **Don't** use Signal Green (`#00AB63`) decoratively. Once it appears on a non-addition, non-Financials, non-brand element, its semantic value is gone.
- **Don't** use `box-shadow` anywhere. The system is flat. If you're reaching for a shadow, solve the hierarchy problem with a background color shift or a 1px Mist Divider border.
- **Don't** use gradient text (`background-clip: text` with a gradient). All text is a single solid color from the palette.
- **Don't** introduce a ninth color without documenting its semantic role and adding it to the palette section above. Off-palette colors are a defect.
- **Don't** render the cover gradient background as flat Authority Blue. The gradient is the one place in the document where visual drama is earned by context; a flat-blue cover looks unfinished.
- **Don't** add decoration that carries no semantic meaning — icons for their own sake, divider ornaments, background textures. Data leads; every element must earn its place.
- **Don't** add a new change-type label in body weight or mixed case. All change type labels (Added, Removed, Modified, Renamed, Shifted) are Label scale, uppercase, weight 700. Consistency here is the semantic contract.
