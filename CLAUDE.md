# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A **Claude Design** export of a single 1600x900 presentation slide: the *RDAS Suite Map*. Used as the ecosystem overview in the July 15, 2026 VP demo covering the RDO Limited Submissions Dashboard, Macro Automation Modernization, Foundations Database, and Grants Monetary Progress Tracker.

Not a normal HTML build. There is no toolchain in this folder - the runtime (`support.js`) is a pre-compiled bundle from an external repo (`dc-runtime/`), and the slide is authored as a `.dc.html` template that gets interpreted at load time. To preview, just open `RDAS Suite Map.dc.html` in a browser.

## Files that matter

- **`RDAS Suite Map.dc.html`** - the slide template. Fixed 1600x900 canvas with absolute-positioned nodes. Everything below the `<x-dc>` root is the slide DOM.
- **`support.js`** - the dc-runtime bundle. **Do not edit by hand.** Header says "GENERATED from `dc-runtime/src/*.ts` - do not edit. Rebuild with `cd dc-runtime && bun run build`." That runtime lives outside this folder.
- **`assets/th-*.png`** - the four thumbnail images used as banners on the product cards inside the RDAS Core (`th-rdo.png`, `th-macro.png`, `th-foundations.png`, `th-grants.png`). Referenced by path from the template.
- **`assets/umd-division-research.jpg`** - UMD Division of Research mark used in the header and footer.
- **`uploads/`** - the source dashboard HTML files this slide represents (`RDO_Dashboard.html`, `Grants_Monetary_Tracker.html`, `LS_Tracker_Demo.html`). They are **not** linked from the slide DOM - they were uploaded to Claude Design as reference material when the slide was authored. Treat as source-of-truth references, not deliverables you edit here.

## Templating conventions (dc-runtime)

The `.dc.html` file uses custom tags interpreted by `support.js` at runtime:

- **`<x-dc>`** - the root wrapper. Everything inside gets rendered by the runtime.
- **`<helmet>`** - contents get moved into `<head>` (used here for Google Fonts + global styles).
- **`<sc-if value="{{ boolProp }}">...</sc-if>`** - conditional block. Renders children when the expression is truthy. The `hint-placeholder-val="{{ true }}"` attribute is a preview default so the editor shows the block enabled.
- **`<sc-for list="..." item="...">`** - loop (not currently used in this slide, but the runtime supports it).
- **`{{ expression }}`** - Handlebars-like interpolation, evaluated against the component's `renderVals()` output.

The component script sits at the bottom in `<script type="text/x-dc" data-dc-script data-props="...">` and looks like:

```js
class Component extends DCLogic {
  renderVals() {
    return {
      showConnectors: this.props.showConnectors ?? true,
      showImpact: this.props.showImpact ?? true,
    };
  }
}
```

Props are declared via the `data-props` JSON on the same script tag. The two current props (`showConnectors`, `showImpact`) are booleans exposed as toggles in the Claude Design editor UI (`section: "Display"`).

## What renders on the slide

Layout is three columns left-to-right, drawn with absolute positioning:

- **Left column (`left:40px`)** - five data-source chip cards at `top:176, 290, 404, 518, 632` (each 280x92): InfoReady, Sponsored Programs Office, Symplectic Elements, Financial Tracker, Foundation Directories.
- **Center (`left:470px, width:660`, top:174)** - the RDAS Core card containing a 2x2 grid of the four product cards. Each product card has a colored thumbnail banner (from `assets/th-*.png`), title, purpose sentence, and 3 capability chips. The colored 3px strip under each banner is the product's brand accent (red for LS, slate for Macro, navy for Foundations, green for Grants).
- **Right column (`left:1280px`)** - four audience chip cards (VP for Research, RDO Leadership, RDO Staff, Deans + Faculty) at `top:191, 327, 463, 599` (each 280x110).
- **Connector overlay** - a full-canvas `<svg>` with cubic-bezier paths drawn between column endpoints. Wrapped in `<sc-if value="{{ showConnectors }}">` so it can be toggled off. Stroke gradients: `#inbound` (UMD red -> navy) for sources-to-core, `#outbound` (green -> navy) for core-to-audiences. Endpoint dots are hardcoded `<circle>` elements at specific coordinates.
- **Ecosystem aura** - two concentric dashed circles + a radial gradient behind the RDAS Core (`z-index:0`) for the "hub" glow effect.
- **Impact strip** (`top:758`) - dark charcoal band with the "Leadership reporting becomes a byproduct" principle plus three stat tiles (3-5 hrs saved, 5 systems unified, 4 audiences). Also wrapped in `<sc-if value="{{ showImpact }}">`.
- **Footer** (`top:856`) - light gray strip with the data-flow one-liner and UMD mark.

## Design system

Copied from the surrounding suite (RDO dashboards + Grants Tracker) so all deliverables feel like one presentation:

- **Palette**: UMD red `#C41230` primary + `#A10E28` dark; gold `#FFD520` reserved for the header hairline and impact strip labels; UMD navy `#2D6A9F`; green `#059669`; slate `#475569`; text `#1E1E2D`; muted `#64748B`; border `#EAEDF2`; canvas `#FAFBFC`.
- **Typography**: DM Sans (400/500/700/800) via Google Fonts. Headings weight 800 with tight negative letter-spacing.
- **Elevation**: soft shadows `0 4px 14px rgba(15,23,42,.06)` for chips, `0 18px 50px rgba(15,23,42,.1)` for the RDAS Core.
- **Corners**: 14px on chips, 20px on the RDAS Core.

## Common operations

- **Preview**: open `RDAS Suite Map.dc.html` directly in a browser. The template renders through `support.js`.
- **Toggle a display panel**: set `showConnectors` or `showImpact` to `false` in the `data-props` JSON on the component script, or via the Claude Design editor if you re-import.
- **Change copy on a product card, source chip, or audience chip**: edit the specific `<div>` block in the template. All content is inline plain HTML strings, no data indirection.
- **Swap a product thumbnail**: drop the new PNG into `assets/` (keep the same 82px-tall aspect) and update the `<img src="assets/th-*.png">` reference. Thumbnails are ~7-11KB PNGs.
- **Re-export to Claude Design**: this file was generated by Claude Design. Any deeper refactor should happen there and be re-exported. Editing `support.js` locally is not supported.

## Do not

- Do not edit `support.js`. It is generated. Rebuild happens in the `dc-runtime/` source repo (not present in this folder).
- Do not delete the `<x-dc>` wrapper, the `<script data-dc-script>` component block, or `data-props` metadata. Removing any of them breaks the runtime's parsing (`parseDcDocument` in support.js).
- Do not move the file paths for `assets/` or rename `.dc.html` - Claude Design's re-import expects these conventions.
