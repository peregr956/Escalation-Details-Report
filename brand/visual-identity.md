# Visual Identity Specification

Use this guide alongside the narrative brand guidelines to ensure every touch-point looks, feels, and performs like Critical Start.

## 1. Logo System

### 1.1 Primary Logo
- File: `assets/critical-start-logo.svg`
- Orientation: Horizontal lockup with Critical Start logomark and wordmark.
- Usage: Default mark for digital, product, and print applications with ample contrast.

### 1.2 Alternate Marks
| Variant | Use Case | File | Notes |
| --- | --- | --- | --- |
| Reversed (white) | Dark or gradient backgrounds | Provide via DAM export | Maintain same clear space rules. |
| Monochrome | One-color production needs | Request from Brand Studio | Use only when full-color reproduction is not feasible. |

### 1.3 Clear Space & Minimum Size
- **Clear Space:** Maintain a buffer equal to the height of the letter "C" on every side. No text, imagery, or UI controls may encroach on this area.
- **Minimum Size (Screen):** 80 px width.
- **Minimum Size (Print):** 1 inch width.

### 1.4 Incorrect Usage
- Do not recolor the logo outside the approved palette or gradients.
- Do not stretch, skew, rotate, or add drop shadows/glows.
- Do not place on busy photography without an approved color field.
- Do not lock up with unapproved taglines or partner marks without Brand approval.

## 2. Color Palette

| Role | Name | HEX | RGB | CMYK | Usage Notes |
| --- | --- | --- | --- | --- | --- |
| Primary 1 | Critical Start Blue | `#009CDE` | `0, 156, 222` | `79, 12, 0, 0` | Hero copy, buttons, key lines. |
| Primary 2 | Deep Navy | `#004C97` | `0, 76, 151` | `100, 75, 8, 12` | Backgrounds, CTAs, gradients. |
| Primary 3 | Charcoal | `#343741` | `52, 55, 65` | `67, 54, 43, 49` | Neutral text, panels. |
| Secondary 1 | Violet | `#702F8A` | `112, 47, 138` | `73, 100, 0, 15` | Highlights, infographics. |
| Secondary 2 | Red | `#EF3340` | `239, 51, 64` | `0, 90, 70, 0` | Alerts, critical states. |
| Secondary 3 | Orange | `#FF6A14` | `255, 106, 20` | `0, 71, 100, 0` | Accent CTAs, icons. |
| Gradient | Blue Sweep | `#009CDE → #004C97` | n/a | n/a | Use when photography/patterns clutter small spaces; default small-space background. |

> Secondary colors may be used for CTA buttons, iconography, or chart emphasis. Keep overall contrast AA+ compliant.

## 3. Typography

| Role | Typeface | Weights | Usage |
| --- | --- | --- | --- |
| Primary | Roboto | Regular, Medium, Bold | Headlines, body copy, UI labels. |
| Fallback Headline | Arial Black | Bold | When Roboto unavailable (PowerPoint, internal docs). |
| Fallback Body | Arial Narrow | Regular, Bold | Long-form copy when Roboto unavailable. |

### 3.1 Typesetting Rules
- Favor sentence case for readability.
- Maintain comfortable line heights (1.3–1.5) for body text.
- Avoid tracking below 0 for digital surfaces.
- Ensure minimum 16 px body size on web and 10 pt in print.

## 4. Imagery & Illustration
- Use photography sparingly in small containers; default to the blue gradient when space is limited.
- When photography is used, prefer authentic, candid shots showing real collaboration.
- Illustration and icon styles should be outlined with a medium stroke to align with the interface aesthetic.

## 5. Iconography
- Employ outlined icons with medium stroke weight; avoid filled shapes unless depicting critical states.
- Color icons with Critical Start Blue or accent colors for emphasis.
- In product experiences, rely on Font Awesome 7 React (fa7) icons for consistency and easy implementation.

## 6. Data Visualization
- Map positive states to Critical Start Blue/Deep Navy, warnings to Orange, critical alerts to Red.
- Use Violet for secondary comparisons and benchmarks.
- Maintain 4.5:1 contrast for all chart labels and data points against their backgrounds.
- Limit pies to 5 slices; prefer line + stacked bar combinations that mirror the MDR dashboard.

## 7. Layout & Components

### 7.1 Grid & Spacing
- Anchor desktop layouts to a 12-column grid with 24 px gutters; tablet/mobile may collapse to 8/4 columns with proportional gutters.
- Maintain generous white space around hero KPIs to reinforce premium positioning.

### 7.2 Modules
| Module | Description | Notes |
| --- | --- | --- |
| Hero | Top-level summary with KPI stack and tier badge. | Use gradient background when photography is absent. |
| Metric Card | Single KPI tile with icon + delta. | Border radius 8 px, 16 px internal padding. |
| Story Panel | Text + illustration pairing for narrative sections. | Keep copy within 550 px width for readability. |

## 8. Accessibility & Compliance
- Follow WCAG 2.1 AA minimum contrast; test gradient overlays to ensure text legibility.
- Provide text alternatives for every decorative or informative image.
- Spell out acronyms on first reference inside graphics and captions, matching the narrative brand rule.

## 9. Asset Inventory
| File | Path | Format | Notes |
| --- | --- | --- | --- |
| Primary Logo | `assets/critical-start-logo.svg` | `.svg` | Use for digital and scalable needs. |
| Gradient Background | (export via design system) | `.png` / `.svg` | Blue sweep default background. |
| Typography Package | Request from Brand | `.ttf` / `.otf` | Includes Roboto weights + Arial guidance. |

## 10. Revision History
| Date | Version | Author | Summary |
| --- | --- | --- | --- |
| 2025-11-17 | v1.0 | AI (ChatGPT) | Converted brand visual guidelines to Markdown. |
