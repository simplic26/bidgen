# Cost Validator Frontend MVP Design

## Goal

Build a frontend-only demonstration MVP for the cost estimate validation system described in the PRD and wireframe. The app should let a user experience the full review flow from file upload through evidence-backed validation results, without requiring real Excel or PDF parsing in this version.

## Scope

This MVP implements the user-facing validation workflow:

- Select procurement type: construction, service, or goods.
- Upload one Excel file and multiple PDF criteria documents.
- Show document recognition and mapping results.
- Show an automatic validation progress screen.
- Show a validation dashboard with summary counts, filters, result list, and evidence panel.
- Show a detailed evidence view for a selected validation result.

The app does not implement real Excel parsing, PDF parsing, backend APIs, authentication, persistent storage, or export features.

## Product Behavior

Uploaded file names are reflected in the UI so the demo feels connected to the user's action. Validation data comes from seeded sample results modeled after the PRD's `ValidationResult` structure.

The primary scenario is construction cost validation. Service and goods remain selectable, and the UI explains that they use common arithmetic checks in this demo.

The result dashboard defaults to surfacing errors and needs-review items first. Selecting a result updates the evidence panel immediately. Status and validation-type filters change the visible result list.

## Screens

### Upload

The first screen is a working tool surface, not a marketing page. It includes procurement type selection, an Excel upload area, a PDF upload area, selected file summaries, and an automatic validation start button.

### Document Recognition

This screen summarizes recognized Excel sheets and extracted PDF criteria. It includes at least one needs-review mapping item to match the wireframe scenario.

### Validation Progress

This screen shows a staged validation process with progress percentage, completed steps, pending steps, and the current cell or item being checked.

### Result Dashboard

The dashboard includes:

- File and criteria summary.
- Status count cards for overall result, errors, needs review, warnings, and normal items.
- Left-side filters for status, validation type, and sheet.
- Center result list ordered by review priority.
- Right-side evidence panel for the selected result.

### Detail Evidence

The detail view includes:

- Excel source information: sheet, cell, formula, referenced cells, input value.
- PDF criteria information: document title, page, applied condition, quote.
- Calculation process with base amount, rate, calculated expected value, input value, and difference.
- Judgment reason, evidence bullets, recommended action, user memo area, and previous/next result controls.

## Data Model

Use frontend sample data shaped to match the PRD:

- `resultId`
- `status`
- `severity`
- `validationType`
- `procurementType`
- `item`
- `excel`
- `expected`
- `difference`
- `summary`
- `reason`
- `evidence`
- `recommendedAction`

The UI should be structured so these sample objects can later be replaced by FastAPI responses with minimal changes.

## Design Direction

The visual style should feel like a calm contract review and audit tool. It should be dense enough for repeated professional use, but not visually heavy.

Use a restrained palette with clear status colors, compact typography, stable panel dimensions, and strong scan paths. Avoid a landing-page hero, oversized decorative cards, and marketing-style layout. The first screen should immediately look usable.

## Component Boundaries

The implementation should separate:

- App step orchestration.
- Sample validation data.
- Shared UI primitives such as buttons, status badges, summary cards, and upload panels.
- Workflow screens.
- Result list and filtering logic.
- Evidence panel and detail view.

`App` should compose the workflow and state, not hold all markup directly.

## Interaction Requirements

- File upload controls update selected file names and metadata.
- Automatic validation advances through recognition, progress, and dashboard screens.
- Result row selection updates the evidence panel.
- Status and validation-type filters update the result list.
- Detail evidence view can navigate back to dashboard.
- Previous and next error item buttons switch selected results.

## Verification

The finished app should be run locally and checked in a browser at desktop and mobile widths. Verification should cover:

- Upload flow works with selected local files.
- Progress screen advances into results.
- Filters update the result list.
- Result selection changes the evidence panel.
- Detail view displays the selected result's Excel, PDF, calculation, and recommendation data.
- Text does not overlap or overflow on mobile.
- Build or type checks pass for the selected toolchain.

## Intentional Deferrals

- Real Excel parsing with `openpyxl`.
- Real PDF extraction with PyMuPDF or pdfplumber.
- FastAPI backend.
- User correction of extracted PDF values.
- Spreadsheet-like full Excel preview.
- PDF page image rendering.
- Exporting results to Excel or CSV.
- Persistence, authentication, permissions, and audit logs.
