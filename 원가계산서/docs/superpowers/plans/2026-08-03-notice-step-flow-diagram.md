# Notice STEP 01~10 Flow Diagram Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the current top SVG and duplicated lower card list with one arrow-connected STEP 01~10 notice-generation diagram in the original diagram position.

**Architecture:** Keep the report as one standalone HTML file. Use semantic HTML cards and CSS Grid for the desktop snake layout, then switch to a single vertical flow on mobile. Preserve the existing status tokens and all downstream detail/checklist content.

**Tech Stack:** Standalone HTML, CSS Grid, CSS pseudo-elements, existing JavaScript checklist logic, Browser visual QA.

## Global Constraints

- STEP copy and status values remain unchanged.
- Desktop order is 01→05, down to 06, then 06→10 right-to-left.
- Mobile order is 01→10 vertically with downward arrows.
- The detail table keeps 10 rows and the improvement checklist keeps 38 items.
- Do not modify the user's unrelated HWPX or React working files.

---

### Task 1: Replace the duplicated flow UI

**Files:**
- Modify: `docs/원가계산서-검증-현황과-개선안.html`

**Interfaces:**
- Consumes: existing `.status-tag`, status color variables, `#diagram` navigation anchor.
- Produces: `.step-flow`, `.step-card`, `.step-arrow` and `.step-turn` layout classes.

- [ ] **Step 1: Run a failing structural assertion**

Assert that the current document does not yet contain `class="step-flow"`, still contains `id="user-flow"`, and contains the STEP cards outside `#diagram`.

- [ ] **Step 2: Verify the assertion fails for the intended reason**

Run a PowerShell HTML assertion and confirm failure reports the missing `step-flow` diagram.

- [ ] **Step 3: Implement the minimal diagram**

Remove the visible and hidden SVG figures, move the STEP 01~10 content into `#diagram`, add arrow connector elements, and remove the lower duplicate heading/cards. Add the desktop snake and mobile vertical CSS rules.

- [ ] **Step 4: Run the structural assertion again**

Confirm one `step-flow`, ten `step-card` items, ten detail rows, and 38 checklist items. Confirm the old `user-flow`, `dg`, and `notice-flow` diagram markup are absent.

- [ ] **Step 5: Commit the HTML change**

Stage only the report HTML and commit with `docs: simplify notice user flow diagram`.

### Task 2: Rendered browser verification

**Files:**
- Verify: `docs/원가계산서-검증-현황과-개선안.html`

**Interfaces:**
- Consumes: the completed `#diagram .step-flow` markup.
- Produces: desktop and mobile screenshot evidence outside the repository.

- [ ] **Step 1: Open the local report in the Browser runtime**

Serve the repository on `127.0.0.1:8765` and navigate directly to `#diagram`.

- [ ] **Step 2: Verify desktop behavior**

At 1440×900 confirm all ten cards are visible, the visual order is 01→05 then 06→10 in reverse columns, arrows connect every transition, and the page has no horizontal overflow.

- [ ] **Step 3: Verify mobile behavior**

At 390×844 confirm cards are one column in STEP order with downward connectors and no document-level horizontal overflow.

- [ ] **Step 4: Run regression checks**

Run `npm.cmd test`, `git diff --check`, and final HTML counts for step cards, detail rows, and checklist items.
