# Shared-chat replay: OfficeCLI feasibility

Source: https://chatgpt.com/s/cx_6ae1bcff28c081919e22559d0f9c375d
Browser title verified: 擬定第四章結果與討論雛形.
All generated artifacts and adapted code belong to this checkout. Old workspaces
and research repositories are read-only sources. Existing uncommitted work is preserved.

## Required operations and evidence

1. Read presentation, research results and source manuscript; map the conversation
   to actual local inputs, preserve original files and record SHA256.
2. Rebuild chapter 4 from the v6 source using OfficeCLI authoring, not a copied
   final manuscript. Include ANN accuracy and step-direction checks, validation
   rho selection, KCI, prediction results and limitations.
3. Replay the intermediate revisions: merge ANN Fig. 9/10; untitled figures;
   No-Gain plus four rho curves for selection; per-case RMSE tables with exact
   minima; baseline colors; signed percentages; four split validation tables.
4. Replay the final gain-pair reorganization, two reference-distribution figures,
   detailed four-pair discussion; retain No-Gain, Gain=0.05, Gain+SS=0.01 and
   Difficult-GradNorm rho=0.01 in subsequent comparisons. Four rho candidates
   remain in the preceding parameter-selection section.
5. Check native EndNote payload preservation, equations and unrelated sections;
   create captions and cross-references through native Word; selectively update
   existing directories; verify both ordinary and complex-script 12 pt sizes.
6. Recalculate values from raw CSV, inspect final Word PDF and all page images,
   record unsupported operations and failures honestly. Keep the known original
   EqJ_322 broken reference separate from replay regressions.
7. Repair demonstrated local tool/skill/configuration problems; document actual
   OfficeCLI-only vs native-adapter boundaries. Record memory guidance locally.
8. Inspect repository publication settings and complete an authorized code-only
   commit/push only after reviewing staged content and upstream identity. Never
   upload ignored manuscript or research materials by implication.

Intermediate reference DOCX files are specifications for content and formatting;
they are not counted as reproduced outputs. Every reproduced stage must show
actual OfficeCLI mutations and independent readback. Python is allowed for
scientific plotting, native Word adapters and QA and is reported separately.

## Current status

- Source chat read in the browser; 90 original source files hash-verified unchanged.
- All nine v7-v15 stages rebuilt from v6 by OfficeCLI and independently checked.
- Native captions/references, every new table's font sizes, numeric/color/bold
  checks, EndNote/math preservation and the final 61-page render review passed.
- Ten regenerated scientific figures inserted into the delivery DOCX by OfficeCLI.
- Project MCP startup/call tests, 23 OfficeCLI integration checks and 58 EndNote
  regression tests passed. Detailed limits and the inherited EqJ_322 defect are
  documented in officecli-chat-replay-results.md.
- Code-only publication is prepared on codex/officecli-chat-replay-20261028;
  publication outcome is recorded separately with the final commit identity.
