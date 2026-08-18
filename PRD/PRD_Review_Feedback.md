# Critical Review — Native Windows Excel Financial-Modeling Add-in PRD v1.0

**Reviewer:** Claude
**Date:** 2026-08-18
**Document reviewed:** `Native_Windows_Excel_Financial_Modeling_AddIn_PRD.docx` (v1.0, dated 2026-08-18)

---

## How to use this document

Every item is numbered and independent. Mark each one:

- `[ ] ACCEPT` — fold into PRD v1.1
- `[ ] REJECT` — leave as-is (optionally note why)
- `[ ] DEFER` — real, but not now

Severity: **S1** = would change the product or sink the project · **S2** = meaningful cost or risk · **S3** = cleanup.

---

## Executive summary — the five things I'd change first

1. **This is an excellent engineering spec and a weak PRD.** ~250 normative statements about *how* to build it; almost nothing about *who buys it, why they switch, or what it costs.* (§A1–A4)
2. **There is no answer to "why not Macabacus."** The feature list is a parity list against an incumbent that is already installed and expensed at most target firms. Without a wedge, the roadmap is a 3-year catch-up. (§A1)
3. **Distribution — not engineering — is the thing most likely to kill this.** The users are analysts on locked-down bank/PE laptops who cannot install a signed XLL without IT. The PRD treats signing as a release checkbox, never as the go-to-market problem it actually is. (§A2)
4. **Phase 1 is roughly a year of work and still isn't a sellable product.** It should be cut roughly in half to get something in an analyst's hands. (§B1)
5. **The application-level undo journal is the highest-cost, highest-risk, lowest-differentiation subsystem in the document,** and one of its core requirements ("never overwrite unrelated later changes") is not implementable as written. (§C1)

---

## A. Product and strategy gaps

### A1 · S1 · No competitive thesis, no wedge

**What I see.** §1 says the product is "inspired by DealMaven and Capital IQ Quick Keys." The word *Macabacus* does not appear. Nor does any competitor, price point, or switching-cost analysis. §1.1 lists user segments but no evidence any were interviewed.

**Why it matters.** Nearly every capability in §6 — formatting cycles, AutoColor, smart copy, sensitivities, model check, compare, PowerPoint snapshot, chart styling — already ships in Macabacus, and at many target firms it's already installed and paid for. A feature list that is 90% parity is not a product decision, it's a clone spec. Everything downstream (roadmap order, what to cut, what to polish) is unanswerable without knowing why someone switches.

**Recommendation.** Add a §1.0 "Why we win" with a defensible wedge and let it reorder the roadmap. Candidates the PRD's own constraints already imply:
- **Price** (incumbent is expensive per-seat; is this $99 vs $1,200?)
- **No account, no cloud, no IT ticket** — the local-first posture is a *sales* advantage, not just an engineering one. Lead with it.
- **Speed / lack of bloat** on huge workbooks, if Phase 0 proves it.
- **Configurability** — the profile system is genuinely deeper than incumbents. That's a real edge, currently buried in §8.

State the wedge, then cut or defer every §6 module that isn't in service of it.

`[ ] ACCEPT  [ ] REJECT  [ ] DEFER`

---

### A2 · S1 · Distribution and procurement are unaddressed

**What I see.** §3.5 and §9.2 require code signing, atomic installers, and rollback. §8.3 mentions organization policy files. That's the extent of it.

**Why it matters.** The stated users work at banks, PE firms, and consultancies with managed Windows estates, application allowlisting, and macro/add-in lockdown. In those environments an individual analyst frequently *cannot* install an XLL at all. That means the buyer is IT/InfoSec, not the analyst — which changes the product: you need a security whitepaper, a SOC-style questionnaire response, an MSI with GPO/Intune deployment, an admin-locked policy file, and a per-seat license model IT can audit. It may also mean a deliberate consumer/prosumer beachhead (independent modelers, small funds, corp dev at non-financial companies) before enterprise.

**Recommendation.** Add a §16 "Deployment and adoption model" covering: target install path (per-user, no admin) vs. managed (MSI/Intune/GPO), the security-review artifacts you'll need, and which segment you sell to first. This should be written *before* Phase 1, because it can invalidate architecture choices.

`[ ] ACCEPT  [ ] REJECT  [ ] DEFER`

---

### A3 · S1 · Success measures don't measure product success

**What I see.** §1.5's five measures are: 90% commands keyboard-invokable, 95% byte-for-byte plan determinism, 0 state-restoration failures, 40% faster in moderated testing, 99.9% crash-free.

**Why it matters.** Four of five are engineering quality gates that belong in §14, not product success. None measure activation, retention, fluency, or revenue. A product can hit all five and have zero users. Two specific problems:
- **"95% byte-for-byte equivalent plans"** — determinism is binary. Accepting 5% non-determinism in a product whose entire pitch is determinism is a contradiction. This is a test-suite pass rate wearing a product metric's clothes.
- **"40% faster in moderated testing"** and **"99.9% crash-free across the corpus"** — with telemetry opt-in-only (§3.5, §11), you have no mechanism to measure either at scale. The moderated-study commitment is expensive and unbudgeted.

**Recommendation.** Move all five to §14 release gates. Replace §1.5 with behavioral measures you can actually observe locally and report voluntarily, e.g.: median commands invoked per active hour; % of week-2 users using ≥10 distinct commands; % invoking via shortcut vs. ribbon (a fluency proxy, and the real leading indicator for this product); trial→paid conversion; 90-day retention. Restate the determinism criterion as 100% on a defined test corpus.

`[ ] ACCEPT  [ ] REJECT  [ ] DEFER`

---

### A4 · S2 · No business model, pricing, or licensing design

**What I see.** §3.5: "Licensing may require periodic network access… offline grace behavior is a commercial decision separate from command execution." §15.1 defers licensing grace policy. Nothing else.

**Why it matters.** Licensing design leaks into architecture (device binding, seat management, offline grace window, what happens mid-command when a license lapses). Deferring it to "implementation qualification" means you'll retrofit it, which is where license-check bugs that block a user's workflow come from.

**Recommendation.** Decide at minimum: perpetual vs. subscription, per-seat vs. site, trial mechanics, and offline grace period (30 days is typical). One page in §16.

`[ ] ACCEPT  [ ] REJECT  [ ] DEFER`

---

### A5 · S1 · Nothing addresses adoption, learning, or switching

**What I see.** §4.2 gives command search and visible shortcut labels. §14.2 requires documentation. That's the whole onboarding story.

**Why it matters.** For a keyboard-first tool, *the product is the muscle memory* — a user who never reaches fluency churns, no matter how good the engine is. And every plausible customer is switching from something (Macabacus, Capital IQ, or a personal VBA library) with fingers already trained.

**Recommendation.** Add these to Phase 1, not "later":
- **Import a Macabacus / Capital IQ / custom keybinding map.** Probably the single highest-leverage feature in the whole product for switchers, and it costs almost nothing given the profile system in §8 already exists. Its absence is the most surprising gap in the document.
- Generated, printable/PDF cheat sheet of the *user's active* bindings.
- A first-run 60-second setup: pick a preset (Macabacus-like / CapIQ-like / clean), done.

Note: §7.2 bans "usage-aware shortcut suggestions," which is fine, but don't let that ban bleed into banning ordinary onboarding.

`[ ] ACCEPT  [ ] REJECT  [ ] DEFER`

---

### A6 · S2 · The absolute AI prohibition is over-committed

**What I see.** §7.2 bans "AI features of any kind in the product's core **or roadmap**." §15 elevates it to a governance constraint requiring product-owner approval to change.

**Why it matters.** I think the *underlying* constraints — no network dependency in a core command, no probabilistic behavior in a mutation path — are correct and should be kept exactly as they are. But those constraints are already stated separately and are sufficient. Banning a category by name additionally forecloses things that don't violate either constraint: a fully local model for natural-language *command search* (a discoverability feature, not a mutation), or an optional, clearly-partitioned companion. Writing the ban into governance means revisiting it requires a PRD amendment rather than a normal product decision — in a market where that stance may become a liability within the document's own lifetime.

There's also an internal blur: §7.1 keeps "usage-aware shortcut suggestions" architecturally anticipated (a heuristic recommender) while §7.2 bans "anomaly inference." The line between them isn't drawn anywhere.

**Recommendation.** Replace the categorical ban with the two constraints that actually protect the product: (a) no core command may depend on a network response; (b) no command that mutates a workbook may be non-deterministic. Then state as *positioning*, not governance: "v1 ships no AI features." Keeps the discipline, loses the handcuffs.

`[ ] ACCEPT  [ ] REJECT  [ ] DEFER`

---

### A7 · S3 · "Deterministic" is conflated with "exact"

**What I see.** §2 forbids "probabilistic inference." §6.5 requires compare alignment to "expose confidence" and use "conservative structural signatures"; §13.6 AC-CMP-03 references an "approved confidence threshold."

**Why it matters.** Heuristic alignment with confidence scores is a *deterministic heuristic* — same input, same output — but it is not exact, and a reader implementing §2 literally could build the wrong thing or think §6.5 violates §2.

**Recommendation.** One sentence in §2: "Deterministic means identical inputs yield identical outputs. It does not mean heuristic-free; heuristics are permitted where they are reproducible, documented, and expose their confidence."

`[ ] ACCEPT  [ ] REJECT  [ ] DEFER`

---

## B. Scope and roadmap

### B1 · S1 · Phase 1 is too large to be a first release

**What I see.** Phase 1 (§12) = command registry + Quick Keys + command search + favorites + profiles + formatting cycles + AutoColor + style library + Smart Copy + spacing/transpose + sign/scale + paste/fill + navigation + row/col/sheet ops + undo journal. No dates, no team size, no estimates anywhere in the document.

**Why it matters.** That's plausibly 9–15 months for a small team, and at the end of it you've shipped no auditing or review capability — nothing that isn't table stakes against the incumbent. You'd be a year in before the first real user signal.

**Recommendation.** Split into **Phase 1a (ship it)** and **Phase 1b**:

| Phase 1a — get to real users | Phase 1b |
|---|---|
| Command registry + contract | Command search (fuzzy) |
| Quick Keys + collision detection | Favorites / command bar |
| ~25 formatting + navigation commands | Full style library + Capture/Apply Style |
| AutoColor (selection + sheet) | AutoColor (workbook, w/ preview) |
| Profile JSON + import/export | Smart Copy / spacing / transpose |
| Application-state guard | Sign/scale transforms |
| *No undo journal* (see C1) | Session undo, narrow scope |

Phase 1a is a defensible standalone tool for a keyboard-speed buyer and gets you feedback while the domain engine is still cheap to redirect. Add target dates and headcount to §12 — a roadmap with neither is a wish list.

`[ ] ACCEPT  [ ] REJECT  [ ] DEFER`

---

### B2 · S1 · Cut workbook compare from v1

**What I see.** §6.5 + AC-CMP-01..05: sheet pairing, row/column alignment via anchors and content signatures, confidence exposure, mechanical-shift separation, parsed-token formula diff, 250k-cell pairs in ≤15s.

**Why it matters.** This is the largest and hardest domain subsystem in the document — alignment against structural edits is a genuinely hard research-flavored problem, and getting it 90% right is worse than not shipping it, because a missed diff in a deal model is a career event for the user. Meanwhile Microsoft ships Spreadsheet Compare with Office Professional Plus, and the incumbents all have a version. High cost, high risk, weak differentiation.

**Recommendation.** Defer to post-v1. If you want a cheap 80% now: same-shape diff only (no structural alignment), refuse when shapes differ, ship in a week instead of a quarter.

`[ ] ACCEPT  [ ] REJECT  [ ] DEFER`

---

### B3 · S2 · Cut or shrink Phase 4 wholesale

**What I see.** Phase 4 = chart formatting + PowerPoint snapshot + deterministic command chains + enterprise policy.

**Why it matters.** **Command chains** (§4.4) alone are a substantial subsystem: ordered validation against a captured context, aggregated preview, inter-step conflict detection, single-transaction execution with a non-atomic fallback, versioned argument serialization, import compatibility validation. That's a lot of engine for a power-user long tail — and Excel already has macros for users who want sequencing. **PowerPoint automation** brings a second COM application, a second failure surface, and a second app's version matrix.

**Recommendation.** Drop chains from the roadmap entirely until users ask; the §4.1 command contract makes adding them later cheap, which is the whole point of the framework. Keep chart formatting (it's cheap, it's a real daily pain). Keep PowerPoint snapshot only if user research says the range→deck loop is a top-3 pain — it usually is, but confirm.

`[ ] ACCEPT  [ ] REJECT  [ ] DEFER`

---

### B4 · S2 · Reduce Name Manager and Link Manager to read-only

**What I see.** §6.6/§6.7 include rename-with-atomic-reference-update, link repointing, and link breaking.

**Why it matters.** The *inventory and navigate* half is cheap, immediately useful, and low-risk. The *mutate* half (atomic multi-location reference rewriting across formulas, charts, validation, print areas, connections) is where all the cost and all the corruption risk lives, and it drags in the mandatory-preview and rollback machinery.

**Recommendation.** v1 = inventory, navigate, report broken/unused/conflicting. Defer rename/repoint/break to a later phase with its own hardening budget.

`[ ] ACCEPT  [ ] REJECT  [ ] DEFER`

---

### B5 · S2 · LBO/PE helpers are the wrong abstraction

**What I see.** §6.9 lists ~13 named helpers (MOIC, IRR/XIRR, sources & uses, debt sweep, PIK, fee amortization…), each requiring finance golden tests (AC-LBO-01).

**Why it matters.** These are formula templates. Hard-coding a bank's house convention for a debt sweep into your product means shipping *someone else's* convention to everyone else, plus you now own finance correctness in perpetuity for thirteen formulas that every user will want to tweak.

**Recommendation.** Build **one** thing instead: a user-authored, profile-portable snippet/template insertion mechanism with named parameter slots. Ship 8–10 built-in snippets as content, not code. Cheaper, more valuable, more defensible (users build libraries and won't leave), and it deletes AC-LBO-01's golden-test burden. Keep the circularity switch tooling (§6.9, last three bullets) — that *is* genuine differentiated engineering.

`[ ] ACCEPT  [ ] REJECT  [ ] DEFER`

---

### B6 · S2 · Drop workbook-persisted bookmarks

**What I see.** §6.11 + §8.4: bookmarks persisted in "a versioned, namespaced metadata location" in the workbook, with §8.4 requiring they not alter formulas, values, formatting, or VBA. AC-NAV-02 requires they survive reopen.

**Why it matters.** Writing custom XML parts marks the workbook dirty, changes its bytes and hash, will interact badly with AutoSave/coauthoring (already an unresolved area — see D5), and in a compliance-controlled environment "your add-in modified my model file" is a conversation you don't want to have during a deal. The value is convenience-tier.

**Recommendation.** Store bookmarks in a **local sidecar** keyed by workbook path + document ID. Loses portability across users; loses all the risk. Revisit only on demand.

`[ ] ACCEPT  [ ] REJECT  [ ] DEFER`

---

### B7 · S3 · Delete the VSTO contingency

**What I see.** §3.2 decides Excel-DNA, then hedges with "retain as contingency." Phase 0 (§12) requires "Excel-DNA vs VSTO revalidation." AC-ARCH-02 mandates an Excel-DNA artifact.

**Why it matters.** The decision is right and the alternative is a .NET-Framework-bound legacy model nobody is going to fall back to. Carrying it costs Phase 0 time and forces the domain layer to hedge against a host it will never have.

**Recommendation.** Delete the contingency; keep §3.2's comparison table as recorded rationale. Reclaim the Phase 0 slot for something that actually retires risk — the formula parser or the installer/trust path.

`[ ] ACCEPT  [ ] REJECT  [ ] DEFER`

---

## C. Architecture and implementation risk

### C1 · S1 · The undo journal is over-engineered and partly unimplementable

**What I see.** §5.3 requires a bounded transaction journal that records before-state, **validates that restoring will not overwrite unrelated subsequent edits**, is **encrypted at rest**, expires by policy, survives unclean shutdown (§9.2), and is cleared on uninstall. AC-CMD-06 makes the overwrite check testable.

**Why it matters.** Two separate problems:

1. **"Will not overwrite unrelated later changes" is not implementable as written.** Detecting whether a cell changed after your write requires continuously tracking every change to the workbook from every source — user edits, recalcs, other add-ins, coauthors, VBA — between your write and the undo. Excel gives you `Worksheet_Change` (which misses recalcs and many programmatic paths) and nothing resembling a reliable versioned change log. The honest options are (a) snapshot-and-compare at undo time — expensive, and can't distinguish "user changed it" from "recalc changed it", or (b) don't promise this. The PRD promises it and then makes it an acceptance criterion, which means Phase 1 gets stuck here.

2. **Encryption, retention policy, and crash recovery are gold-plating** on a feature whose realistic scope is "undo the formatting command I just ran."

**Recommendation.** Reduce to: **in-session, in-memory, N=20 receipts, formatting-property changes only, no encryption, cleared on workbook close.** Replace the overwrite-safety guarantee with a cheap, honest precondition check (target still exists, plan hash matches a re-read of the affected properties) that *refuses* on mismatch. State plainly in the UI and docs: "Add-in commands may clear Excel's native undo stack; add-in undo covers the current session only." That's what the incumbents do, users accept it, and it removes what is probably a full quarter of work and the document's largest correctness risk.

`[ ] ACCEPT  [ ] REJECT  [ ] DEFER`

---

### C2 · S2 · Performance targets are asserted, not derived — and one path may be unreachable

**What I see.** §9's table: ≤750ms added startup; ≤100ms on 1,000 cells; ≤500ms on 10,000 cells; Model Check on 100k cells ≤3s; workbook scan 250k cells/20 sheets ≤8s; compare two 250k-cell workbooks ≤15s; ≤250MB incremental working set. §9.1 fails builds on a 15% P95 regression.

**Why it matters.** These read as normative with no stated derivation. Two look questionable before Phase 0 says anything: **≤750ms added startup** for a .NET XLL loading a runtime, a Ribbon, a several-hundred-entry command registry, and profile JSON is tight (and the perceived cost is worse on cold start with AV scanning); **≤15s compare** is aggressive when the COM snapshot of formulas + formats across 40 sheets may dominate the entire budget on its own. Publishing unvalidated numbers as requirements means either the gates get waived (and stop meaning anything) or the team burns weeks chasing an arbitrary figure.

**Recommendation.** Mark all of §9 **"provisional — to be re-baselined at Phase 0 exit against the reference corpus."** Make the Phase 0 harness produce the real numbers, then freeze them. Also define the waiver process §9.1 alludes to (who approves, for how long).

`[ ] ACCEPT  [ ] REJECT  [ ] DEFER`

---

### C3 · S2 · Comparing unopened files implies an undeclared second read path

**What I see.** §6.5 permits comparing "two workbook files selected locally" — i.e. not open in Excel. §3.1's Excel adapter is described purely in terms of COM snapshots.

**Why it matters.** Reading a closed .xlsx without Excel means an OpenXML read path: a second implementation of formula/format extraction, a second set of fidelity bugs, and almost certainly a third-party dependency. None of that is in §3.1 or §3.5. (Per project convention, that dependency needs an explicit decision, not a discovery during Phase 3.)

**Recommendation.** Either (a) open comparison sources invisibly in Excel and delete the "files selected locally" language, or (b) add an OpenXML read path to §3.1 with its own fidelity requirements and dependency approval. (a) is much cheaper. Moot if B2 is accepted.

`[ ] ACCEPT  [ ] REJECT  [ ] DEFER`

---

### C4 · S2 · Formula parser strategy is deferred but is the critical path

**What I see.** §15.1 defers "Excel parser integration, custom parser, or hybrid" to implementation qualification. Phase 0 lists "formula parser" as a spike.

**Why it matters.** Nearly every differentiated capability — Smart Copy, transpose, sign/scale, IFERROR toggle, inspector, model check normalization, compare token diff — sits on this parser. It has to handle locale-dependent separators and function names, R1C1, structured table references, external workbook refs, dynamic arrays and spills, `@` implicit intersection, sheet names requiring quoting, and defined names. It's the single largest source of latent correctness bugs, and AC-FORM-02 already demands a test suite that rejects raw string replacement.

**Recommendation.** Promote it from a §15.1 deferral to the **primary Phase 0 exit gate**, with an explicit coverage matrix (which reference forms are in scope for v1, which explicitly refuse). Declaring a *narrow* v1 coverage set with hard refusals outside it is far better than a broad set with silent corruption.

`[ ] ACCEPT  [ ] REJECT  [ ] DEFER`

---

### C5 · S3 · `Formula2` and dynamic arrays aren't in the compatibility matrix

**What I see.** §9.1 mandates block reads of `Formula2`/`Value2`. §2.1 leaves the minimum build to release management.

**Why it matters.** `Formula2` and dynamic-array/spill semantics only exist on newer builds. If LTSC 2019 is in the support matrix you need a dual path, and spill behavior changes what several commands (transpose, transforms, sensitivities) can safely do.

**Recommendation.** Make the dynamic-array baseline an explicit §2.1 decision, not an implicit one, and note the `Formula`/`Formula2` fallback if pre-DA builds are supported.

`[ ] ACCEPT  [ ] REJECT  [ ] DEFER`

---

## D. Specification defects

### D1 · S2 · The scale-transform naming is ambiguous and the example contradicts the label

**What I see.** §6.2 lists transforms "x 1,000, / 1,000" and then illustrates: "=A1+B1 **scaled by 1,000** becomes =(A1+B1)**/1000**." AC-FORM-03 repeats the division.

**Why it matters.** "Scale by 1,000" means multiply to one reader and "display in thousands" (divide) to another. The document's own example uses the second reading under a label that suggests the first. This is exactly the kind of ambiguity that ships as a bug and generates support tickets on a command that silently changes numbers in a financial model.

**Recommendation.** Name commands by **intent**, not operator: "Convert to thousands (÷1,000)", "Convert from thousands (×1,000)". Fix the §6.2 example and AC-FORM-03 to match.

`[ ] ACCEPT  [ ] REJECT  [ ] DEFER`

---

### D2 · S3 · The formula-to-value prohibition needs an explicit carve-out list

**What I see.** §5.1: "Never replace a formula with its calculated value unless the command explicitly states that it converts." §7.2 removes "formula-to-hardcode snapshots." But §6.7 break-links converts formulas to values, and §6.2 has "paste values only."

**Why it matters.** Not actually contradictory — the "unless" clause covers it — but a reader hitting §7.2's flat prohibition first will read §6.7 as a violation. Ambiguity in a governance-tier rule is worth ten words to remove.

**Recommendation.** In §5.1, name the complete set of commands permitted to convert formulas to values (break links, paste values, and any others), and state that additions to that set require a PRD change.

`[ ] ACCEPT  [ ] REJECT  [ ] DEFER`

---

### D3 · S3 · "Every important command" vs. "at least 90%"

**What I see.** §2: "Every important command has an assignable shortcut." §1.5: "At least 90% of released commands are invokable without opening the ribbon or task pane."

**Why it matters.** Two different promises, and "important" is undefined, so the 90% is unfalsifiable.

**Recommendation.** Pick one. Suggest: "Every command with a fixed parameter set is shortcut-assignable; commands requiring parameter entry open a dialog reachable from the keyboard." That's testable and drops the arbitrary percentage.

`[ ] ACCEPT  [ ] REJECT  [ ] DEFER`

---

### D4 · S2 · Locally-stored Model Check ignores don't survive handoff

**What I see.** §8.4 stores ignore fingerprints locally, as a consequence of §7.2 banning embedded reviewer state.

**Why it matters.** The review workflow in practice is: an associate scans a model, dismisses 40 known-fine findings, then hands the model to a VP — who scans it and sees all 40 again. That's a real friction point created deliberately by the §7.2 ban, and the PRD never acknowledges the tradeoff.

**Recommendation.** Acknowledge it explicitly in §6.4, and permit ignore sets to be **exported/imported as a standalone file** alongside the workbook. Keeps §7.2's "no embedded reviewer state" rule intact while making the workflow survivable.

`[ ] ACCEPT  [ ] REJECT  [ ] DEFER`

---

### D5 · S1 · AutoSave and coauthoring are named once and never resolved

**What I see.** §2.1: commands that "cannot guarantee a safe transaction under those states must refuse or degrade to a documented mode." No such mode is documented anywhere, and **no acceptance criterion in §13 covers AutoSave or coauthoring at all.**

**Why it matters.** AutoSave-on is the default for workbooks on OneDrive/SharePoint, which is where a growing share of target-user files live. Under AutoSave, your mutations are persisted continuously and your undo journal's before-state may already be on disk and replicated to other users. This is a first-order correctness issue for §5.3, not an edge case — and it's currently a one-line hand-wave.

**Recommendation.** Decide and write the behavior: I'd suggest detecting AutoSave and either (a) refusing high-impact commands with a clear message, or (b) offering to pause AutoSave for the transaction and restore it after. Add ACs: `AC-CMD-09` (documented behavior under AutoSave-on for each impact tier) and `AC-CMD-10` (documented behavior under active coauthoring).

`[ ] ACCEPT  [ ] REJECT  [ ] DEFER`

---

### D6 · S3 · AC-PRIV-01 needs a licensing carve-out

**What I see.** §11: "Core execution makes zero outbound network calls." AC-PRIV-01: "zero outbound requests while exercising the complete core command suite." §3.5 permits licensing to require periodic network access.

**Why it matters.** Consistent in intent, but a literal reading of AC-PRIV-01 fails the moment a license heartbeat fires during the test run. Testers will hit this.

**Recommendation.** Amend AC-PRIV-01: "…with licensing and update services disabled or stubbed. Licensing and update traffic is separately enumerated and inspected."

`[ ] ACCEPT  [ ] REJECT  [ ] DEFER`

---

### D7 · S3 · Crash-free and speed measurement have no mechanism

**What I see.** §1.5 requires >99.9% crash-free "across the supported workbook corpus" and 40% faster in moderated testing; §14.2 requires crash-free >99.9% "in pre-release usage." Telemetry is opt-in and content-free (§3.5, §11).

**Why it matters.** With opt-in telemetry and no stated beta program, neither number is measurable. A gate you can't measure is a gate that gets waived.

**Recommendation.** Either define the mechanism (sized beta cohort with opt-in crash reporting; a named moderated-study protocol with N and task list) or convert both to internal corpus-based measures. Related to A3.

`[ ] ACCEPT  [ ] REJECT  [ ] DEFER`

---

## E. Acceptance-criteria coverage gaps

§13 is unusually thorough on the engine and unusually thin on everything the user touches. Missing ACs:

| Area | Gap |
|---|---|
| Quick Keys | No AC for collision detection, reserved-key warnings, multi-stroke timeout, or Escape-cancel (all required by §4.2/§10.1) |
| Command search | No AC at all |
| Favorites / command bar | No AC at all |
| Edit-mode safety | §10.1's "must not steal typing in cell edit mode" — the highest-severity UX failure mode in the product — has no AC |
| Accessibility | §10.3 requirements have no AC |
| Localization | Separator/function-name/date locale testing has no AC despite §14.1 calling for it |
| AutoSave / coauthoring | See D5 |
| Pane behavior | §10.2 (no continuous scanning, focus return, close-doesn't-disable) has no AC |

**Recommendation.** Add a §13.10 "Invocation surfaces and UX" block. At minimum: `AC-KEY-01` (typing in cell edit mode is never intercepted except by commands explicitly declaring edit-mode support), `AC-KEY-02` (binding collisions detected and surfaced at assignment time), `AC-KEY-03` (multi-stroke timeout leaves workbook unchanged), `AC-UX-01` (accessible names + keyboard focus order on all panes), `AC-LOC-01` (golden formula tests pass under comma and semicolon list separators).

`[ ] ACCEPT  [ ] REJECT  [ ] DEFER`

---

## F. Document and process

### F1 · S2 · A .docx cannot be the authoritative baseline for a software product

**What I see.** §15 declares this the single product baseline. The repo contains the .docx and nothing else — no code, no README, no ADRs.

**Why it matters.** No diffs, no line-level review, no blame, no PR workflow, no agent- or tool-readability. §15 requires written change control against a format that can't support it.

**Recommendation.** Convert to Markdown in-repo. Keep the .docx as an export for external readers if needed.

`[ ] ACCEPT  [ ] REJECT  [ ] DEFER`

---

### F2 · S2 · Three documents are welded into one

**What I see.** The document simultaneously serves as PRD (§1, §7, §12), engineering design spec (§3, §4, §5, §9), command catalog (§6), and test plan (§13, §14).

**Why it matters.** This is the root cause of the imbalance flagged throughout §A: the engineering sections are exhaustive because engineering had a place to write, and the product sections are thin because they're competing for the same document. It also means §6 — the part that will change weekly once users show up — is locked behind §15's change-control ceremony.

**Recommendation.** Split into four:
- `PRD.md` — users, jobs, wedge, positioning, success measures, roadmap, governance (§1, §2, §7, §12, §15, plus new §A1/A2/A4 content)
- `ARCHITECTURE.md` — §3, §4, §5, §8, §9, §11
- `COMMANDS.md` — §6, as a living catalog with a lightweight change process
- `ACCEPTANCE.md` — §13, §14

Keep §15's heavyweight change control for the PRD and ARCHITECTURE only.

`[ ] ACCEPT  [ ] REJECT  [ ] DEFER`

---

### F3 · S3 · Phase 0 has no exit criteria that can fail

**What I see.** §12 lists Phase 0 activities (spike, prototype, harness…). §12.1 gives general gates.

**Why it matters.** Activities aren't gates. "Excel-DNA host spike" is done when someone says it's done.

**Recommendation.** Give Phase 0 falsifiable exits, e.g.: add-in loads in ≤X ms measured on the reference machine; parser passes the v1 coverage matrix at 100%; state guard survives 100% of injected-exception cases; a signed installer installs on a clean managed VM without admin rights (this one also tests A2); block-write of 10,000 cells measured and recorded. If any fails, the corresponding §9 number gets re-baselined rather than the gate being waived.

`[ ] ACCEPT  [ ] REJECT  [ ] DEFER`

---

## Appendix — what a cut-down v1 looks like

Applying B1–B6 and C1, the scope that survives:

**In v1**
- Command framework, registry, contract, impact tiers, state guard (§4, §5.1, §5.2)
- Quick Keys, collision detection, keybinding import from incumbents (§4.2 + A5)
- Formatting cycles, AutoColor, style library (§6.1)
- Formula transforms: smart copy, spacing, transpose, sign/scale, IFERROR (§6.2)
- Precedents/dependents + trace pane + formula inspector (§6.3)
- Model Check (§6.4)
- Navigation, rows/columns/sheets, local bookmarks (§6.11, modified per B6)
- Profiles: local JSON, export/import, org policy (§8)
- Session-scoped undo (§5.3, reduced per C1)

**Deferred**
- Workbook/worksheet compare (B2)
- Command chains (B3)
- Name/link *mutation* — inventory only in v1 (B4)
- LBO helpers as code → snippet mechanism instead (B5)
- Sensitivity builder (§6.8) — good feature, no reason it's ahead of user feedback
- Chart formatting, PowerPoint snapshot (B3) — pending user research
- Workbook-persisted bookmarks, encrypted persistent undo journal (B6, C1)

**Rough effect:** cuts the largest and riskiest subsystems (compare alignment, persistent undo with overwrite detection, chain transaction aggregation, atomic reference rewriting) while keeping everything that constitutes the daily-speed value proposition. The §4.1 command contract is what makes deferring all of this cheap — that part of the design is genuinely good and should not be compromised.

---

## What the document does well

Worth preserving explicitly through any revision:

- **The unified command contract (§4.1)** is the best idea in the document. Identity / CanExecute / Plan / Preview / Execute / Undo with an immutable approved plan is exactly right, and it's what makes aggressive scope cuts safe.
- **Property-scoped mutation (§2, §5.1)** — the discipline that number formatting never touches values, and that reference transforms parse rather than string-replace, is the difference between a tool analysts trust with a live deal model and one they uninstall.
- **AC-ARCH-03/04** (domain engine has zero COM references; instrumentation fails tests when proxies cross threads) are excellent, enforceable architectural constraints. Most specs assert layering; this one tests it.
- **The offline/local-first posture (§11)** is coherent, testable, and — per A1 — probably the commercial wedge, not just an engineering stance.
- **Explainable findings over opaque scores (§2, §6.4, AC-CHECK-06)** is the right call and correctly resisted in the removed-features list.
