# copilot-agent-eval — Agent Evaluation Skill

A reusable AI skill that helps a development / system-analyst team **evaluate and improve custom Copilot agents** (Q&A bots, technical summarizers, and similar) using a lightweight, human-in-the-loop process. Give it an agent name and a one-line description, and it generates a scoring checklist and sample test items matched to that agent's domain. Ask for the full package and it delivers an evaluation plan, a process diagram, and a fully reconciled Excel tracking workbook.

## What this skill does

| You ask | You get |
|---|---|
| "Create an evaluation checklist for my *Invoice Q&A Bot*" | A 5-point Pass/Partial/Fail scoring checklist rewritten in your agent's domain language, plus 8–12 ready-to-use sample test items (including trap questions) |
| "Create the full evaluation plan / package" | `copilot-agent-evaluation-plan.md` (framework, phased plan, metrics, troubleshooting playbook) |
| "Give me the process diagram" | `copilot-agent-evaluation-flow.svg` — 6-step flow your whole team can read at a glance |
| "Build / update the tracking workbook" | `copilot-agent-evaluation-tracker.xlsx` — 4 sheets (Agent Tracker, Iteration Log, Scoring Checklist, Test Item Log) where all pass rates are formulas fed from item-level scores, so the numbers always reconcile |

## The method in one paragraph

Build a small frozen **golden dataset** (10–20 documents, with trap questions) → run a **baseline** → **score by hand** with a 3-point Pass/Partial/Fail scale → **map failures** to prompt weaknesses and fix only the top 2–3 → change **one thing per round** and re-run until ≥90% pass → **pilot and monitor**, feeding every real failure back into the test set. Q&A agents and summary agents are always scored with different criteria. Automated LLM-as-a-judge scoring comes last, never first.

## File structure

```
copilot-agent-eval/
├── SKILL.md                                  # The skill prompt (role, workflow, guardrails)
├── README.md                                 # This file
├── assets/
│   └── copilot-agent-evaluation-flow.svg     # Reference process diagram (browser / draw.io)
└── scripts/
    └── build_tracker.py                      # Generates the 4-sheet Excel tracker (openpyxl)
```

`scripts/build_tracker.py` is self-contained: `pip install openpyxl`, then `python build_tracker.py [output.xlsx]`. Edit the `rows`, `log`, and `items` data blocks to describe your own agents — formatting, dropdowns, conditional colors, and cross-sheet formulas are already wired. It prints a pass-rate self-check after saving.

---

## Installation

### Claude Code (native support — recommended)

Skills are folders containing a `SKILL.md`. Copy this **whole folder** to either location:

| Scope | Location | Effect |
|---|---|---|
| One project | `<project>/.claude/skills/copilot-agent-eval/` | Everyone who clones the repo gets it |
| All projects on your machine | `~/.claude/skills/copilot-agent-eval/` (Windows: `C:\Users\<you>\.claude\skills\...`) | Personal, machine-wide |

Then in any Claude Code session: type **`/copilot-agent-eval`**, or just ask naturally ("evaluate my new HR bot") — the `description:` frontmatter lets Claude trigger it automatically.

```bash
# quickest install from this repo
git clone <this-repo-url>
cp -r copilot-agent-eval ~/.claude/skills/
```

### Claude.ai / Claude Desktop (Projects)

Create a Project and paste the body of `SKILL.md` (everything below the `---` frontmatter) into **Project instructions**. Upload the SVG and `build_tracker.py` as project files so Claude can reference them.

### OpenAI Codex CLI

Codex reads instructions from `AGENTS.md` files (global and per-repo):

1. **Per-repo:** copy this folder into the repo (e.g. `skills/copilot-agent-eval/`), then add to the repo-root `AGENTS.md`:
   ```markdown
   ## Skill: Copilot agent evaluation
   When asked to evaluate, score, or improve a custom Copilot/AI agent,
   follow the instructions in skills/copilot-agent-eval/SKILL.md exactly.
   Use skills/copilot-agent-eval/scripts/build_tracker.py to generate the
   Excel tracker instead of writing spreadsheet code from scratch.
   ```
2. **Global:** append the same pointer (or the full SKILL.md body) to `~/.codex/AGENTS.md`.

### GitHub Copilot (VS Code / github.com)

Two supported mechanisms — use both for best results:

1. **Repository custom instructions** — create `.github/copilot-instructions.md` in the repo and add a pointer:
   ```markdown
   When the user asks to evaluate, score, or improve a custom Copilot/AI agent,
   follow skills/copilot-agent-eval/SKILL.md in this repository.
   ```
   Copilot Chat automatically includes this file as context for every chat in the repo.
2. **Prompt file (reusable slash command)** — create `.github/prompts/agent-eval.prompt.md` (VS Code: enable prompt files in settings) containing the full body of `SKILL.md`. Team members then run it in Copilot Chat with **`/agent-eval`** plus their agent name:
   ```
   /agent-eval Invoice Q&A Bot — answers supplier invoice questions from the AP knowledge base
   ```

> Note: Copilot Chat cannot execute Python itself. Team members run `scripts/build_tracker.py` in their own terminal after Copilot edits the data blocks.

### Other assistants (Cursor, Windsurf, Gemini CLI, etc.)

Any assistant that supports project rules/instructions files works the same way: keep this folder in the repo and point the assistant's rules file (`.cursor/rules`, `.windsurfrules`, `GEMINI.md`, ...) at `skills/copilot-agent-eval/SKILL.md`. As a last resort, paste the SKILL.md body as the first message of a new conversation.

---

## Typical team workflow

1. **Install** the skill (above) and open your AI assistant in the repo.
2. **Register your agent:** "My agent is *Warehouse SOP Summarizer* — it condenses SOP documents into 1-page briefs. Create my scoring checklist and sample test items."
3. **Review the generated checklist** with a domain expert; adjust wording, keep the Pass/Partial/Fail scale.
4. **Add the agent to the tracker:** ask the assistant to append it to `build_tracker.py` and rebuild the workbook.
5. **Run the phases** (0–4) recorded in the workbook: score items in Sheet 4, log each fix round in Sheet 2 — Sheets 1–2 percentages update automatically.
6. **Go live** when you hit ≥90% pass on accuracy checks, and keep feeding pilot failures back into Sheet 4.

## Quality gates baked into the skill

- Any hallucinated fact = automatic **Fail** (target hallucination rate <2%)
- Trap questions must get an honest "I don't know" (target 100%)
- Numbers/units/codes in summaries: zero-tolerance transcription accuracy
- One change per iteration round, full re-run, keep only non-regressing improvements
- No automated LLM-judge scoring until it agrees with human scoring ≥85% of the time
