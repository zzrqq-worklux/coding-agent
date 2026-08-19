# AGENTS.md

This workspace is a **coding-dedicated agent**. It bundles a curated set of coding skills (installed from GitHub) chained into a single workflow, plus a default agent (`coder`) that runs that workflow.

## How to use

- Open opencode in this directory (or a subdirectory). The `coder` agent is the default.
- Ask for any coding task: feature, bugfix, refactor, review, security audit, commit. The agent routes the work through the pipeline below automatically.
- Commands are not required — the agent detects which skill applies from your request.

## The chained workflow

| # | Stage | Skill(s) | When it runs |
|---|-------|----------|--------------|
| 1 | Clarify | `brainstorming` | Vague requirements, "let's build X", design questions |
| 2 | Plan | `writing-plans` | Any non-trivial task, after clarifying |
| 3 | Isolate | `using-git-worktrees` | Multi-file/multi-task work in a git repo |
| 4 | Build | `executing-plans` + `test-driven-development` | Implementing the plan; TDD for every feature/bugfix |
| 5 | Debug | `systematic-debugging` | Any bug, error, or unexpected behavior |
| 6 | Review | `code-review-checklist`, `requesting-code-review`, `receiving-code-review` | After implementation, before commit |
| 7 | Security | `owasp-security` | Auth, crypto, access control, APIs, k8s, AI/LLM features |
| 8 | Quality | `clean-code`, `lint-and-validate` | While writing code and after every modification |
| 9 | Verify | `verification-before-completion` | Before claiming anything is done |
| 10 | Finish | `commit-smart`, `finishing-a-development-branch` | Before commit and at branch wrap-up |

Folding back: verification failure → debug; review/security findings → rebuild. Never claim done without verification; never commit unless asked.

## Installed skills

All skills live in `.opencode/skills/` and load via the skill tool.

**From [obra/superpowers](https://github.com/obra/superpowers)** (canonical workflow skills):
`brainstorming`, `writing-plans`, `executing-plans`, `test-driven-development`, `systematic-debugging`, `requesting-code-review`, `receiving-code-review`, `verification-before-completion`, `using-git-worktrees`, `finishing-a-development-branch`

**From [davila7/claude-code-templates](https://github.com/davila7/claude-code-templates)**:
`owasp-security`, `code-review-checklist`, `commit-smart`, `lint-and-validate`, `clean-code`

## Skill naming note

The superpowers skills were written for Claude Code and reference each other as `superpowers:name`. In opencode the skills are loaded by bare name — treat `superpowers:test-driven-development` as the skill `test-driven-development`.

## Config

- `opencode.json` — project config; sets `default_agent: coder`.
- `.opencode/agent/coder.md` — the coding agent prompt that defines the chained pipeline.
- Skills may be removed/added freely; update the pipeline in `coder.md` and this file to match.