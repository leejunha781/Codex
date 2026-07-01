# Installing GitHub app

- **UUID:** `9f13e200-8de1-4c73-a3ba-a434a3755890`
- **Created:** 2026-06-05 19:42 UTC
- **Updated:** 2026-06-05 19:42 UTC
- **Messages:** 2

> **Conversation Overview**
> 
> The person typed the slash command `/install-github-app` in the Claude chat interface. Claude clarified that this command is specific to Claude Code, the agentic command-line tool, and does not function in the standard chat interface. Claude explained that the command must be run within an active Claude Code terminal session and outlined the prerequisite steps: installing Claude Code via npm (requiring Node.js 18+) and launching it in a project directory before executing the command.
> 
> Claude inferred from context that the person has an existing PLM project with a GitHub Actions CI pipeline configured in a lint→test→smoke sequence, and hypothesized they may be looking to integrate Claude into that repository's workflow for automated code review or PR assistance. Claude offered to walk through the GitHub app setup and how it would fit alongside the existing CI configuration, and asked whether the person wanted current official setup documentation pulled up to verify accuracy. The conversation ended with Claude's clarifying questions; no further response was provided by the person.

---

### 🧑 Human  
*2026-06-05 19:42 UTC*

/install-github-app\

---

### 🤖 Assistant  
*2026-06-05 19:42 UTC*

It looks like you typed a slash command, `/install-github-app`, but I should clarify something: I'm Claude in the web/mobile chat interface, and I don't process slash commands like that. That command is specific to **Claude Code**, the agentic command-line coding tool.

If you're trying to install the Claude GitHub app to use Claude Code with GitHub Actions, here's what you'd want to do:

In **Claude Code** (the terminal tool), running `/install-github-app` walks you through installing the GitHub app and setting up the Claude Code GitHub Actions workflow on a repository. You need to run it from within an actual Claude Code session, not here in chat.

If you don't have Claude Code set up yet, the broad steps are:

1. Install Claude Code (requires Node.js 18+): `npm install -g @anthropic-ai/claude-code`
2. Launch it in your project directory: `claude`
3. Then run `/install-github-app` inside that session.

Given your PLM project with its GitHub Actions CI pipeline (lint→test→smoke), I suspect you're looking to wire up Claude into that repo's workflow for automated code review or PR assistance. Is that what you're after? If so, I can walk you through the GitHub app setup and how it'd slot alongside your existing CI.

Want me to pull up the current official setup docs to make sure the steps are accurate?

---
