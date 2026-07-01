#!/usr/bin/env python3
"""Convert a Claude data export (raw JSON) into readable Markdown.

Reads the raw export files under ``claude-data/raw`` and writes:

- ``claude-data/conversations/*.md`` — one transcript per conversation
- ``claude-data/projects/*.md``      — one page per project (incl. docs)
- ``claude-data/memories.md``        — the account memory
- ``claude-data/README.md``          — an index of everything

Usage:
    python3 claude-data/scripts/export_to_markdown.py
"""

import json
import re
from datetime import datetime, timezone
from pathlib import Path

# claude-data/ (parent of scripts/)
ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "raw"


def load_json(path: Path):
    with path.open(encoding="utf-8") as fh:
        return json.load(fh)


def fmt_time(value: str) -> str:
    """Format an ISO-8601 timestamp for display; fall back to the raw string."""
    if not value:
        return ""
    try:
        cleaned = value.replace("Z", "+00:00")
        return datetime.fromisoformat(cleaned).strftime("%Y-%m-%d %H:%M UTC")
    except ValueError:
        return value


def slugify(text: str, fallback: str) -> str:
    """ASCII-only slug. Non-ASCII (e.g. Korean) titles fall back to the
    provided identifier, since Korean filenames fail to open on some devices."""
    text = (text or "").strip().lower()
    # Drop anything that is not ASCII alphanumeric, whitespace, or hyphen.
    text = re.sub(r"[^a-z0-9\s-]", "", text)
    text = re.sub(r"[\s_-]+", "-", text).strip("-")
    return text or fallback


def render_block(block: dict) -> str:
    """Render a single content block to Markdown."""
    btype = block.get("type")

    if btype == "text":
        return block.get("text", "").strip()

    if btype == "thinking":
        thinking = (block.get("thinking") or "").strip()
        if not thinking:
            return ""
        return "<details>\n<summary>💭 Thinking</summary>\n\n" + thinking + "\n\n</details>"

    if btype == "tool_use":
        name = block.get("name", "tool")
        tool_input = block.get("input")
        try:
            rendered = json.dumps(tool_input, ensure_ascii=False, indent=2)
        except (TypeError, ValueError):
            rendered = str(tool_input)
        return f"**🔧 Tool call — `{name}`**\n\n```json\n{rendered}\n```"

    if btype == "tool_result":
        name = block.get("name", "tool")
        content = block.get("content")
        parts = []
        if isinstance(content, list):
            for item in content:
                if isinstance(item, dict) and item.get("type") == "text":
                    parts.append(item.get("text", ""))
                else:
                    parts.append(json.dumps(item, ensure_ascii=False))
        elif content is not None:
            parts.append(str(content))
        body = "\n".join(p for p in parts if p).strip()
        if len(body) > 4000:
            body = body[:4000] + "\n… (truncated)"
        prefix = "⚠️ Tool error" if block.get("is_error") else "📄 Tool result"
        if not body:
            return f"**{prefix} — `{name}`** *(empty)*"
        return f"**{prefix} — `{name}`**\n\n```\n{body}\n```"

    # Unknown block type: keep raw JSON so nothing is lost.
    return f"```json\n{json.dumps(block, ensure_ascii=False, indent=2)}\n```"


def render_message(msg: dict) -> str:
    sender = msg.get("sender", "unknown")
    label = {"human": "🧑 Human", "assistant": "🤖 Assistant"}.get(sender, sender)
    when = fmt_time(msg.get("created_at", ""))

    pieces = []
    blocks = msg.get("content") or []
    for block in blocks:
        rendered = render_block(block)
        if rendered:
            pieces.append(rendered)

    if not pieces:
        text = (msg.get("text") or "").strip()
        if text:
            pieces.append(text)

    header = f"### {label}"
    if when:
        header += f"  \n*{when}*"

    body = "\n\n".join(pieces) if pieces else "*(no content)*"
    return f"{header}\n\n{body}"


def write_conversations(conversations: list) -> list:
    out_dir = ROOT / "conversations"
    out_dir.mkdir(parents=True, exist_ok=True)

    index = []
    seen = {}
    # Oldest first for a stable, chronological archive.
    ordered = sorted(conversations, key=lambda c: c.get("created_at", ""))

    for conv in ordered:
        name = conv.get("name") or "untitled"
        created = conv.get("created_at", "")
        base = f"{created[:10]}-{slugify(name, conv.get('uuid', 'conversation')[:8])}"
        seen[base] = seen.get(base, 0) + 1
        if seen[base] > 1:
            base = f"{base}-{seen[base]}"
        filename = f"{base}.md"

        lines = [
            f"# {name}",
            "",
            f"- **UUID:** `{conv.get('uuid', '')}`",
            f"- **Created:** {fmt_time(created)}",
            f"- **Updated:** {fmt_time(conv.get('updated_at', ''))}",
            f"- **Messages:** {len(conv.get('chat_messages', []))}",
        ]
        summary = (conv.get("summary") or "").strip()
        if summary:
            lines += ["", "> " + summary.replace("\n", "\n> ")]
        lines += ["", "---", ""]

        for msg in conv.get("chat_messages", []):
            lines.append(render_message(msg))
            lines.append("\n---\n")

        (out_dir / filename).write_text("\n".join(lines), encoding="utf-8")
        index.append((name, created, len(conv.get("chat_messages", [])), f"conversations/{filename}"))

    return index


def write_projects(project_files: list) -> list:
    out_dir = ROOT / "projects"
    out_dir.mkdir(parents=True, exist_ok=True)

    index = []
    for path in project_files:
        project = load_json(path)
        name = project.get("name") or "untitled-project"
        filename = f"{slugify(name, project.get('uuid', 'project')[:8])}.md"

        lines = [
            f"# {name}",
            "",
            f"- **UUID:** `{project.get('uuid', '')}`",
            f"- **Created:** {fmt_time(project.get('created_at', ''))}",
            f"- **Updated:** {fmt_time(project.get('updated_at', ''))}",
            f"- **Private:** {project.get('is_private')}",
        ]
        description = (project.get("description") or "").strip()
        if description:
            lines += ["", "## Description", "", description]

        template = (project.get("prompt_template") or "").strip()
        if template:
            lines += ["", "## Prompt template", "", template]

        docs = project.get("docs") or []
        if docs:
            lines += ["", f"## Documents ({len(docs)})", ""]
            for doc in docs:
                lines.append(f"### {doc.get('filename', 'document')}")
                lines.append("")
                lines.append((doc.get("content") or "").strip())
                lines.append("\n---\n")

        (out_dir / filename).write_text("\n".join(lines), encoding="utf-8")
        index.append((name, len(docs), f"projects/{filename}"))

    return index


def write_memories(memories: list) -> bool:
    if not memories:
        return False
    lines = ["# Account memory", ""]
    for entry in memories:
        text = (entry.get("conversations_memory") or "").strip()
        if text:
            lines.append(text)
            lines.append("")
    (ROOT / "memories.md").write_text("\n".join(lines), encoding="utf-8")
    return True


def write_index(users, conv_index, proj_index, has_memories):
    lines = [
        "# Claude data export",
        "",
        "Human-readable archive generated from a Claude data export.",
        "Raw JSON lives in [`raw/`](raw/); regenerate the Markdown with",
        "`python3 claude-data/scripts/export_to_markdown.py`.",
        "",
        "## Account",
        "",
    ]
    for user in users:
        name = user.get("full_name", "")
        email = user.get("email_address", "")
        lines.append(f"- **{name}** — {email}")
    now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    lines += ["", f"_Generated: {now}_", ""]

    if has_memories:
        lines += ["## Memory", "", "- [Account memory](memories.md)", ""]

    if proj_index:
        lines += ["## Projects", ""]
        for name, doc_count, rel in sorted(proj_index):
            lines.append(f"- [{name}]({rel}) — {doc_count} doc(s)")
        lines.append("")

    lines += ["## Conversations", "", f"Total: {len(conv_index)}", ""]
    lines.append("| Date | Conversation | Messages |")
    lines.append("| --- | --- | --- |")
    for name, created, msg_count, rel in sorted(conv_index, key=lambda x: x[1], reverse=True):
        display = (name or "untitled").replace("|", "\\|")
        lines.append(f"| {created[:10]} | [{display}]({rel}) | {msg_count} |")
    lines.append("")

    (ROOT / "README.md").write_text("\n".join(lines), encoding="utf-8")


def main():
    users = load_json(RAW / "users.json") if (RAW / "users.json").exists() else []
    memories = load_json(RAW / "memories.json") if (RAW / "memories.json").exists() else []
    conversations = load_json(RAW / "conversations.json") if (RAW / "conversations.json").exists() else []
    project_files = sorted((RAW / "projects").glob("*.json")) if (RAW / "projects").exists() else []

    conv_index = write_conversations(conversations)
    proj_index = write_projects(project_files)
    has_memories = write_memories(memories)
    write_index(users, conv_index, proj_index, has_memories)

    print(f"Wrote {len(conv_index)} conversation(s), {len(proj_index)} project(s), "
          f"memories={'yes' if has_memories else 'no'}.")


if __name__ == "__main__":
    main()
