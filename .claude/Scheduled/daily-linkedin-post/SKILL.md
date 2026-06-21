---
name: daily-linkedin-post
description: Generates one English LinkedIn post per day on a rotating industrial/marine engineering topic, ready at 6 PM.
---

You are preparing today's daily LinkedIn post draft for Joonha Lee. Produce ONE ready-to-publish post in professional English on the day's rotating topic, save it to a file, and present it. The user reviews and posts it manually at https://www.linkedin.com/feed/ ("Start a post").

STEP 1 — Determine today's topic by rotation.
Run a bash command to get today's date and day-of-year, e.g.: `date +"%Y-%m-%d (day %j)"`. Take the day-of-year number (strip leading zeros) and compute: index = day_of_year % 13.
Map index to topic using this ordered list (0-based):
0 = Marine Engineering
1 = PLM (Product Lifecycle Management)
2 = Digital Twin
3 = AVEVA
4 = Engineering Consulting
5 = Business Development
6 = Solution Architect
7 = Technical Sales
8 = Defense & Shipbuilding
9 = Global Account Management
10 = HART Protocol
11 = Control Valve
12 = Shipbuilding Data Integration

STEP 2 — Write the post (professional English).
- 130–220 words.
- Strong, specific first line as a hook (no "Today I want to talk about..."). Lead with a tension, a sharp claim, or a concrete scenario.
- 2–4 short paragraphs, or a tight mini-list, with ONE concrete, practical insight grounded in real industrial / marine / shipbuilding engineering reality (AVEVA, PLM, digital twin, control valves, HART protocol, shipbuilding data integration, etc.).
- Voice: senior practitioner / solution architect — confident, useful, never salesy.
- End with one reflective question or a light call-to-action to invite comments.
- 4–6 relevant hashtags blending the day's topic with evergreen tags (e.g. #DigitalTwin #PLM #AVEVA #Shipbuilding #MarineEngineering #IndustrialSoftware #ControlValve #HART #DefenseTech).
- No fabricated statistics or fake citations. Any number must be clearly illustrative. Use no emojis, or at most one professional emoji.
- Each day's post must feel distinct — vary the hook style and structure from a typical post. Do not reuse the same opening pattern every day.

STEP 3 — Save the draft.
Write the post to this folder: C:\Users\namma\AppData\Roaming\Claude\local-agent-mode-sessions\dc934002-beb5-48c7-84bf-566d0d5aefca\d3453b4f-497b-4e7e-bbff-aa1b470046d5\local_ec2fa5d9-7315-476a-b5b2-5d99e69c367f\outputs\
Filename: linkedin-post-YYYY-MM-DD.md (use today's date).
File contents:
  # LinkedIn Post — YYYY-MM-DD
  **Topic:** <topic>
  **Paste at:** https://www.linkedin.com/feed/ → "Start a post"
  ---
  ## POST TEXT
  <the full post text, ready to copy-paste>
  ---
  *Word count: ~N. Review, then post.*

STEP 4 — Present it.
Call mcp__cowork__present_files with the saved file path so the user can open it. Then write a 1–2 sentence summary: today's topic and a reminder to review and post at https://www.linkedin.com/feed/. Keep it concise.