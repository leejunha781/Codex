---
name: daily-linkedin-post
description: Generates one English LinkedIn post daily, types it into the LinkedIn Windows app composer, and stops for the user to click Post.
---

You are preparing and staging today's daily LinkedIn post for Joonha Lee. Generate ONE post in professional English on the day's rotating topic, save a copy, then open the LinkedIn WINDOWS APP and type the post into the composer — but DO NOT click "Post". The user reviews it in the app and clicks Post themselves.

STEP 1 — Determine today's topic by rotation.
Run a bash command to get day-of-year, e.g. `date +"%Y-%m-%d (day %j)"`. Take the day-of-year number (strip leading zeros) and compute: index = day_of_year % 13.
Ordered topic list (0-based):
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
- Strong, specific first-line hook (no "Today I want to talk about..."). Lead with a tension, sharp claim, or concrete scenario.
- 2–4 short paragraphs or a tight mini-list, ONE concrete practical insight grounded in real industrial / marine / shipbuilding engineering reality (AVEVA, PLM, digital twin, control valves, HART protocol, shipbuilding data integration, etc.).
- Voice: senior practitioner / solution architect — confident, useful, never salesy.
- End with one reflective question or light CTA.
- 4–6 relevant hashtags blending the day's topic with evergreen tags (#DigitalTwin #PLM #AVEVA #Shipbuilding #MarineEngineering #IndustrialSoftware #ControlValve #HART #DefenseTech).
- No fabricated statistics. No emojis, or at most one professional emoji. Vary the hook style day to day.

STEP 3 — Save a copy (record + fallback).
Write the post to this folder: C:\Users\namma\AppData\Roaming\Claude\local-agent-mode-sessions\dc934002-beb5-48c7-84bf-566d0d5aefca\d3453b4f-497b-4e7e-bbff-aa1b470046d5\local_ec2fa5d9-7315-476a-b5b2-5d99e69c367f\outputs\
Filename: linkedin-post-YYYY-MM-DD.md (today's date), with: a "# LinkedIn Post — YYYY-MM-DD" heading, "**Topic:** <topic>", a "## POST TEXT" section containing the full ready-to-post text, and a word-count note.

STEP 4 — Stage it in the LinkedIn Windows app.
- Load the computer-use tools via ToolSearch (query "computer-use", max_results 30) if they are not already loaded.
- Call request_access for the LinkedIn Windows app (application name "LinkedIn"). 
- Use open_application to bring the LinkedIn app to the front. Take a screenshot to see the current state.
- Open the post composer: click the "Start a post" box (or the equivalent compose/create-post control). Take a screenshot to confirm the composer is open and focused.
- Type the full post text into the composer using the type tool.
- Take a screenshot to verify the text is entered correctly.
- IMPORTANT: STOP HERE. DO NOT click "Post", "Publish", or any send/submit control. Leave the composer open with the text ready so the user can review and click Post themselves. Posting publicly is the user's action.
- If the LinkedIn Windows app cannot be found or opened, or the composer can't be located, skip the app step gracefully.

STEP 5 — Notify.
Call mcp__cowork__present_files with the saved draft file path. Then write a 1–2 sentence summary: today's topic, that the post has been typed into the LinkedIn app composer and is waiting for the user to review and click Post (or, if the app step failed, that the draft is saved and ready to paste manually). Keep it concise.