---
name: daily-linkedin-post
description: Daily English LinkedIn post on a rotating topic + a matching branded image; types into the LinkedIn Windows app for the user to click Post.
---

You are preparing and staging today's daily LinkedIn post for Joonha Lee: ONE professional English post on the day's rotating topic, PLUS a matching branded image. Save both, type the text into the LinkedIn Windows app composer, attach the image if possible, and STOP before clicking "Post". The user reviews and clicks Post themselves.

STEP 1 — Determine today's topic by rotation.
In bash run: date +"%Y-%m-%d (day %j)". Take day-of-year (strip leading zeros), compute index = day_of_year % 13.
Ordered topics (0-based):
0 Marine Engineering | 1 PLM (Product Lifecycle Management) | 2 Digital Twin | 3 AVEVA | 4 Engineering Consulting | 5 Business Development | 6 Solution Architect | 7 Technical Sales | 8 Defense & Shipbuilding | 9 Global Account Management | 10 HART Protocol | 11 Control Valve | 12 Shipbuilding Data Integration

STEP 2 — Write the post (professional English).
- 130–220 words. Strong specific first-line hook (no "Today I want to talk about..."). 2–4 short paragraphs or a tight mini-list, ONE concrete practical insight grounded in real industrial / marine / shipbuilding engineering reality (AVEVA, PLM, digital twin, control valves, HART protocol, shipbuilding data integration, etc.). Senior practitioner / solution-architect voice, useful not salesy. End with one reflective question or light CTA. 4–6 relevant hashtags. No fabricated statistics. No emojis (or at most one). Vary the hook day to day.
- Also craft a SHORT image HEADLINE: a punchy 6–14 word distillation of the post's core idea (often a tightened version of the hook), suitable for a graphic.

STEP 3 — Generate the matching image.
The reusable generator lives in the persistent outputs folder at: C:\Users\namma\AppData\Roaming\Claude\local-agent-mode-sessions\dc934002-beb5-48c7-84bf-566d0d5aefca\d3453b4f-497b-4e7e-bbff-aa1b470046d5\local_ec2fa5d9-7315-476a-b5b2-5d99e69c367f\outputs\make_linkedin_image.py
In bash, locate the mounted outputs folder and run the generator:
  OUT=$(ls -d /sessions/*/mnt/outputs 2>/dev/null | head -1)
  python3 -c "import PIL" 2>/dev/null || pip install pillow --break-system-packages -q
  python3 "$OUT/make_linkedin_image.py" --topic "<TOPIC>" --headline "<IMAGE HEADLINE>" --tags "<3-4 of the hashtags, separated by 3 spaces>" --date "YYYY-MM-DD" --out "$OUT/linkedin-image-YYYY-MM-DD.png"
If make_linkedin_image.py is missing for any reason, skip the image gracefully and note it. Use the Read tool to view the generated PNG and confirm it looks correct (text fits, no overflow).

STEP 4 — Save the text draft (record + fallback).
Write to the outputs folder, filename linkedin-post-YYYY-MM-DD.md, containing: "# LinkedIn Post — YYYY-MM-DD", "**Topic:** <topic>", "**Image:** linkedin-image-YYYY-MM-DD.png", a "## POST TEXT" section with the full ready-to-post text, and a word-count note.

STEP 5 — Stage in the LinkedIn Windows app (best effort).
- Load computer-use tools via ToolSearch (query "computer-use", max_results 30) if not loaded. Call request_access for the "LinkedIn" app. open_application "LinkedIn". Screenshot.
- If the app content is BLANK/white or fails to load after ~15s and a reload, do NOT keep retrying — skip to STEP 6 and report that the app didn't render so the user should post manually using the saved files.
- If the feed renders: click "Start a post", confirm the composer is open, and type the full post text.
- Try to attach the image: click the photo/media button in the composer, and in the file dialog paste the image path (the Windows path is the outputs folder + linkedin-image-YYYY-MM-DD.png) and confirm/open. If the media flow is unclear or fails, leave the text typed and note the image must be added manually.
- IMPORTANT: never click "Post"/"Publish"/send. Leave it staged for the user.

STEP 6 — Notify.
Call mcp__cowork__present_files with BOTH the saved .md draft and the .png image. Then write 1–2 concise sentences: today's topic; whether the post was typed into the app and image attached, or (if the app was blank) that the text + image are saved and ready for the user to post manually at the LinkedIn app or https://www.linkedin.com/feed/.