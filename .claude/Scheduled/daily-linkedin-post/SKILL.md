---
name: daily-linkedin-post
description: Daily English LinkedIn post on a rotating topic (16-topic rotation incl. Space/Satellite & HF/VHF/UHF SATCOM) + a matching branded image; stages the post for the user to click Post.
---

You are preparing and staging today's daily LinkedIn post for Joonha Lee: ONE professional English post on the day's rotating topic, PLUS a matching branded image. Save both, stage the post (type the text into the composer and attach the image where possible), and STOP before clicking "Post". The user reviews and clicks Post themselves.

STEP 1 — Determine today's topic by rotation.
In bash run: date +"%Y-%m-%d (day %j)". Take day-of-year (strip leading zeros), compute index = day_of_year % 16.
Ordered topics (0-based):
0 Marine Engineering | 1 PLM (Product Lifecycle Management) | 2 Digital Twin | 3 AVEVA | 4 Engineering Consulting | 5 Business Development | 6 Solution Architect | 7 Technical Sales | 8 Defense & Shipbuilding | 9 Global Account Management | 10 HART Protocol | 11 Control Valve | 12 Shipbuilding Data Integration | 13 Space & Satellite Systems | 14 HF/VHF/UHF Radio Communications | 15 Satellite Communication (SATCOM)

STEP 2 — Write the post (professional English).
- 130–220 words. Strong specific first-line hook (no "Today I want to talk about..."). 2–4 short paragraphs or a tight mini-list, ONE concrete practical insight grounded in real industrial / marine / shipbuilding / space / RF-communications engineering reality (AVEVA, PLM, digital twin, control valves, HART protocol, shipbuilding data integration, space & satellite systems, HF/VHF/UHF radio links, SATCOM link budgets, antennas, ground stations, etc.). Senior practitioner / solution-architect voice, useful not salesy. End with one reflective question or light CTA. 4–6 relevant hashtags. No fabricated statistics. No emojis (or at most one). Vary the hook day to day.
- Also craft a SHORT image HEADLINE: a punchy 6–14 word distillation of the post's core idea (often a tightened version of the hook), suitable for a graphic.

STEP 3 — Generate the matching branded image.
The reusable generator lives in the persistent outputs folder at: C:\Users\namma\AppData\Roaming\Claude\local-agent-mode-sessions\dc934002-beb5-48c7-84bf-566d0d5aefca\d3453b4f-497b-4e7e-bbff-aa1b470046d5\local_ec2fa5d9-7315-476a-b5b2-5d99e69c367f\outputs\make_linkedin_image.py
In bash, locate the mounted outputs folder and run the generator:
  OUT=$(ls -d /sessions/*/mnt/outputs 2>/dev/null | head -1)
  python3 -c "import PIL" 2>/dev/null || pip install pillow --break-system-packages -q
  python3 "$OUT/make_linkedin_image.py" --topic "<TOPIC>" --headline "<IMAGE HEADLINE>" --tags "<3-4 of the hashtags, separated by 3 spaces>" --date "YYYY-MM-DD" --out "$OUT/linkedin-image-YYYY-MM-DD.png"
Preferred premium style: a polished 1200x1200 card on a deep navy gradient with a teal accent — an eyebrow topic label, a bold headline (accent the key phrase), three condensed at-a-glance takeaway rows with simple line-icon badges that summarize the post, subtle marine/industrial motif, a footer with 3 hashtags, and a cursive "Joonha Lee" signature bottom-right (URW Chancery L / Z003 font). If a generator script is missing, recreate a clean Pillow-based generator (or render an SVG via cairosvg) that produces this layout, save it to OUT for reuse, and note that you recreated it. Use the Read tool to view the generated PNG and confirm text fits with no overflow.

STEP 4 — Save the text draft (record + fallback).
Write to the outputs folder, filename linkedin-post-YYYY-MM-DD.md, containing: "# LinkedIn Post — YYYY-MM-DD", "**Topic:** <topic>", "**Image:** linkedin-image-YYYY-MM-DD.png", a "## POST TEXT" section with the full ready-to-post text, and a word-count note.

STEP 5 — Stage the post (best effort).
- Preferred path is the LinkedIn website via the Claude-in-Chrome extension (the LinkedIn Windows app's webview often renders BLANK in screen capture and can't be driven reliably). If a Chrome browser is connected (list_connected_browsers), open a tab to https://www.linkedin.com/feed/, click "Start a post" (locate it via the find tool and click its ref), and type the full post text into the composer.
- Image attach caveat: LinkedIn's "Upload from computer" uploader runs in an iframe / native OS picker that the browser tools usually cannot reach (file_upload needs a file-input ref that isn't exposed, and the OS picker is browser-owned/read-only). So the image typically must be attached by the user. Leave the composer with the text staged and tell the user the image (outputs\linkedin-image-YYYY-MM-DD.png) is ready to attach.
- If only the Windows app is available: request_access for "LinkedIn", open it, screenshot; if the content is blank/white after ~15s + one reload, do NOT keep retrying — report that the app didn't render and the user should post manually using the saved files.
- IMPORTANT: never click "Post"/"Publish"/send. Leave it staged for the user.

STEP 6 — Notify.
Call mcp__cowork__present_files with BOTH the saved .md draft and the .png image. Then write 1–2 concise sentences: today's topic; whether the text was typed into the composer; and that the image is saved and ready for the user to attach + click Post (at https://www.linkedin.com/feed/ or the LinkedIn app).