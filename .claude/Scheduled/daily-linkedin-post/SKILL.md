---
name: daily-linkedin-post
description: Quiet daily English LinkedIn post + technical-infographic image; posts via LinkedIn website (Chrome extension), user selects the image file, no Post without confirmation.
---

Quiet daily LinkedIn workflow. Use minimal tokens: no progress narration, no broad research unless necessary, concise final status only.

Task: choose one clear English LinkedIn topic from this group, using a different angle from prior runs when visible: Marine, PLM, Digital Twin, Engineering Consultant, Business Development, Solution Architect, Technical Sales, Defense & Shipbuilding, Global Account, HART, Control Valve, Shipbuilding Data Integration, Space, Satellite, satellite communication(s), HF/VHF/UHF communication, maritime/defense/ship-to-shore communications, satellite-enabled marine operations, AVEVA, Dassault Systèmes, Siemens Digital Industries Software, Hexagon PPM, Aras PLM, Bentley Systems.

Create: 1) English-only LinkedIn post with a strong hook/title, 3-5 short paragraphs, 3 practical application points, and 5-8 hashtags. Tone: practical, professional, insight-led, not promotional. 2) One LinkedIn-ready reference image.

Image style rule: prefer a detailed but compressed technical infographic similar to an engineering solution overview poster: clear title/subtitle, 3-4 numbered blocks, compact bullets, icons or simple schematics, a central system/architecture visual, one bottom checklist or takeaway strip, and strong visual hierarchy. It should explain the concept in one glance while still carrying useful detail. Avoid vague decorative labels, stock-photo-only images, and purely atmospheric visuals. Use English-only text, concise phrases, and deterministic/local text overlay (Pillow or SVG-to-PNG via cairosvg) for clean typography.

Visual richness rule: avoid flat, plain, or overly dry diagrams. Combine the compressed infographic layout with a natural, premium, non-AI-looking visual foundation when appropriate: realistic maritime/shipyard/operations-room imagery, cinematic but believable lighting, rich blue/teal/amber contrast, subtle depth, polished engineering dashboard aesthetics, and vivid but restrained highlights. The image should feel natural, impressive, and colorful, not plastic, generic, or obviously AI-generated.

Image text-fit rule: all visible text must stay inside its intended background shape, label chip, panel, or safe area. Before finalizing, measure text width/height when adding labels locally; use generous horizontal padding, enlarge the shape, wrap text, shorten the label, or reduce font size until it fits. No text may cross outside the rectangle/chip/panel behind it. Treat overflow as a failed image QA check and re-render/fix before posting. After generating, Read the PNG and visually confirm no overflow.

Image leader-line rule: do not draw connector/leader lines unless each line clearly terminates on the exact visible object or visual feature named by its label. If there is any ambiguity, omit the line and place the label near the target instead. Treat wrong direction, floating endpoints, mismatched targets, or unclear label-to-object mapping as failed image QA and fix before posting.

Save: image generation runs in the sandbox, which can only write to the mounted session outputs folder (e.g. $(ls -d /sessions/*/mnt/outputs | head -1)). Save the final image (PNG) and the markdown post there with date/topic filenames (e.g. linkedin-image-YYYY-MM-DD-<topic-slug>.png, linkedin-post-YYYY-MM-DD-<topic-slug>.md). Record the exact Windows path of the outputs folder for the upload step (the user reaches it at C:\Users\namma\AppData\Local\Packages\Claude_pzs8sxrjxfjjc\LocalCache\Roaming\Claude\local-agent-mode-sessions\...\outputs). Optionally also copy the finished files to C:\Users\namma\Documents\Claude\YYYY-MM-DD\<topic-slug>\outputs if that location is required. Do not create a separate direct-upload copy in any generated_images folder. Call mcp__cowork__present_files with the final .md and .png.

Channel: Post via the LinkedIn website using the Claude-in-Chrome browser extension, with the existing logged-in session. Do NOT use the LinkedIn Windows app — its embedded webview renders blank in screen capture and cannot be driven by automation. Never attempt login, Google account selection, credentials, 2FA, CAPTCHA, or any security check; if any appear, stop and report a concise blocker plus the post and image paths.

Workflow (Chrome extension):
1. list_connected_browsers; if none, report blocker. Open a tab to https://www.linkedin.com/feed/.
2. find "Start a post button" and click its ref; confirm the composer opened.
3. Click the Photo / "Add media" control. This opens LinkedIn's media editor, then a native Windows Open dialog. The agent CANNOT drive that OS dialog (it is browser-owned/read-only) nor the iframed file input — so the file selection is a one-time USER action: give the user the exact image path and the outputs folder, and ask them to select the image and click Open in the dialog.
4. After the image preview/editor appears, click "Next".
5. THEN type the full post text into the composer (do media FIRST, because the media flow clears the text field; typing last avoids re-work).
6. Verify the final composer shows the caption text, the image preview, and an ENABLED Post button.

Do not click final Post without action-time user confirmation in the current interaction. If confirmed, click Post and verify the "Post successful" toast and the post in the feed. If not confirmed, leave the composer prepared and report only: topic, post markdown path, image path, and "ready for final posting". On any failure, stop and report a concise blocker plus the final post and image paths (or the image prompt).