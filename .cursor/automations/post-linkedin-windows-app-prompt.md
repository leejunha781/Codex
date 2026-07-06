# LinkedIn Windows App Posting (scheduled auto-post)

Use this workflow after `mirror-linkedin-runs.ps1` has mirrored today's run to:

`C:\Users\namma\Documents\Codex\YYYY-MM-DD\<topic-slug>\`

## Inputs

1. Read `ready-for-posting.json` in the mirrored topic folder (or `C:\Users\namma\.cursor\automations\cache\linkedin-mirror\latest-mirror-runs.json`).
2. Read post text from `linkedin-post.md` in the same folder.
3. Use `imagePath` from the manifest for media upload.

## Rules

- Post via **LinkedIn Windows app only** with the existing logged-in session.
- Never use Chrome and never attempt login, Google selection, credentials, 2FA, CAPTCHA, or security checks.
- For this **scheduled automation**, click **Post** automatically after composer verification. Do not wait for user confirmation.
- If LinkedIn is already posted for this topic/date (check memory or duplicate composer), skip and report `already posted`.
- On any blocker, stop and report: topic, post path, image path, blocker.

## Workflow

1. Open/focus LinkedIn Windows app.
2. Click **Start a post**.
3. Paste post text from `linkedin-post.md` into **Share your thoughts** (use click/type_text if set_value fails).
4. Click **Add media**.
5. In the Windows **Open** dialog, enter the full `imagePath` in **File name** and press Open/Return.
6. When image preview/editor appears, click **Next**.
7. Verify composer has text, image preview, and enabled **Post** button.
8. Click **Post** and verify success toast or feed appearance.
9. Append run summary to `.codex/automations/daily-linkedin-marine-plm-post/memory.md` with `LinkedIn status: posted`.

## Output paths (fixed)

- Folder: `C:\Users\namma\Documents\Codex\YYYY-MM-DD\<topic-slug>\`
- Post: `...\linkedin-post.md`
- Image: `...\`<topic-slug>`-infographic.png`
