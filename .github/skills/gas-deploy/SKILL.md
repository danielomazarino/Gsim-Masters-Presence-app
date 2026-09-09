---
name: gas-deploy
description: Deploy local Index.html / Kod.gs to the Göteborg Sim Apps Script web app via the built-in browser, with verification. Use when the user asks to push, deploy, sync, or publish changes to the Script_Training_presence Apps Script project (a.k.a. Gsim Masters Presence app).
---

# Apps Script Deploy via Built-in Browser (clasp blocked)

## Context

- Project: `/home/lm/Dev/Gsim-Masters-Presence-app` — `Index.html` (frontend) + `Kod.gs` (backend).
- Remote: Apps Script project `Script_Training_presence`, script ID `1X3gV-3Wq0NvhaN8EZTUvfM9onboCBe3F7dBpLiHhoWOa_qE-NNd2ZJKx`.
- **clasp does NOT work** for this Workspace account (`goteborgsim.se`): Google rejects every refresh token with `invalid_grant / rapt_required` (Workspace security policy on clasp's public OAuth client). Do not retry clasp login.
- Web app URL (stable across versions): `https://script.google.com/a/macros/goteborgsim.se/s/AKfycby-YQowvpANxzysEErwhF2cP_AkfDJXWRBmZJ3wNJ29_pFD1uCF-kMqeGmDxVTmVPI3/exec`
- Editor URL: `https://script.google.com/u/0/home/projects/1X3gV-3Wq0NvhaN8EZTUvfM9onboCBe3F7dBpLiHhoWOa_qE-NNd2ZJKx/edit`
- Login: `masters@goteborgsim.se` — session usually alive in the built-in browser. If a password prompt appears, STOP and let the user type it (never handle credentials).

## Prerequisites

- Local file(s) already edited and mirrored to `appsscript/` (`cp Index.html appsscript/`).
- Apps Script editor open in the built-in browser, logged in.

## Procedure

### 0. Read remote state first (always)

Monaco is virtualized — never scrape `.view-lines`. Use the model API:

```js
return await page.evaluate(() => {
  const models = monaco.editor.getModels();
  const target = models.find(m => m.uri.path.endsWith('.html')); // '.js' for Kod.gs
  const t = target.getValue();
  let h = 0; for (let i = 0; i < t.length; i++) { h = ((h<<5)-h+t.charCodeAt(i))|0; }
  return { len: t.length, hash: h };
});
```

Compare `len`/hash with the local file. If identical → nothing to deploy.

Fingerprint reference (2026-09-09): Kod.gs = 42112 chars, hash -710490996.

### 1. Put the local file on the system clipboard

```bash
xclip -selection clipboard -i <file> && xclip -selection clipboard -o | md5sum   # verify vs md5sum <file>
```

(xclip is installed; session is X11.)

### 2. Select the correct file tab, then paste

**CRITICAL — the paste-into-wrong-file trap:** `Ctrl+A` + `Ctrl+V` go to whichever model is active. Always explicitly click the file tab first and verify by checking the model content prefix afterwards:

```js
await page.getByRole('option', { name: 'Index.html' }).first().click();
await page.waitForTimeout(800);
const editorArea = page.locator('.monaco-editor .view-lines').first();
await editorArea.click();
await page.waitForTimeout(300);
await page.keyboard.press('Control+a');
await page.waitForTimeout(300);
await page.keyboard.press('Control+v');
await page.waitForTimeout(3000);
```

Then verify: model `len` equals local file length exactly. **If the wrong model got overwritten, `model.undo()` restores it — verify with the fingerprint from step 0.** Never save while any model has wrong content.

Do NOT inject content by chunked `pushEditOperations` — chunk-boundary duplication errors are easy to make and hard to spot (this caused a real near-miss; the paste path is exact).

### 3. Save

```js
await page.keyboard.press('Control+s');
await page.waitForTimeout(4000);
```

Verify: save button (`aria-label` contains "Spara projekt") is `disabled` and page text contains `Har sparats på Drive`.

### 4. Deploy new version

```js
await page.getByRole('button', { name: 'Implementera' }).first().click();
await page.waitForTimeout(1000);
await page.getByRole('menuitem', { name: 'Hantera implementeringar' }).click();
await page.waitForTimeout(2000);
await page.getByRole('option', { name: 'v11' }).first().click(); // highest vN = active deployment
await page.waitForTimeout(1000);
await page.getByRole('button', { name: 'Redigera' }).first().click();
await page.waitForTimeout(1500);
await page.getByRole('combobox', { name: /Version/ }).first().click();
await page.waitForTimeout(800);
await page.getByRole('option', { name: 'Ny version' }).click();
await page.waitForTimeout(800);
// DEPLOY — get user approval BEFORE this point
await page.getByRole('dialog').getByRole('button', { name: 'Implementera' }).first().click();
await page.waitForTimeout(5000);
```

Success indicator: text `Implementeringen har uppdaterats. Version N den <date>` inside the dialog. Click `Klar` to close.

To inspect WITHOUT deploying: do everything except the final click; `Escape` closes the version dropdown and dialog with no changes.

### 5. Verify live

```js
await page.setViewportSize({ width: 390, height: 844 }); // iPhone 13 view when testing mobile layout
await page.goto('<web app URL>');
await page.waitForTimeout(6000);
const appFrame = page.frames().find(f => f.name() === 'userHtmlFrame');
return await appFrame.evaluate(() => { /* check the actual fix */ });
```

Hard-refresh may be needed on the phone (Apps Script serves cached HTML briefly after deploy).

## Gotchas

- **Saving a file does NOT update the deployed web app** — a new version must be deployed (step 4). This bit us once: user refreshed their phone after editor-save only and saw no change.
- The web app is served inside nested iframes; app code lives in the frame named `userHtmlFrame`. Cross-origin — use `page.frames()` + `frame.evaluate`, never `document.querySelector` from the top page.
- `page.evaluate(() => document...)` runs in the TOP page; Monaco APIs are available there (editor chrome), app DOM is not.
- Version dropdown shows both deployments and versions; pick the option literally named `Ny version`, not `Version N`.
- After deploy, old version stays on the version list (rollbacks possible via the same dialog).
