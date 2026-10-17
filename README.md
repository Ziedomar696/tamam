# تمام الملاك — مراسي

App link: **https://ziedomar696.github.io/tamam/**

- `apps-script/` — the app itself (Google Apps Script): paste `Code.gs` and `Index.html` into a new Apps Script project, set `SETUP_CODE`, run `setup`, deploy as a Web app (Execute as: Me, Access: Anyone).
  Data goes to the Google Sheet **تمام الملاك** (sheets: سجل التمام، سجل التعديلات، المستخدمين، القطاعات، الإعدادات).
- `index.html`, `manifest.webmanifest`, `sw.js`, `icons/` — the installable wrapper (PWA) served by GitHub Pages. Put the `/exec` URL in `APP_URL` in `index.html`.

Each sector submits a headcount (owners / tenants / delegates) four times a day (6am, 12pm, 6pm, midnight) from its own account.
The supervisor name comes from the account, every change is logged, late submissions are flagged, and admins/managers get one e-mail per round listing sectors that did not submit.
