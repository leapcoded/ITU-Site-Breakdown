TERRA — Read Me (Non-technical Guide)

Welcome to TERRA — a small browser app for editing roster-style tables and creating printable PDF reports. This document explains how to use the app from the user interface; you do not need any technical skills or the browser console.

Quick overview

- Add categories to group files for a specific report section.
- Upload CSV or XLSX files (for tables) and image files (for photos) directly into a category.
- Add notes and attach reusable library items (files, images, notes) to categories.
- Use the Customize screen to pick which columns to include, reorder them, and add simple filters.
- Set a report title, subtitle, and logo, then click Generate PDF to download the finished report.

Start here (step-by-step)

1) Open the app

- Launch the app in a modern browser (Chrome, Edge, Firefox, Safari). No installation is required.

2) Create categories

- Click "+ Add Category" to create a new category box. Use categories to group files that belong together in the report (for example: "Site A", "Site B", or "Electrical Team").

3) Upload files

- Click inside a category's upload area (or drag-and-drop) to add files. Supported file types:
  - CSV (.csv) — comma-separated values for tables
  - Excel (.xlsx) — spreadsheet-style tables
  - Image files (.png, .jpg, .jpeg, .webp, .svg) — for photos or logos

- After you upload a CSV/XLSX file the app will parse it into a table you can preview and edit.

4) Add notes and library attachments

- Use the "+ Note" button inside a category to attach textual notes that will appear in the report.
- Use the "Library" button to reuse previously uploaded files, images, or notes across categories.

5) Customize your report

- Click "Customize Report" to open the customization screen.
- Choose which columns to include in the final report, reorder columns by dragging, and add simple filter rules (for example: show only rows where Role contains "Technician").

6) Set report title, subtitle and logo

- Use the project controls (top-right or inside the Customize screen) to enter a report title and subtitle.
- Upload an organization logo to brand the PDF. The logo is added client-side and embedded into the generated PDF.

7) Generate PDF

- When ready, click "Generate PDF Report". The app compiles your selected columns and filters, embeds any attached images, and prompts a PDF download. Large images may increase generation time.

Helpful details and tips

- Editing parsed tables: click "Edit" on a parsed file to open the inline table editor. You can hide rows, correct values, and change the file's detected locale for date parsing.
- Image recommendations: for best PDF quality, use images between 800–2000 pixels on the long side. If you upload very large originals, the app may compress them to keep the report size reasonable.
- Reusing content: items added to the Library are available to attach to any category.
- Local storage: all data you add (categories, files, images, notes, and your theme) is stored in your browser on this computer only. No data is sent to any server unless you explicitly share the generated PDF yourself.

Settings and clearing data

- Open Settings to change the app theme or clear saved data.
- If you want to remove everything and start fresh, use Settings -> Clear saved data. This removes stored categories, library items, and other saved state from your browser.

Troubleshooting (no console required)

- The page looks broken or a button is unresponsive:
  - Try refreshing the page (reload).
  - If the problem persists, open Settings and choose "Clear saved data" then reload the page.

- Uploaded table doesn't look right or columns are wrong:
  - Make sure your file is a valid CSV or XLSX file.
  - Open the file's Edit view to correct headings or values before generating your report.

- Images are low-quality in the PDF:
  - Upload a higher-resolution original image. TIFF/BMP files are not recommended — use PNG or JPG.

- The PDF generation seems slow:
  - Large images increase processing time. Try reducing the number or size of images, or allow a few extra seconds for the app to prepare the file.

Accessibility & keyboard tips

- You can navigate the interface using Tab / Shift+Tab and activate buttons with Enter or Space.
- Most interactive elements (buttons, inputs) follow standard browser keyboard behavior.

Privacy and security

- The app stores your data in your browser (local storage / IndexedDB). Data stays on your machine.
- Generated PDFs are downloaded to your device — share them only with people you trust.

Frequently asked questions

Q: Do I need to install anything?
A: No. Open the app URL in a modern browser and use it.

Q: Will my data be uploaded anywhere?
A: No — files and images you add stay in your browser. The app only uses the resources in your local browser session.

Q: Can I edit a CSV before generating the report?
A: Yes. Use the file's Edit control to adjust values and hide rows.

Need more help?

If something still isn't working, tell the person who set up this app for you (support contact or IT) what you tried and which step failed. Provide a short description such as "Uploading a CSV fails" or "Generate PDF button doesn't respond" and include the browser and version if possible.

Credits & version

- This tool is a small, client-side web app that parses tables and generates PDFs entirely in your browser.

---

Cleanup performed on 2025-10-12:
- Removed legacy/duplicate JS files not referenced by `index.html`.
- Removed Node artifacts (`node_modules/` and `package-lock.json`) to trim the repository. To restore the Node environment, run `npm install`.

Additionally on 2025-10-12:
- Removed ESLint and Prettier configuration files and devDependencies from `package.json` because linting/formatting is not required in this browser-only repo. If you want to re-enable them, restore `package.json` or add the devDependencies then run `npm install` and re-create configs.

Date handling change (2025-10-12):
- Canonical internal date representation has been switched to ISO format (YYYY-MM-DD). The app still displays dates to users according to the selected display locale (UK or US) but all internal storage and alert/filter matching use ISO to avoid ambiguity.

If you'd like, I can also:
- Convert this Read Me into an in-app full page (instead of a modal) and link the "Read Me" nav item to it.
- Add a printable PDF of this Read Me accessible from the app UI.

Tell me what you prefer and I'll add it.