# RiceMate – Tamil/English voice entry for rice sales

Runs fully in the browser. No backend, account or paid API. Data stays on the device (IndexedDB).

## Run (Windows CMD)
```
cd ricemate
npm install
npm run sample
npm test
npm run dev
```
Open the printed address. To use it on a phone, open `http://<PC-IP>:5173` on the same Wi-Fi. Voice needs HTTPS or localhost, so for real phone use run `npm run build` and host the `dist` folder on any HTTPS static host (Netlify, GitHub Pages, etc.).

## Using it
1. **Excel** tab: pick the company workbook, check worksheet, header row and column mapping.
2. **Voice** tab: tap mic (Chrome on Android recommended; needs internet), or type. Review/edit fields, **Save Entry** (asks to confirm).
3. **Entries** tab: **Export Updated Excel** downloads a NEW file `name_updated_YYYYMMDD_HHMMSS.xlsx`. The original is never touched. The exported copy becomes the base for later entries.

## Data safety
- Statuses: Pending (draft) → Confirmed → Exported.
- Phone numbers stored as text. Quantity/price written as numbers. Blank stays blank; nothing is invented.
- Duplicate entries (same name, phone, variety, quantity, date) are rejected.
- Browser data can be wiped if the user clears site data: export regularly.
- Export is blocked for `.xlsm`/`.xlsb` (macros). Formatting, charts, images, validation may be lost in other files (free SheetJS). Formulas in existing rows are kept; new rows get no auto-formulas. **Test with a COPY of the real workbook before relying on it.**
- `xlsx@0.18.5` (npm) has known advisories (prototype pollution/ReDoS on untrusted files). Only open your own workbooks.

## Tests
`npm test` checks parsing, header detection, column mapping, append, round-trip reopen, phone-as-text, formula retention. `sample/RiceMate_sample.xlsx` is a test file only.
Not automated (test by hand): real browser speech recognition, phone file picker, Tamil speech accuracy, your real workbook's formatting.
