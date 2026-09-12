# Excel Intelligence Dashboard

A browser-only Excel analytics dashboard. Upload an `.xlsx`, `.xls`, or `.csv` file and the app automatically detects useful fields, builds KPIs/charts, and provides date and dimension filters.

## Run locally

```bash
npm install
npm run dev
```

Then open the local Vite URL.

## Build for GitHub Pages / static hosting

```bash
npm run build
```

The production files are created in `dist/`.

Because the app uses `base: "./"`, it is suitable for static hosting such as GitHub Pages.

## Privacy

Excel files are parsed in the browser. There is no backend and no API key. Your workbook is not uploaded by this app.

## Supported data

Best results come from a table with a header row and columns such as:
- Date / Invoice Date / Order Date
- Sales / Revenue / Amount
- Quantity / Units
- Product / Customer / Region / Department / Status

The dashboard also works with other numeric and categorical columns and adapts to the detected schema.

## Notes

- For multiple sheets, the app lets you choose the sheet to analyze.
- Date filtering appears automatically when a usable date column is detected.
- The dashboard is responsive for mobile and desktop.
