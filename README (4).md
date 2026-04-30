# Tournament Gantt Planner

Tournament Gantt Planner turns a tournament schedule spreadsheet into a prep timeline. It helps debate teams see when to confirm teams, book travel, and submit budget requests for each tournament.

## What it does

- Upload an Excel or CSV schedule.
- Match your spreadsheet columns to tournament name, date, location, transport, and number of debaters.
- Automatically calculate prep tasks before each tournament.
- Show the schedule as both a table and a visual Gantt chart.
- Export the finished plan as an Excel workbook.

## Spreadsheet format

The app works best with these columns:

- Tournament Name
- Date
- Location
- Transport
- Debaters

Only tournament name and date are required. You can also download a blank template from the app.

## How to use it

1. Open `index.html` in a browser.
2. Set the budget request deadline and lead time.
3. Upload your spreadsheet.
4. Review or adjust the column mapping.
5. Check the preview and visual timeline.
6. Export the Gantt chart to Excel.

## Running locally

You can open `index.html` directly, or run a local server:

```bash
npm run dev
```

The app uses browser-based libraries loaded from CDNs for spreadsheet handling and chart rendering, so an internet connection is needed for the full experience.
