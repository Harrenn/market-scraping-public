# Contributor instructions

Read docs/project-preferences.md and the relevant architecture, schema and tests before changing scraper behavior. Treat each dealer as a separate source; preserve provenance and do not invent relationships.

This is a public source-only repository. Never commit working workbooks, generated reports, captured state, logs, credentials or browser profiles. Use synthetic fixtures in tests and redact tokens from any evidence intended for sharing.

Run `python -m unittest discover -s tests -v` for offline validation. Live scraping is separate: review prior local evidence and the current website, stage output, compare with the last known-good workbook, investigate anomalies and log the result before promoting data. Do not overwrite known-good datasets with partial or blocked results.
