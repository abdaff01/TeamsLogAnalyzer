# TeamsLogAnalyzer

TeamsLogAnalyzer converts Microsoft Teams Admin Center logs into concise, human-readable descriptions so administrators and support engineers can quickly understand events, diagnose issues, and produce reports.

- Input: logs exported from the Teams Admin Center (JSON, CSV, or other Teams export formats)
- Output: human-friendly text summaries, JSON reports, or simple CSV/TSV summaries suitable for sharing

Features
- Parse Teams Admin Center event logs and translate technical fields into plain-language descriptions
- Summarize common events (call failures, sign-in issues, policy changes, device registration)
- Filter and group events by user, device, event type, time range, or severity
- Export readable reports for non-technical stakeholders or for ticketing systems
- Configurable output formats and verbosity

Quick start

1. Clone the repository
   git clone https://github.com/abdaff01/TeamsLogAnalyzer.git
   cd TeamsLogAnalyzer

2. Install prerequisites
   - The project may require a runtime or dependencies (Node, Python, or similar). Check the repository files (package.json, requirements.txt) for exact requirements and install them.
   - If a package manager is used:
     - Node/npm: npm install
     - Python/venv: python -m venv .venv && source .venv/bin/activate && pip install -r requirements.txt

3. Run the analyzer
   The command below is a generic example. Replace `teams-log-analyzer` with the actual entrypoint (script, binary, or CLI) provided by the repository.

   # Analyze a single log file and print a summary to stdout
   teams-log-analyzer -i teams-logs.json

   # Analyze and write a human-readable Markdown report
   teams-log-analyzer -i teams-logs.json -o report.md --format markdown

   # Filter by user and time range, produce JSON structured report
   teams-log-analyzer -i teams-logs.json -o report.json --format json --user user@example.com --from 2025-01-01 --to 2025-01-31

If the repository provides a different entrypoint (for example, `node index.js`, `python main.py`, or a compiled binary), use that instead:
- Node: node index.js -i teams-logs.json -o report.md
- Python: python main.py -i teams-logs.json -o report.md

Input formats
- TeamsLogAnalyzer expects logs exported from the Teams Admin Center. Common formats are JSON and CSV. The analyzer attempts to detect the input format automatically, but you can specify the format explicitly with `--input-format`.
- If your logs are ZIP archives or packaged exports, extract them first or use the tool's archive option (if available).

Output formats
- Human-readable Markdown or plain text: good for sharing with non-technical stakeholders.
- JSON: structured output suitable for programmatic consumption or ingestion into dashboards.
- CSV/TSV: for spreadsheet analysis and filtering.

Common options (examples)
- -i, --input <path>           Input log file or directory
- -o, --output <path>          Output file (defaults to stdout)
- --format <markdown|text|json|csv>
- --input-format <auto|json|csv>
- --user <email>               Filter by user email or user id
- --event-type <type>          Filter by event type (e.g., Call, SignIn, PolicyChange)
- --from <ISO-date>            Only include events on or after this date
- --to <ISO-date>              Only include events on or before this date
- --verbose                    Include raw event fields in the output
- --help                       Show help and exit

Example summary (human-readable)
- 2026-01-10 14:32 UTC — User alice@example.com — Call failed: ICE negotiation timeout — Device: Yealink T46S — Suggestion: check network port 3478 and STUN/TURN reachability.
- 2026-01-09 09:15 UTC — Policy change — User bob@example.com had CallingPolicy set to Disabled by admin carol@example.com
- 2026-01-08 18:03 UTC — Sign-in failure — User david@example.com — Reason: Invalid credentials — Suggestion: reset password or confirm MFA status

Tips and troubleshooting
- If output is empty, verify that the input file contains Teams Admin Center event records and that the input format option (if used) matches the file.
- For very large logs, try filtering by date range or user to reduce memory use.
- If event names or fields look unfamiliar, open a small sample of the raw log to inspect the schema and match field names (some exports include nested objects or different property names).

Extending or customizing
- Add new translators/mappings for event types to produce clearer messages.
- Add sinks to export results directly to ticketing systems, SIEMs, or dashboards.
- Implement plugins to understand custom or vendor-specific log fields.

Contributing
Contributions are welcome. Typical contribution flow:
1. Fork the repository
2. Create a branch for your change: git checkout -b feature/your-feature
3. Add tests or sample logs if you add parsing logic
4. Open a pull request describing your changes

When opening issues or PRs, include:
- sample log snippets (anonymized)
- the exact command you ran
- the expected vs. actual output

License
Specify a license for this project (e.g., MIT, Apache-2.0). If you haven't chosen one yet, consider adding a LICENSE file.

Contact / Issues
- Report bugs or request features using the repository Issues page: https://github.com/abdaff01/TeamsLogAnalyzer/issues

Acknowledgements
- Built to aid administrators and support teams in interpreting Teams Admin Center logs quickly and accurately.

---

If you want, I can:
- tailor the README with exact install/run instructions after I inspect the repository files (package.json, setup.py, etc.),
- add examples using real sample log snippets (anonymized) from your repo,
- or create a CONTRIBUTING.md and LICENSE file for this project.

Which of those would you like me to do next?
