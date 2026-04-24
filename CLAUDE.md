# CLAUDE.md

## Outcomes
Deliver a production-safe Google Sheets contract management toolkit that can be configured for any client without exposing private data.

## Repo Boundaries
- Keep this repository template-safe: no client names, no real contract values, no live credentials.
- Use placeholder data for examples and tests.
- Preserve Google Apps Script + Sheets workflow compatibility.

## Primary Entry Points
- `Code.gs`: core contract lifecycle logic, reporting, and alerts.
- `config.gs`: deployment-time branding placeholders.
- `FunctionRunner.gs`: sheet-driven function execution.
- `setup.js`: client setup/deployment wizard.

## Guardrails
- Never commit `.clasp.json`, `.env*`, credential files, or private keys.
- Keep all contact info under TNDS-controlled addresses.
- If adding sample data, use synthetic company names and synthetic values.

## Definition of Done
- README is accurate and copy-paste runnable.
- LICENSE and `.gitignore` remain present.
- Secret scan and PII scan pass before publish.
