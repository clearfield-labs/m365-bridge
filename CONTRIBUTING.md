# Contributing

Thanks for helping. Issues and pull requests are welcome.

1. `npm install`, then `npm test`. Tests run against a mock Microsoft Graph (`test/mock-graph.js`), so you need no tenant.
2. `npm run demo` starts the real page against a fake Microsoft 365 at http://localhost:4366.
3. Keep changes small and add a test for each behaviour you change. Anything that writes to SharePoint needs a test that proves it cannot overwrite a colleague's work.
4. Add a line to `changelog.md`.

Please report security issues privately (see [SECURITY.md](SECURITY.md)).
