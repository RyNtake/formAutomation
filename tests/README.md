# Local GAS mock tests

These tests run in VS Code's integrated terminal with Node.js 18 or newer:

```sh
node --test tests/*.test.cjs
```

They use Node's built-in test runner and mock the Apps Script services they need.
They do not connect to Google Sheets, Drive, Discord, or any other external service.
The `.cjs` test files are not among the `.js`/`.gs` extensions configured for clasp.

The existing `test.gs.js` is an Apps Script function that accesses a real Drive
folder; it is separate from this local mock test suite.
