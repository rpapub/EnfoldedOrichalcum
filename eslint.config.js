export default [
  {
    files: ["docs/addins/**/*.js", "docs/shared/*.js"],
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: "script",
      globals: {
        // Browser
        window: "readonly", document: "readonly", console: "readonly",
        fetch: "readonly", localStorage: "readonly", sessionStorage: "readonly",
        URL: "readonly", URLSearchParams: "readonly", crypto: "readonly",
        navigator: "readonly", setTimeout: "readonly", location: "readonly",
        module: "readonly",
        // Office.js
        Office: "readonly",
        // graph-utils.js globals (loaded before each taskpane.js)
        getCachedToken: "readonly", cacheToken: "readonly",
        clearCachedToken: "readonly", graphError: "readonly",
        readGraphExtension: "readonly", writeGraphExtension: "readonly",
        deleteGraphExtension: "readonly",
      }
    },
    rules: {
      "no-unused-vars": "warn",
      "no-undef": "error",
    }
  }
];
