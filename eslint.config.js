import tseslint from "typescript-eslint";

export default tseslint.config(
  // Base: recommended rules (catches real TS bugs without requiring full type resolution)
  ...tseslint.configs.recommended,

  // Prudent overrides
  {
    rules: {
      // Warn on explicit `any` — needed in API response layers but should be conscious
      "@typescript-eslint/no-explicit-any": "warn",
      // Allow non-null assertions — useful for DOM-like patterns and array access
      "@typescript-eslint/no-non-null-assertion": "warn",
      // Unused vars: error, but allow underscore-prefixed intentional ignores
      "@typescript-eslint/no-unused-vars": [
        "error",
        { argsIgnorePattern: "^_", varsIgnorePattern: "^_" },
      ],
    },
  },

  // Ignore build output and non-source files
  {
    ignores: ["dist/", "coverage/", "node_modules/"],
  }
);
