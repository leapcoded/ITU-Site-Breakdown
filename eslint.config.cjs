module.exports = {
  from: "--config",
  ignores: ["node_modules/**", "dist/**", ".cache/**"],
  languageOptions: {
    ecmaVersion: 2021,
    sourceType: "module",
  },
  plugins: {
    html: require('eslint-plugin-html')
  },
  overrides: [
    {
      files: ["**/*.html"],
      processor: 'html/html-processor'
    }
  ],
  rules: {
    'no-console': 'off',
    'no-unused-vars': ['warn', { args: 'none', ignoreRestSiblings: true }]
  }
};
