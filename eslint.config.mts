import tseslint from 'typescript-eslint';
import obsidianmd from "eslint-plugin-obsidianmd";
import globals from "globals";
import { globalIgnores } from "eslint/config";

// Extract the @microsoft/sdl plugin object bundled inside obsidianmd recommended
// so we can reference it in our rule-severity overrides.
const _sdlPlugin = [...obsidianmd.configs.recommended]
	.find(c => c.plugins?.['@microsoft/sdl'])
	?.plugins['@microsoft/sdl'];

export default tseslint.config(
	{
		languageOptions: {
			globals: {
				...globals.browser,
			},
			parserOptions: {
				projectService: {
					allowDefaultProject: [
						'eslint.config.js',
						'manifest.json'
					]
				},
				tsconfigRootDir: import.meta.dirname,
				extraFileExtensions: ['.json']
			},
		},
	},
	...obsidianmd.configs.recommended,
	{
		files: ['**/*.ts', '**/*.tsx'],
		plugins: {
			'@typescript-eslint': tseslint.plugin,
		},
		rules: {
			'no-console': 'error',
			// Downgraded to warn: these stem from interacting with untyped
			// third-party libraries (pptxviewjs, pdfjs-dist) and will be
			// addressed incrementally.
			'@typescript-eslint/no-explicit-any': 'warn',
			'@typescript-eslint/no-unsafe-assignment': 'warn',
			'@typescript-eslint/no-unsafe-member-access': 'warn',
			'@typescript-eslint/no-unsafe-call': 'warn',
			'@typescript-eslint/no-unsafe-argument': 'warn',
			'@typescript-eslint/no-unsafe-return': 'warn',
			// execCommand has no modern replacement for undo/redo in
			// contentEditable — keep as warning until Clipboard API adoption.
			'@typescript-eslint/no-deprecated': 'warn',
		},
	},
	// Rule severity overrides for plugins provided by obsidianmd recommended.
	// Plugins must be re-declared here because tseslint.config() scopes them
	// to the config object where they are first defined.
	{
		plugins: {
			obsidianmd,
			...(_sdlPlugin ? { '@microsoft/sdl': _sdlPlugin } : {}),
		},
		rules: {
			// Dynamic style manipulation is integral to the plugin UI.
			'obsidianmd/no-static-styles-assignment': 'warn',
			// Acronyms (PDF, CSV) are flagged by sentence-case but are correct.
			'obsidianmd/ui/sentence-case': 'warn',
			// innerHTML usage is required for the DOCX viewer rendering pipeline.
			'@microsoft/sdl/no-inner-html': 'warn',
		},
	},
	globalIgnores([
		"node_modules",
		"dist",
		"esbuild.config.mjs",
		"eslint.config.js",
		"version-bump.mjs",
		"versions.json",
		"main.js",
		"scripts/check-css.js",
	]),
);
