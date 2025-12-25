module.exports = {
	preset: 'ts-jest',
	testEnvironment: 'node',
	roots: ['<rootDir>'],
	testMatch: ['**/__tests__/**/*.ts', '**/?(*.)+(spec|test).ts'],
	transform: {
		'^.+\\.ts$': ['ts-jest', {
			tsconfig: 'tsconfig.test.json',
		}],
	},
	collectCoverageFrom: [
		'nodes/**/*.ts',
		'!nodes/**/*.node.ts',
		'!**/*.d.ts',
		'!**/*.test.ts',
	],
	coverageDirectory: 'coverage',
	coverageReporters: ['text', 'lcov', 'html'],
	moduleFileExtensions: ['ts', 'js', 'json'],
	testTimeout: 10000,
	verbose: true,
};
