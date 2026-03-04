/**
 * Jest configuration for teams-meeting-fix
 * Issue #15205 — blank pop-up dialog on first app load in ad-hoc Teams meeting
 *
 * Configuration:
 * - testEnvironment: jsdom (required for DOM testing with @testing-library/react)
 * - Module mapper for @microsoft/teams-js → mock
 * - Module mapper for @fluentui/react → identity mock (Spinner renders its label)
 * - TypeScript support via ts-jest with JSX transformation
 */

module.exports = {
  preset: 'ts-jest',
  testEnvironment: 'jsdom',
  setupFilesAfterFramework: ['@testing-library/jest-dom'],

  moduleNameMapper: {
    // Map Teams JS SDK to the manual mock
    '^@microsoft/teams-js$':
      '<rootDir>/packages/components-meeting-apis/src/__mocks__/teams-js.ts',

    // Map Fluent UI components to a simple identity mock
    // Allows tests to work without mocking each component individually
    '^@fluentui/react$': '<rootDir>/jest.mocks/fluentui-react.js',
  },

  testMatch: [
    '**/src/__tests__/**/*.test.{ts,tsx}',
    '**/src/__tests__/**/*.spec.{ts,tsx}',
  ],

  transform: {
    '^.+\\.(ts|tsx)$': [
      'ts-jest',
      {
        tsconfig: {
          jsx: 'react-jsx',
          esModuleInterop: true,
          allowSyntheticDefaultImports: true,
        },
      },
    ],
  },

  moduleFileExtensions: ['ts', 'tsx', 'js', 'jsx', 'json', 'node'],

  collectCoverageFrom: [
    'packages/*/src/**/*.{ts,tsx}',
    '!packages/*/src/**/*.d.ts',
    '!packages/*/src/__tests__/**',
    '!packages/*/src/__mocks__/**',
  ],

  testPathIgnorePatterns: ['/node_modules/', '/dist/'],
  transformIgnorePatterns: ['node_modules/(?!(@microsoft|@fluentui)/)'],
};
