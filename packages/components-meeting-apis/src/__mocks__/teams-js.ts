/**
 * Manual mock for @microsoft/teams-js
 *
 * Used by jest via moduleNameMapper to replace the real SDK in tests.
 * Each method is a jest.fn() so tests can spy and assert call order,
 * which is essential for verifying AC2 (notifySuccess after getContext).
 */

const mockContext = {
  meeting: { id: 'mock-meeting-id' },
  user: { id: 'mock-user-object-id' },
  app: {
    locale: 'en-us',
    theme: 'default',
  },
  page: {
    subPageId: undefined as string | undefined,
  },
};

export const app = {
  initialize: jest.fn(() => Promise.resolve()),
  getContext: jest.fn(() => Promise.resolve(mockContext)),
  notifySuccess: jest.fn(),
  notifyFailure: jest.fn(),
  notifyAppLoaded: jest.fn(),
};

/** Allow tests to customise the resolved context per-test. */
export function __setMockContext(
  overrides: Partial<typeof mockContext>
): void {
  Object.assign(mockContext, overrides);
  // Also handle nested page overrides
  if ('page' in overrides && overrides.page) {
    Object.assign(mockContext.page, overrides.page);
  }
}

/** Reset all mocks to their default state between tests. */
export function __resetMocks(): void {
  app.initialize.mockClear();
  app.getContext.mockClear();
  app.notifySuccess.mockClear();
  app.notifyFailure.mockClear();
  app.notifyAppLoaded.mockClear();
  mockContext.page.subPageId = undefined;
}
