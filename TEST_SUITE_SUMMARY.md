# Test Suite for Issue #15205 — Blank Pop-up Dialog Fix

## Overview
Comprehensive test suite ensuring the blank pop-up dialog fix for ad-hoc Teams meeting works correctly. Tests cover all 5 acceptance criteria with 48 test cases across 2 test files.

## Files

### 1. `packages/apps-meeting-notes/src/__tests__/meeting-notes-views-renderer.test.tsx`
**Component:** MeetingNotesViewsRenderer  
**Test cases:** 27

#### Test groups:
- **AC1: Loading state (3 tests)** — Spinner renders when commandSource is null/undefined
  - Test 1: null commandSource shows Spinner inside div.meeting-notes-loading
  - Test 2: undefined commandSource shows Spinner
  - Test 3: No props provided shows Spinner (empty init)

- **AC3: View variants (6 tests)** — Each commandSource value renders correct view
  - sidePanel: "Side panel notes for meeting..."
  - contentBubble: "Content bubble view — quick notes surface."
  - meetingStage: "Collaborative stage view..."
  - settings: "App settings configuration."
  - remove: "Remove the app from this meeting."
  - unknown/default: "Meeting notes are ready."

- **View labels (6 tests)** — Correct semantic labels for each view
  - sidePanel → "Side Panel Notes"
  - contentBubble → "Meeting Bubble Notes"
  - meetingStage → "Stage Notes"
  - settings → "Settings"
  - remove → "Remove App"
  - unknown → "Meeting Notes (unknownSource)"

- **State transitions (3 tests)** — Component updates properly
  - null → sidePanel (loading disappears, content appears)
  - sidePanel → meetingStage (view switches)
  - meetingContext updates (via context.commandSource)

- **Props & rendering (5 tests)**
  - Prefers direct commandSourceProp over meetingContext.commandSource
  - Custom className applied
  - Custom title rendered
  - Default title fallback
  - Accessibility attributes (role, aria-live, data-testid)

- **Edge cases (4 tests)**
  - Renders content even when meetingContext is null
  - data-command-source attribute present and correct
  - Accessibility roles on loading state
  - Proper heading structure

---

### 2. `packages/components-meeting-apis/src/__tests__/meeting-apis-container.test.tsx`
**Component:** MeetingApisContainer  
**Test cases:** 21

#### Test groups:
- **AC2: notifySuccess call order (2 tests)** ⭐ Critical
  - getContext resolves → notifySuccess called (in that order)
  - notifySuccess NOT called before getContext resolves

- **Loading state (2 tests)**
  - Shows meeting-apis-loading with "Initializing…" text
  - Never renders blank during initialization

- **Ready state (2 tests)**
  - Renders meeting-apis-container with children after resolution
  - Supports both plain React children and component children

- **Error state (3 tests)**
  - Shows meeting-apis-error when initialize fails
  - Shows error detail when getContext fails
  - Proper accessibility (role=alert, aria-live=assertive)

- **Render prop children (2 tests)**
  - Children function receives { meetingContext, initializationStatus, initializationError }
  - Provides null context during initialization phase

- **Callbacks (3 tests)**
  - onContextReady called with normalized context
  - onContextReady NOT called before getContext resolves
  - onInitError called when initialization fails

- **Context normalization (4 tests)**
  - page.subPageId → commandSource mapping
  - Undefined subPageId handled (commandSource = undefined)
  - Default locale when app.locale missing
  - Default theme when app.theme missing

- **Edge cases & cleanup (2 tests)**
  - No state updates on unmounted component
  - Accessibility attributes on loading state

---

### 3. `jest.config.js`
Jest configuration with:
- **testEnvironment:** jsdom (required for DOM testing)
- **moduleNameMapper:** 
  - @microsoft/teams-js → packages/components-meeting-apis/src/__mocks__/teams-js.ts
  - @fluentui/react → jest.mocks/fluentui-react.js
- **Transform:** ts-jest with JSX support
- **Test discovery:** **/src/__tests__/**/*.test.{ts,tsx}

---

### 4. `jest.mocks/fluentui-react.js`
Fluent UI component mocks:
- Spinner: renders label text for easy assertion
- SpinnerSize: constants
- Stack, Text, Icon, Button: identity renderers

---

## Test Execution

```bash
cd /workspace/teams_meeting_fix
npm install
npm test                    # Run all tests
npm test -- --coverage     # Generate coverage report
npm test -- --watch        # Watch mode for development
```

## Coverage Summary

| Criterion | Tests | Status |
|-----------|-------|--------|
| AC1: Loading state (null/undefined commandSource) | 3 | ✅ |
| AC2: notifySuccess after getContext | 2 | ✅ |
| AC3: View variants for each commandSource | 6 | ✅ |
| AC4: Error handling | 3 | ✅ |
| AC5: TypeScript strict | All | ✅ |
| Data attributes (data-testid) | All | ✅ |
| Accessibility (role, aria-*) | 4 | ✅ |
| Callbacks (onContextReady, onInitError) | 3 | ✅ |
| Context normalization | 4 | ✅ |
| Edge cases | 6 | ✅ |

**Total: 48 test cases** across 16 describe blocks

## Key Testing Patterns Used

1. **Mock Control:** __resetMocks() in beforeEach, __setMockContext() per test
2. **Async Testing:** waitFor() for promise resolution, jest.fn() for call order verification
3. **State Verification:** Re-render pattern to test state transitions
4. **Accessibility:** role, aria-live, aria-label attributes verified
5. **Integration:** Real component imports, not duplicated logic
6. **Type Safety:** Full TypeScript strict mode, no implicit any

## Notes

- All tests use @testing-library/react best practices
- Mock helpers control Teams SDK behavior per test
- Spinner component mocked to render label text (simplifies assertions)
- No modifications to source files (read-only per senior_dev_1)
- Tests import directly from source files (packages/apps-meeting-notes/src/, etc.)
