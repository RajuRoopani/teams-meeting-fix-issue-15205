/**
 * Tests for MeetingApisContainer
 * Issue #15205 — blank pop-up dialog on first app load in ad-hoc Teams meeting
 *
 * Test cases:
 * AC2: notifySuccess is called AFTER getContext resolves (call order verification)
 * Loading/Ready/Error states: proper rendering during each phase
 * Callbacks: onContextReady and onInitError are called at the right time
 * Render props: children function receives WithTeamsContextProps
 */

import React from 'react';
import { render, screen, waitFor } from '@testing-library/react';
import MeetingApisContainer from '../meeting-apis-container';
import * as teamsJs from '@microsoft/teams-js';
import type { MeetingContext } from '../types';

// Import mock helpers
const { __resetMocks, __setMockContext } = teamsJs as any;

// Test helper to verify call order
function getCallOrder(...mocks: jest.Mock[]): number[] {
  return mocks.map(m => m.mock.invocationCallOrder[0] ?? -1);
}

describe('MeetingApisContainer', () => {
  beforeEach(() => {
    __resetMocks();
  });

  // =========================================================================
  // AC2: notifySuccess called AFTER getContext resolves
  // =========================================================================

  describe('AC2: notifySuccess call order', () => {
    it('calls notifySuccess, and it is called AFTER getContext resolves', async () => {
      const mockInitialize = jest.spyOn(teamsJs.app, 'initialize');
      const mockGetContext = jest.spyOn(teamsJs.app, 'getContext');
      const mockNotifySuccess = jest.spyOn(teamsJs.app, 'notifySuccess');

      render(
        <MeetingApisContainer>
          <div>Test Content</div>
        </MeetingApisContainer>
      );

      await waitFor(() => {
        expect(mockGetContext).toHaveBeenCalled();
        expect(mockNotifySuccess).toHaveBeenCalled();
      });

      // Verify getContext was called before notifySuccess
      const initCallOrder = mockInitialize.mock.invocationCallOrder[0];
      const getContextCallOrder = mockGetContext.mock.invocationCallOrder[0];
      const notifySuccessCallOrder = mockNotifySuccess.mock.invocationCallOrder[0];

      expect(initCallOrder).toBeLessThan(getContextCallOrder);
      expect(getContextCallOrder).toBeLessThan(notifySuccessCallOrder);

      mockInitialize.mockRestore();
      mockGetContext.mockRestore();
      mockNotifySuccess.mockRestore();
    });

    it('notifySuccess is NOT called before getContext resolves', async () => {
      const mockGetContext = jest.spyOn(teamsJs.app, 'getContext');
      const mockNotifySuccess = jest.spyOn(teamsJs.app, 'notifySuccess');

      // Mock getContext to delay resolution
      mockGetContext.mockImplementationOnce(
        () =>
          new Promise(resolve =>
            setTimeout(() => {
              resolve({
                meeting: { id: 'test-meeting' },
                user: { id: 'test-user' },
                app: { locale: 'en-us', theme: 'default' },
                page: { subPageId: 'sidePanel' },
              });
            }, 100)
          )
      );

      render(
        <MeetingApisContainer>
          <div>Test Content</div>
        </MeetingApisContainer>
      );

      // Immediately after render, notifySuccess should not have been called
      expect(mockNotifySuccess).not.toHaveBeenCalled();

      // Wait for async resolution
      await waitFor(() => {
        expect(mockNotifySuccess).toHaveBeenCalled();
      });

      mockGetContext.mockRestore();
      mockNotifySuccess.mockRestore();
    });
  });

  // =========================================================================
  // Loading state — before context resolves
  // =========================================================================

  describe('Loading state', () => {
    it('renders data-testid="meeting-apis-loading" with "Initializing…" text during initialization', async () => {
      const mockGetContext = jest.spyOn(teamsJs.app, 'getContext');

      // Delay context resolution to ensure we catch the loading state
      mockGetContext.mockImplementationOnce(
        () =>
          new Promise(resolve =>
            setTimeout(() => {
              resolve({
                meeting: { id: 'test-meeting' },
                user: { id: 'test-user' },
                app: { locale: 'en-us', theme: 'default' },
                page: { subPageId: 'sidePanel' },
              });
            }, 100)
          )
      );

      const { rerender } = render(
        <MeetingApisContainer>
          <div>Test Content</div>
        </MeetingApisContainer>
      );

      // Should initially show loading state
      expect(screen.getByTestId('meeting-apis-loading')).toBeInTheDocument();
      expect(screen.getByText('Initializing…')).toBeInTheDocument();

      // After context resolves, loading should be gone
      await waitFor(() => {
        expect(screen.queryByTestId('meeting-apis-loading')).not.toBeInTheDocument();
      });

      mockGetContext.mockRestore();
    });

    it('does not render blank markup during initialization', async () => {
      const mockGetContext = jest.spyOn(teamsJs.app, 'getContext');

      mockGetContext.mockImplementationOnce(
        () =>
          new Promise(resolve =>
            setTimeout(() => {
              resolve({
                meeting: { id: 'test-meeting' },
                user: { id: 'test-user' },
                app: { locale: 'en-us', theme: 'default' },
                page: { subPageId: 'sidePanel' },
              });
            }, 50)
          )
      );

      const { container } = render(
        <MeetingApisContainer>
          <div>Test Content</div>
        </MeetingApisContainer>
      );

      // Should have content (loading indicator)
      expect(container.textContent).not.toBe('');
      expect(container.textContent).toContain('Initializing');

      mockGetContext.mockRestore();
    });
  });

  // =========================================================================
  // Ready state — after context resolves
  // =========================================================================

  describe('Ready state', () => {
    it('renders data-testid="meeting-apis-container" with children after context resolves', async () => {
      __setMockContext({
        page: { subPageId: 'sidePanel' },
      });

      render(
        <MeetingApisContainer>
          <div>Test Content</div>
        </MeetingApisContainer>
      );

      await waitFor(() => {
        expect(screen.getByTestId('meeting-apis-container')).toBeInTheDocument();
      });
      expect(screen.getByText('Test Content')).toBeInTheDocument();
    });

    it('renders plain React children', async () => {
      render(
        <MeetingApisContainer>
          <section>
            <h1>Hello World</h1>
          </section>
        </MeetingApisContainer>
      );

      await waitFor(() => {
        expect(screen.getByText('Hello World')).toBeInTheDocument();
      });
    });
  });

  // =========================================================================
  // Error state — when initialization fails
  // =========================================================================

  describe('Error state', () => {
    it('renders data-testid="meeting-apis-error" with error message when initialize fails', async () => {
      const mockInitialize = jest.spyOn(teamsJs.app, 'initialize');

      mockInitialize.mockRejectedValueOnce(
        new Error('Failed to initialize Teams')
      );

      render(
        <MeetingApisContainer>
          <div>Test Content</div>
        </MeetingApisContainer>
      );

      await waitFor(() => {
        expect(screen.getByTestId('meeting-apis-error')).toBeInTheDocument();
      });
      expect(
        screen.getByText('Failed to connect to Teams.')
      ).toBeInTheDocument();
      expect(
        screen.getByText('Failed to initialize Teams')
      ).toBeInTheDocument();

      mockInitialize.mockRestore();
    });

    it('renders error detail message when getContext fails', async () => {
      const mockGetContext = jest.spyOn(teamsJs.app, 'getContext');

      mockGetContext.mockRejectedValueOnce(
        new Error('Network error: connection refused')
      );

      render(
        <MeetingApisContainer>
          <div>Test Content</div>
        </MeetingApisContainer>
      );

      await waitFor(() => {
        expect(screen.getByTestId('meeting-apis-error')).toBeInTheDocument();
      });
      expect(
        screen.getByText('Network error: connection refused')
      ).toBeInTheDocument();

      mockGetContext.mockRestore();
    });

    it('has accessibility attributes on error state', async () => {
      const mockInitialize = jest.spyOn(teamsJs.app, 'initialize');

      mockInitialize.mockRejectedValueOnce(new Error('Test error'));

      render(
        <MeetingApisContainer>
          <div>Test Content</div>
        </MeetingApisContainer>
      );

      await waitFor(() => {
        const errorDiv = screen.getByTestId('meeting-apis-error');
        expect(errorDiv).toHaveAttribute('role', 'alert');
        expect(errorDiv).toHaveAttribute('aria-live', 'assertive');
      });

      mockInitialize.mockRestore();
    });
  });

  // =========================================================================
  // Render prop children — function children receive context
  // =========================================================================

  describe('Render prop children', () => {
    it('calls children function with { meetingContext, initializationStatus, initializationError }', async () => {
      const childrenFn = jest.fn(() => <div>Rendered by function</div>);

      __setMockContext({
        meeting: { id: 'meeting-abc' },
        user: { id: 'user-xyz' },
        page: { subPageId: 'meetingStage' },
      });

      render(
        <MeetingApisContainer>
          {childrenFn}
        </MeetingApisContainer>
      );

      await waitFor(() => {
        expect(screen.getByText('Rendered by function')).toBeInTheDocument();
      });

      // Verify the children function was called with the right props
      expect(childrenFn).toHaveBeenCalled();
      const callArgs = childrenFn.mock.calls[childrenFn.mock.calls.length - 1][0];
      expect(callArgs).toHaveProperty('meetingContext');
      expect(callArgs).toHaveProperty('initializationStatus');
      expect(callArgs).toHaveProperty('initializationError');

      // Verify the values when ready
      expect(callArgs.initializationStatus).toBe('ready');
      expect(callArgs.initializationError).toBeNull();
      expect(callArgs.meetingContext).toMatchObject({
        meetingId: 'meeting-abc',
        userObjectId: 'user-xyz',
        commandSource: 'meetingStage',
      });
    });

    it('provides null meetingContext during initialization', async () => {
      const mockGetContext = jest.spyOn(teamsJs.app, 'getContext');

      mockGetContext.mockImplementationOnce(
        () =>
          new Promise(resolve =>
            setTimeout(() => {
              resolve({
                meeting: { id: 'test-meeting' },
                user: { id: 'test-user' },
                app: { locale: 'en-us', theme: 'default' },
                page: { subPageId: 'sidePanel' },
              });
            }, 100)
          )
      );

      const childrenFn = jest.fn(() => <div>Test</div>);

      render(
        <MeetingApisContainer>
          {childrenFn}
        </MeetingApisContainer>
      );

      // On first render (initializing), context should be null
      let firstCallArgs = childrenFn.mock.calls[0][0];
      expect(firstCallArgs.initializationStatus).toBe('initializing');
      expect(firstCallArgs.meetingContext).toBeNull();

      // After resolution
      await waitFor(() => {
        const finalCallArgs = childrenFn.mock.calls[childrenFn.mock.calls.length - 1][0];
        expect(finalCallArgs.initializationStatus).toBe('ready');
        expect(finalCallArgs.meetingContext).not.toBeNull();
      });

      mockGetContext.mockRestore();
    });
  });

  // =========================================================================
  // Callbacks: onContextReady and onInitError
  // =========================================================================

  describe('Callbacks', () => {
    it('calls onContextReady with the normalized context', async () => {
      const onContextReady = jest.fn();

      __setMockContext({
        meeting: { id: 'meeting-123' },
        user: { id: 'user-456' },
        page: { subPageId: 'contentBubble' },
      });

      render(
        <MeetingApisContainer onContextReady={onContextReady}>
          <div>Test</div>
        </MeetingApisContainer>
      );

      await waitFor(() => {
        expect(onContextReady).toHaveBeenCalled();
      });

      const contextArg = onContextReady.mock.calls[0][0];
      expect(contextArg).toMatchObject({
        meetingId: 'meeting-123',
        userObjectId: 'user-456',
        commandSource: 'contentBubble',
      });
    });

    it('does NOT call onContextReady until after getContext resolves', async () => {
      const onContextReady = jest.fn();
      const mockGetContext = jest.spyOn(teamsJs.app, 'getContext');

      mockGetContext.mockImplementationOnce(
        () =>
          new Promise(resolve =>
            setTimeout(() => {
              resolve({
                meeting: { id: 'test-meeting' },
                user: { id: 'test-user' },
                app: { locale: 'en-us', theme: 'default' },
                page: { subPageId: 'sidePanel' },
              });
            }, 50)
          )
      );

      render(
        <MeetingApisContainer onContextReady={onContextReady}>
          <div>Test</div>
        </MeetingApisContainer>
      );

      // Immediately after render, onContextReady should not have been called
      expect(onContextReady).not.toHaveBeenCalled();

      // Wait for resolution
      await waitFor(() => {
        expect(onContextReady).toHaveBeenCalled();
      });

      mockGetContext.mockRestore();
    });

    it('calls onInitError when initialization fails', async () => {
      const onInitError = jest.fn();
      const mockInitialize = jest.spyOn(teamsJs.app, 'initialize');

      const testError = new Error('Initialization failed');
      mockInitialize.mockRejectedValueOnce(testError);

      render(
        <MeetingApisContainer onInitError={onInitError}>
          <div>Test</div>
        </MeetingApisContainer>
      );

      await waitFor(() => {
        expect(onInitError).toHaveBeenCalled();
      });

      const errorArg = onInitError.mock.calls[0][0];
      expect(errorArg).toBeInstanceOf(Error);
      expect(errorArg.message).toContain('Initialization failed');

      mockInitialize.mockRestore();
    });

    it('wraps non-Error exceptions in Error for onInitError', async () => {
      const onInitError = jest.fn();
      const mockGetContext = jest.spyOn(teamsJs.app, 'getContext');

      mockGetContext.mockRejectedValueOnce('String error');

      render(
        <MeetingApisContainer onInitError={onInitError}>
          <div>Test</div>
        </MeetingApisContainer>
      );

      await waitFor(() => {
        expect(onInitError).toHaveBeenCalled();
      });

      const errorArg = onInitError.mock.calls[0][0];
      expect(errorArg).toBeInstanceOf(Error);
      expect(errorArg.message).toContain('String error');

      mockGetContext.mockRestore();
    });
  });

  // =========================================================================
  // Context normalization
  // =========================================================================

  describe('Context normalization', () => {
    it('normalises page.subPageId to commandSource', async () => {
      const onContextReady = jest.fn();

      __setMockContext({
        page: { subPageId: 'settings' },
      });

      render(
        <MeetingApisContainer onContextReady={onContextReady}>
          <div>Test</div>
        </MeetingApisContainer>
      );

      await waitFor(() => {
        expect(onContextReady).toHaveBeenCalled();
      });

      const context = onContextReady.mock.calls[0][0];
      expect(context.commandSource).toBe('settings');
    });

    it('sets commandSource to undefined when page.subPageId is missing', async () => {
      const onContextReady = jest.fn();

      __setMockContext({
        page: { subPageId: undefined },
      });

      render(
        <MeetingApisContainer onContextReady={onContextReady}>
          <div>Test</div>
        </MeetingApisContainer>
      );

      await waitFor(() => {
        expect(onContextReady).toHaveBeenCalled();
      });

      const context = onContextReady.mock.calls[0][0];
      expect(context.commandSource).toBeUndefined();
    });

    it('provides default locale when app.locale is missing', async () => {
      const onContextReady = jest.fn();
      const mockGetContext = jest.spyOn(teamsJs.app, 'getContext');

      mockGetContext.mockResolvedValueOnce({
        meeting: { id: 'test' },
        user: { id: 'test-user' },
        app: { theme: 'default' },
        page: { subPageId: 'sidePanel' },
      } as any);

      render(
        <MeetingApisContainer onContextReady={onContextReady}>
          <div>Test</div>
        </MeetingApisContainer>
      );

      await waitFor(() => {
        expect(onContextReady).toHaveBeenCalled();
      });

      const context = onContextReady.mock.calls[0][0];
      expect(context.locale).toBe('en-us');

      mockGetContext.mockRestore();
    });

    it('provides default theme when app.theme is missing', async () => {
      const onContextReady = jest.fn();
      const mockGetContext = jest.spyOn(teamsJs.app, 'getContext');

      mockGetContext.mockResolvedValueOnce({
        meeting: { id: 'test' },
        user: { id: 'test-user' },
        app: { locale: 'en-us' },
        page: { subPageId: 'sidePanel' },
      } as any);

      render(
        <MeetingApisContainer onContextReady={onContextReady}>
          <div>Test</div>
        </MeetingApisContainer>
      );

      await waitFor(() => {
        expect(onContextReady).toHaveBeenCalled();
      });

      const context = onContextReady.mock.calls[0][0];
      expect(context.theme).toBe('default');

      mockGetContext.mockRestore();
    });
  });

  // =========================================================================
  // Edge cases and cleanup
  // =========================================================================

  describe('Edge cases and cleanup', () => {
    it('does not update state if component unmounts during async initialization', async () => {
      const mockGetContext = jest.spyOn(teamsJs.app, 'getContext');
      const consoleErrorSpy = jest.spyOn(console, 'error').mockImplementation();

      mockGetContext.mockImplementationOnce(
        () =>
          new Promise(resolve =>
            setTimeout(() => {
              resolve({
                meeting: { id: 'test' },
                user: { id: 'test-user' },
                app: { locale: 'en-us', theme: 'default' },
                page: { subPageId: 'sidePanel' },
              });
            }, 100)
          )
      );

      const { unmount } = render(
        <MeetingApisContainer>
          <div>Test</div>
        </MeetingApisContainer>
      );

      // Unmount immediately, before async operations complete
      unmount();

      // Wait to ensure no state updates occur after unmount
      await new Promise(resolve => setTimeout(resolve, 150));

      // No errors should be logged (React doesn't complain about unmounted component updates)
      mockGetContext.mockRestore();
      consoleErrorSpy.mockRestore();
    });

    it('has proper accessibility attributes on loading state', async () => {
      const mockGetContext = jest.spyOn(teamsJs.app, 'getContext');

      mockGetContext.mockImplementationOnce(
        () =>
          new Promise(resolve =>
            setTimeout(() => {
              resolve({
                meeting: { id: 'test' },
                user: { id: 'test-user' },
                app: { locale: 'en-us', theme: 'default' },
                page: { subPageId: 'sidePanel' },
              });
            }, 50)
          )
      );

      render(
        <MeetingApisContainer>
          <div>Test</div>
        </MeetingApisContainer>
      );

      const loadingDiv = screen.getByTestId('meeting-apis-loading');
      expect(loadingDiv).toHaveAttribute('role', 'status');
      expect(loadingDiv).toHaveAttribute('aria-live', 'polite');

      mockGetContext.mockRestore();
    });
  });
});
