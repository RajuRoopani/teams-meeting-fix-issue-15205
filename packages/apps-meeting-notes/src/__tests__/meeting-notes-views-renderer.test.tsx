/**
 * Tests for MeetingNotesViewsRenderer
 * Issue #15205 — blank pop-up dialog on first app load in ad-hoc Teams meeting
 *
 * Test cases:
 * AC1 (Loading state): when commandSource is null/undefined, render Spinner
 * AC3 (View variants): render correct content for each commandSource value
 * Transitions: state changes from loading to loaded are handled correctly
 */

import React from 'react';
import { render, screen, waitFor } from '@testing-library/react';
import MeetingNotesViewsRenderer from '../meeting-notes-views-renderer';
import type { MeetingContext, CommandSource } from '../../components-meeting-apis/src/types';

// Mock Fluent UI Spinner to simplify testing (render just the label)
jest.mock('@fluentui/react', () => ({
  Spinner: ({ label, size, ariaLabel }: any) => (
    <div data-testid="spinner" role="status" aria-label={ariaLabel}>
      {label}
    </div>
  ),
  SpinnerSize: { medium: 'medium' },
}));

describe('MeetingNotesViewsRenderer', () => {
  // =========================================================================
  // AC1: Loading state — null commandSource
  // =========================================================================

  describe('AC1: Loading state when commandSource is null', () => {
    it('renders a Spinner inside div.meeting-notes-loading when commandSource is null', () => {
      render(
        <MeetingNotesViewsRenderer
          commandSource={null as unknown as CommandSource}
        />
      );

      const loadingDiv = screen.getByTestId('meeting-notes-loading');
      expect(loadingDiv).toBeInTheDocument();
      expect(loadingDiv).toHaveClass('meeting-notes-loading');

      const spinner = screen.getByTestId('spinner');
      expect(spinner).toBeInTheDocument();
      expect(spinner).toHaveTextContent('Loading meeting notes');
    });

    it('renders a Spinner when commandSource is undefined', () => {
      render(
        <MeetingNotesViewsRenderer
          commandSource={undefined}
        />
      );

      const loadingDiv = screen.getByTestId('meeting-notes-loading');
      expect(loadingDiv).toBeInTheDocument();

      const spinner = screen.getByTestId('spinner');
      expect(spinner).toBeInTheDocument();
    });

    it('renders a Spinner when no commandSource is provided at all', () => {
      render(<MeetingNotesViewsRenderer />);

      const loadingDiv = screen.getByTestId('meeting-notes-loading');
      expect(loadingDiv).toBeInTheDocument();

      const spinner = screen.getByTestId('spinner');
      expect(spinner).toBeInTheDocument();
    });

    it('does NOT return null or empty JSX when commandSource is null', () => {
      const { container } = render(
        <MeetingNotesViewsRenderer commandSource={null as unknown as CommandSource} />
      );

      // Container should have at least one element (the loading div)
      expect(container.children.length).toBeGreaterThan(0);
      expect(container.textContent).not.toBe('');
    });
  });

  // =========================================================================
  // AC3: View variants — each commandSource value renders its view
  // =========================================================================

  describe('AC3: View variants for initialized commandSource', () => {
    it('renders sidePanel view when commandSource is "sidePanel"', () => {
      const mockContext: MeetingContext = {
        meetingId: 'meeting-123',
        userObjectId: 'user-456',
        locale: 'en-us',
        commandSource: 'sidePanel',
        theme: 'default',
      };

      render(
        <MeetingNotesViewsRenderer
          commandSource="sidePanel"
          meetingContext={mockContext}
        />
      );

      expect(screen.getByTestId('meeting-notes-content')).toBeInTheDocument();
      expect(
        screen.getByText(/Side panel notes for meeting/)
      ).toBeInTheDocument();
      expect(screen.queryByTestId('spinner')).not.toBeInTheDocument();
    });

    it('renders contentBubble view when commandSource is "contentBubble"', () => {
      render(
        <MeetingNotesViewsRenderer
          commandSource="contentBubble"
        />
      );

      expect(screen.getByTestId('meeting-notes-content')).toBeInTheDocument();
      expect(
        screen.getByText(/Content bubble view/)
      ).toBeInTheDocument();
    });

    it('renders meetingStage view when commandSource is "meetingStage"', () => {
      render(
        <MeetingNotesViewsRenderer
          commandSource="meetingStage"
        />
      );

      expect(screen.getByTestId('meeting-notes-content')).toBeInTheDocument();
      expect(
        screen.getByText(/Collaborative stage view/)
      ).toBeInTheDocument();
    });

    it('renders settings view when commandSource is "settings"', () => {
      render(
        <MeetingNotesViewsRenderer
          commandSource="settings"
        />
      );

      expect(screen.getByTestId('meeting-notes-content')).toBeInTheDocument();
      expect(
        screen.getByText(/App settings configuration/)
      ).toBeInTheDocument();
    });

    it('renders remove view when commandSource is "remove"', () => {
      render(
        <MeetingNotesViewsRenderer
          commandSource="remove"
        />
      );

      expect(screen.getByTestId('meeting-notes-content')).toBeInTheDocument();
      expect(
        screen.getByText(/Remove the app from this meeting/)
      ).toBeInTheDocument();
    });

    it('renders default view when commandSource is an unknown string', () => {
      render(
        <MeetingNotesViewsRenderer
          commandSource="unknownSource"
        />
      );

      expect(screen.getByTestId('meeting-notes-content')).toBeInTheDocument();
      expect(
        screen.getByText(/Meeting notes are ready/)
      ).toBeInTheDocument();
    });
  });

  // =========================================================================
  // View label test — AC3 semantic verification
  // =========================================================================

  describe('View label rendering', () => {
    it('displays correct view label for sidePanel', () => {
      render(
        <MeetingNotesViewsRenderer
          commandSource="sidePanel"
        />
      );

      const viewLabel = screen.getByTestId('meeting-notes-view-label');
      expect(viewLabel).toHaveTextContent('Side Panel Notes');
    });

    it('displays correct view label for contentBubble', () => {
      render(
        <MeetingNotesViewsRenderer
          commandSource="contentBubble"
        />
      );

      const viewLabel = screen.getByTestId('meeting-notes-view-label');
      expect(viewLabel).toHaveTextContent('Meeting Bubble Notes');
    });

    it('displays correct view label for meetingStage', () => {
      render(
        <MeetingNotesViewsRenderer
          commandSource="meetingStage"
        />
      );

      const viewLabel = screen.getByTestId('meeting-notes-view-label');
      expect(viewLabel).toHaveTextContent('Stage Notes');
    });

    it('displays correct view label for settings', () => {
      render(
        <MeetingNotesViewsRenderer
          commandSource="settings"
        />
      );

      const viewLabel = screen.getByTestId('meeting-notes-view-label');
      expect(viewLabel).toHaveTextContent('Settings');
    });

    it('displays correct view label for remove', () => {
      render(
        <MeetingNotesViewsRenderer
          commandSource="remove"
        />
      );

      const viewLabel = screen.getByTestId('meeting-notes-view-label');
      expect(viewLabel).toHaveTextContent('Remove App');
    });

    it('displays fallback label for unknown commandSource', () => {
      render(
        <MeetingNotesViewsRenderer
          commandSource="unknownSource"
        />
      );

      const viewLabel = screen.getByTestId('meeting-notes-view-label');
      expect(viewLabel).toHaveTextContent('Meeting Notes (unknownSource)');
    });
  });

  // =========================================================================
  // Transition test — from loading to loaded state
  // =========================================================================

  describe('State transitions', () => {
    it('transitions from loading (null commandSource) to content view (sidePanel)', async () => {
      const { rerender } = render(
        <MeetingNotesViewsRenderer commandSource={null as unknown as CommandSource} />
      );

      // Should show loading state
      expect(screen.getByTestId('meeting-notes-loading')).toBeInTheDocument();
      expect(screen.getByTestId('spinner')).toBeInTheDocument();

      // Update to initialized state
      rerender(
        <MeetingNotesViewsRenderer commandSource="sidePanel" />
      );

      // Loading should be gone, content should appear
      await waitFor(() => {
        expect(screen.queryByTestId('meeting-notes-loading')).not.toBeInTheDocument();
        expect(screen.queryByTestId('spinner')).not.toBeInTheDocument();
      });
      expect(screen.getByTestId('meeting-notes-content')).toBeInTheDocument();
      expect(
        screen.getByText(/Side panel notes for meeting/)
      ).toBeInTheDocument();
    });

    it('transitions between different view types (sidePanel → meetingStage)', async () => {
      const { rerender } = render(
        <MeetingNotesViewsRenderer commandSource="sidePanel" />
      );

      expect(
        screen.getByText(/Side panel notes for meeting/)
      ).toBeInTheDocument();

      rerender(
        <MeetingNotesViewsRenderer commandSource="meetingStage" />
      );

      await waitFor(() => {
        expect(
          screen.queryByText(/Side panel notes for meeting/)
        ).not.toBeInTheDocument();
      });
      expect(
        screen.getByText(/Collaborative stage view/)
      ).toBeInTheDocument();
    });

    it('handles meetingContext prop updates in addition to direct commandSource prop', async () => {
      const context1: MeetingContext = {
        meetingId: 'meeting-1',
        userObjectId: 'user-1',
        locale: 'en-us',
        commandSource: 'sidePanel',
        theme: 'default',
      };

      const context2: MeetingContext = {
        ...context1,
        commandSource: 'meetingStage',
      };

      const { rerender } = render(
        <MeetingNotesViewsRenderer meetingContext={context1} />
      );

      expect(
        screen.getByText(/Side panel notes for meeting/)
      ).toBeInTheDocument();

      rerender(
        <MeetingNotesViewsRenderer meetingContext={context2} />
      );

      await waitFor(() => {
        expect(
          screen.queryByText(/Side panel notes for meeting/)
        ).not.toBeInTheDocument();
      });
      expect(
        screen.getByText(/Collaborative stage view/)
      ).toBeInTheDocument();
    });
  });

  // =========================================================================
  // Props and rendering behavior
  // =========================================================================

  describe('Component props and rendering', () => {
    it('prefers commandSource prop over meetingContext.commandSource', () => {
      const context: MeetingContext = {
        meetingId: 'meeting-1',
        userObjectId: 'user-1',
        locale: 'en-us',
        commandSource: 'sidePanel',
        theme: 'default',
      };

      render(
        <MeetingNotesViewsRenderer
          commandSource="meetingStage"
          meetingContext={context}
        />
      );

      // Should use commandSourceProp, not context's sidePanel
      expect(
        screen.getByText(/Collaborative stage view/)
      ).toBeInTheDocument();
      expect(
        screen.queryByText(/Side panel notes for meeting/)
      ).not.toBeInTheDocument();
    });

    it('applies custom className to root div when provided', () => {
      render(
        <MeetingNotesViewsRenderer
          commandSource="sidePanel"
          className="custom-class-name"
        />
      );

      const rootDiv = screen.getByTestId('meeting-notes-views');
      expect(rootDiv).toHaveClass('custom-class-name');
      expect(rootDiv).toHaveClass('meeting-notes-views');
    });

    it('uses custom title when provided', () => {
      render(
        <MeetingNotesViewsRenderer
          commandSource="sidePanel"
          title="My Custom Title"
        />
      );

      expect(screen.getByText('My Custom Title')).toBeInTheDocument();
    });

    it('uses default title "Meeting Notes" when not provided', () => {
      render(
        <MeetingNotesViewsRenderer
          commandSource="sidePanel"
        />
      );

      expect(screen.getByText('Meeting Notes')).toBeInTheDocument();
    });
  });

  // =========================================================================
  // Edge cases
  // =========================================================================

  describe('Edge cases', () => {
    it('handles meetingContext being null while commandSource is provided', () => {
      render(
        <MeetingNotesViewsRenderer
          commandSource="sidePanel"
          meetingContext={null}
        />
      );

      expect(screen.getByTestId('meeting-notes-content')).toBeInTheDocument();
      // Should still render content even though meetingContext is null
      expect(
        screen.getByText(/Side panel notes for meeting/)
      ).toBeInTheDocument();
    });

    it('renders data-command-source attribute with correct value', () => {
      render(
        <MeetingNotesViewsRenderer
          commandSource="meetingStage"
        />
      );

      const rootDiv = screen.getByTestId('meeting-notes-views');
      expect(rootDiv).toHaveAttribute('data-command-source', 'meetingStage');
    });

    it('has proper accessibility attributes on loading state', () => {
      render(
        <MeetingNotesViewsRenderer
          commandSource={null as unknown as CommandSource}
        />
      );

      const loadingDiv = screen.getByTestId('meeting-notes-loading');
      expect(loadingDiv).toHaveAttribute('role', 'status');
      expect(loadingDiv).toHaveAttribute('aria-live', 'polite');
      expect(loadingDiv).toHaveAttribute('aria-label', 'Loading meeting notes');
    });

    it('has proper accessibility attributes on content state', () => {
      render(
        <MeetingNotesViewsRenderer
          commandSource="sidePanel"
        />
      );

      const rootDiv = screen.getByTestId('meeting-notes-views');
      expect(rootDiv).toBeInTheDocument();
      // Verify structure includes header with proper heading
      const heading = screen.getByRole('heading', { level: 2 });
      expect(heading).toBeInTheDocument();
    });
  });
});
