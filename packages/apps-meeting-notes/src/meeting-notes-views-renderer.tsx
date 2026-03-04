/**
 * MeetingNotesViewsRenderer
 *
 * Fix for issue #15205 — "blank pop-up dialog on first app load in ad-hoc Teams meeting"
 *
 * Root cause (before fix):
 *   The component returned `null` / empty JSX when `commandSource` was
 *   undefined/null on first load. Ad-hoc meetings don't set `commandSource`
 *   until the Teams context resolves, so the dialog was blank for the entire
 *   SDK initialisation window.
 *
 * Fix:
 *   AC1 — When `commandSource` is null/undefined, render a visible loading
 *   indicator (Fluent UI Spinner) instead of returning null.
 *   AC3 — All existing view logic for initialised state is preserved.
 *   AC5 — Strict TypeScript; no implicit any.
 */

import React, { useEffect, useState } from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react';

import type { CommandSource, MeetingContext } from '../../components-meeting-apis/src/types';

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

export interface MeetingNotesViewsRendererProps {
  /**
   * The Teams meeting context, supplied by MeetingApisContainer (or a parent).
   * While null, the component shows a loading state (AC1).
   */
  meetingContext?: MeetingContext | null;
  /**
   * Directly pass a commandSource to override / bypass context resolution.
   * Useful in Storybook and unit tests.
   */
  commandSource?: CommandSource;
  /**
   * Optional title displayed in the meeting notes header.
   * Defaults to "Meeting Notes".
   */
  title?: string;
  /**
   * Optional class name applied to the root element.
   */
  className?: string;
}

// ---------------------------------------------------------------------------
// View map — maps a commandSource value to the view name shown in the header.
// Extend this record as new meeting note views are added.
// ---------------------------------------------------------------------------

const VIEW_LABEL_MAP: Record<string, string> = {
  sidePanel: 'Side Panel Notes',
  contentBubble: 'Meeting Bubble Notes',
  meetingStage: 'Stage Notes',
  settings: 'Settings',
  remove: 'Remove App',
};

function resolveViewLabel(source: CommandSource): string {
  return VIEW_LABEL_MAP[source] ?? `Meeting Notes (${source})`;
}

// ---------------------------------------------------------------------------
// Component
// ---------------------------------------------------------------------------

/**
 * Renders the correct meeting-notes view for the given `commandSource`.
 *
 * Key invariant: this component NEVER returns null or empty JSX.
 * - Uninitialised state  → Fluent UI Spinner (AC1)
 * - Initialised state    → Meeting notes content for the resolved source (AC3)
 */
const MeetingNotesViewsRenderer: React.FC<MeetingNotesViewsRendererProps> = ({
  meetingContext,
  commandSource: commandSourceProp,
  title = 'Meeting Notes',
  className,
}) => {
  // Prefer the directly-passed prop; fall back to what's in the context.
  const [resolvedCommandSource, setResolvedCommandSource] = useState<
    CommandSource | null | undefined
  >(commandSourceProp ?? meetingContext?.commandSource ?? null);

  // Re-sync whenever the parent provides a new context or direct prop.
  useEffect(() => {
    const incoming = commandSourceProp ?? meetingContext?.commandSource ?? null;
    setResolvedCommandSource(incoming);
  }, [commandSourceProp, meetingContext]);

  // ---------------------------------------------------------------------------
  // AC1: commandSource is null/undefined → render a visible loading indicator.
  //      NEVER return null here — that is the original bug.
  // ---------------------------------------------------------------------------
  if (resolvedCommandSource === null || resolvedCommandSource === undefined) {
    return (
      <div
        className="meeting-notes-loading"
        role="status"
        aria-live="polite"
        aria-label="Loading meeting notes"
        data-testid="meeting-notes-loading"
      >
        <Spinner
          size={SpinnerSize.medium}
          label="Loading meeting notes…"
          ariaLabel="Loading meeting notes"
        />
      </div>
    );
  }

  // ---------------------------------------------------------------------------
  // AC3: commandSource is populated → render the actual meeting notes view.
  // ---------------------------------------------------------------------------
  const viewLabel = resolveViewLabel(resolvedCommandSource);
  const rootClass = ['meeting-notes-views', className].filter(Boolean).join(' ');

  return (
    <div
      className={rootClass}
      data-testid="meeting-notes-views"
      data-command-source={resolvedCommandSource}
    >
      <header className="meeting-notes-views__header">
        <h2 className="meeting-notes-views__title">{title}</h2>
        <span
          className="meeting-notes-views__view-label"
          data-testid="meeting-notes-view-label"
        >
          {viewLabel}
        </span>
      </header>

      <main
        className="meeting-notes-views__content meeting-notes-content"
        data-testid="meeting-notes-content"
      >
        <MeetingNotesContent
          commandSource={resolvedCommandSource}
          meetingContext={meetingContext ?? null}
        />
      </main>
    </div>
  );
};

// ---------------------------------------------------------------------------
// MeetingNotesContent — renders the body for each commandSource variant.
// Kept as a separate private component so MeetingNotesViewsRenderer stays
// focused on the loading-vs-loaded branching logic.
// ---------------------------------------------------------------------------

interface MeetingNotesContentProps {
  commandSource: CommandSource;
  meetingContext: MeetingContext | null;
}

const MeetingNotesContent: React.FC<MeetingNotesContentProps> = ({
  commandSource,
  meetingContext,
}) => {
  switch (commandSource) {
    case 'sidePanel':
      return (
        <section className="meeting-notes-content__side-panel">
          <p>Side panel notes for meeting <strong>{meetingContext?.meetingId ?? 'unknown'}</strong>.</p>
        </section>
      );

    case 'contentBubble':
      return (
        <section className="meeting-notes-content__bubble">
          <p>Content bubble view — quick notes surface.</p>
        </section>
      );

    case 'meetingStage':
      return (
        <section className="meeting-notes-content__stage">
          <p>Collaborative stage view — all participants can see this.</p>
        </section>
      );

    case 'settings':
      return (
        <section className="meeting-notes-content__settings">
          <p>App settings configuration.</p>
        </section>
      );

    case 'remove':
      return (
        <section className="meeting-notes-content__remove">
          <p>Remove the app from this meeting.</p>
        </section>
      );

    default:
      return (
        <section className="meeting-notes-content__default">
          <p>Meeting notes are ready.</p>
        </section>
      );
  }
};

// ---------------------------------------------------------------------------
// Exports
// ---------------------------------------------------------------------------

export { resolveViewLabel };
export default MeetingNotesViewsRenderer;
