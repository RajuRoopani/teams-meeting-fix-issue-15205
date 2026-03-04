/**
 * MeetingApisContainer
 *
 * Fix for issue #15205 — "blank pop-up dialog on first app load in ad-hoc Teams meeting"
 *
 * Root cause (before fix):
 *   `notifySuccess()` was called unconditionally on mount, before the Teams SDK
 *   had resolved `getContext()`. This caused the host shell to mark the app as
 *   "ready" while the React tree was still blank, resulting in users seeing an
 *   empty dialog.
 *
 * Fix:
 *   AC2 — `notifySuccess()` is now called ONLY inside the `.then()` of
 *   `getContext()`, guaranteeing the context is fully resolved first.
 *   Until initialization is complete, a visible "Initializing…" state is
 *   rendered so the dialog is never blank.
 */

import React, { useCallback, useEffect, useState } from 'react';
import * as microsoftTeams from '@microsoft/teams-js';

import type {
  InitializationStatus,
  MeetingContext,
  WithTeamsContextProps,
} from './types';

export interface MeetingApisContainerProps {
  children?: React.ReactNode | ((ctx: WithTeamsContextProps) => React.ReactNode);
  /**
   * Optional callback invoked once the Teams context is successfully resolved.
   * Receives the normalised MeetingContext.
   */
  onContextReady?: (ctx: MeetingContext) => void;
  /**
   * Optional callback invoked if Teams SDK initialisation fails.
   */
  onInitError?: (error: Error) => void;
}

/**
 * Normalises a raw microsoftTeams.app.Context into our internal MeetingContext
 * shape, applying safe defaults for fields that may be absent on ad-hoc meetings.
 */
function normaliseMeetingContext(
  raw: microsoftTeams.app.Context
): MeetingContext {
  return {
    meetingId: raw.meeting?.id,
    userObjectId: raw.user?.id,
    locale: raw.app?.locale ?? 'en-us',
    // commandSource lives at page.subPageId in Teams JS SDK v2.
    // On ad-hoc meetings it can be undefined on first load — that is the root
    // cause of issue #15205 and is handled in MeetingNotesViewsRenderer.
    commandSource: (raw.page as { subPageId?: string } | undefined)?.subPageId,
    theme: (raw.app?.theme as MeetingContext['theme']) ?? 'default',
  };
}

/**
 * Provides Teams SDK context to the subtree.
 *
 * Renders a visible initialising indicator while the SDK is starting up,
 * and calls `notifySuccess()` only after context is fully resolved (AC2).
 *
 * Supports both render-prop children (receives `WithTeamsContextProps`) and
 * plain React children for backwards compatibility.
 */
const MeetingApisContainer: React.FC<MeetingApisContainerProps> = ({
  children,
  onContextReady,
  onInitError,
}) => {
  const [meetingContext, setMeetingContext] = useState<MeetingContext | null>(
    null
  );
  const [initializationStatus, setInitializationStatus] =
    useState<InitializationStatus>('idle');
  const [initializationError, setInitializationError] = useState<string | null>(
    null
  );

  const handleInitError = useCallback(
    (err: unknown): void => {
      const error =
        err instanceof Error ? err : new Error(String(err ?? 'Unknown error'));
      console.error('[MeetingApisContainer] Initialisation failed:', error);
      setInitializationStatus('error');
      setInitializationError(error.message);
      onInitError?.(error);
    },
    [onInitError]
  );

  useEffect(() => {
    let cancelled = false;

    const initialise = async (): Promise<void> => {
      setInitializationStatus('initializing');

      try {
        // Step 1 — Initialise the Teams JS SDK.
        await microsoftTeams.app.initialize();

        if (cancelled) return;

        // Step 2 — Fetch the Teams context.
        const rawCtx = await microsoftTeams.app.getContext();

        if (cancelled) return;

        const ctx = normaliseMeetingContext(rawCtx);
        setMeetingContext(ctx);
        setInitializationStatus('ready');

        // AC2: notifySuccess() is called ONLY after context is fully resolved.
        // Previously this was called unconditionally on mount, which caused
        // the blank dialog bug on first load of an ad-hoc meeting.
        microsoftTeams.app.notifySuccess();

        onContextReady?.(ctx);
      } catch (err) {
        if (!cancelled) {
          handleInitError(err);
        }
      }
    };

    initialise();

    return () => {
      cancelled = true;
    };
  }, [handleInitError, onContextReady]);

  // --- Render ---

  if (initializationStatus === 'error') {
    return (
      <div
        className="meeting-apis-error"
        role="alert"
        aria-live="assertive"
        data-testid="meeting-apis-error"
      >
        <p>Failed to connect to Teams.</p>
        {initializationError !== null && (
          <p className="meeting-apis-error__detail">{initializationError}</p>
        )}
      </div>
    );
  }

  if (initializationStatus !== 'ready') {
    // Visible initialising indicator — never renders blank (AC1 support).
    return (
      <div
        className="meeting-apis-loading"
        role="status"
        aria-live="polite"
        aria-label="Initializing Teams connection"
        data-testid="meeting-apis-loading"
      >
        <span className="meeting-apis-loading__text">Initializing…</span>
      </div>
    );
  }

  const ctxProps: WithTeamsContextProps = {
    meetingContext,
    initializationStatus,
    initializationError,
  };

  // Support both render-prop and plain children patterns.
  const resolvedChildren =
    typeof children === 'function' ? children(ctxProps) : children;

  return (
    <div
      className="meeting-apis-container"
      data-testid="meeting-apis-container"
    >
      {resolvedChildren}
    </div>
  );
};

export default MeetingApisContainer;
