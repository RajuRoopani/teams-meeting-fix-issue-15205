/**
 * Shared types for Teams meeting app packages.
 * Issue #15205 — types used across components-meeting-apis and apps-meeting-notes.
 */

/**
 * The source that triggered the meeting notes command.
 * Maps to microsoftTeams.app.Context.page.subPageId or the tab entity context
 * that determines which view to render in the meeting notes app.
 */
export type CommandSource =
  | 'sidePanel'
  | 'contentBubble'
  | 'meetingStage'
  | 'settings'
  | 'remove'
  | string; // Allow future/unknown sources without breaking the type

/**
 * Current initialization lifecycle state of the Teams SDK / meeting context.
 */
export type InitializationStatus = 'idle' | 'initializing' | 'ready' | 'error';

/**
 * Shape of the Teams app context relevant to the meeting notes app.
 * This is a subset of microsoftTeams.app.Context used internally.
 */
export interface MeetingContext {
  /** The meeting ID (chatId of the meeting chat thread). */
  meetingId: string | undefined;
  /** The user's Microsoft Azure AD object ID. */
  userObjectId: string | undefined;
  /** The locale set in the Teams client. */
  locale: string;
  /**
   * The command source that opened the task module / content bubble.
   * May be undefined on first load of an ad-hoc meeting (issue #15205).
   */
  commandSource: CommandSource | undefined;
  /** The Teams theme (default | dark | contrast). */
  theme: 'default' | 'dark' | 'contrast';
}

/**
 * Props shared by components that need to communicate Teams init state to children.
 */
export interface WithTeamsContextProps {
  /** The resolved Teams meeting context, or null while initializing. */
  meetingContext: MeetingContext | null;
  /** Current initialization status of the Teams SDK. */
  initializationStatus: InitializationStatus;
  /** Error message if initialization failed, otherwise null. */
  initializationError: string | null;
}
