/**
 * Mock for @fluentui/react
 *
 * Provides simple identity mocks of Fluent UI components so tests can render
 * without complex dependencies.
 *
 * The Spinner component is mocked to render its label text for assertion in tests.
 */

const React = require('react');

module.exports = {
  Spinner: ({ label, ariaLabel, ...props }: any) => {
    return React.createElement(
      'div',
      {
        ...props,
        'data-testid': props['data-testid'] || 'spinner',
        'aria-label': ariaLabel,
      },
      label
    );
  },

  SpinnerSize: {
    xSmall: 'xSmall',
    small: 'small',
    medium: 'medium',
    large: 'large',
  },

  Stack: ({ children, ...props }: any) => {
    return React.createElement('div', props, children);
  },

  Text: ({ children, ...props }: any) => {
    return React.createElement('span', props, children);
  },

  Icon: ({ iconName, ...props }: any) => {
    return React.createElement('i', { ...props, className: `ms-Icon ms-Icon--${iconName}` });
  },

  Button: ({ children, ...props }: any) => {
    return React.createElement('button', props, children);
  },

  DefaultButton: ({ children, ...props }: any) => {
    return React.createElement('button', props, children);
  },

  PrimaryButton: ({ children, ...props }: any) => {
    return React.createElement('button', props, children);
  },
};
