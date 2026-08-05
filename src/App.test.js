import { render, screen } from '@testing-library/react';
import App from './App';

test('renders quiz title and team score summary', () => {
  render(<App />);

  expect(screen.getByText(/quiz league/i)).toBeInTheDocument();
  expect(screen.getByText(/team a/i)).toBeInTheDocument();
  expect(screen.getByText(/sport/i)).toBeInTheDocument();
  expect(screen.getByText(/history/i)).toBeInTheDocument();
});
