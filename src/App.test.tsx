import React from 'react';
import { render, screen } from '@testing-library/react';
import App from './App';
import { PublicClientApplication } from '@azure/msal-browser';

// Mock MSAL configuration
const msalConfig = {
  auth: {
    clientId: 'test-client-id',
    authority: 'https://login.microsoftonline.com/common',
    redirectUri: 'http://localhost:3000'
  },
  cache: {
    cacheLocation: 'sessionStorage',
    storeAuthStateInCookie: false,
  }
};

let pca: PublicClientApplication;

beforeEach(() => {
  // Initialize the PublicClientApplication before each test
  pca = new PublicClientApplication(msalConfig);
});

test('renders learn react link', () => {
  render(<App pca={pca}/>);
  const linkElement = screen.getByText(/learn react/i);
  expect(linkElement).toBeInTheDocument();
});