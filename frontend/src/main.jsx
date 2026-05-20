import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import { MsalProvider } from "@azure/msal-react";
import { msalInstance } from "./authConfig";
import './index.css'
import App from './App.jsx'

import { BrowserRouter } from 'react-router-dom'

createRoot(document.getElementById('root')).render(
  <StrictMode>
    <MsalProvider instance={msalInstance}>
      <BrowserRouter>
        <App />
      </BrowserRouter>
    </MsalProvider>
  </StrictMode>,
)
