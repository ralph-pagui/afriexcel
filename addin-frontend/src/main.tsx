import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './App';
import './index.css';

/* global Office */

let appRendered = false;

const renderApp = () => {
  if (appRendered) return;
  appRendered = true;
  const rootElement = document.getElementById('root');
  if (rootElement) {
    ReactDOM.createRoot(rootElement).render(
      <React.StrictMode>
        <App />
      </React.StrictMode>
    );
  }
};

// Initialisation robuste d'Office.js pour Excel Desktop (2016 à 365) et Web
if (typeof Office !== 'undefined') {
  // Définir Office.initialize si ce n'est pas déjà fait
  const existingInit = Office.initialize;
  Office.initialize = function (reason) {
    if (typeof existingInit === 'function') {
      try { existingInit(reason); } catch (_) {}
    }
    renderApp();
  };

  if (Office.onReady) {
    Office.onReady((info) => {
      console.log("Office.onReady confirmé :", info?.host, info?.platform);
      renderApp();
    });
  }

  // Filet de sécurité si ouvert hors d'Office
  setTimeout(renderApp, 400);
} else {
  // Mode Web autonome direct
  renderApp();
}
