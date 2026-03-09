import React from 'react';
import ReactDOM from 'react-dom/client';
import './index.css';
import App from './App';

const root = ReactDOM.createRoot(document.getElementById('root'));
root.render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);

// Enregistrement du Service Worker pour la PWA
if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => {
    navigator.serviceWorker
      .register('/service-worker.js')
      .then((registration) => {
        console.log('[PWA] Service Worker enregistré :', registration.scope);
        registration.onupdatefound = () => {
          const worker = registration.installing;
          worker.onstatechange = () => {
            if (worker.state === 'installed' && navigator.serviceWorker.controller) {
              if (window.confirm('Une mise à jour est disponible. Recharger ?')) {
                window.location.reload();
              }
            }
          };
        };
      })
      .catch((err) => console.warn('[PWA] Échec service worker :', err));
  });
}
