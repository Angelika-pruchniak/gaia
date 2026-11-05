import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import App from './App'
import { primeAudio } from './fx'
import './install'
import './styles.css'

primeAudio()
createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <App />
  </StrictMode>,
)

if ('serviceWorker' in navigator && import.meta.env.PROD) {
  window.addEventListener('load', () => navigator.serviceWorker.register('./sw.js').catch(() => {}))
}
