import { useState } from 'react'
import { useInstall } from '../install'
import { Lilek } from './Lilek'

export function InstallCard() {
  const { installed, canPrompt, ios, prompt } = useInstall()
  const [hidden, setHidden] = useState(() => {
    try { return localStorage.getItem('install-hidden') === '1' } catch { return false }
  })
  const [howto, setHowto] = useState(false)
  if (installed || hidden || (!canPrompt && !ios)) return null
  const hide = () => { setHidden(true); try { localStorage.setItem('install-hidden', '1') } catch { /* ignoruj */ } }
  return (
    <div className="install-card">
      <Lilek mood="wave" size={48} />
      <div className="install-text">
        <b>Zainstaluj Akademię na telefonie</b>
        {howto ? (
          <span>W Safari dotknij <b>Udostępnij</b> (kwadrat ze strzałką), a potem <b>Do ekranu początkowego</b>.</span>
        ) : (
          <span>Ikona Lilka na ekranie, działa też offline.</span>
        )}
      </div>
      <div className="install-btns">
        <button className="btn small" onClick={() => (canPrompt ? prompt() : setHowto(true))}>Zainstaluj</button>
        <button className="icon-btn" aria-label="Ukryj" onClick={hide}>✕</button>
      </div>
    </div>
  )
}
