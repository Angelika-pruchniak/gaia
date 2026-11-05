import { useState } from 'react'
import { Lilek, Speech } from '../components/Lilek'
import { fx } from '../fx'
import { update } from '../store'

const GOALS = [
  { xp: 10, label: 'Na luzie', time: '5 min dziennie' },
  { xp: 20, label: 'Regularnie', time: '10 min dziennie' },
  { xp: 30, label: 'Na serio', time: '15 min dziennie' },
  { xp: 50, label: 'Turbo', time: '25 min dziennie' },
]

export function Onboarding({ onFinish }: { onFinish: (mode: 'zero' | 'test') => void }) {
  const [step, setStep] = useState(0)
  const [name, setName] = useState('')
  const [goal, setGoal] = useState(20)

  const next = () => { fx('tap'); setStep(step + 1) }
  const end = (mode: 'zero' | 'test') => {
    fx('win')
    update((s) => { s.name = name.trim(); s.dailyGoal = goal; s.onboarded = true })
    onFinish(mode)
  }

  return (
    <div className="onboarding">
      {step === 0 && (
        <div className="ob-step center">
          <Lilek mood="wave" size={170} className="bounce" />
          <h1>Akademia Guliwera</h1>
          <p className="lead">Język biznesu w 5 minut dziennie. EBITDA, CAPEX, macierz BCG, MECE i wszystko, o czym mówią na zarządzie.</p>
          <button className="btn big" onClick={next}>Zaczynamy</button>
        </div>
      )}
      {step === 1 && (
        <div className="ob-step">
          <Speech mood="happy" text="Cześć! Jestem Lilek. Mały, ale łebski. Jak mam do Ciebie mówić?" />
          <input className="input" placeholder="Twoje imię" value={name} autoFocus maxLength={30}
            onChange={(e) => setName(e.target.value)} onKeyDown={(e) => e.key === 'Enter' && name.trim() && next()} />
          <button className="btn big" disabled={!name.trim()} onClick={next}>Dalej</button>
        </div>
      )}
      {step === 2 && (
        <div className="ob-step">
          <Speech mood="think" text={`Miło Cię poznać, ${name.trim()}! Ile czasu dziennie dajesz sobie na naukę?`} />
          <div className="goal-list">
            {GOALS.map((g) => (
              <button key={g.xp} className={`opt goal-opt ${goal === g.xp ? 'sel' : ''}`} onClick={() => { fx('tap'); setGoal(g.xp) }}>
                <b>{g.label}</b><span>{g.time} · {g.xp} XP</span>
              </button>
            ))}
          </div>
          <button className="btn big" onClick={next}>Dalej</button>
        </div>
      )}
      {step === 3 && (
        <div className="ob-step">
          <Speech mood="happy" text="Ostatnie pytanie: od czego zaczynamy?" />
          <button className="opt level-opt" onClick={() => end('zero')}>
            <b>Zaczynam od zera</b>
            <span>Pierwsza lekcja: przychód, koszt, zysk</span>
          </button>
          <button className="opt level-opt" onClick={() => end('test')}>
            <b>Coś już wiem: sprawdź mnie</b>
            <span>Krótki test (ok. 3 min) dopasuje poziom</span>
          </button>
        </div>
      )}
      <div className="dots">{[0, 1, 2, 3].map((i) => <span key={i} className={i === step ? 'on' : ''} />)}</div>
    </div>
  )
}
