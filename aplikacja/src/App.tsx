import { useEffect, useState } from 'react'
import { flatLessons, pick, quickExercises, units, type Item } from './content'
import { Lilek } from './components/Lilek'
import { Coin, Confetti, Flame, Heart } from './components/ui'
import { fx } from './fx'
import { Home } from './screens/Home'
import { Onboarding } from './screens/Onboarding'
import { Placement } from './screens/Placement'
import { Practice, practiceItems } from './screens/Practice'
import { Profile } from './screens/Profile'
import { Result } from './screens/Result'
import { Session, type SessionResult } from './screens/Session'
import {
  buyHearts, completeLesson, completeReview, didToday, openChest, passTest, PRICE_HEARTS,
  tickHearts, useStore, visibleStreak, type Rewards,
} from './store'

type Tab = 'path' | 'practice' | 'profile'
type View =
  | { name: 'tabs' }
  | { name: 'placement' }
  | { name: 'lesson'; flat: number }
  | { name: 'review'; items: Item[] }
  | { name: 'test'; unit: number; items: Item[] }
  | { name: 'result'; rewards: Rewards; accuracy: number | null; seconds: number | null; title: string; heading?: string }
  | { name: 'info'; mood: 'sad' | 'cheer'; title: string; text: string }

export default function App() {
  const s = useStore()
  const [tab, setTab] = useState<Tab>('path')
  const [view, setView] = useState<View>({ name: 'tabs' })
  const [noHearts, setNoHearts] = useState(false)

  useEffect(() => {
    const t = setInterval(tickHearts, 20_000)
    return () => clearInterval(t)
  }, [])

  const home = () => { setView({ name: 'tabs' }); setTab('path') }
  const accuracy = (r: SessionResult) => (r.total ? (r.total - r.firstTryWrong.length) / r.total : 1)

  if (!s.onboarded) {
    return <Onboarding onFinish={(m) => setView(m === 'test' ? { name: 'placement' } : { name: 'tabs' })} />
  }

  switch (view.name) {
    case 'placement':
      return (
        <Placement
          onExit={home}
          onDone={(u) => {
            const rewards = passTest(u, true)
            setView({
              name: 'result', rewards, accuracy: null, seconds: null, heading: 'Poziom sprawdzony!',
              title: u === 0
                ? 'Zaczynamy od podstaw: dział 1, „Język pieniędzy”. Lilek poprowadzi Cię krok po kroku.'
                : `Twój start: dział ${u + 1}, „${units[u].title}”. Działy 1–${u} masz odblokowane, możesz do nich wracać.`,
            })
          }}
        />
      )
    case 'lesson': {
      const l = flatLessons[view.flat]
      return (
        <Session
          key={`l-${view.flat}`}
          title={l.title}
          items={l.items}
          mode="lesson"
          onExit={home}
          onDone={(r) => {
            if (r.aborted) return home()
            const rewards = completeLesson(l.id, r)
            setView({ name: 'result', rewards, accuracy: accuracy(r), seconds: r.seconds, title: l.title })
          }}
        />
      )
    }
    case 'review':
      return (
        <Session
          title="Powtórka"
          items={view.items}
          mode="review"
          onExit={home}
          onDone={(r) => {
            const rewards = completeReview(r)
            setView({ name: 'result', rewards, accuracy: accuracy(r), seconds: r.seconds, heading: 'Powtórka zrobiona!', title: 'Pamięć odświeżona. W nagrodę +1 serce.' })
          }}
        />
      )
    case 'test':
      return (
        <Session
          title="Sprawdzian"
          items={view.items}
          mode="test"
          maxMistakes={2}
          onExit={home}
          onDone={(r) => {
            if (r.mistakes > 2) {
              return setView({ name: 'info', mood: 'sad', title: 'Tym razem się nie udało', text: 'Spokojnie, przejdź kilka lekcji i spróbuj ponownie. Sprawdzian można powtarzać bez limitu.' })
            }
            const rewards = passTest(view.unit, false)
            setView({ name: 'result', rewards, accuracy: accuracy(r), seconds: r.seconds, heading: 'Sprawdzian zdany!', title: `Przeskakujesz do działu „${units[view.unit].title}”.` })
          }}
        />
      )
    case 'result':
      return <Result {...view} onContinue={home} />
    case 'info':
      return (
        <div className="result">
          {view.mood === 'cheer' && <Confetti />}
          <Lilek mood={view.mood} size={140} className={view.mood === 'cheer' ? 'jump' : ''} />
          <h1>{view.title}</h1>
          <p className="lead">{view.text}</p>
          <button className="btn big result-btn" onClick={home}>Dalej</button>
        </div>
      )
  }

  const startLesson = (flat: number) => {
    if (s.hearts <= 0) return setNoHearts(true)
    setView({ name: 'lesson', flat })
  }
  const startSkip = (u: number) => {
    const pool = units.slice(s.skipTo, u).flatMap((_, i) => quickExercises(s.skipTo + i))
    setView({ name: 'test', unit: u, items: pick(pool, 8) })
  }
  const chest = (u: number) => {
    const amount = openChest(u)
    fx('coin')
    setView({ name: 'info', mood: 'cheer', title: `+${amount} monet!`, text: `Skrzynia działu „${units[u].title}” otwarta. Wydaj monety w sklepiku Lilka.` })
  }

  const streak = visibleStreak(s)
  return (
    <div className="app">
      <header className="topbar">
        <span className="stat" title="Seria dni"><Flame off={!didToday(s)} /> {streak}</span>
        <span className="brand">Akademia <i>Guliwera</i></span>
        <span className="stat" title="Monety"><Coin /> {s.coins}</span>
        <span className="stat" title="Serca"><Heart off={s.hearts === 0} /> {s.hearts}</span>
      </header>
      <main className="content">
        {tab === 'path' && <Home onLesson={startLesson} onSkip={startSkip} onChest={chest} />}
        {tab === 'practice' && <Practice onStart={() => setView({ name: 'review', items: practiceItems(s) })} />}
        {tab === 'profile' && <Profile />}
      </main>
      <nav className="bottomnav">
        {([['path', 'Ścieżka'], ['practice', 'Powtórki'], ['profile', 'Profil']] as const).map(([id, label]) => (
          <button key={id} className={tab === id ? 'on' : ''} onClick={() => { fx('tap'); setTab(id) }}>
            <NavIcon id={id} />{label}
          </button>
        ))}
      </nav>
      {noHearts && (
        <div className="modal-bg" onClick={() => setNoHearts(false)}>
          <div className="modal" onClick={(e) => e.stopPropagation()}>
            <Lilek mood="sleep" size={90} />
            <h3>Brak serc</h3>
            <p>Jedno serce wraca co 30 minut. Zrób powtórkę (+1 serce) albo uzupełnij zapas.</p>
            <button className="btn big" disabled={s.coins < PRICE_HEARTS} onClick={() => { if (buyHearts()) { fx('coin'); setNoHearts(false) } }}>
              Uzupełnij za {PRICE_HEARTS} monet
            </button>
            <button className="btn ghost" onClick={() => { setNoHearts(false); setTab('practice') }}>Idę na powtórkę</button>
          </div>
        </div>
      )}
    </div>
  )
}

function NavIcon({ id }: { id: Tab }) {
  const common = { width: 26, height: 26, viewBox: '0 0 24 24', fill: 'none', stroke: 'currentColor', strokeWidth: 2.2, strokeLinecap: 'round' as const, strokeLinejoin: 'round' as const }
  if (id === 'path') return <svg {...common}><circle cx="6" cy="18" r="2.5" /><circle cx="18" cy="6" r="2.5" /><path d="M8.5 18h5a3.5 3.5 0 0 0 0-7h-3a3.5 3.5 0 0 1 0-7h5" /></svg>
  if (id === 'practice') return <svg {...common}><path d="M20 11a8 8 0 0 0-14.5-4.5M4 4v4h4" /><path d="M4 13a8 8 0 0 0 14.5 4.5M20 20v-4h-4" /></svg>
  return <svg {...common}><circle cx="12" cy="8" r="4" /><path d="M4 21c1.5-4 4.5-6 8-6s6.5 2 8 6" /></svg>
}
