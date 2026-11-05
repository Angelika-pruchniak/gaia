import { useEffect, useRef } from 'react'
import { flatLessons, units } from '../content'
import { InstallCard } from '../components/InstallCard'
import { Lilek } from '../components/Lilek'
import { Progress, Star } from '../components/ui'
import { fx } from '../fx'
import { chestReady, currentLesson, dayKey, lessonStatus, unitReached, useStore } from '../store'

const OFFSETS = [0, 52, 78, 52, 0, -52, -78, -52]

type Props = {
  onLesson: (flatIndex: number) => void
  onSkip: (unitIndex: number) => void
  onChest: (unitIndex: number) => void
}

export function Home({ onLesson, onSkip, onChest }: Props) {
  const s = useStore()
  const cur = currentLesson(s)
  const curRef = useRef<HTMLDivElement>(null)
  const todayXp = s.xpByDay[dayKey()] ?? 0

  useEffect(() => {
    curRef.current?.scrollIntoView({ block: 'center', behavior: 'smooth' })
  }, [])

  return (
    <div className="home">
      <InstallCard />
      <div className="goal-card">
        <Lilek mood={todayXp >= s.dailyGoal ? 'cheer' : 'happy'} size={54} />
        <div className="goal-text">
          <b>{todayXp >= s.dailyGoal ? 'Cel dnia zrobiony!' : `Cel dnia: ${todayXp}/${s.dailyGoal} XP`}</b>
          <Progress value={todayXp / s.dailyGoal} color="var(--yellow)" />
        </div>
      </div>

      {units.map((u, ui) => {
        const reached = unitReached(s, ui)
        const lessons = flatLessons.filter((l) => l.unitIndex === ui)
        return (
          <section key={u.id} className="unit">
            <div className={`unit-banner c-${u.color}`}>
              <div>
                <div className="unit-no">Dział {ui + 1}</div>
                <h2>{u.title}</h2>
                <div className="unit-sub">{u.subtitle}</div>
              </div>
              {!reached && (
                <button className="btn small light" onClick={() => { fx('tap'); onSkip(ui) }}>Przeskocz</button>
              )}
            </div>
            <div className="path">
              {lessons.map((l, i) => {
                const st = lessonStatus(s, l.flatIndex)
                const isCur = l.flatIndex === cur
                const times = s.done[l.id] ?? 0
                return (
                  <div key={l.id} className="node-wrap" style={{ transform: `translateX(${OFFSETS[(l.flatIndex) % OFFSETS.length]}px)` }}
                    ref={isCur ? curRef : undefined}>
                    {isCur && <div className="start-bubble">START</div>}
                    <button
                      className={`node ${st} c-${u.color} ${isCur ? 'current' : ''}`}
                      disabled={st === 'locked'}
                      aria-label={`${l.title}${st === 'locked' ? ' (zablokowana)' : ''}`}
                      onClick={() => { fx('tap'); onLesson(l.flatIndex) }}
                    >
                      {st === 'locked' ? '🔒' : st === 'done' ? '✓' : i + 1}
                    </button>
                    <div className="node-title">{l.title}</div>
                    {times > 0 && (
                      <div className="node-stars">
                        {[1, 2, 3].map((n) => <Star key={n} size={14} off={times < n} />)}
                      </div>
                    )}
                    {isCur && (
                      <div className={`node-mascot ${OFFSETS[l.flatIndex % OFFSETS.length] > 0 ? 'left' : ''}`}>
                        <Lilek mood="wave" size={64} />
                      </div>
                    )}
                  </div>
                )
              })}
              <div className="node-wrap" style={{ transform: `translateX(${OFFSETS[(lessons[lessons.length - 1].flatIndex + 1) % OFFSETS.length]}px)` }}>
                <button
                  className={`chest ${chestReady(s, ui) ? 'ready' : s.chests[u.id] ? 'open' : ''}`}
                  disabled={!chestReady(s, ui)}
                  aria-label="Skrzynia z monetami"
                  onClick={() => onChest(ui)}
                >
                  <ChestIcon open={!!s.chests[u.id]} />
                </button>
              </div>
            </div>
          </section>
        )
      })}
      <div className="path-end">
        <Lilek mood="think" size={80} />
        <p>Kolejne działy w przygotowaniu. Lilek już pisze!</p>
      </div>
    </div>
  )
}

function ChestIcon({ open }: { open: boolean }) {
  return (
    <svg width="54" height="46" viewBox="0 0 54 46" aria-hidden>
      <rect x="4" y="18" width="46" height="26" rx="5" fill="#E8A92E" />
      <rect x="4" y="18" width="46" height="8" fill="#C9861A" />
      {open ? (
        <path d="M4 18 L10 4 H44 L50 18Z" fill="#FFC84A" />
      ) : (
        <path d="M4 20 Q27 2 50 20 Z" fill="#FFC84A" />
      )}
      <rect x="23" y="20" width="8" height="10" rx="2" fill="#1E2A4A" />
    </svg>
  )
}
