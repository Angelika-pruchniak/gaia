import { useEffect, useState } from 'react'
import { Lilek } from '../components/Lilek'
import { Coin, Confetti, Flame, Star } from '../components/ui'
import { fx } from '../fx'
import { BADGES, type Rewards } from '../store'

type Panel = 'main' | 'streak' | 'level' | 'badges'

export function Result({ rewards, accuracy, seconds, title, heading, onContinue }: {
  rewards: Rewards
  heading?: string
  accuracy: number | null
  seconds: number | null
  title: string
  onContinue: () => void
}) {
  const panels: Panel[] = ['main']
  if (rewards.streakUp) panels.push('streak')
  if (rewards.levelUp) panels.push('level')
  if (rewards.badges.length) panels.push('badges')
  const [i, setI] = useState(0)
  const p = panels[i]

  useEffect(() => {
    fx(p === 'main' ? 'win' : p === 'streak' ? 'streak' : p === 'level' ? 'level' : 'coin')
  }, [p])

  const next = () => (i + 1 < panels.length ? setI(i + 1) : onContinue())

  return (
    <div className="result">
      <Confetti key={p} />
      {p === 'main' && (
        <>
          <Lilek mood="cheer" size={150} className="jump" />
          <h1>{heading ?? (accuracy === 1 ? 'Bezbłędnie!' : 'Lekcja zaliczona!')}</h1>
          <p className="lead">{title}</p>
          <div className="tiles">
            <div className="tile t-yellow"><span>XP</span><b>+{rewards.xp}</b></div>
            {accuracy !== null && <div className="tile t-green"><span>Celność</span><b>{Math.round(accuracy * 100)}%</b></div>}
            {seconds !== null && <div className="tile t-lilac"><span>Czas</span><b>{Math.floor(seconds / 60)}:{String(seconds % 60).padStart(2, '0')}</b></div>}
            {rewards.coins > 0 && <div className="tile t-peach"><span>Monety</span><b><Coin size={18} /> +{rewards.coins}</b></div>}
          </div>
          {rewards.goalReached && <div className="pill">Cel dnia osiągnięty!</div>}
        </>
      )}
      {p === 'streak' && (
        <>
          <div className="big-flame"><Flame size={140} /></div>
          <h1 className="streak-num">{rewards.streak}</h1>
          <p className="lead">{rewards.streak === 1 ? 'Pierwszy dzień serii! Wróć jutro, żeby ją podtrzymać.' : `dni z rzędu! Lilek jest pod wrażeniem.`}</p>
          <div className="week">
            {Array.from({ length: 7 }, (_, d) => (
              <span key={d} className={d < Math.min(rewards.streak, 7) ? 'on' : ''}><Flame size={22} off={d >= Math.min(rewards.streak, 7)} /></span>
            ))}
          </div>
        </>
      )}
      {p === 'level' && (
        <>
          <Lilek mood="cheer" size={130} className="jump" />
          <div className="level-badge"><Star size={28} /> Nowy poziom</div>
          <h1>{rewards.levelUp}</h1>
          <p className="lead">Awans! Jeszcze trochę i będziesz tłumaczyć EBITDA na firmowej wigilii.</p>
        </>
      )}
      {p === 'badges' && (
        <>
          <h1>Nowa odznaka!</h1>
          <div className="badge-list">
            {rewards.badges.map((id) => {
              const b = BADGES.find((x) => x.id === id)!
              return (
                <div key={id} className="badge on big"><div className="badge-ico"><Star size={34} /></div><b>{b.name}</b><span>{b.desc}</span></div>
              )
            })}
          </div>
        </>
      )}
      <button className="btn big result-btn" onClick={() => { fx('tap'); next() }}>Dalej</button>
    </div>
  )
}
