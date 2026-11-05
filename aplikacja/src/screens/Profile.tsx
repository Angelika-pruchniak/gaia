import { useState } from 'react'
import { Lilek } from '../components/Lilek'
import { Coin, Flame, Heart, Progress, Snow, Star } from '../components/ui'
import { fx } from '../fx'
import {
  BADGES, buyFreeze, buyHearts, dayKey, levelOf, MAX_HEARTS, PRICE_FREEZE, PRICE_HEARTS,
  resetAll, update, useStore, visibleStreak,
} from '../store'

const DAYS = ['nd', 'pn', 'wt', 'śr', 'cz', 'pt', 'sb']

export function Profile() {
  const s = useStore()
  const lvl = levelOf(s.xp)
  const [confirm, setConfirm] = useState(false)
  const week = Array.from({ length: 7 }, (_, i) => {
    const d = new Date()
    d.setDate(d.getDate() - (6 - i))
    return { label: DAYS[d.getDay()], xp: s.xpByDay[dayKey(d)] ?? 0 }
  })
  const maxXp = Math.max(s.dailyGoal, ...week.map((w) => w.xp))

  return (
    <div className="page">
      <div className="profile-head">
        <Lilek mood="happy" size={86} />
        <div>
          <h1>{s.name || 'Manager'}</h1>
          <div className="lvl-name">Poziom {lvl.index}: {lvl.name}</div>
          <Progress value={lvl.progress} color="var(--lilac)" />
          <div className="muted small">{lvl.next ? `${s.xp} / ${lvl.next} XP do kolejnego poziomu` : `${s.xp} XP: najwyższy poziom!`}</div>
        </div>
      </div>

      <div className="tiles">
        <div className="tile t-peach"><span>Seria</span><b><Flame size={18} /> {visibleStreak(s)}</b></div>
        <div className="tile t-yellow"><span>Razem XP</span><b>{s.xp}</b></div>
        <div className="tile t-lilac"><span>Najdłuższa seria</span><b>{s.longestStreak}</b></div>
        <div className="tile t-pink"><span>Lekcje</span><b>{Object.keys(s.done).length}</b></div>
      </div>

      <div className="card">
        <h3>Ostatnie 7 dni</h3>
        <div className="bars">
          {week.map((w, i) => (
            <div key={i} className="bar-col">
              <div className="bar-val">{w.xp || ''}</div>
              <div className="bar" style={{ height: `${(w.xp / maxXp) * 90 + 4}px`, background: w.xp >= s.dailyGoal ? 'var(--coral)' : 'var(--peach)' }} />
              <div className="bar-label">{w.label}</div>
            </div>
          ))}
        </div>
      </div>

      <div className="card">
        <h3>Sklepik Lilka <span className="coins"><Coin size={18} /> {s.coins}</span></h3>
        <div className="shop-row">
          <Snow size={30} />
          <div><b>Zamrożenie serii</b><div className="muted small">Ratuje serię, gdy opuścisz dzień. Masz: {s.freezes}/2</div></div>
          <button className="btn small" disabled={s.coins < PRICE_FREEZE || s.freezes >= 2}
            onClick={() => buyFreeze() && fx('coin')}><Coin size={14} /> {PRICE_FREEZE}</button>
        </div>
        <div className="shop-row">
          <Heart size={30} />
          <div><b>Pełne serca</b><div className="muted small">Masz {s.hearts}/{MAX_HEARTS}</div></div>
          <button className="btn small" disabled={s.coins < PRICE_HEARTS || s.hearts >= MAX_HEARTS}
            onClick={() => buyHearts() && fx('coin')}><Coin size={14} /> {PRICE_HEARTS}</button>
        </div>
      </div>

      <div className="card">
        <h3>Odznaki</h3>
        <div className="badge-grid">
          {BADGES.map((b) => {
            const on = s.badges.includes(b.id)
            return (
              <div key={b.id} className={`badge ${on ? 'on' : ''}`}>
                <div className="badge-ico"><Star size={26} off={!on} /></div>
                <b>{b.name}</b><span>{b.desc}</span>
              </div>
            )
          })}
        </div>
      </div>

      <div className="card">
        <h3>Ustawienia</h3>
        <label className="toggle">
          <span>Dźwięki</span>
          <input type="checkbox" checked={s.settings.sound} onChange={(e) => { update((x) => { x.settings.sound = e.target.checked }); fx('pop') }} />
        </label>
        <label className="toggle">
          <span>Wibracje <span className="muted small">(Android)</span></span>
          <input type="checkbox" checked={s.settings.vibrate} onChange={(e) => { update((x) => { x.settings.vibrate = e.target.checked }); fx('pop') }} />
        </label>
        <label className="toggle">
          <span>Cel dnia</span>
          <select value={s.dailyGoal} onChange={(e) => update((x) => { x.dailyGoal = Number(e.target.value) })}>
            {[10, 20, 30, 50].map((g) => <option key={g} value={g}>{g} XP</option>)}
          </select>
        </label>
        {confirm ? (
          <div className="row">
            <button className="btn red small" onClick={() => { resetAll(); setConfirm(false) }}>Tak, usuń postęp</button>
            <button className="btn ghost small" onClick={() => setConfirm(false)}>Anuluj</button>
          </div>
        ) : (
          <button className="btn ghost small" onClick={() => setConfirm(true)}>Zacznij od nowa</button>
        )}
      </div>

      <div className="card about">
        <h3>O aplikacji</h3>
        <p>Akademia Guliwera to autorskie lekcje <b>Guliwera 4</b> („strategia, produkty cyfrowe i dobry humor”). Zakres działów odpowiada typowym przedmiotom studiów ekonomicznych i menedżerskich: mikro- i makroekonomii, podstawom finansów, rachunkowości oraz zarządzaniu strategicznemu.</p>
        <p className="muted small">Nazwy modeli (np. macierz BCG, model 7S, 5 sił Portera) opisują ogólnie znane koncepcje. Aplikacja nie jest powiązana z Boston Consulting Group, McKinsey &amp; Company ani żadną uczelnią. Treści mają charakter edukacyjny i nie są poradą finansową. Postęp jest zapisany tylko na tym urządzeniu.</p>
        <p className="muted small">guliwera4.pl</p>
      </div>
    </div>
  )
}
