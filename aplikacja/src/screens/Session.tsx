import { useRef, useState } from 'react'
import { isExercise, type Item, type TeachCard } from '../content'
import { ExerciseCard } from '../components/ExerciseCard'
import { Lilek, Speech } from '../components/Lilek'
import { Heart, Progress } from '../components/ui'
import { fx } from '../fx'
import { buyHearts, loseHeart, PRICE_HEARTS, useStore } from '../store'

export type SessionMode = 'lesson' | 'review' | 'test'
export type SessionResult = {
  exerciseIds: string[]
  firstTryWrong: string[]
  maxCombo: number
  mistakes: number
  total: number
  seconds: number
  aborted?: boolean
}

type Props = {
  title: string
  items: Item[]
  mode: SessionMode
  maxMistakes?: number
  onExit: () => void
  onDone: (r: SessionResult) => void
}

export function Session({ title, items, mode, maxMistakes = 2, onExit, onDone }: Props) {
  const s = useStore()
  const [queue, setQueue] = useState<Item[]>(items)
  const [pos, setPos] = useState(0)
  const [doneIds, setDoneIds] = useState<string[]>([])
  const [combo, setCombo] = useState(0)
  const [toast, setToast] = useState<string | null>(null)
  const [confirmExit, setConfirmExit] = useState(false)
  const [noHearts, setNoHearts] = useState(false)
  const stats = useRef({ firstTryWrong: new Set<string>(), mistakes: 0, maxCombo: 0, start: Date.now(), lastOk: false })

  const total = new Set(items.map((i) => i.id)).size
  const item = queue[pos]
  const exIds = items.filter(isExercise).map((e) => e.id)

  const finish = (aborted = false) => {
    const st = stats.current
    onDone({
      exerciseIds: exIds,
      firstTryWrong: [...st.firstTryWrong],
      maxCombo: st.maxCombo,
      mistakes: st.mistakes,
      total: exIds.length,
      seconds: Math.round((Date.now() - st.start) / 1000),
      aborted,
    })
  }

  const handleResult = (ok: boolean) => {
    const st = stats.current
    st.lastOk = ok
    if (!ok) {
      st.mistakes++
      if (!doneIds.includes(item.id)) st.firstTryWrong.add(item.id)
      setCombo(0)
      if (mode === 'lesson') loseHeart()
      return
    }
    const c = combo + 1
    setCombo(c)
    st.maxCombo = Math.max(st.maxCombo, c)
    if (c === 3 || c === 5 || (c >= 8 && c % 4 === 0)) {
      setTimeout(() => fx('combo'), 380)
      setToast(c >= 8 ? `${c} z rzędu! Jesteś w transie!` : `${c} z rzędu!`)
      setTimeout(() => setToast(null), 1600)
    }
  }

  const advance = () => {
    const st = stats.current
    const cur = queue[pos]
    let nextQueue = queue
    let nextDone = doneIds
    if (!isExercise(cur) || st.lastOk) {
      if (!doneIds.includes(cur.id)) nextDone = [...doneIds, cur.id]
    } else if (mode !== 'test') {
      nextQueue = [...queue, cur] // błędne pytanie wraca na koniec
    } else {
      nextDone = [...doneIds, cur.id]
    }
    setDoneIds(nextDone)
    setQueue(nextQueue)
    st.lastOk = false

    if (mode === 'test' && st.mistakes > maxMistakes) return finish()
    if (mode === 'lesson' && s.hearts <= 0) { setNoHearts(true); return }
    if (pos + 1 >= nextQueue.length) return finish()
    setPos(pos + 1)
  }

  const progress = doneIds.length / total
  return (
    <div className="session">
      <header className="session-top">
        <button className="icon-btn" aria-label="Zamknij" onClick={() => setConfirmExit(true)}>✕</button>
        <Progress value={progress} color={combo >= 3 ? 'var(--yellow)' : 'var(--coral)'} />
        {mode === 'lesson' ? (
          <span className="stat"><Heart off={s.hearts === 0} /> {s.hearts}</span>
        ) : mode === 'test' ? (
          <span className="stat small">{Math.max(0, maxMistakes - stats.current.mistakes + 1)} szans</span>
        ) : <span className="stat small">Powtórka</span>}
      </header>
      {combo >= 3 && <div className="combo-chip">🔥 {combo}</div>}
      {toast && <div className="toast">{toast}</div>}

      <div className="session-main" key={`${pos}-${item.id}`}>
        {item.kind === 'teach' ? (
          <TeachView card={item} title={title} onNext={advance} />
        ) : (
          <ExerciseCard ex={item} onResult={handleResult} onContinue={advance} />
        )}
      </div>

      {confirmExit && (
        <div className="modal-bg">
          <div className="modal">
            <Lilek mood="sad" size={90} />
            <h3>Już uciekasz?</h3>
            <p>Jeśli wyjdziesz teraz, postęp tej lekcji przepadnie.</p>
            <button className="btn big" onClick={() => setConfirmExit(false)}>Zostaję!</button>
            <button className="btn ghost" onClick={onExit}>Zakończ</button>
          </div>
        </div>
      )}
      {noHearts && (
        <div className="modal-bg">
          <div className="modal">
            <Lilek mood="sleep" size={90} />
            <h3>Skończyły się serca</h3>
            <p>Serca odnawiają się same: jedno co 30 minut. Możesz też zrobić powtórkę (+1 serce) albo kupić pełen zapas.</p>
            <button className="btn big" disabled={s.coins < PRICE_HEARTS}
              onClick={() => { if (buyHearts()) { fx('coin'); setNoHearts(false); if (pos + 1 < queue.length) setPos(pos + 1); else finish() } }}>
              Uzupełnij za {PRICE_HEARTS} monet
            </button>
            <button className="btn ghost" onClick={() => finish(true)}>Zakończ lekcję</button>
          </div>
        </div>
      )}
    </div>
  )
}

function TeachView({ card, title, onNext }: { card: TeachCard; title: string; onNext: () => void }) {
  return (
    <div className="exercise">
      <div className="ex-body">
        <Speech mood="wave" text={`Najpierw krótka ściąga: ${title.toLowerCase()}.`} size={72} />
        <div className="teach">
          <h2>{card.title}</h2>
          {card.body.map((p, i) => <p key={i}>{p}</p>)}
          {card.formula && <pre className="formula">{card.formula}</pre>}
          {card.example && (
            <div className="example"><span className="hand">Przykład z życia</span>{card.example}</div>
          )}
        </div>
      </div>
      <div className="footer">
        <button className="btn big" onClick={() => { fx('tap'); onNext() }}>Rozumiem, dalej</button>
      </div>
    </div>
  )
}
