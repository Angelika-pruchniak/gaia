import { useMemo, useState } from 'react'
import { shuffle, type Exercise } from '../content'
import { fx } from '../fx'
import { Lilek, oops, praise } from './Lilek'

const LABEL: Record<Exercise['kind'], string> = {
  choice: 'Wybierz odpowiedź',
  tf: 'Prawda czy fałsz?',
  fill: 'Uzupełnij lukę',
  match: 'Połącz w pary',
  order: 'Ułóż we właściwej kolejności',
  sort: 'Rozłóż do szuflad',
}

type Props = {
  ex: Exercise
  onResult: (correct: boolean) => void
  onContinue: () => void
  allowSkip?: boolean
}

export function ExerciseCard({ ex, onResult, onContinue, allowSkip }: Props) {
  const [phase, setPhase] = useState<'answer' | 'feedback'>('answer')
  const [correct, setCorrect] = useState(false)
  const [ready, setReady] = useState(false)
  const [checker, setChecker] = useState<() => boolean>(() => () => false)
  const [msg, setMsg] = useState('')

  const reveal = (ok: boolean) => {
    setCorrect(ok)
    setMsg(ok ? praise() : oops())
    setPhase('feedback')
    fx(ok ? 'correct' : 'wrong')
    onResult(ok)
  }

  const register = (isReady: boolean, check: () => boolean) => {
    setReady(isReady)
    setChecker(() => check)
  }

  const solution = solutionText(ex)
  const locked = phase === 'feedback'

  return (
    <div className="exercise">
      <div className="ex-body">
        <div className="ex-label">{LABEL[ex.kind]}</div>
        <h2 className="ex-q">{renderQ(ex)}</h2>
        {(ex.kind === 'choice' || ex.kind === 'fill') && <ChoiceView ex={ex} locked={locked} register={register} />}
        {ex.kind === 'tf' && <TFView answer={ex.answer} locked={locked} register={register} />}
        {ex.kind === 'order' && <OrderView tokens={ex.tokens} locked={locked} register={register} />}
        {ex.kind === 'sort' && <SortView ex={ex} locked={locked} register={register} />}
        {ex.kind === 'match' && <MatchView pairs={ex.pairs} done={() => reveal(true)} />}
      </div>

      <div className={`footer ${phase === 'feedback' ? (correct ? 'ok' : 'bad') : ''}`}>
        {phase === 'feedback' ? (
          <>
            <div className="fb">
              <Lilek mood={correct ? 'cheer' : 'sad'} size={56} />
              <div>
                <div className="fb-title">{msg}</div>
                {!correct && solution && <div className="fb-sol">Poprawnie: <b>{solution}</b></div>}
                {ex.explain && <div className="fb-exp">{ex.explain}</div>}
              </div>
            </div>
            <button className={`btn big ${correct ? 'green' : 'red'}`} onClick={() => { fx('tap'); onContinue() }} autoFocus>
              Dalej
            </button>
          </>
        ) : ex.kind === 'match' ? (
          <div className="hint">Dotknij pojęcia, a potem jego pary.</div>
        ) : (
          <div className="row">
            {allowSkip && (
              <button className="btn ghost" onClick={() => reveal(false)}>Nie wiem</button>
            )}
            <button className="btn big" disabled={!ready} onClick={() => reveal(checker())}>
              Sprawdź
            </button>
          </div>
        )}
      </div>
    </div>
  )
}

function renderQ(ex: Exercise) {
  if (ex.kind !== 'fill') return ex.q
  const [a, b] = ex.q.split('___')
  return <>{a}<span className="blank">____</span>{b}</>
}

function solutionText(ex: Exercise): string | null {
  switch (ex.kind) {
    case 'choice':
    case 'fill': return ex.options[0]
    case 'tf': return ex.answer ? 'Prawda' : 'Fałsz'
    case 'order': return ex.tokens.join(' → ')
    case 'sort': return ex.items.map(([t, b]) => `${t}: ${ex.buckets[b]}`).join(' · ')
    default: return null
  }
}

type Reg = (ready: boolean, check: () => boolean) => void

function ChoiceView({ ex, locked, register }: { ex: Exercise & { options: string[] }; locked: boolean; register: Reg }) {
  const opts = useMemo(() => shuffle(ex.options), [ex.id])
  const [sel, setSel] = useState<number | null>(null)
  return (
    <div className="options">
      {opts.map((o, i) => {
        const isRight = o === ex.options[0]
        const cls = locked ? (isRight ? 'right' : sel === i ? 'wrong' : 'dim') : sel === i ? 'sel' : ''
        return (
          <button
            key={o}
            className={`opt ${cls}`}
            disabled={locked}
            onClick={() => { fx('tap'); setSel(i); register(true, () => opts[i] === ex.options[0]) }}
          >
            <span className="opt-key">{i + 1}</span>{o}
          </button>
        )
      })}
    </div>
  )
}

function TFView({ answer, locked, register }: { answer: boolean; locked: boolean; register: Reg }) {
  const [sel, setSel] = useState<boolean | null>(null)
  return (
    <div className="tf">
      {[true, false].map((v) => {
        const cls = locked ? (v === answer ? 'right' : sel === v ? 'wrong' : 'dim') : sel === v ? 'sel' : ''
        return (
          <button key={String(v)} className={`opt tf-btn ${cls}`} disabled={locked}
            onClick={() => { fx('tap'); setSel(v); register(true, () => v === answer) }}>
            {v ? '✓ Prawda' : '✗ Fałsz'}
          </button>
        )
      })}
    </div>
  )
}

function OrderView({ tokens, locked, register }: { tokens: string[]; locked: boolean; register: Reg }) {
  const bank = useMemo(() => {
    let s = shuffle(tokens.map((_, i) => i))
    while (tokens.length > 1 && s.every((v, i) => v === i)) s = shuffle(s)
    return s
  }, [tokens])
  const [picked, setPicked] = useState<number[]>([])
  const set = (next: number[]) => {
    setPicked(next)
    register(next.length === tokens.length, () => next.every((v, i) => v === i))
  }
  return (
    <div className="order">
      <ol className="order-line">
        {picked.length === 0 && <li className="placeholder">Dotykaj kafelków poniżej po kolei</li>}
        {picked.map((t, pos) => (
          <li key={t}>
            <button className={`chip ${locked ? (t === pos ? 'right' : 'wrong') : ''}`} disabled={locked}
              onClick={() => { fx('tap'); set(picked.filter((x) => x !== t)) }}>
              <span className="chip-n">{pos + 1}</span>{tokens[t]}
            </button>
          </li>
        ))}
      </ol>
      <div className="bank">
        {bank.map((t) => (
          <button key={t} className={`chip ${picked.includes(t) ? 'used' : ''}`}
            disabled={locked || picked.includes(t)} onClick={() => { fx('pop'); set([...picked, t]) }}>
            {tokens[t]}
          </button>
        ))}
      </div>
    </div>
  )
}

function SortView({ ex, locked, register }: { ex: Exercise & { kind: 'sort' }; locked: boolean; register: Reg }) {
  const items = useMemo(() => shuffle(ex.items), [ex.id])
  const [choice, setChoice] = useState<Record<string, number>>({})
  return (
    <div className="sort">
      {items.map(([text, bucket]) => (
        <div key={text} className="sort-row">
          <div className="sort-text">{text}</div>
          <div className="sort-btns">
            {ex.buckets.map((b, bi) => {
              const chosen = choice[text] === bi
              const cls = locked ? (bi === bucket ? (chosen ? 'right' : 'should') : chosen ? 'wrong' : 'dim') : chosen ? 'sel' : ''
              return (
                <button key={b} className={`seg ${cls}`} disabled={locked}
                  onClick={() => {
                    fx('tap')
                    const next = { ...choice, [text]: bi }
                    setChoice(next)
                    register(Object.keys(next).length === items.length, () => items.every(([t, bk]) => next[t] === bk))
                  }}>
                  {b}
                </button>
              )
            })}
          </div>
        </div>
      ))}
    </div>
  )
}

function MatchView({ pairs, done }: { pairs: [string, string][]; done: () => void }) {
  const left = useMemo(() => shuffle(pairs.map((_, i) => i)), [pairs])
  const right = useMemo(() => shuffle(pairs.map((_, i) => i)), [pairs])
  const [selL, setSelL] = useState<number | null>(null)
  const [selR, setSelR] = useState<number | null>(null)
  const [matched, setMatched] = useState<number[]>([])
  const [shake, setShake] = useState<[number, number] | null>(null)

  const attempt = (l: number | null, r: number | null) => {
    if (l === null || r === null) return
    if (l === r) {
      const next = [...matched, l]
      setMatched(next)
      setSelL(null); setSelR(null)
      if (next.length === pairs.length) done()
      else fx('pop')
    } else {
      fx('wrong')
      setShake([l, r])
      setTimeout(() => { setShake(null); setSelL(null); setSelR(null) }, 450)
    }
  }

  const cls = (side: 'l' | 'r', i: number) => {
    if (matched.includes(i)) return 'matched'
    if (shake && (side === 'l' ? shake[0] : shake[1]) === i) return 'wrong shake'
    return (side === 'l' ? selL : selR) === i ? 'sel' : ''
  }

  return (
    <div className="match">
      <div className="match-col">
        {left.map((i) => (
          <button key={i} className={`opt match-btn ${cls('l', i)}`} disabled={matched.includes(i)}
            onClick={() => { fx('tap'); setSelL(i); attempt(i, selR) }}>
            {pairs[i][0]}
          </button>
        ))}
      </div>
      <div className="match-col">
        {right.map((i) => (
          <button key={i} className={`opt match-btn ${cls('r', i)}`} disabled={matched.includes(i)}
            onClick={() => { fx('tap'); setSelR(i); attempt(selL, i) }}>
            {pairs[i][1]}
          </button>
        ))}
      </div>
    </div>
  )
}
