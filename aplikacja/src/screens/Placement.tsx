import { useMemo, useRef, useState } from 'react'
import { pick, quickExercises, units } from '../content'
import { ExerciseCard } from '../components/ExerciseCard'
import { Progress } from '../components/ui'

const PER_UNIT = 2

/** Adaptacyjny test: 2 pytania na dział. Oba dobrze → kolejny dział, błąd → tu zaczynasz. */
export function Placement({ onDone, onExit }: { onDone: (unit: number) => void; onExit: () => void }) {
  const [unit, setUnit] = useState(0)
  const [step, setStep] = useState(0)
  const ok = useRef(true)
  const qs = useMemo(() => pick(quickExercises(unit), PER_UNIT), [unit])
  const ex = qs[step]

  const next = () => {
    if (!ok.current) return onDone(unit)
    if (step + 1 < PER_UNIT) return setStep(step + 1)
    if (unit + 1 >= units.length) return onDone(units.length - 1)
    setUnit(unit + 1)
    setStep(0)
  }

  return (
    <div className="session">
      <header className="session-top">
        <button className="icon-btn" aria-label="Zamknij" onClick={onExit}>✕</button>
        <Progress value={(unit * PER_UNIT + step) / (units.length * PER_UNIT)} color="var(--lilac)" />
        <span className="stat small">Dział {unit + 1}/{units.length}</span>
      </header>
      <div className="placement-unit">{units[unit].title}</div>
      <div className="session-main" key={`${unit}-${step}`}>
        <ExerciseCard ex={ex} allowSkip onResult={(r) => { ok.current = r }} onContinue={next} />
      </div>
    </div>
  )
}
