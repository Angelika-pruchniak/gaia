import { exerciseById, flatLessons, isExercise, pick, type Exercise } from '../content'
import { Lilek, Speech } from '../components/Lilek'
import { Heart } from '../components/ui'
import { currentLesson, dueReviewIds, useStore } from '../store'

export function practiceItems(s: ReturnType<typeof useStore>) {
  const due = dueReviewIds(s).map((id) => exerciseById.get(id)).filter((e): e is Exercise => !!e)
  if (due.length >= 5) return due.slice(0, 10)
  let learned = flatLessons.filter((l) => s.done[l.id]).flatMap((l) => l.items.filter(isExercise))
  if (learned.length === 0) learned = flatLessons[Math.max(0, currentLesson(s))].items.filter(isExercise)
  const extra = pick(learned.filter((e) => !due.includes(e)), 8 - due.length)
  return [...due, ...extra]
}

export function Practice({ onStart }: { onStart: () => void }) {
  const s = useStore()
  const due = dueReviewIds(s).length
  const anyDone = Object.keys(s.done).length > 0
  return (
    <div className="page">
      <h1 className="page-title">Powtórki</h1>
      {anyDone ? (
        <>
          <Speech mood="think" text={due > 0
            ? `Mam dla Ciebie ${due} ${due === 1 ? 'pytanie' : due < 5 ? 'pytania' : 'pytań'} do odświeżenia. Pamięć lubi powroty!`
            : 'Wszystko świeże! Możesz zrobić trening z losowych pytań.'} />
          <div className="card">
            <h3>Jak to działa?</h3>
            <p>Pytania wracają po 1, 3, 7, 14 i 30 dniach. Te z błędem wracają szybciej. Powtórka nie zabiera serc, a za jej ukończenie dostajesz <b>+1 <Heart size={16} /></b>.</p>
          </div>
          <button className="btn big" onClick={onStart}>{due > 0 ? 'Zacznij powtórkę' : 'Trening losowy'}</button>
        </>
      ) : (
        <div className="center">
          <Lilek mood="sleep" size={120} />
          <p>Tu pojawią się powtórki po Twojej pierwszej lekcji. Na razie możesz potrenować pytania z najbliższej lekcji (bez utraty serc).</p>
          <button className="btn big" onClick={onStart}>Trening na rozgrzewkę</button>
        </div>
      )}
    </div>
  )
}
