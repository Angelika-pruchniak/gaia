import { units } from './kurs'
import { isExercise, type Exercise, type Lesson } from './types'

export { units }
export * from './types'

export type FlatLesson = Lesson & { unitIndex: number; indexInUnit: number; flatIndex: number }

export const flatLessons: FlatLesson[] = units.flatMap((u, unitIndex) =>
  u.lessons.map((l, indexInUnit) => ({ ...l, unitIndex, indexInUnit, flatIndex: 0 })),
).map((l, i) => ({ ...l, flatIndex: i }))

export const exerciseById = new Map<string, Exercise>()
for (const l of flatLessons) for (const it of l.items) if (isExercise(it)) exerciseById.set(it.id, it)

export const unitExercises = (unitIndex: number): Exercise[] =>
  units[unitIndex].lessons.flatMap((l) => l.items.filter(isExercise))

/** Szybkie pytania (bez przeciągania) do testu poziomującego i sprawdzianów. */
export const quickExercises = (unitIndex: number): Exercise[] =>
  unitExercises(unitIndex).filter((e) => e.kind === 'choice' || e.kind === 'tf' || e.kind === 'fill')

export function shuffle<T>(arr: T[]): T[] {
  const a = arr.slice()
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1))
    ;[a[i], a[j]] = [a[j], a[i]]
  }
  return a
}

export const pick = <T,>(arr: T[], n: number): T[] => shuffle(arr).slice(0, n)
