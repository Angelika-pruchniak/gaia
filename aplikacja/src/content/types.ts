// Konwencja: w pytaniach "choice" i "fill" PIERWSZA opcja jest poprawna.
// Aplikacja i tak tasuje kolejność przy wyświetlaniu.

export type Choice = { kind: 'choice'; q: string; options: string[]; explain?: string }
export type TrueFalse = { kind: 'tf'; q: string; answer: boolean; explain?: string }
export type Fill = { kind: 'fill'; q: string; options: string[]; explain?: string }
export type Match = { kind: 'match'; q: string; pairs: [string, string][]; explain?: string }
export type Order = { kind: 'order'; q: string; tokens: string[]; explain?: string }
export type Sort = {
  kind: 'sort'
  q: string
  buckets: string[]
  items: [string, number][]
  explain?: string
}
export type Teach = {
  kind: 'teach'
  title: string
  body: string[]
  formula?: string
  example?: string
}

export type RawExercise = Choice | TrueFalse | Fill | Match | Order | Sort
export type Exercise = RawExercise & { id: string }
export type TeachCard = Teach & { id: string }
export type Item = Exercise | TeachCard

export type Lesson = { id: string; title: string; items: Item[] }
export type Unit = {
  id: string
  title: string
  subtitle: string
  color: string
  lessons: Lesson[]
}

export const isExercise = (i: Item): i is Exercise => i.kind !== 'teach'
