import { describe, expect, it } from 'vitest'
import { flatLessons, quickExercises, units } from './content'
import { isExercise } from './content/types'

describe('treść kursu', () => {
  const all = flatLessons.flatMap((l) => l.items)

  it('ma unikalne identyfikatory', () => {
    expect(new Set(all.map((i) => i.id)).size).toBe(all.length)
  })

  it('każda lekcja zaczyna się od ściągi i ma min. 4 ćwiczenia', () => {
    for (const l of flatLessons) {
      expect(l.items[0].kind, l.id).toBe('teach')
      expect(l.items.filter(isExercise).length, l.id).toBeGreaterThanOrEqual(4)
    }
  })

  it('ćwiczenia są poprawnie zbudowane', () => {
    for (const e of all.filter(isExercise)) {
      if (e.kind === 'choice' || e.kind === 'fill') {
        expect(e.options.length, e.id).toBeGreaterThanOrEqual(3)
        expect(new Set(e.options).size, e.id).toBe(e.options.length)
      }
      if (e.kind === 'fill') expect(e.q.split('___').length, e.id).toBe(2)
      if (e.kind === 'match') {
        expect(new Set(e.pairs.map((p) => p[0])).size, e.id).toBe(e.pairs.length)
        expect(new Set(e.pairs.map((p) => p[1])).size, e.id).toBe(e.pairs.length)
      }
      if (e.kind === 'order') expect(new Set(e.tokens).size, e.id).toBe(e.tokens.length)
      if (e.kind === 'sort') {
        expect(new Set(e.items.map((i) => i[0])).size, e.id).toBe(e.items.length)
        for (const [, b] of e.items) expect(b, e.id).toBeLessThan(e.buckets.length)
        // każda szuflada użyta co najmniej raz
        e.buckets.forEach((_, bi) => expect(e.items.some(([, b]) => b === bi), e.id).toBe(true))
      }
    }
  })

  it('każdy dział ma pytania do testu poziomującego', () => {
    units.forEach((_, i) => expect(quickExercises(i).length).toBeGreaterThanOrEqual(4))
  })
})
