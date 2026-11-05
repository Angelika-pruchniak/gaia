import { useSyncExternalStore } from 'react'
import { flatLessons, units } from './content'

export const MAX_HEARTS = 5
export const HEART_REGEN_MS = 30 * 60 * 1000
export const PRICE_FREEZE = 50
export const PRICE_HEARTS = 30
const SRS_DAYS = [0, 1, 3, 7, 14, 30]

export type State = {
  v: 1
  name: string
  onboarded: boolean
  dailyGoal: number
  xp: number
  coins: number
  hearts: number
  heartsAt: number
  streak: number
  longestStreak: number
  lastDay: string | null
  freezes: number
  xpByDay: Record<string, number>
  done: Record<string, number>
  perfect: number
  skipTo: number
  chests: Record<string, boolean>
  srs: Record<string, { box: number; due: string }>
  badges: string[]
  settings: { sound: boolean; vibrate: boolean }
}

const fresh = (): State => ({
  v: 1, name: '', onboarded: false, dailyGoal: 20,
  xp: 0, coins: 40, hearts: MAX_HEARTS, heartsAt: Date.now(),
  streak: 0, longestStreak: 0, lastDay: null, freezes: 0, xpByDay: {},
  done: {}, perfect: 0, skipTo: 0, chests: {}, srs: {}, badges: [],
  settings: { sound: true, vibrate: true },
})

const KEY = 'akademia-guliwera-v1'

function load(): State {
  try {
    const raw = localStorage.getItem(KEY)
    if (raw) return { ...fresh(), ...JSON.parse(raw) }
  } catch { /* brak pamięci – zaczynamy od zera */ }
  return fresh()
}

let state: State = regen(load())
const listeners = new Set<() => void>()

function save() {
  try { localStorage.setItem(KEY, JSON.stringify(state)) } catch { /* ignoruj */ }
}

export function update(fn: (s: State) => void) {
  const s = structuredClone(state)
  fn(s)
  state = s
  save()
  listeners.forEach((l) => l())
}

export const getState = () => state
export function useStore(): State {
  return useSyncExternalStore((l) => { listeners.add(l); return () => listeners.delete(l) }, () => state)
}

// ───── daty ─────
export const dayKey = (d = new Date()) =>
  `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`
export function dayDiff(a: string, b: string) {
  const [y1, m1, d1] = a.split('-').map(Number)
  const [y2, m2, d2] = b.split('-').map(Number)
  return Math.round((Date.UTC(y2, m2 - 1, d2) - Date.UTC(y1, m1 - 1, d1)) / 86400000)
}
const addDays = (n: number) => { const d = new Date(); d.setDate(d.getDate() + n); return dayKey(d) }

// ───── serca ─────
function regen(s: State): State {
  if (s.hearts >= MAX_HEARTS) return { ...s, heartsAt: Date.now() }
  const gained = Math.floor((Date.now() - s.heartsAt) / HEART_REGEN_MS)
  if (gained <= 0) return s
  const hearts = Math.min(MAX_HEARTS, s.hearts + gained)
  return { ...s, hearts, heartsAt: hearts >= MAX_HEARTS ? Date.now() : s.heartsAt + gained * HEART_REGEN_MS }
}
export function tickHearts() {
  const n = regen(state)
  if (n !== state && n.hearts !== state.hearts) update((s) => { s.hearts = n.hearts; s.heartsAt = n.heartsAt })
}
export function loseHeart() {
  update((s) => {
    if (s.hearts >= MAX_HEARTS) s.heartsAt = Date.now()
    s.hearts = Math.max(0, s.hearts - 1)
  })
}
export function buyHearts(): boolean {
  if (state.coins < PRICE_HEARTS) return false
  update((s) => { s.coins -= PRICE_HEARTS; s.hearts = MAX_HEARTS; s.heartsAt = Date.now() })
  return true
}
export function buyFreeze(): boolean {
  if (state.coins < PRICE_FREEZE || state.freezes >= 2) return false
  update((s) => { s.coins -= PRICE_FREEZE; s.freezes += 1 })
  return true
}

// ───── seria dni ─────
/** Seria widoczna dziś (0, jeśli przerwana i nie ma czym jej uratować). */
export function visibleStreak(s: State) {
  if (!s.lastDay) return 0
  const gap = dayDiff(s.lastDay, dayKey())
  if (gap <= 1) return s.streak
  return s.freezes >= gap - 1 ? s.streak : 0
}
export const didToday = (s: State) => s.lastDay === dayKey()

function touchStreak(s: State): boolean {
  const t = dayKey()
  if (s.lastDay === t) return false
  if (!s.lastDay) s.streak = 1
  else {
    const gap = dayDiff(s.lastDay, t)
    if (gap === 1) s.streak += 1
    else if (gap > 1 && s.freezes >= gap - 1) { s.freezes -= gap - 1; s.streak += 1 }
    else s.streak = 1
  }
  s.lastDay = t
  s.longestStreak = Math.max(s.longestStreak, s.streak)
  return true
}

// ───── poziomy ─────
export const LEVELS: [number, string][] = [
  [0, 'Liliput'], [60, 'Praktykant'], [150, 'Asystent'], [300, 'Analityk'], [500, 'Koordynator'],
  [750, 'Kierownik'], [1050, 'Manager'], [1400, 'Dyrektor'], [1850, 'Strateg'],
  [2400, 'Członek Zarządu'], [3100, 'Prezes'], [4000, 'Guliwer'],
]
export function levelOf(xp: number) {
  let i = 0
  while (i + 1 < LEVELS.length && xp >= LEVELS[i + 1][0]) i++
  const [from, name] = LEVELS[i]
  const next = LEVELS[i + 1]?.[0] ?? null
  return { index: i + 1, name, from, next, progress: next ? (xp - from) / (next - from) : 1 }
}

// ───── odznaki ─────
export const BADGES: { id: string; name: string; desc: string; test: (s: State) => boolean }[] = [
  { id: 'first', name: 'Pierwszy krok', desc: 'Ukończ pierwszą lekcję', test: (s) => Object.keys(s.done).length >= 1 },
  { id: 'placed', name: 'Wiem, co wiem', desc: 'Zrób test poziomujący', test: (s) => s.badges.includes('placed') },
  { id: 'streak3', name: 'Rozgrzewka', desc: 'Seria 3 dni', test: (s) => s.longestStreak >= 3 },
  { id: 'streak7', name: 'Tydzień w siodle', desc: 'Seria 7 dni', test: (s) => s.longestStreak >= 7 },
  { id: 'streak30', name: 'Nawyk lidera', desc: 'Seria 30 dni', test: (s) => s.longestStreak >= 30 },
  { id: 'perfect5', name: 'Bez pudła', desc: '5 lekcji bez błędu', test: (s) => s.perfect >= 5 },
  { id: 'money', name: 'Mówię po finansowemu', desc: 'Ukończ dział „Język pieniędzy”', test: (s) => unitDone(s, 0) },
  { id: 'consult', name: 'Głowa konsultanta', desc: 'Ukończ dział „Myśl jak konsultant”', test: (s) => unitDone(s, 4) },
  { id: 'xp1000', name: 'Tysiącznik', desc: 'Zdobądź 1000 XP', test: (s) => s.xp >= 1000 },
  { id: 'all', name: 'Guliwer w biznesie', desc: 'Ukończ wszystkie działy', test: (s) => units.every((_, i) => unitDone(s, i)) },
]
function unitDone(s: State, u: number) {
  return units[u].lessons.every((l) => s.done[l.id])
}
function grantBadges(s: State): string[] {
  const fresh = BADGES.filter((b) => !s.badges.includes(b.id) && b.test(s)).map((b) => b.id)
  s.badges.push(...fresh)
  return fresh
}

// ───── ścieżka ─────
export type LessonStatus = 'done' | 'open' | 'locked'
export function lessonStatus(s: State, flatIndex: number): LessonStatus {
  const l = flatLessons[flatIndex]
  if (s.done[l.id]) return 'done'
  if (flatIndex === 0) return 'open'
  if (l.unitIndex < s.skipTo) return 'open'
  if (l.unitIndex === s.skipTo && l.indexInUnit === 0) return 'open'
  return s.done[flatLessons[flatIndex - 1].id] ? 'open' : 'locked'
}
export function currentLesson(s: State): number {
  const firstAhead = flatLessons.find((l) => l.unitIndex >= s.skipTo && lessonStatus(s, l.flatIndex) === 'open')
  if (firstAhead) return firstAhead.flatIndex
  const any = flatLessons.find((l) => lessonStatus(s, l.flatIndex) === 'open')
  return any ? any.flatIndex : -1
}
export const unitReached = (s: State, u: number) =>
  u <= s.skipTo || units[u - 1].lessons.every((l) => s.done[l.id])
export const chestReady = (s: State, u: number) =>
  !s.chests[units[u].id] && (u < s.skipTo || unitDone(s, u))

// ───── nagrody ─────
export type Rewards = {
  xp: number
  coins: number
  streakUp: boolean
  streak: number
  levelUp: string | null
  badges: string[]
  goalReached: boolean
}

type Summary = { exerciseIds: string[]; firstTryWrong: string[]; maxCombo: number }

function updateSrs(s: State, sum: Summary) {
  for (const id of sum.exerciseIds) {
    const cur = s.srs[id] ?? { box: 0, due: dayKey() }
    const box = sum.firstTryWrong.includes(id) ? 0 : Math.min(cur.box + 1, SRS_DAYS.length - 1)
    s.srs[id] = { box, due: addDays(SRS_DAYS[box]) }
  }
}

function finish(s: State, xp: number, coins: number): Rewards {
  const lvlBefore = levelOf(s.xp).index
  const t = dayKey()
  const before = s.xpByDay[t] ?? 0
  s.xp += xp
  s.coins += coins
  s.xpByDay[t] = before + xp
  const streakUp = touchStreak(s)
  if (streakUp && (s.streak === 7 || s.streak % 30 === 0)) s.coins += 25
  const lvl = levelOf(s.xp)
  return {
    xp, coins, streakUp, streak: s.streak,
    levelUp: lvl.index > lvlBefore ? lvl.name : null,
    badges: grantBadges(s),
    goalReached: before < s.dailyGoal && before + xp >= s.dailyGoal,
  }
}

export function completeLesson(lessonId: string, sum: Summary): Rewards {
  let r!: Rewards
  update((s) => {
    const perfect = sum.firstTryWrong.length === 0
    const repeat = !!s.done[lessonId]
    s.done[lessonId] = (s.done[lessonId] ?? 0) + 1
    if (perfect) s.perfect += 1
    updateSrs(s, sum)
    const xp = (repeat ? 6 : 10) + (perfect ? 5 : 0) + Math.min(5, Math.floor(sum.maxCombo / 3))
    r = finish(s, xp, (repeat ? 3 : 5) + (perfect ? 5 : 0))
  })
  return r
}

export function completeReview(sum: Summary): Rewards {
  let r!: Rewards
  update((s) => {
    updateSrs(s, sum)
    s.hearts = Math.min(MAX_HEARTS, s.hearts + 1)
    r = finish(s, 5 + (sum.exerciseIds.length - sum.firstTryWrong.length), 3)
  })
  return r
}

export function passTest(toUnit: number, placement: boolean): Rewards {
  let r!: Rewards
  update((s) => {
    s.skipTo = Math.max(s.skipTo, toUnit)
    const isNew = placement && !s.badges.includes('placed')
    if (isNew) s.badges.push('placed')
    r = finish(s, placement ? 10 : 15, placement ? 0 : 10)
    if (isNew) r.badges.unshift('placed')
  })
  return r
}

export function openChest(u: number): number {
  const amount = 15 + Math.floor(Math.random() * 16)
  update((s) => { s.chests[units[u].id] = true; s.coins += amount })
  return amount
}

export function dueReviewIds(s: State): string[] {
  const t = dayKey()
  return Object.entries(s.srs).filter(([, v]) => v.due <= t).sort((a, b) => a[1].box - b[1].box).map(([id]) => id)
}

export function resetAll() {
  state = fresh()
  save()
  listeners.forEach((l) => l())
}
