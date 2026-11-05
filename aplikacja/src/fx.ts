import { getState } from './store'

// Wszystkie dźwięki są syntezowane na żywo (Web Audio API), więc są w 100% nasze.
let ctx: AudioContext | null = null
function ac(): AudioContext | null {
  try {
    if (!ctx) {
      const C = window.AudioContext ?? (window as unknown as { webkitAudioContext: typeof AudioContext }).webkitAudioContext
      ctx = new C()
    }
    if (ctx.state === 'suspended') void ctx.resume()
    return ctx
  } catch {
    return null
  }
}

function tone(freq: number, at: number, dur: number, type: OscillatorType = 'sine', vol = 0.18) {
  const a = ac()
  if (!a) return
  const t = a.currentTime + at
  const o = a.createOscillator()
  const g = a.createGain()
  o.type = type
  o.frequency.setValueAtTime(freq, t)
  g.gain.setValueAtTime(0.0001, t)
  g.gain.exponentialRampToValueAtTime(vol, t + 0.012)
  g.gain.exponentialRampToValueAtTime(0.0001, t + dur)
  o.connect(g).connect(a.destination)
  o.start(t)
  o.stop(t + dur + 0.02)
}

const bell = (f: number, at: number, dur = 0.35, vol = 0.16) => {
  tone(f, at, dur, 'sine', vol)
  tone(f * 2, at, dur * 0.6, 'triangle', vol * 0.18)
}

const sounds = {
  tap: () => tone(520, 0, 0.05, 'triangle', 0.06),
  pop: () => { tone(740, 0, 0.07, 'sine', 0.12); tone(990, 0.04, 0.08, 'sine', 0.08) },
  correct: () => { bell(880, 0, 0.18); bell(1318.5, 0.09, 0.4) },
  wrong: () => { tone(233, 0, 0.16, 'triangle', 0.14); tone(185, 0.13, 0.28, 'triangle', 0.14) },
  combo: () => [1046.5, 1318.5, 1568].forEach((f, i) => bell(f, i * 0.06, 0.22, 0.1)),
  win: () => {
    ;[523.25, 659.25, 783.99, 1046.5].forEach((f, i) => bell(f, i * 0.11, 0.3))
    ;[523.25, 659.25, 783.99, 1046.5].forEach((f) => bell(f, 0.5, 0.9, 0.08))
  },
  level: () => [392, 523.25, 659.25, 783.99, 1046.5, 1318.5].forEach((f, i) => bell(f, i * 0.08, 0.35, 0.12)),
  coin: () => { bell(1318.5, 0, 0.1, 0.1); bell(1975.5, 0.07, 0.3, 0.1) },
  streak: () => { tone(330, 0, 0.5, 'sawtooth', 0.03); [659.25, 880, 1174.7].forEach((f, i) => bell(f, 0.1 + i * 0.1, 0.4, 0.12)) },
}

export type Sound = keyof typeof sounds

const patterns: Record<string, number | number[]> = {
  tap: 8, pop: 15, correct: 35, wrong: [70, 50, 70], combo: [25, 30, 25],
  win: [40, 40, 40, 40, 160], level: [60, 40, 60, 40, 200], coin: 20, streak: [30, 30, 30, 30, 120],
}

export function fx(name: Sound) {
  const s = getState().settings
  if (s.sound) sounds[name]()
  if (s.vibrate && 'vibrate' in navigator) {
    try { navigator.vibrate(patterns[name]) } catch { /* iOS nie obsługuje wibracji */ }
  }
}

/** Odblokowanie audio na iOS: pierwsze dotknięcie ekranu. */
export function primeAudio() {
  const once = () => { ac(); window.removeEventListener('pointerdown', once) }
  window.addEventListener('pointerdown', once)
}
