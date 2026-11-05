import { useMemo } from 'react'

export function Flame({ size = 22, off = false }: { size?: number; off?: boolean }) {
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" aria-hidden>
      <path d="M12 2c1 4 6 6 6 12a6 6 0 0 1-12 0c0-3 1.5-5 3-6 0 2 1 3 2 3-1-3 0-6 1-9z" fill={off ? '#CFC4D2' : '#FF6B3D'} />
      <path d="M12 12c.5 2 3 3 3 5.5a3 3 0 0 1-6 0c0-1.5 1-2.5 1.6-3 .2 1 .6 1.4 1 1.5-.3-1.6 0-3 .4-4z" fill={off ? '#E6DEE8' : '#FFC84A'} />
    </svg>
  )
}
export function Heart({ size = 22, off = false }: { size?: number; off?: boolean }) {
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" aria-hidden>
      <path d="M12 21s-8-5.2-8-11a4.5 4.5 0 0 1 8-2.8A4.5 4.5 0 0 1 20 10c0 5.8-8 11-8 11z" fill={off ? '#CFC4D2' : '#F4799F'} />
      <ellipse cx="8.5" cy="9" rx="1.8" ry="1.2" fill="#fff" opacity=".6" />
    </svg>
  )
}
export function Coin({ size = 22 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" aria-hidden>
      <circle cx="12" cy="12" r="10" fill="#FFC84A" />
      <circle cx="12" cy="12" r="7" fill="none" stroke="#E8A92E" strokeWidth="1.6" />
      <text x="12" y="16" textAnchor="middle" fontSize="10" fontWeight="700" fill="#B07A10" fontFamily="DM Sans, sans-serif">4</text>
    </svg>
  )
}
export function Star({ size = 22, off = false }: { size?: number; off?: boolean }) {
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" aria-hidden>
      <path d="M12 2.5l2.9 6 6.6.8-4.9 4.6 1.3 6.6L12 17.3 6.1 20.5l1.3-6.6L2.5 9.3l6.6-.8z" fill={off ? '#CFC4D2' : '#FFC84A'} />
    </svg>
  )
}
export function Snow({ size = 22 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" aria-hidden stroke="#7FA7E8" strokeWidth="2" strokeLinecap="round">
      <path d="M12 2v20M3.3 7l17.4 10M3.3 17L20.7 7" />
    </svg>
  )
}

const COLORS = ['#FF6B3D', '#F4799F', '#B9A5E3', '#FFC84A', '#FFD8C2', '#1E2A4A']

export function Confetti({ count = 70 }: { count?: number }) {
  const bits = useMemo(
    () =>
      Array.from({ length: count }, (_, i) => ({
        left: Math.random() * 100,
        delay: Math.random() * 0.6,
        dur: 1.8 + Math.random() * 1.6,
        rot: Math.random() * 360,
        color: COLORS[i % COLORS.length],
        w: 6 + Math.random() * 6,
        round: Math.random() > 0.6,
      })),
    [count],
  )
  return (
    <div className="confetti" aria-hidden>
      {bits.map((b, i) => (
        <span
          key={i}
          style={{
            left: `${b.left}%`,
            background: b.color,
            width: b.w,
            height: b.round ? b.w : b.w * 0.45,
            borderRadius: b.round ? '50%' : 2,
            animationDelay: `${b.delay}s`,
            animationDuration: `${b.dur}s`,
            transform: `rotate(${b.rot}deg)`,
          }}
        />
      ))}
    </div>
  )
}

export function Progress({ value, color = 'var(--coral)' }: { value: number; color?: string }) {
  return (
    <div className="progress">
      <div className="progress-fill" style={{ width: `${Math.max(4, Math.min(100, value * 100))}%`, background: color }} />
    </div>
  )
}
