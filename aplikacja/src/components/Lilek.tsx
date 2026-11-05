export type Mood = 'happy' | 'cheer' | 'sad' | 'think' | 'wave' | 'sleep'

const INK = '#3A2340'

/** Lilek – maskotka Akademii: mały Liliput w trójgraniastym kapeluszu i krawacie managera. */
export function Lilek({ mood = 'happy', size = 120, className = '' }: { mood?: Mood; size?: number; className?: string }) {
  const closed = mood === 'cheer' || mood === 'sleep'
  const pupilDy = mood === 'think' ? -3 : mood === 'sad' ? 2 : 0
  const pupilDx = mood === 'think' ? 2 : 0
  return (
    <svg
      viewBox="0 0 120 132"
      width={size}
      height={size * 1.1}
      className={`lilek lilek-${mood} ${className}`}
      role="img"
      aria-label="Lilek, maskotka Akademii"
    >
      <ellipse cx="60" cy="127" rx="30" ry="4" fill={INK} opacity="0.12" />
      {/* stopy */}
      <ellipse cx="46" cy="122" rx="10" ry="5" fill="#E8512A" />
      <ellipse cx="74" cy="122" rx="10" ry="5" fill="#E8512A" />
      {/* ręce */}
      {mood === 'cheer' ? (
        <g className="lilek-arms-up">
          <ellipse cx="17" cy="56" rx="7" ry="14" fill="#E8512A" transform="rotate(-28 17 56)" />
          <ellipse cx="103" cy="56" rx="7" ry="14" fill="#E8512A" transform="rotate(28 103 56)" />
        </g>
      ) : (
        <>
          <g className="lilek-arm-l">
            <ellipse cx="20" cy="88" rx="7" ry="12" fill="#E8512A" transform="rotate(18 20 88)" />
          </g>
          <g className="lilek-arm-r">
            <ellipse cx="100" cy="88" rx="7" ry="12" fill="#E8512A" transform="rotate(-18 100 88)" />
          </g>
        </>
      )}
      {/* ciało */}
      <path d="M60 32C88 32 101 54 101 81C101 108 85 123 60 123C35 123 19 108 19 81C19 54 32 32 60 32Z" fill="#FF6B3D" />
      <ellipse cx="60" cy="100" rx="25" ry="19" fill="#FFF6EE" />
      {/* krawat */}
      <path d="M55 86h10l-2 4 4 16-7 6-7-6 4-16z" fill="#B9A5E3" />
      <path d="M55 86h10l-2 4h-6z" fill="#8F79C9" />
      {/* policzki */}
      <ellipse cx="34" cy="80" rx="6" ry="3.5" fill="#F4799F" opacity="0.75" />
      <ellipse cx="86" cy="80" rx="6" ry="3.5" fill="#F4799F" opacity="0.75" />
      {/* oczy */}
      {closed ? (
        mood === 'cheer' ? (
          <g stroke={INK} strokeWidth="3.4" fill="none" strokeLinecap="round">
            <path d="M39 70Q46 61 53 70" />
            <path d="M67 70Q74 61 81 70" />
          </g>
        ) : (
          <g stroke={INK} strokeWidth="3" fill="none" strokeLinecap="round">
            <path d="M39 68Q46 73 53 68" />
            <path d="M67 68Q74 73 81 68" />
          </g>
        )
      ) : (
        <g className="lilek-eyes">
          <ellipse cx="46" cy="67" rx="9" ry="10" fill="#fff" />
          <ellipse cx="74" cy="67" rx="9" ry="10" fill="#fff" />
          <circle cx={47 + pupilDx} cy={69 + pupilDy} r="4.8" fill={INK} />
          <circle cx={73 + pupilDx} cy={69 + pupilDy} r="4.8" fill={INK} />
          <circle cx={48.6 + pupilDx} cy={67 + pupilDy} r="1.6" fill="#fff" />
          <circle cx={74.6 + pupilDx} cy={67 + pupilDy} r="1.6" fill="#fff" />
        </g>
      )}
      {/* brwi */}
      {mood === 'sad' && (
        <g stroke={INK} strokeWidth="2.6" strokeLinecap="round">
          <path d="M38 55l12 3" />
          <path d="M82 55l-12 3" />
        </g>
      )}
      {mood === 'think' && (
        <g stroke={INK} strokeWidth="2.6" strokeLinecap="round" fill="none">
          <path d="M38 54q8-4 15 0" />
          <path d="M68 52q8-2 14 2" />
        </g>
      )}
      {/* usta */}
      {(mood === 'happy' || mood === 'wave') && (
        <path d="M51 83Q60 91 69 83" stroke={INK} strokeWidth="3" fill="none" strokeLinecap="round" />
      )}
      {mood === 'cheer' && (
        <g>
          <path d="M48 80Q60 98 72 80Z" fill={INK} />
          <path d="M53 88Q60 95 67 88Q60 85 53 88Z" fill="#F4799F" />
        </g>
      )}
      {mood === 'sad' && <path d="M52 88Q60 81 68 88" stroke={INK} strokeWidth="3" fill="none" strokeLinecap="round" />}
      {mood === 'think' && <circle cx="62" cy="85" r="3" fill={INK} />}
      {mood === 'sleep' && <path d="M55 84q5 3 10 0" stroke={INK} strokeWidth="2.6" fill="none" strokeLinecap="round" />}
      {/* kapelusz trójgraniasty z piórkiem */}
      <path d="M77 30C83 14 95 7 106 4C102 15 92 27 80 33Z" fill="#FFC84A" />
      <path d="M86 21C91 15 97 11 103 8" stroke="#E8A92E" strokeWidth="1.5" fill="none" />
      <path d="M36 40C37 18 83 18 84 40Z" fill="#1E2A4A" />
      <path d="M37 36h46v5H37z" fill="#B9A5E3" />
      <path d="M14 42Q60 24 106 42Q98 50 60 45Q22 50 14 42Z" fill="#1E2A4A" />
      <circle cx="60" cy="40" r="3.4" fill="#FFC84A" />
      {mood === 'sleep' && (
        <text x="92" y="62" fontFamily="Caveat, cursive" fontSize="18" fill="#7A6580">z z</text>
      )}
      {mood === 'sad' && <path d="M84 74q3 5 0 8q-3-3 0-8z" fill="#B9A5E3" />}
    </svg>
  )
}

export function Speech({ mood = 'happy', text, size = 84 }: { mood?: Mood; text: string; size?: number }) {
  return (
    <div className="speech">
      <Lilek mood={mood} size={size} />
      <div className="bubble">{text}</div>
    </div>
  )
}

const PRAISE = ['Brawo!', 'Świetnie!', 'Tak jest!', 'Jak z podręcznika!', 'Prezes byłby dumny!', 'Mistrzowsko!', 'Bingo!']
const OOPS = ['Prawie!', 'Nic straconego!', 'Zerknij na wyjaśnienie.', 'Każdy Guliwer kiedyś się mylił.']
export const praise = () => PRAISE[Math.floor(Math.random() * PRAISE.length)]
export const oops = () => OOPS[Math.floor(Math.random() * OOPS.length)]
