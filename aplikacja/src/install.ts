import { useEffect, useState } from 'react'

type PromptEvent = Event & { prompt: () => Promise<void>; userChoice: Promise<{ outcome: string }> }
let deferred: PromptEvent | null = null
const subs = new Set<() => void>()

window.addEventListener('beforeinstallprompt', (e) => {
  e.preventDefault()
  deferred = e as PromptEvent
  subs.forEach((f) => f())
})
window.addEventListener('appinstalled', () => { deferred = null; subs.forEach((f) => f()) })

export const isStandalone = () =>
  window.matchMedia('(display-mode: standalone)').matches || (navigator as unknown as { standalone?: boolean }).standalone === true
export const isIOS = () => /iphone|ipad|ipod/i.test(navigator.userAgent)

export function useInstall() {
  const [, force] = useState(0)
  useEffect(() => {
    const f = () => force((n) => n + 1)
    subs.add(f)
    return () => { subs.delete(f) }
  }, [])
  return {
    installed: isStandalone(),
    canPrompt: !!deferred,
    ios: isIOS(),
    prompt: async () => {
      if (!deferred) return
      await deferred.prompt()
      await deferred.userChoice
      deferred = null
      force((n) => n + 1)
    },
  }
}
