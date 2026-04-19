import { useEffect } from 'react'
import { useUIStore } from '@/stores/ui-store'
import { usePlatform } from './use-platform'

export function useKeyboard() {
  const { modKey } = usePlatform()
  const ui = useUIStore()

  useEffect(() => {
    function handler(e: KeyboardEvent) {
      const mod = modKey === 'Meta' ? e.metaKey : e.ctrlKey

      if (mod && e.key === 'k') {
        e.preventDefault()
        ui.setCommandPaletteOpen(!ui.isCommandPaletteOpen)
      }
      if (mod && e.key === ',') {
        e.preventDefault()
        ui.setActiveView('settings')
      }
      if (e.key === 'Escape') {
        if (ui.isCommandPaletteOpen) ui.setCommandPaletteOpen(false)
        else if (ui.viewingPresentationId) ui.setViewingPresentation(null)
      }
    }

    window.addEventListener('keydown', handler)
    return () => window.removeEventListener('keydown', handler)
  }, [modKey, ui])
}
