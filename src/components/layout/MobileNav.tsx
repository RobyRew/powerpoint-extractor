import {
  FileUp,
  Download,
  Settings,
  Eye,
} from 'lucide-react'
import { useUIStore, type ActiveView } from '@/stores/ui-store'
import { useSettingsStore } from '@/stores/settings-store'
import { useFileStore } from '@/stores/file-store'
import { getTranslations } from '@/i18n'
import { cn } from '@/lib/utils'

const TABS: { view: ActiveView; icon: typeof FileUp; labelKey: string }[] = [
  { view: 'extractor', icon: FileUp, labelKey: 'extractor' },
  { view: 'viewer', icon: Eye, labelKey: 'viewer' },
  { view: 'export', icon: Download, labelKey: 'export_' },
  { view: 'settings', icon: Settings, labelKey: 'settings' },
]

export function MobileNav() {
  const { activeView, setActiveView } = useUIStore()
  const { settings } = useSettingsStore()
  const presentations = useFileStore((s) => s.presentations)
  const t = getTranslations(settings.language)

  return (
    <nav className="fixed bottom-0 left-0 right-0 z-30 bg-surface/90 backdrop-blur-xl border-t border-border safe-bottom">
      <div className="flex items-center justify-around h-12">
        {TABS.map(({ view, icon: Icon, labelKey }) => {
          const disabled = (view === 'viewer' || view === 'export') && presentations.length === 0
          return (
            <button
              key={view}
              onClick={() => !disabled && setActiveView(view)}
              disabled={disabled}
              className={cn(
                'flex flex-col items-center gap-0.5 px-3 py-1 transition-colors min-w-[48px]',
                activeView === view ? 'text-accent' : disabled ? 'text-text-3 opacity-40' : 'text-text-3',
              )}
            >
              <Icon size={20} />
              <span className="text-[10px] font-medium">{t[labelKey as keyof typeof t] as string}</span>
            </button>
          )
        })}
      </div>
    </nav>
  )
}
