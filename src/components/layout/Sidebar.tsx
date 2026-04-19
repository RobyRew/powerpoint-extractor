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
import { APP_VERSION } from '@/lib/constants'

const NAV_ITEMS: { view: ActiveView; icon: typeof FileUp; labelKey: string }[] = [
  { view: 'extractor', icon: FileUp, labelKey: 'extractor' },
  { view: 'viewer', icon: Eye, labelKey: 'viewer' },
  { view: 'export', icon: Download, labelKey: 'export_' },
  { view: 'settings', icon: Settings, labelKey: 'settings' },
]

export function Sidebar() {
  const { activeView, setActiveView, isSidebarOpen } = useUIStore()
  const { settings } = useSettingsStore()
  const presentations = useFileStore((s) => s.presentations)
  const t = getTranslations(settings.language)

  return (
    <aside
      className={cn(
        'bg-surface-2 border-r border-border flex flex-col shrink-0 transition-all duration-200 overflow-hidden',
        isSidebarOpen ? 'w-56' : 'w-0 border-0',
      )}
    >
      <nav className="flex-1 py-2 px-2 space-y-0.5">
        {NAV_ITEMS.map(({ view, icon: Icon, labelKey }) => {
          const disabled = (view === 'viewer' || view === 'export') && presentations.length === 0
          return (
            <button
              key={view}
              onClick={() => !disabled && setActiveView(view)}
              disabled={disabled}
              className={cn(
                'w-full flex items-center gap-3 px-3 py-2 rounded-lg text-sm font-medium transition-all duration-150',
                activeView === view
                  ? 'bg-accent/15 text-accent'
                  : disabled
                    ? 'text-text-3 opacity-40 cursor-not-allowed'
                    : 'text-text-2 hover:bg-surface-3 active:bg-surface-3',
              )}
            >
              <Icon size={18} />
              <span>{t[labelKey as keyof typeof t] as string}</span>
              {view === 'viewer' && presentations.length > 0 && (
                <span className="ml-auto text-xs px-1.5 py-0.5 rounded-full bg-accent/15 text-accent">
                  {presentations.length}
                </span>
              )}
            </button>
          )
        })}
      </nav>

      <div className="p-3 border-t border-border">
        <p className="text-xs text-text-3 text-center">PPTExtract v{APP_VERSION}</p>
      </div>
    </aside>
  )
}
