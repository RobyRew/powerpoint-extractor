import { Menu, Search, X } from 'lucide-react'
import { useUIStore } from '@/stores/ui-store'
import { useSettingsStore } from '@/stores/settings-store'
import { getTranslations } from '@/i18n'
import { useIsMobile } from '@/hooks/use-media-query'

const logoUrl = new URL('/favicon.svg', import.meta.url).href

export function Header() {
  const {
    toggleSidebar,
    searchQuery,
    setSearchQuery,
  } = useUIStore()
  const { settings } = useSettingsStore()
  const t = getTranslations(settings.language)
  const isMobile = useIsMobile()
  const isSearchOpen = searchQuery.length > 0

  return (
    <header className="h-12 flex items-center gap-2 px-3 border-b border-border bg-surface shrink-0 safe-top">
      {!isMobile && (
        <button
          onClick={toggleSidebar}
          className="p-1.5 rounded-lg hover:bg-surface-2 text-text-2 transition-colors"
        >
          <Menu size={20} />
        </button>
      )}

      <img src={logoUrl} alt="" className="w-6 h-6 shrink-0" />
      <h1 className="text-base font-semibold text-text flex-1 truncate">{t.appName}</h1>

      <div className="flex items-center gap-1">
        {!isMobile && (
          <div className="flex items-center gap-2 mr-2">
            <div className="relative">
              <Search size={14} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-text-3" />
              <input
                type="text"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                placeholder={t.search}
                className="h-8 w-48 rounded-lg bg-surface-2 pl-8 pr-8 text-xs text-text placeholder:text-text-3 outline-none focus:ring-2 focus:ring-accent/20 transition-all"
              />
              {isSearchOpen && (
                <button
                  onClick={() => setSearchQuery('')}
                  className="absolute right-2 top-1/2 -translate-y-1/2 text-text-3 hover:text-text"
                >
                  <X size={12} />
                </button>
              )}
            </div>
          </div>
        )}
      </div>
    </header>
  )
}
