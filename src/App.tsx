import { AppShell } from '@/components/layout/AppShell'
import { CommandPalette } from '@/components/common/CommandPalette'
import { DropZone } from '@/components/DropZone'
import { FileList } from '@/components/FileList'
import { DataViewer } from '@/components/DataViewer'
import { ExportPanel } from '@/components/ExportPanel'
import { SettingsPanel } from '@/components/settings/SettingsPanel'
import { useUIStore } from '@/stores/ui-store'
import { useFileStore } from '@/stores/file-store'
import { useSettingsStore } from '@/stores/settings-store'
import { useKeyboard } from '@/hooks/use-keyboard'
import { getTranslations } from '@/i18n'
import { Shield, FileUp, Download } from 'lucide-react'

function AppContent() {
  useKeyboard()
  const activeView = useUIStore((s) => s.activeView)
  const error = useFileStore((s) => s.error)
  const { settings } = useSettingsStore()
  const t = getTranslations(settings.language)

  return (
    <AppShell>
      <CommandPalette />

      {/* Toast / error */}
      {error && (
        <div className="mx-4 mt-4 p-3 rounded-xl bg-danger/10 border border-danger/30 text-sm text-danger">
          {error}
        </div>
      )}

      <ToastOverlay />

      {activeView === 'extractor' && <ExtractorView t={t} />}
      {activeView === 'viewer' && <DataViewer />}
      {activeView === 'export' && <ExportPanel />}
      {activeView === 'settings' && <SettingsPanel />}
    </AppShell>
  )
}

function ExtractorView({ t }: { t: ReturnType<typeof getTranslations> }) {
  const presentations = useFileStore((s) => s.presentations)

  return (
    <div className="max-w-3xl mx-auto p-4 pb-24 md:pb-4 space-y-6 animate-fade-in">
      {/* Hero */}
      <div className="text-center py-4">
        <h2 className="text-2xl md:text-3xl font-bold text-text">{t.extractData}</h2>
        <p className="text-text-3 mt-1">{t.supportedFormats}</p>
      </div>

      <DropZone />
      <FileList />

      {/* Feature cards when empty */}
      {presentations.length === 0 && (
        <div className="grid grid-cols-1 md:grid-cols-3 gap-3 pt-4">
          {([
            { icon: FileUp, title: t.featureExtract, desc: t.featureExtractDesc },
            { icon: Download, title: t.featureExport, desc: t.featureExportDesc },
            { icon: Shield, title: t.featureMedia, desc: t.featureMediaDesc },
          ] as const).map(({ icon: Icon, title, desc }) => (
            <div key={title} className="p-4 rounded-xl bg-surface-2/50 text-center">
              <div className="w-12 h-12 mx-auto mb-3 rounded-xl bg-accent/10 flex items-center justify-center">
                <Icon size={24} className="text-accent" />
              </div>
              <h3 className="font-semibold text-text text-sm">{title}</h3>
              <p className="text-xs text-text-3 mt-1">{desc}</p>
            </div>
          ))}
        </div>
      )}
    </div>
  )
}

function ToastOverlay() {
  const toast = useUIStore((s) => s.toastMessage)
  const toastType = useUIStore((s) => s.toastType)

  if (!toast) return null

  const colors = {
    success: 'bg-success text-white',
    error: 'bg-danger text-white',
    info: 'bg-surface-3 text-text',
  }

  return (
    <div className="fixed top-4 right-4 z-[60] animate-slide-down">
      <div className={`px-4 py-2 rounded-xl shadow-lg text-sm font-medium ${colors[toastType]}`}>
        {toast}
      </div>
    </div>
  )
}

export default AppContent
