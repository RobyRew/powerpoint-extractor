import { X, FileType, Loader2, CheckCircle, Eye } from 'lucide-react'
import { cn, formatFileSize } from '@/lib/utils'
import { useFileStore } from '@/stores/file-store'
import { useUIStore } from '@/stores/ui-store'
import { useSettingsStore } from '@/stores/settings-store'
import { getTranslations } from '@/i18n'
import { Button } from '@/components/ui/Button'
import { Badge } from '@/components/ui/Badge'

export function FileList() {
  const { files, presentations, processingFile, removeFile, clearAll } = useFileStore()
  const { setActiveView, setViewingPresentation, showToast } = useUIStore()
  const { settings } = useSettingsStore()
  const t = getTranslations(settings.language)

  if (files.length === 0) return null

  const getStatus = (name: string) => {
    if (processingFile === name) return 'processing'
    if (presentations.find((p) => p.fileName === name)) return 'done'
    return 'pending'
  }

  const handleView = (fileName: string) => {
    const pres = presentations.find((p) => p.fileName === fileName)
    if (pres) {
      setViewingPresentation(pres.id)
      setActiveView('viewer')
    }
  }

  const handleRemove = (index: number) => {
    if (settings.confirmBeforeDelete && !confirm(t.removeFile + '?')) return
    removeFile(index)
  }

  const handleClearAll = () => {
    if (settings.confirmBeforeClear && !confirm(t.clearAllConfirm)) return
    clearAll()
    showToast(t.filesCleared, 'info')
  }

  return (
    <div className="rounded-xl bg-surface-2/50 p-4">
      <div className="flex items-center justify-between mb-3">
        <h3 className="font-semibold text-text text-sm">
          {t.extractedData} ({files.length})
        </h3>
        <div className="flex items-center gap-2">
          <span className="text-xs text-text-3">
            {presentations.length}/{files.length}
          </span>
          <Button variant="ghost" size="sm" onClick={handleClearAll}>
            {t.clearAll}
          </Button>
        </div>
      </div>

      <div className="space-y-2">
        {files.map((file, index) => {
          const status = getStatus(file.name)
          const extracted = presentations.find((p) => p.fileName === file.name)
          const isPPTX = file.name.toLowerCase().endsWith('.pptx')

          return (
            <div
              key={`${file.name}-${index}`}
              className={cn(
                'flex items-center gap-3 p-3 rounded-xl transition-colors',
                status === 'done' ? 'bg-success/5' : 'bg-surface',
              )}
            >
              <div
                className={cn(
                  'w-10 h-10 rounded-xl flex items-center justify-center shrink-0',
                  isPPTX ? 'bg-accent' : 'bg-info',
                )}
              >
                <FileType size={20} className="text-white" />
              </div>

              <div className="flex-1 min-w-0">
                <p className="font-medium text-text text-sm truncate">{file.name}</p>
                <div className="flex items-center gap-1.5 text-xs text-text-3">
                  {settings.showFileSize && <span>{formatFileSize(file.size)}</span>}
                  <Badge variant="default">{isPPTX ? 'PPTX' : 'PPT'}</Badge>
                  {extracted && settings.showSlideCount && (
                    <span>{extracted.slides.length} {t.slides.toLowerCase()}</span>
                  )}
                </div>
              </div>

              <div className="flex items-center gap-1">
                {status === 'processing' && (
                  <Loader2 size={18} className="animate-spin text-accent" />
                )}
                {status === 'done' && (
                  <>
                    <CheckCircle size={18} className="text-success" />
                    <Button
                      variant="ghost"
                      size="sm"
                      onClick={() => handleView(file.name)}
                    >
                      <Eye size={14} /> {t.viewData}
                    </Button>
                  </>
                )}
                {status === 'pending' && (
                  <span className="text-xs text-text-3">{t.processing}</span>
                )}
                <button
                  onClick={() => handleRemove(index)}
                  className="p-1.5 rounded-lg hover:bg-danger/10 transition-colors"
                >
                  <X size={16} className="text-text-3" />
                </button>
              </div>
            </div>
          )
        })}
      </div>
    </div>
  )
}
