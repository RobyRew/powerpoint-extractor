import { useState } from 'react'
import {
  Download,
  FileJson,
  FileCode,
  FileSpreadsheet,
  FileText,
  Globe,
  Image as ImageIcon,
  Package,
  Loader2,
  Check,
} from 'lucide-react'
import { cn } from '@/lib/utils'
import { useFileStore } from '@/stores/file-store'
import { useSettingsStore } from '@/stores/settings-store'
import { getTranslations } from '@/i18n'
import { Button } from '@/components/ui/Button'
import {
  exportToJSON,
  exportToXML,
  exportToCSV,
  exportToText,
  exportToHTML,
  exportToPDF,
  downloadFile,
  downloadMediaAsZip,
  downloadAllAsZip,
} from '@/lib'
import type { ExportFormatId } from '@/types'

const FORMAT_META: Record<ExportFormatId, { name: string; icon: typeof FileJson }> = {
  json: { name: 'JSON', icon: FileJson },
  xml: { name: 'XML', icon: FileCode },
  csv: { name: 'CSV', icon: FileSpreadsheet },
  txt: { name: 'Text', icon: FileText },
  html: { name: 'HTML', icon: Globe },
  pdf: { name: 'PDF', icon: FileText },
}
const FORMAT_IDS = Object.keys(FORMAT_META) as ExportFormatId[]

export function ExportPanel() {
  const presentations = useFileStore((s) => s.presentations)
  const { settings } = useSettingsStore()
  const t = getTranslations(settings.language)
  const [selected, setSelected] = useState<Set<ExportFormatId>>(
    new Set(settings.defaultExportFormats),
  )
  const [exporting, setExporting] = useState(false)
  const [success, setSuccess] = useState<string | null>(null)

  const hasMedia = presentations.some((p) => p.media.length > 0)
  const totalMedia = presentations.reduce((a, p) => a + p.media.length, 0)

  const toggleFormat = (id: ExportFormatId) => {
    const next = new Set(selected)
    next.has(id) ? next.delete(id) : next.add(id)
    setSelected(next)
  }

  const getFilename = (ext: string) => {
    const ts = new Date().toISOString().split('T')[0]
    if (presentations.length === 1) {
      const base = presentations[0]!.fileName.replace(/\.(pptx?|ppt)$/i, '')
      if (settings.exportFilenamePattern === 'original') return `${base}.${ext}`
      if (settings.exportFilenamePattern === 'timestamp') return `export-${ts}.${ext}`
      return `${base}-${ts}.${ext}`
    }
    return `presentations-${ts}.${ext}`
  }

  const handleExport = async (fmt: ExportFormatId) => {
    setExporting(true)
    try {
      switch (fmt) {
        case 'json':
          downloadFile(exportToJSON(presentations), getFilename('json'), 'application/json')
          break
        case 'xml':
          downloadFile(exportToXML(presentations), getFilename('xml'), 'application/xml')
          break
        case 'csv':
          downloadFile(exportToCSV(presentations), getFilename('csv'), 'text/csv')
          break
        case 'txt':
          downloadFile(exportToText(presentations), getFilename('txt'), 'text/plain')
          break
        case 'html':
          downloadFile(exportToHTML(presentations), getFilename('html'), 'text/html')
          break
        case 'pdf': {
          const pdf = exportToPDF(presentations)
          pdf.save(getFilename('pdf'))
          break
        }
      }
      setSuccess(fmt)
      setTimeout(() => setSuccess(null), 2000)
    } catch (err) {
      console.error('Export error:', err)
    } finally {
      setExporting(false)
    }
  }

  const handleMediaExport = async () => {
    setExporting(true)
    try {
      await downloadMediaAsZip(presentations)
      setSuccess('media')
      setTimeout(() => setSuccess(null), 2000)
    } catch (err) {
      console.error('Media export error:', err)
    } finally {
      setExporting(false)
    }
  }

  const handleExportAll = async () => {
    if (selected.size === 0) return
    setExporting(true)
    try {
      await downloadAllAsZip(presentations, Array.from(selected))
      setSuccess('all')
      setTimeout(() => setSuccess(null), 2000)
    } catch (err) {
      console.error('Export all error:', err)
    } finally {
      setExporting(false)
    }
  }

  if (presentations.length === 0) {
    return (
      <div className="flex flex-col items-center justify-center h-full gap-4 p-8 text-center">
        <Download size={48} className="text-text-3" />
        <p className="font-semibold text-text">{t.noDataToExport}</p>
      </div>
    )
  }

  return (
    <div className="max-w-xl mx-auto p-4 pb-24 md:pb-4 space-y-4">
      <div className="flex items-center gap-3">
        <Download size={20} className="text-accent" />
        <h2 className="font-semibold text-text">{t.exportData}</h2>
        <span className="text-sm text-text-3">
          {presentations.length} {t.presentations}
        </span>
      </div>

      {/* Format grid */}
      <div className="grid grid-cols-3 sm:grid-cols-6 gap-2">
        {FORMAT_IDS.map((id) => {
          const { name, icon: Icon } = FORMAT_META[id]
          const isSel = selected.has(id)
          const isDone = success === id
          return (
            <button
              key={id}
              onClick={() => toggleFormat(id)}
              disabled={exporting}
              className={cn(
                'p-3 rounded-xl border text-center transition-all',
                isSel ? 'border-accent bg-accent/10' : 'border-border hover:border-text-3',
                isDone && 'bg-success/10 border-success',
              )}
            >
              <Icon size={20} className={cn('mx-auto mb-1', isSel && 'text-accent')} />
              <span className="text-xs font-medium text-text">{name}</span>
              {isDone && <Check size={14} className="mx-auto mt-1 text-success" />}
            </button>
          )
        })}
      </div>

      {/* Actions */}
      <div className="flex flex-wrap gap-2">
        {Array.from(selected).map((id) => {
          const { name, icon: Icon } = FORMAT_META[id]
          return (
            <Button key={id} variant="secondary" size="sm" onClick={() => handleExport(id)} disabled={exporting}>
              <Icon size={14} /> {t.exportSelected} {name}
            </Button>
          )
        })}

        {hasMedia && (
          <Button variant="secondary" size="sm" onClick={handleMediaExport} disabled={exporting}>
            <ImageIcon size={14} /> {t.exportMedia} ({totalMedia})
          </Button>
        )}

        {selected.size > 0 && (
          <Button variant="primary" size="sm" onClick={handleExportAll} disabled={exporting} className="ml-auto">
            {exporting ? <Loader2 size={14} className="animate-spin" /> : <Package size={14} />}
            {t.downloadAllZip}
          </Button>
        )}
      </div>

      <p className="text-xs text-text-3">{t.selectFormats}</p>
    </div>
  )
}
