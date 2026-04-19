import { useCallback, useRef, useState } from 'react'
import { Upload, FileUp } from 'lucide-react'
import { cn } from '@/lib/utils'
import { useSettingsStore } from '@/stores/settings-store'
import { useFileStore } from '@/stores/file-store'
import { useUIStore } from '@/stores/ui-store'
import { getTranslations } from '@/i18n'

export function DropZone() {
  const inputRef = useRef<HTMLInputElement>(null)
  const [isDragging, setIsDragging] = useState(false)
  const { settings } = useSettingsStore()
  const { addFiles, processingFile } = useFileStore()
  const { showToast } = useUIStore()
  const t = getTranslations(settings.language)

  const handleFiles = useCallback(
    async (fileList: FileList | File[]) => {
      const files = Array.from(fileList)
      const valid = files.filter((f) => {
        const ext = f.name.toLowerCase()
        return ext.endsWith('.ppt') || ext.endsWith('.pptx')
      })
      if (valid.length === 0) {
        showToast(t.invalidFileType, 'error')
        return
      }
      await addFiles(valid)
    },
    [addFiles, showToast, t],
  )

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault()
    setIsDragging(false)
    handleFiles(e.dataTransfer.files)
  }

  return (
    <div
      onDragOver={(e) => {
        e.preventDefault()
        setIsDragging(true)
      }}
      onDragLeave={() => setIsDragging(false)}
      onDrop={handleDrop}
      onClick={() => inputRef.current?.click()}
      className={cn(
        'relative flex flex-col items-center justify-center gap-3 p-8 rounded-2xl border-2 border-dashed cursor-pointer transition-all duration-200',
        isDragging
          ? 'border-accent bg-accent/5 scale-[1.01]'
          : 'border-border hover:border-accent/50 hover:bg-surface-2/50',
        processingFile && 'pointer-events-none opacity-60',
      )}
    >
      <div
        className={cn(
          'w-14 h-14 rounded-2xl flex items-center justify-center transition-colors',
          isDragging ? 'bg-accent/15 text-accent' : 'bg-surface-2 text-text-3',
        )}
      >
        {isDragging ? <Upload size={28} /> : <FileUp size={28} />}
      </div>

      <div className="text-center">
        <p className="text-base font-semibold text-text">{t.dropFiles}</p>
        <p className="text-sm text-text-3 mt-1">{t.dropFilesDescription}</p>
      </div>

      <span className="text-xs text-text-3 bg-surface-2 px-3 py-1 rounded-full">
        {t.supportedFormats}
      </span>

      <input
        ref={inputRef}
        type="file"
        accept=".ppt,.pptx"
        multiple
        className="hidden"
        onChange={(e) => e.target.files && handleFiles(e.target.files)}
      />
    </div>
  )
}
