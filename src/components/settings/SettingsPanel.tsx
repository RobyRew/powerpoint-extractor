import { Select } from '@/components/ui/Select'
import { Toggle } from '@/components/ui/Toggle'
import { Button } from '@/components/ui/Button'
import { useSettingsStore } from '@/stores/settings-store'
import { getTranslations } from '@/i18n'
import { getStorageSize, clearAppData } from '@/lib/storage'
import { APP_VERSION } from '@/lib/constants'
import { useState, useEffect } from 'react'
import { RotateCcw, Trash2, HardDrive, Shield } from 'lucide-react'

export function SettingsPanel() {
  const { settings, set, resetAll } = useSettingsStore()
  const t = getTranslations(settings.language)
  const [storageUsed, setStorageUsed] = useState('')

  useEffect(() => {
    setStorageUsed(getStorageSize())
  }, [])

  const handleClearData = () => {
    if (confirm(t.clearAppDataConfirm)) {
      clearAppData()
      location.reload()
    }
  }

  const handleReset = () => {
    if (confirm(t.resetSettingsConfirm)) {
      resetAll()
    }
  }

  return (
    <div className="max-w-xl mx-auto p-4 pb-24 md:pb-4 space-y-6">
      <h2 className="text-xl font-bold text-text">{t.settings}</h2>

      {/* Appearance */}
      <section className="space-y-3">
        <h3 className="text-sm font-semibold text-text-2 uppercase tracking-wider">{t.appearance}</h3>
        <Select
          label={t.theme}
          options={[
            { value: 'system', label: t.systemTheme },
            { value: 'light', label: t.lightTheme },
            { value: 'dark', label: t.darkTheme },
            { value: 'oled', label: t.oledTheme },
          ]}
          value={settings.theme}
          onChange={(e) => set('theme', e.target.value as typeof settings.theme)}
        />
        <Select
          label={t.language}
          options={[
            { value: 'en', label: 'English' },
            { value: 'es', label: 'Español' },
            { value: 'de', label: 'Deutsch' },
            { value: 'fr', label: 'Français' },
          ]}
          value={settings.language}
          onChange={(e) => set('language', e.target.value as typeof settings.language)}
        />
        <Toggle
          label={t.compactView}
          description={t.compactViewDesc}
          checked={settings.compactView}
          onChange={(v) => set('compactView', v)}
        />
      </section>

      {/* Viewer Settings */}
      <section className="space-y-3">
        <h3 className="text-sm font-semibold text-text-2 uppercase tracking-wider">{t.viewer}</h3>
        <Toggle
          label={t.autoExpandSlides}
          description={t.autoExpandSlidesDesc}
          checked={settings.autoExpandSlides}
          onChange={(v) => set('autoExpandSlides', v)}
        />
        <Toggle
          label={t.showFileSize}
          checked={settings.showFileSize}
          onChange={(v) => set('showFileSize', v)}
        />
        <Toggle
          label={t.showSlideCount}
          checked={settings.showSlideCount}
          onChange={(v) => set('showSlideCount', v)}
        />
        <Select
          label={t.sortSlidesBy}
          options={[
            { value: 'number', label: t.sortByNumber },
            { value: 'title', label: t.sortByTitle },
          ]}
          value={settings.sortSlidesBy}
          onChange={(e) => set('sortSlidesBy', e.target.value as typeof settings.sortSlidesBy)}
        />
        <Select
          label={t.mediaPreviewSize}
          options={[
            { value: 'small', label: t.previewSmall },
            { value: 'medium', label: t.previewMedium },
            { value: 'large', label: t.previewLarge },
          ]}
          value={settings.mediaPreviewSize}
          onChange={(e) => set('mediaPreviewSize', e.target.value as typeof settings.mediaPreviewSize)}
        />
        <Select
          label={t.maxInitialSlides}
          options={[
            { value: '10', label: '10' },
            { value: '20', label: '20' },
            { value: '50', label: '50' },
            { value: '100', label: '100' },
          ]}
          value={String(settings.maxInitialSlides)}
          onChange={(e) => set('maxInitialSlides', parseInt(e.target.value))}
        />
      </section>

      {/* Extraction Settings */}
      <section className="space-y-3">
        <h3 className="text-sm font-semibold text-text-2 uppercase tracking-wider">{t.extraction}</h3>
        <Toggle
          label={t.includeNotes}
          checked={settings.includeNotes}
          onChange={(v) => set('includeNotes', v)}
        />
        <Toggle
          label={t.includeMedia}
          checked={settings.includeMedia}
          onChange={(v) => set('includeMedia', v)}
        />
        <Toggle
          label={t.includeMetadata}
          checked={settings.includeMetadata}
          onChange={(v) => set('includeMetadata', v)}
        />
        <Toggle
          label={t.includeThemes}
          checked={settings.includeThemes}
          onChange={(v) => set('includeThemes', v)}
        />
        <Toggle
          label={t.includeShapes}
          checked={settings.includeShapes}
          onChange={(v) => set('includeShapes', v)}
        />
        <Toggle
          label={t.includeTables}
          checked={settings.includeTables}
          onChange={(v) => set('includeTables', v)}
        />
        <Toggle
          label={t.includeCustomProperties}
          checked={settings.includeCustomProperties}
          onChange={(v) => set('includeCustomProperties', v)}
        />
      </section>

      {/* Export Settings */}
      <section className="space-y-3">
        <h3 className="text-sm font-semibold text-text-2 uppercase tracking-wider">{t.exportSettings}</h3>
        <Select
          label={t.pdfPageSize}
          options={[
            { value: 'a4', label: 'A4' },
            { value: 'letter', label: 'Letter' },
            { value: 'legal', label: 'Legal' },
          ]}
          value={settings.pdfPageSize}
          onChange={(e) => set('pdfPageSize', e.target.value as typeof settings.pdfPageSize)}
        />
        <Select
          label={t.pdfOrientation}
          options={[
            { value: 'portrait', label: t.portrait },
            { value: 'landscape', label: t.landscape },
          ]}
          value={settings.pdfOrientation}
          onChange={(e) => set('pdfOrientation', e.target.value as typeof settings.pdfOrientation)}
        />
        <Select
          label={t.csvDelimiter}
          options={[
            { value: ',', label: 'Comma (,)' },
            { value: ';', label: 'Semicolon (;)' },
            { value: '\t', label: 'Tab' },
            { value: '|', label: 'Pipe (|)' },
          ]}
          value={settings.csvDelimiter}
          onChange={(e) => set('csvDelimiter', e.target.value as typeof settings.csvDelimiter)}
        />
        <Select
          label={t.jsonIndent}
          options={[
            { value: '2', label: '2 spaces' },
            { value: '4', label: '4 spaces' },
          ]}
          value={String(settings.jsonIndent)}
          onChange={(e) => set('jsonIndent', parseInt(e.target.value) as typeof settings.jsonIndent)}
        />
        <Toggle
          label={t.htmlIncludeStyles}
          description={t.htmlIncludeStylesDesc}
          checked={settings.htmlIncludeStyles}
          onChange={(v) => set('htmlIncludeStyles', v)}
        />
        <Select
          label={t.exportFilenamePattern}
          options={[
            { value: 'original', label: t.filenameOriginal },
            { value: 'timestamp', label: t.filenameTimestamp },
            { value: 'both', label: t.filenameBoth },
          ]}
          value={settings.exportFilenamePattern}
          onChange={(e) => set('exportFilenamePattern', e.target.value as typeof settings.exportFilenamePattern)}
        />
      </section>

      {/* Behavior */}
      <section className="space-y-3">
        <h3 className="text-sm font-semibold text-text-2 uppercase tracking-wider">{t.behavior}</h3>
        <Toggle
          label={t.confirmBeforeDelete}
          checked={settings.confirmBeforeDelete}
          onChange={(v) => set('confirmBeforeDelete', v)}
        />
        <Toggle
          label={t.confirmBeforeClear}
          checked={settings.confirmBeforeClear}
          onChange={(v) => set('confirmBeforeClear', v)}
        />
      </section>

      {/* Data Management */}
      <section className="space-y-3">
        <h3 className="text-sm font-semibold text-text-2 uppercase tracking-wider">{t.dataManagement}</h3>

        <div className="flex items-center gap-2 text-sm text-text-2">
          <HardDrive size={16} /> {t.storageUsed}: {storageUsed || '—'}
        </div>

        <div className="flex flex-wrap gap-2 pt-2 border-t border-border">
          <Button variant="ghost" size="sm" onClick={handleReset}>
            <RotateCcw size={16} /> {t.resetSettings}
          </Button>
          <Button variant="danger" size="sm" onClick={handleClearData}>
            <Trash2 size={16} /> {t.clearAppData}
          </Button>
        </div>
      </section>

      {/* About */}
      <section className="space-y-3 pt-4 border-t border-border">
        <div className="text-center">
          <p className="text-sm text-text-2 font-medium">{t.appName} v{APP_VERSION}</p>
          <p className="text-xs text-text-3 mt-1">{t.appDescription}</p>
        </div>
        <div className="flex items-start gap-2 p-3 rounded-lg bg-surface-2">
          <Shield size={16} className="text-success shrink-0 mt-0.5" />
          <p className="text-xs text-text-3">{t.privacyNote}</p>
        </div>
      </section>
    </div>
  )
}
