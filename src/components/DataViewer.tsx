import { useState, useMemo } from 'react'
import {
  FileText,
  User,
  Calendar,
  Layers,
  MessageSquare,
  Table2,
  Image,
  Palette,
  ChevronDown,
  ChevronRight,
  ArrowLeft,
} from 'lucide-react'
import { cn, formatFileSize, formatDate } from '@/lib/utils'
import { useFileStore } from '@/stores/file-store'
import { useUIStore } from '@/stores/ui-store'
import { useSettingsStore } from '@/stores/settings-store'
import { getTranslations } from '@/i18n'
import { Button } from '@/components/ui/Button'
import { Badge } from '@/components/ui/Badge'
import { Tabs } from '@/components/ui/Tabs'
import type { SlideContent, ExtractedPresentation, AppSettings } from '@/types'

export function DataViewer() {
  const { settings } = useSettingsStore()
  const t = getTranslations(settings.language)
  const { viewingPresentationId, setViewingPresentation } = useUIStore()
  const presentations = useFileStore((s) => s.presentations)

  const presentation = presentations.find((p) => p.id === viewingPresentationId)

  if (!presentation) {
    return (
      <div className="flex flex-col items-center justify-center h-full gap-4 p-8 text-center">
        <Layers size={48} className="text-text-3" />
        <div>
          <p className="font-semibold text-text">{t.noDataToExport}</p>
          <p className="text-sm text-text-3 mt-1">{t.dropFiles}</p>
        </div>
        {presentations.length > 0 && (
          <div className="flex flex-wrap gap-2 mt-4">
            {presentations.map((p) => (
              <Button
                key={p.id}
                variant="secondary"
                size="sm"
                onClick={() => setViewingPresentation(p.id)}
              >
                {p.fileName}
              </Button>
            ))}
          </div>
        )}
      </div>
    )
  }

  return (
    <PresentationDetail
      presentation={presentation}
      t={t}
      settings={settings}
      onBack={() => setViewingPresentation(null)}
    />
  )
}

function PresentationDetail({
  presentation,
  t,
  settings,
  onBack,
}: {
  presentation: ExtractedPresentation
  t: ReturnType<typeof getTranslations>
  settings: AppSettings
  onBack: () => void
}) {
  const initialExpanded = useMemo(() => {
    if (settings.autoExpandSlides) return new Set(presentation.slides.map((s) => s.slideNumber))
    return new Set([1])
  }, [settings.autoExpandSlides, presentation.slides])

  const [expandedSlides, setExpandedSlides] = useState<Set<number>>(initialExpanded)

  const sortedSlides = useMemo(() => {
    const slides = [...presentation.slides]
    if (settings.sortSlidesBy === 'title') {
      slides.sort((a, b) => (a.title || '').localeCompare(b.title || ''))
    }
    return slides
  }, [presentation.slides, settings.sortSlidesBy])

  const toggleSlide = (num: number) => {
    const next = new Set(expandedSlides)
    next.has(num) ? next.delete(num) : next.add(num)
    setExpandedSlides(next)
  }

  const tabs = [
    { id: 'slides', label: t.slides, icon: <Layers size={14} /> },
    { id: 'metadata', label: t.metadata, icon: <FileText size={14} /> },
    { id: 'themes', label: t.themes, icon: <Palette size={14} /> },
    { id: 'media', label: t.media, icon: <Image size={14} />, badge: presentation.media.length },
  ]

  return (
    <div className="max-w-4xl mx-auto p-4 pb-24 md:pb-4">
      {/* Header */}
      <div className="flex items-center gap-3 mb-4">
        <Button variant="ghost" size="sm" onClick={onBack}>
          <ArrowLeft size={16} />
        </Button>
        <div className="w-10 h-10 rounded-xl bg-accent flex items-center justify-center shrink-0">
          <FileText size={20} className="text-white" />
        </div>
        <div className="min-w-0">
          <h2 className="font-semibold text-text truncate">{presentation.fileName}</h2>
          <p className="text-xs text-text-3">
            {formatFileSize(presentation.fileSize)} · {presentation.fileType.toUpperCase()} · {presentation.slides.length} {t.slides.toLowerCase()}
          </p>
        </div>
      </div>

      <Tabs tabs={tabs}>
        {(activeTab) => (
          <>
            {activeTab === 'slides' && (
              <div className="space-y-2">
                <div className="flex justify-end gap-2 mb-3">
                  <Button
                    variant="ghost"
                    size="sm"
                    onClick={() => setExpandedSlides(new Set(presentation.slides.map((s) => s.slideNumber)))}
                  >
                    {t.expandAll}
                  </Button>
                  <Button variant="ghost" size="sm" onClick={() => setExpandedSlides(new Set())}>
                    {t.collapseAll}
                  </Button>
                </div>
                {sortedSlides.map((slide) => (
                  <SlideCard
                    key={slide.slideNumber}
                    slide={slide}
                    expanded={expandedSlides.has(slide.slideNumber)}
                    onToggle={() => toggleSlide(slide.slideNumber)}
                    compact={settings.compactView}
                    t={t}
                  />
                ))}
              </div>
            )}

            {activeTab === 'metadata' && (
              <div>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                  <MetaItem icon={FileText} label={t.title} value={presentation.metadata.title} />
                  <MetaItem icon={User} label={t.creator} value={presentation.metadata.creator} />
                  <MetaItem icon={User} label={t.lastModifiedBy} value={presentation.metadata.lastModifiedBy} />
                  <MetaItem icon={Calendar} label={t.created} value={formatDate(presentation.metadata.created)} />
                  <MetaItem icon={Calendar} label={t.modified} value={formatDate(presentation.metadata.modified)} />
                  <MetaItem icon={FileText} label={t.subject} value={presentation.metadata.subject} />
                  <MetaItem icon={FileText} label={t.category} value={presentation.metadata.category} />
                  <MetaItem icon={FileText} label={t.keywords} value={presentation.metadata.keywords} />
                  <MetaItem icon={FileText} label={t.description} value={presentation.metadata.description} />
                  <MetaItem icon={Layers} label={t.application} value={presentation.metadata.application} />
                  <MetaItem icon={FileText} label={t.appVersion} value={presentation.metadata.appVersion} />
                  <MetaItem icon={FileText} label={t.company} value={presentation.metadata.company} />
                  <MetaItem icon={FileText} label={t.totalSlides} value={String(presentation.metadata.totalSlides)} />
                  <MetaItem icon={FileText} label={t.totalWords} value={String(presentation.metadata.totalWords)} />
                  <MetaItem icon={FileText} label={t.format} value={presentation.metadata.presentationFormat} />
                </div>

                {Object.keys(presentation.customProperties).length > 0 && (
                  <div className="mt-6">
                    <h3 className="font-semibold text-text mb-3">{t.customProperties}</h3>
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-2">
                      {Object.entries(presentation.customProperties).map(([key, value]) => (
                        <div key={key} className="p-3 rounded-xl bg-surface-2">
                          <span className="text-sm text-text-3">{key}:</span>
                          <span className="ml-2 font-medium text-text">{value}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </div>
            )}

            {activeTab === 'themes' && (
              <div>
                {presentation.themes.length === 0 ? (
                  <p className="text-center text-text-3 py-8">{t.noThemes}</p>
                ) : (
                  <div className="space-y-4">
                    {presentation.themes.map((theme, i) => (
                      <div key={i} className="p-4 rounded-xl bg-surface-2">
                        <h4 className="font-semibold text-text mb-3">{theme.name}</h4>
                        {theme.fonts.length > 0 && (
                          <div className="mb-3">
                            <p className="text-sm text-text-3 mb-1">{t.fonts}</p>
                            <div className="flex flex-wrap gap-1">
                              {theme.fonts.map((f, j) => (
                                <Badge key={j}>{f}</Badge>
                              ))}
                            </div>
                          </div>
                        )}
                        {theme.colors.length > 0 && (
                          <div>
                            <p className="text-sm text-text-3 mb-1">{t.colors}</p>
                            <div className="flex flex-wrap gap-1">
                              {theme.colors.map((c, j) => (
                                <span key={j} className="text-xs p-1 px-2 rounded-lg bg-surface">{c}</span>
                              ))}
                            </div>
                          </div>
                        )}
                      </div>
                    ))}

                    {presentation.masterSlides.length > 0 && (
                      <div>
                        <h3 className="font-semibold text-text mb-3">{t.masterSlides}</h3>
                        <div className="flex flex-wrap gap-1">
                          {presentation.masterSlides.map((m, i) => (
                            <Badge key={i}>{m}</Badge>
                          ))}
                        </div>
                      </div>
                    )}
                  </div>
                )}
              </div>
            )}

            {activeTab === 'media' && (
              <div>
                {presentation.media.length === 0 ? (
                  <p className="text-center text-text-3 py-8">{t.noMedia}</p>
                ) : (
                  <div className="space-y-2">
                    {presentation.media.map((media, i) => {
                      const previewSize =
                        settings.mediaPreviewSize === 'small' ? 'w-12 h-12' :
                        settings.mediaPreviewSize === 'large' ? 'w-24 h-24' : 'w-16 h-16'
                      return (
                        <div key={i} className="flex items-center gap-3 p-3 rounded-xl bg-surface-2">
                          <div
                            className={cn(
                              'w-10 h-10 rounded-xl flex items-center justify-center shrink-0',
                              media.type === 'image' ? 'bg-success' : media.type === 'video' ? 'bg-info' : 'bg-accent',
                            )}
                          >
                            <Image size={20} className="text-white" />
                          </div>
                          <div className="flex-1 min-w-0">
                            <p className="font-medium text-text text-sm truncate">{media.name}</p>
                            <p className="text-xs text-text-3">
                              {media.type} · {formatFileSize(media.size)} · .{media.extension}
                            </p>
                          </div>
                          {media.data && media.type === 'image' && (
                            <img
                              src={`data:image/${media.extension};base64,${media.data}`}
                              alt={media.name}
                              className={cn(previewSize, 'object-cover rounded-lg')}
                            />
                          )}
                        </div>
                      )
                    })}
                  </div>
                )}
              </div>
            )}
          </>
        )}
      </Tabs>
    </div>
  )
}

function SlideCard({
  slide,
  expanded,
  onToggle,
  compact,
  t,
}: {
  slide: SlideContent
  expanded: boolean
  onToggle: () => void
  compact: boolean
  t: ReturnType<typeof getTranslations>
}) {
  return (
    <div className="rounded-xl bg-surface-2 overflow-hidden">
      <button
        onClick={onToggle}
        className={cn(
          'w-full flex items-center gap-3 text-left transition-colors hover:bg-surface-3/50',
          compact ? 'p-2' : 'p-3',
        )}
      >
        {expanded ? <ChevronDown size={18} /> : <ChevronRight size={18} />}
        <div className="w-8 h-8 rounded-full bg-accent text-white flex items-center justify-center text-sm font-semibold shrink-0">
          {slide.slideNumber}
        </div>
        <div className="flex-1 min-w-0">
          <p className="font-medium text-text truncate">{slide.title || t.untitledSlide}</p>
          <p className="text-xs text-text-3">
            {slide.textContent.length} {t.textBlocks}
            {slide.tables.length > 0 && ` · ${slide.tables.length} ${t.tables.toLowerCase()}`}
            {slide.shapes.length > 0 && ` · ${slide.shapes.length} ${t.shapes.toLowerCase()}`}
            {slide.notes && ` · ${t.hasNotes}`}
          </p>
        </div>
      </button>

      {expanded && (
        <div className="px-4 pb-4 border-t border-border space-y-4">
          {slide.textContent.length > 0 && (
            <div className="mt-3">
              <h5 className="text-xs font-semibold text-text-2 uppercase mb-2">{t.content}</h5>
              <div className="space-y-1">
                {slide.textContent.map((text, i) => (
                  <div key={i} className="p-2 rounded-lg bg-surface text-sm text-text">{text}</div>
                ))}
              </div>
            </div>
          )}

          {slide.notes && (
            <div>
              <h5 className="text-xs font-semibold text-text-2 uppercase mb-2 flex items-center gap-1">
                <MessageSquare size={14} /> {t.speakerNotes}
              </h5>
              <div className="p-3 rounded-lg bg-warning/10 border-l-4 border-warning text-sm text-text">
                {slide.notes}
              </div>
            </div>
          )}

          {slide.tables.length > 0 && (
            <div>
              <h5 className="text-xs font-semibold text-text-2 uppercase mb-2 flex items-center gap-1">
                <Table2 size={14} /> {t.tables}
              </h5>
              {slide.tables.map((table, i) => (
                <div key={i} className="overflow-x-auto mt-2">
                  <table className="w-full text-sm border-collapse">
                    <tbody>
                      {table.cells.map((row, ri) => (
                        <tr key={ri}>
                          {row.map((cell, ci) => (
                            <td
                              key={ci}
                              className={cn(
                                'p-2 border border-border',
                                ri === 0 ? 'bg-surface-2 font-medium text-text' : 'bg-surface text-text',
                              )}
                            >
                              {cell}
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              ))}
            </div>
          )}

          {slide.shapes.filter((s) => s.text).length > 0 && (
            <div>
              <h5 className="text-xs font-semibold text-text-2 uppercase mb-2">{t.shapes}</h5>
              <div className="flex flex-wrap gap-1">
                {slide.shapes
                  .filter((s) => s.text)
                  .map((shape, i) => (
                    <Badge key={i}>
                      {shape.type}: {shape.text.substring(0, 50)}
                      {shape.text.length > 50 ? '…' : ''}
                    </Badge>
                  ))}
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  )
}

function MetaItem({ icon: Icon, label, value }: { icon: React.ElementType; label: string; value: string }) {
  return (
    <div className="p-3 rounded-xl bg-surface-2">
      <div className="flex items-center gap-2 text-text-3 mb-1">
        <Icon size={14} />
        <span className="text-xs">{label}</span>
      </div>
      <p className="font-medium text-text text-sm">{value || 'N/A'}</p>
    </div>
  )
}
