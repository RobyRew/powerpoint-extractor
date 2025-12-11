/**
 * Data Viewer Modal - Display extracted presentation data
 */

import { X, FileText, User, Calendar, Layers, MessageSquare, Table2, Image, Palette, ChevronDown, ChevronRight } from 'lucide-react';
import { useState } from 'react';
import type { ExtractedPresentation, SlideContent } from '../types';

interface DataViewerProps {
  presentation: ExtractedPresentation;
  onClose: () => void;
}

export function DataViewer({ presentation, onClose }: DataViewerProps) {
  const [expandedSlides, setExpandedSlides] = useState<Set<number>>(new Set([1]));
  const [activeTab, setActiveTab] = useState<'slides' | 'metadata' | 'themes' | 'media'>('slides');

  const toggleSlide = (slideNum: number) => {
    const newExpanded = new Set(expandedSlides);
    if (newExpanded.has(slideNum)) {
      newExpanded.delete(slideNum);
    } else {
      newExpanded.add(slideNum);
    }
    setExpandedSlides(newExpanded);
  };

  const expandAll = () => {
    setExpandedSlides(new Set(presentation.slides.map(s => s.slideNumber)));
  };

  const collapseAll = () => {
    setExpandedSlides(new Set());
  };

  const formatDate = (dateStr: string) => {
    if (!dateStr) return 'N/A';
    try {
      return new Date(dateStr).toLocaleString();
    } catch {
      return dateStr;
    }
  };

  const formatFileSize = (bytes: number): string => {
    if (bytes === 0) return '0 B';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return `${parseFloat((bytes / Math.pow(k, i)).toFixed(1))} ${sizes[i]}`;
  };

  return (
    <>
      <div className="modal-overlay" onClick={onClose} />
      <div className="modal-content w-full max-w-4xl animate-slide-up">
        <div className="card card-elevated max-h-[90vh] flex flex-col">
          {/* Header */}
          <div className="flex items-center justify-between p-4 border-b border-[rgb(var(--border))]">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-lg bg-[rgb(var(--primary))] flex items-center justify-center">
                <FileText className="w-5 h-5 text-[rgb(var(--primary-foreground))]" />
              </div>
              <div>
                <h2 className="font-semibold text-[rgb(var(--foreground))]">
                  {presentation.fileName}
                </h2>
                <p className="text-xs text-[rgb(var(--muted-foreground))]">
                  {formatFileSize(presentation.fileSize)} • {presentation.fileType.toUpperCase()} • {presentation.slides.length} slides
                </p>
              </div>
            </div>
            <button
              onClick={onClose}
              className="p-2 rounded-lg hover:bg-[rgb(var(--accent))] transition-colors"
            >
              <X className="w-5 h-5" />
            </button>
          </div>

          {/* Tabs */}
          <div className="flex border-b border-[rgb(var(--border))]">
            {[
              { id: 'slides', label: 'Slides', icon: Layers },
              { id: 'metadata', label: 'Metadata', icon: FileText },
              { id: 'themes', label: 'Themes', icon: Palette },
              { id: 'media', label: 'Media', icon: Image },
            ].map(tab => {
              const Icon = tab.icon;
              return (
                <button
                  key={tab.id}
                  onClick={() => setActiveTab(tab.id as typeof activeTab)}
                  className={`
                    flex items-center gap-2 px-4 py-3 text-sm font-medium transition-colors
                    ${activeTab === tab.id 
                      ? 'text-[rgb(var(--foreground))] border-b-2 border-[rgb(var(--primary))]' 
                      : 'text-[rgb(var(--muted-foreground))] hover:text-[rgb(var(--foreground))]'
                    }
                  `}
                >
                  <Icon className="w-4 h-4" />
                  {tab.label}
                  {tab.id === 'media' && presentation.media.length > 0 && (
                    <span className="badge">{presentation.media.length}</span>
                  )}
                </button>
              );
            })}
          </div>

          {/* Content */}
          <div className="flex-1 overflow-auto p-4">
            {activeTab === 'slides' && (
              <div className="space-y-3">
                <div className="flex justify-end gap-2 mb-4">
                  <button onClick={expandAll} className="btn btn-ghost text-sm py-1">
                    Expand All
                  </button>
                  <button onClick={collapseAll} className="btn btn-ghost text-sm py-1">
                    Collapse All
                  </button>
                </div>

                {presentation.slides.map(slide => (
                  <SlideCard
                    key={slide.slideNumber}
                    slide={slide}
                    expanded={expandedSlides.has(slide.slideNumber)}
                    onToggle={() => toggleSlide(slide.slideNumber)}
                  />
                ))}
              </div>
            )}

            {activeTab === 'metadata' && (
              <div className="space-y-4">
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <MetadataItem icon={FileText} label="Title" value={presentation.metadata.title} />
                  <MetadataItem icon={User} label="Creator" value={presentation.metadata.creator} />
                  <MetadataItem icon={User} label="Last Modified By" value={presentation.metadata.lastModifiedBy} />
                  <MetadataItem icon={Calendar} label="Created" value={formatDate(presentation.metadata.created)} />
                  <MetadataItem icon={Calendar} label="Modified" value={formatDate(presentation.metadata.modified)} />
                  <MetadataItem icon={FileText} label="Subject" value={presentation.metadata.subject} />
                  <MetadataItem icon={FileText} label="Category" value={presentation.metadata.category} />
                  <MetadataItem icon={FileText} label="Keywords" value={presentation.metadata.keywords} />
                  <MetadataItem icon={FileText} label="Description" value={presentation.metadata.description} />
                  <MetadataItem icon={Layers} label="Application" value={presentation.metadata.application} />
                  <MetadataItem icon={FileText} label="App Version" value={presentation.metadata.appVersion} />
                  <MetadataItem icon={FileText} label="Company" value={presentation.metadata.company} />
                  <MetadataItem icon={FileText} label="Total Slides" value={String(presentation.metadata.totalSlides)} />
                  <MetadataItem icon={FileText} label="Total Words" value={String(presentation.metadata.totalWords)} />
                  <MetadataItem icon={FileText} label="Format" value={presentation.metadata.presentationFormat} />
                </div>

                {Object.keys(presentation.customProperties).length > 0 && (
                  <div className="mt-6">
                    <h3 className="font-semibold mb-3">Custom Properties</h3>
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-2">
                      {Object.entries(presentation.customProperties).map(([key, value]) => (
                        <div key={key} className="p-3 rounded-lg bg-[rgb(var(--secondary))]">
                          <span className="text-sm text-[rgb(var(--muted-foreground))]">{key}:</span>
                          <span className="ml-2 font-medium">{value}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                )}
              </div>
            )}

            {activeTab === 'themes' && (
              <div className="space-y-4">
                {presentation.themes.length === 0 ? (
                  <p className="text-center text-[rgb(var(--muted-foreground))] py-8">
                    No theme information available
                  </p>
                ) : (
                  presentation.themes.map((theme, index) => (
                    <div key={index} className="p-4 rounded-lg bg-[rgb(var(--secondary))]">
                      <h4 className="font-semibold mb-3">{theme.name}</h4>
                      
                      {theme.fonts.length > 0 && (
                        <div className="mb-3">
                          <p className="text-sm text-[rgb(var(--muted-foreground))] mb-1">Fonts</p>
                          <div className="flex flex-wrap gap-2">
                            {theme.fonts.map((font, i) => (
                              <span key={i} className="badge">{font}</span>
                            ))}
                          </div>
                        </div>
                      )}

                      {theme.colors.length > 0 && (
                        <div>
                          <p className="text-sm text-[rgb(var(--muted-foreground))] mb-1">Colors</p>
                          <div className="flex flex-wrap gap-2">
                            {theme.colors.map((color, i) => (
                              <span key={i} className="text-xs p-1 px-2 rounded bg-[rgb(var(--background))]">
                                {color}
                              </span>
                            ))}
                          </div>
                        </div>
                      )}
                    </div>
                  ))
                )}

                {presentation.masterSlides.length > 0 && (
                  <div>
                    <h3 className="font-semibold mb-3">Master Slides</h3>
                    <div className="flex flex-wrap gap-2">
                      {presentation.masterSlides.map((master, i) => (
                        <span key={i} className="badge">{master}</span>
                      ))}
                    </div>
                  </div>
                )}
              </div>
            )}

            {activeTab === 'media' && (
              <div className="space-y-3">
                {presentation.media.length === 0 ? (
                  <p className="text-center text-[rgb(var(--muted-foreground))] py-8">
                    No media files found in this presentation
                  </p>
                ) : (
                  presentation.media.map((media, index) => (
                    <div key={index} className="flex items-center gap-3 p-3 rounded-lg bg-[rgb(var(--secondary))]">
                      <div className={`
                        w-10 h-10 rounded-lg flex items-center justify-center
                        ${media.type === 'image' ? 'bg-green-500' : media.type === 'video' ? 'bg-blue-500' : 'bg-purple-500'}
                      `}>
                        <Image className="w-5 h-5 text-white" />
                      </div>
                      <div className="flex-1">
                        <p className="font-medium">{media.name}</p>
                        <p className="text-xs text-[rgb(var(--muted-foreground))]">
                          {media.type} • {formatFileSize(media.size)} • .{media.extension}
                        </p>
                      </div>
                      {media.data && media.type === 'image' && (
                        <img
                          src={`data:image/${media.extension};base64,${media.data}`}
                          alt={media.name}
                          className="w-16 h-16 object-cover rounded"
                        />
                      )}
                    </div>
                  ))
                )}
              </div>
            )}
          </div>
        </div>
      </div>
    </>
  );
}

function SlideCard({ 
  slide, 
  expanded, 
  onToggle 
}: { 
  slide: SlideContent; 
  expanded: boolean; 
  onToggle: () => void;
}) {
  return (
    <div className="rounded-lg bg-[rgb(var(--secondary))] overflow-hidden">
      <button
        onClick={onToggle}
        className="w-full flex items-center gap-3 p-3 hover:bg-[rgb(var(--accent))] transition-colors text-left"
      >
        {expanded ? (
          <ChevronDown className="w-5 h-5 shrink-0" />
        ) : (
          <ChevronRight className="w-5 h-5 shrink-0" />
        )}
        <div className="w-8 h-8 rounded-full bg-[rgb(var(--primary))] text-[rgb(var(--primary-foreground))] flex items-center justify-center text-sm font-semibold shrink-0">
          {slide.slideNumber}
        </div>
        <div className="flex-1 min-w-0">
          <p className="font-medium truncate">{slide.title || 'Untitled Slide'}</p>
          <p className="text-xs text-[rgb(var(--muted-foreground))]">
            {slide.textContent.length} text blocks
            {slide.tables.length > 0 && ` • ${slide.tables.length} tables`}
            {slide.shapes.length > 0 && ` • ${slide.shapes.length} shapes`}
            {slide.notes && ' • Has notes'}
          </p>
        </div>
      </button>

      {expanded && (
        <div className="px-4 pb-4 border-t border-[rgb(var(--border))]">
          {/* Text Content */}
          {slide.textContent.length > 0 && (
            <div className="mt-4">
              <h5 className="text-sm font-medium text-[rgb(var(--muted-foreground))] mb-2">Content</h5>
              <div className="space-y-2">
                {slide.textContent.map((text, i) => (
                  <div key={i} className="p-2 rounded bg-[rgb(var(--background))] text-sm">
                    {text}
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* Notes */}
          {slide.notes && (
            <div className="mt-4">
              <h5 className="text-sm font-medium text-[rgb(var(--muted-foreground))] mb-2 flex items-center gap-2">
                <MessageSquare className="w-4 h-4" />
                Speaker Notes
              </h5>
              <div className="p-3 rounded bg-amber-50 dark:bg-amber-900/20 border-l-4 border-amber-500 text-sm">
                {slide.notes}
              </div>
            </div>
          )}

          {/* Tables */}
          {slide.tables.length > 0 && (
            <div className="mt-4">
              <h5 className="text-sm font-medium text-[rgb(var(--muted-foreground))] mb-2 flex items-center gap-2">
                <Table2 className="w-4 h-4" />
                Tables
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
                              className={`
                                p-2 border border-[rgb(var(--border))]
                                ${ri === 0 ? 'bg-[rgb(var(--accent))] font-medium' : 'bg-[rgb(var(--background))]'}
                              `}
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

          {/* Shapes */}
          {slide.shapes.length > 0 && (
            <div className="mt-4">
              <h5 className="text-sm font-medium text-[rgb(var(--muted-foreground))] mb-2">Shapes</h5>
              <div className="flex flex-wrap gap-2">
                {slide.shapes.filter(s => s.text).map((shape, i) => (
                  <span key={i} className="badge">
                    {shape.type}: {shape.text.substring(0, 50)}{shape.text.length > 50 ? '...' : ''}
                  </span>
                ))}
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

function MetadataItem({ 
  icon: Icon, 
  label, 
  value 
}: { 
  icon: React.ElementType; 
  label: string; 
  value: string;
}) {
  return (
    <div className="p-3 rounded-lg bg-[rgb(var(--secondary))]">
      <div className="flex items-center gap-2 text-[rgb(var(--muted-foreground))] mb-1">
        <Icon className="w-4 h-4" />
        <span className="text-sm">{label}</span>
      </div>
      <p className="font-medium text-[rgb(var(--foreground))]">
        {value || 'N/A'}
      </p>
    </div>
  );
}
