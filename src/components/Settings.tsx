/**
 * Settings Panel Component
 * Slide-out settings panel matching calendar-event-generator style
 */

import { useState, useEffect, useRef, useCallback } from 'react';
import { 
  X, 
  Sun, 
  Moon, 
  Smartphone, 
  Layers, 
  Globe, 
  Trash2, 
  Check, 
  AlertTriangle,
  HardDrive
} from 'lucide-react';
import { useI18n } from '../context';
import type { ThemeId } from '../styles/themes';
import { THEMES } from '../styles/themes';

interface SettingsProps {
  isOpen: boolean;
  onClose: () => void;
  theme: ThemeId;
  onThemeChange: (theme: ThemeId) => void;
}

const themeIcons: Record<ThemeId, typeof Sun> = {
  light: Sun,
  dark: Moon,
  oled: Smartphone,
  neumorphic: Layers,
};

export function Settings({ isOpen, onClose, theme, onThemeChange }: SettingsProps) {
  const { t, language, setLanguage, languageNames, availableLanguages } = useI18n();
  const [showClearConfirm, setShowClearConfirm] = useState(false);
  const [storageSize, setStorageSize] = useState<string>('0 KB');
  const panelRef = useRef<HTMLDivElement>(null);

  // Calculate storage size
  useEffect(() => {
    if (isOpen) {
      let totalSize = 0;
      for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (key) {
          const value = localStorage.getItem(key);
          if (value) {
            totalSize += key.length + value.length;
          }
        }
      }
      // Convert to KB
      const sizeInKB = totalSize / 1024;
      setStorageSize(sizeInKB < 1 ? '< 1 KB' : `${sizeInKB.toFixed(1)} KB`);
    }
  }, [isOpen]);

  // Handle escape key
  useEffect(() => {
    const handleEscape = (e: KeyboardEvent) => {
      if (e.key === 'Escape' && isOpen) {
        onClose();
      }
    };

    document.addEventListener('keydown', handleEscape);
    return () => document.removeEventListener('keydown', handleEscape);
  }, [isOpen, onClose]);

  // Handle click outside
  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (panelRef.current && !panelRef.current.contains(e.target as Node)) {
        onClose();
      }
    };

    if (isOpen) {
      // Add slight delay to prevent immediate closing when opening
      const timeout = setTimeout(() => {
        document.addEventListener('mousedown', handleClickOutside);
      }, 100);
      return () => {
        clearTimeout(timeout);
        document.removeEventListener('mousedown', handleClickOutside);
      };
    }
  }, [isOpen, onClose]);

  // Clear all data
  const handleClearData = useCallback(() => {
    // Clear localStorage items related to the app
    const keysToRemove = [];
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i);
      if (key && key.startsWith('pptx-extractor')) {
        keysToRemove.push(key);
      }
    }
    keysToRemove.forEach(key => localStorage.removeItem(key));
    
    setShowClearConfirm(false);
    setStorageSize('0 KB');
    
    // Reload to reset app state
    window.location.reload();
  }, []);

  if (!isOpen) return null;

  return (
    <>
      {/* Backdrop */}
      <div 
        className="fixed inset-0 bg-black/50 z-40 backdrop-blur-sm transition-opacity"
        aria-hidden="true"
      />
      
      {/* Panel */}
      <div
        ref={panelRef}
        className={`
          fixed right-0 top-0 h-full w-full max-w-md z-50
          bg-[rgb(var(--background))] border-l border-[rgb(var(--border))]
          shadow-2xl transform transition-transform duration-300 ease-out
          ${isOpen ? 'translate-x-0' : 'translate-x-full'}
        `}
        role="dialog"
        aria-modal="true"
        aria-label={t.settings}
      >
        {/* Header */}
        <div className="flex items-center justify-between p-4 border-b border-[rgb(var(--border))]">
          <h2 className="text-xl font-semibold text-[rgb(var(--foreground))]">
            {t.settings}
          </h2>
          <button
            onClick={onClose}
            className="btn btn-ghost p-2 rounded-lg hover:bg-[rgb(var(--muted))]"
            aria-label={t.close}
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Content */}
        <div className="overflow-y-auto h-[calc(100%-64px)] p-4 space-y-6">
          {/* Appearance Section */}
          <section>
            <h3 className="text-sm font-medium text-[rgb(var(--muted-foreground))] uppercase tracking-wider mb-3">
              {t.appearance}
            </h3>
            
            {/* Theme Selection */}
            <div className="space-y-2">
              <label className="text-sm font-medium text-[rgb(var(--foreground))]">
                {t.theme}
              </label>
              <div className="grid grid-cols-2 gap-2">
                {THEMES.map((themeOption) => {
                  const Icon = themeIcons[themeOption.id];
                  const isSelected = theme === themeOption.id;
                  
                  return (
                    <button
                      key={themeOption.id}
                      onClick={() => onThemeChange(themeOption.id)}
                      className={`
                        flex items-center gap-3 p-3 rounded-lg border-2 transition-all
                        ${isSelected 
                          ? 'border-[rgb(var(--primary))] bg-[rgb(var(--primary))]/10' 
                          : 'border-[rgb(var(--border))] hover:border-[rgb(var(--primary))]/50'
                        }
                      `}
                    >
                      <Icon className={`w-5 h-5 ${isSelected ? 'text-[rgb(var(--primary))]' : 'text-[rgb(var(--muted-foreground))]'}`} />
                      <span className={`text-sm font-medium ${isSelected ? 'text-[rgb(var(--primary))]' : 'text-[rgb(var(--foreground))]'}`}>
                        {t[themeOption.id as keyof typeof t] || themeOption.name}
                      </span>
                      {isSelected && (
                        <Check className="w-4 h-4 ml-auto text-[rgb(var(--primary))]" />
                      )}
                    </button>
                  );
                })}
              </div>
            </div>
          </section>

          {/* Language Section */}
          <section>
            <h3 className="text-sm font-medium text-[rgb(var(--muted-foreground))] uppercase tracking-wider mb-3">
              {t.language}
            </h3>
            
            <div className="space-y-2">
              <div className="flex items-center gap-2 text-sm text-[rgb(var(--muted-foreground))] mb-2">
                <Globe className="w-4 h-4" />
                <span>{t.language}</span>
              </div>
              <div className="grid grid-cols-2 gap-2">
                {availableLanguages.map((lang) => {
                  const isSelected = language === lang;
                  
                  return (
                    <button
                      key={lang}
                      onClick={() => setLanguage(lang)}
                      className={`
                        flex items-center justify-between p-3 rounded-lg border-2 transition-all
                        ${isSelected 
                          ? 'border-[rgb(var(--primary))] bg-[rgb(var(--primary))]/10' 
                          : 'border-[rgb(var(--border))] hover:border-[rgb(var(--primary))]/50'
                        }
                      `}
                    >
                      <span className={`text-sm font-medium ${isSelected ? 'text-[rgb(var(--primary))]' : 'text-[rgb(var(--foreground))]'}`}>
                        {languageNames[lang]}
                      </span>
                      {isSelected && (
                        <Check className="w-4 h-4 text-[rgb(var(--primary))]" />
                      )}
                    </button>
                  );
                })}
              </div>
            </div>
          </section>

          {/* Storage Section */}
          <section>
            <h3 className="text-sm font-medium text-[rgb(var(--muted-foreground))] uppercase tracking-wider mb-3">
              {t.storage}
            </h3>
            
            {/* Storage Info */}
            <div className="p-3 rounded-lg bg-[rgb(var(--muted))]/50 border border-[rgb(var(--border))] mb-3">
              <div className="flex items-center justify-between">
                <div className="flex items-center gap-2">
                  <HardDrive className="w-4 h-4 text-[rgb(var(--muted-foreground))]" />
                  <span className="text-sm text-[rgb(var(--foreground))]">{t.storageUsed}</span>
                </div>
                <span className="text-sm font-medium text-[rgb(var(--foreground))]">{storageSize}</span>
              </div>
            </div>

            {/* Clear Data */}
            {!showClearConfirm ? (
              <button
                onClick={() => setShowClearConfirm(true)}
                className="w-full flex items-center justify-center gap-2 p-3 rounded-lg border-2 border-red-500/50 text-red-500 hover:bg-red-500/10 transition-all"
              >
                <Trash2 className="w-4 h-4" />
                <span className="text-sm font-medium">{t.clearData}</span>
              </button>
            ) : (
              <div className="p-4 rounded-lg border-2 border-red-500 bg-red-500/10">
                <div className="flex items-start gap-3 mb-3">
                  <AlertTriangle className="w-5 h-5 text-red-500 flex-shrink-0 mt-0.5" />
                  <p className="text-sm text-[rgb(var(--foreground))]">
                    {t.clearDataConfirm}
                  </p>
                </div>
                <div className="flex gap-2">
                  <button
                    onClick={() => setShowClearConfirm(false)}
                    className="flex-1 p-2 rounded-lg border border-[rgb(var(--border))] text-sm font-medium hover:bg-[rgb(var(--muted))] transition-all"
                  >
                    {t.cancel}
                  </button>
                  <button
                    onClick={handleClearData}
                    className="flex-1 p-2 rounded-lg bg-red-500 text-white text-sm font-medium hover:bg-red-600 transition-all"
                  >
                    {t.confirm}
                  </button>
                </div>
              </div>
            )}
          </section>

          {/* Version Info */}
          <section className="pt-4 border-t border-[rgb(var(--border))]">
            <div className="flex items-center justify-between text-sm text-[rgb(var(--muted-foreground))]">
              <span>{t.version}</span>
              <span className="font-mono">1.0.0</span>
            </div>
          </section>
        </div>
      </div>
    </>
  );
}
