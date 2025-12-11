/**
 * Header Component
 * Matching calendar-event-generator style
 */

import { Settings as SettingsIcon, FileSpreadsheet } from 'lucide-react';
import { useI18n } from '../context';

interface HeaderProps {
  onSettingsClick: () => void;
}

export function Header({ onSettingsClick }: HeaderProps) {
  const { t } = useI18n();

  return (
    <header className="sticky top-0 z-40 border-b border-[rgb(var(--border))] bg-[rgb(var(--background))] backdrop-blur-sm">
      <div className="container mx-auto px-4 h-16 flex items-center justify-between">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-xl bg-[rgb(var(--primary))] flex items-center justify-center">
            <FileSpreadsheet className="w-5 h-5 text-[rgb(var(--primary-foreground))]" />
          </div>
          <div>
            <h1 className="text-lg font-semibold text-[rgb(var(--foreground))]">
              {t.appName}
            </h1>
            <p className="text-xs text-[rgb(var(--muted-foreground))] hidden sm:block">
              {t.appDescription}
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2">
          <button
            onClick={onSettingsClick}
            className="btn btn-ghost p-2 rounded-lg hover:bg-[rgb(var(--muted))] transition-colors"
            aria-label={t.settings}
          >
            <SettingsIcon className="w-5 h-5" />
          </button>
        </div>
      </div>
    </header>
  );
}
