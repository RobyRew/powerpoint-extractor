/**
 * Header Component
 */

import { Sun, Moon, Smartphone, Layers, FileSpreadsheet, Github } from 'lucide-react';
import type { ThemeId } from '../styles/themes';
import { THEMES } from '../styles/themes';

interface HeaderProps {
  theme: ThemeId;
  onThemeChange: (theme: ThemeId) => void;
}

const themeIcons = {
  light: Sun,
  dark: Moon,
  oled: Smartphone,
  neumorphic: Layers,
};

export function Header({ theme, onThemeChange }: HeaderProps) {
  const currentTheme = THEMES.find(t => t.id === theme) || THEMES[0];
  const ThemeIcon = themeIcons[theme];

  const cycleTheme = () => {
    const currentIndex = THEMES.findIndex(t => t.id === theme);
    const nextIndex = (currentIndex + 1) % THEMES.length;
    onThemeChange(THEMES[nextIndex].id);
  };

  return (
    <header className="sticky top-0 z-40 border-b border-[rgb(var(--border))] bg-[rgb(var(--background))] backdrop-blur-sm">
      <div className="container mx-auto px-4 h-16 flex items-center justify-between">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-xl bg-[rgb(var(--primary))] flex items-center justify-center">
            <FileSpreadsheet className="w-5 h-5 text-[rgb(var(--primary-foreground))]" />
          </div>
          <div>
            <h1 className="text-lg font-semibold text-[rgb(var(--foreground))]">
              PowerPoint Extractor
            </h1>
            <p className="text-xs text-[rgb(var(--muted-foreground))] hidden sm:block">
              Extract data from PPT & PPTX files
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2">
          <a
            href="https://github.com/RobyRew/powerpoint-extractor"
            target="_blank"
            rel="noopener noreferrer"
            className="btn btn-ghost p-2"
            title="View on GitHub"
          >
            <Github className="w-5 h-5" />
          </a>
          
          <button
            onClick={cycleTheme}
            className="btn btn-ghost p-2 relative group"
            title={`Theme: ${currentTheme.name}`}
          >
            <ThemeIcon className="w-5 h-5" />
            <span className="absolute -bottom-8 left-1/2 -translate-x-1/2 px-2 py-1 text-xs rounded bg-[rgb(var(--popover))] border border-[rgb(var(--border))] opacity-0 group-hover:opacity-100 transition-opacity whitespace-nowrap pointer-events-none">
              {currentTheme.name}
            </span>
          </button>
        </div>
      </div>
    </header>
  );
}
