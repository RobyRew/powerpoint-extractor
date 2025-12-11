/**
 * Footer Component
 * Matching calendar-event-generator style with mobile/desktop variants
 */

import { Heart, Github } from 'lucide-react';
import { useI18n } from '../context';

export function Footer() {
  const { t } = useI18n();
  const currentYear = new Date().getFullYear();

  return (
    <footer className="border-t border-[rgb(var(--border))] py-4 mt-auto">
      <div className="container mx-auto px-4">
        {/* Mobile Footer */}
        <div className="flex md:hidden flex-col items-center gap-2">
          <div className="flex items-center gap-1 text-sm text-[rgb(var(--muted-foreground))]">
            <span>{t.madeWith}</span>
            <Heart className="w-4 h-4 text-red-500 fill-red-500" />
            <span>{t.by}</span>
            <a
              href="https://github.com/RobyRew"
              target="_blank"
              rel="noopener noreferrer"
              className="font-medium text-[rgb(var(--foreground))] hover:underline"
            >
              RobyRew
            </a>
          </div>
          <p className="text-xs text-[rgb(var(--muted-foreground))]">
            © {currentYear} · v1.0.0
          </p>
        </div>

        {/* Desktop Footer */}
        <div className="hidden md:flex items-center justify-between">
          <div className="flex items-center gap-4">
            <a
              href="https://github.com/RobyRew/powerpoint-extractor"
              target="_blank"
              rel="noopener noreferrer"
              className="flex items-center gap-2 text-sm text-[rgb(var(--muted-foreground))] hover:text-[rgb(var(--foreground))] transition-colors"
            >
              <Github className="w-4 h-4" />
              <span>GitHub</span>
            </a>
            <span className="text-sm text-[rgb(var(--muted-foreground))]">
              {t.supportsPPT}
            </span>
          </div>
          
          <div className="flex items-center gap-4">
            <div className="flex items-center gap-1 text-sm text-[rgb(var(--muted-foreground))]">
              <span>{t.madeWith}</span>
              <Heart className="w-4 h-4 text-red-500 fill-red-500" />
              <span>{t.by}</span>
              <a
                href="https://github.com/RobyRew"
                target="_blank"
                rel="noopener noreferrer"
                className="font-medium text-[rgb(var(--foreground))] hover:underline"
              >
                RobyRew
              </a>
            </div>
            <span className="text-sm text-[rgb(var(--muted-foreground))]">
              © {currentYear} · v1.0.0
            </span>
          </div>
        </div>
      </div>
    </footer>
  );
}
