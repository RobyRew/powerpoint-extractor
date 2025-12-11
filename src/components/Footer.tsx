/**
 * Footer Component
 */

import { Heart } from 'lucide-react';

export function Footer() {
  return (
    <footer className="border-t border-[rgb(var(--border))] py-6 mt-auto">
      <div className="container mx-auto px-4 text-center">
        <p className="text-sm text-[rgb(var(--muted-foreground))] flex items-center justify-center gap-1">
          Made with <Heart className="w-4 h-4 text-red-500 fill-red-500" /> by{' '}
          <a
            href="https://github.com/RobyRew"
            target="_blank"
            rel="noopener noreferrer"
            className="font-medium hover:underline"
          >
            RobyRew
          </a>
        </p>
        <p className="text-xs text-[rgb(var(--muted-foreground))] mt-2">
          Supports PPT (97-2003) and PPTX (2007+) formats
        </p>
      </div>
    </footer>
  );
}
