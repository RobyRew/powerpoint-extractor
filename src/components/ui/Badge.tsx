import { cn } from '@/lib/utils'

interface BadgeProps {
  children: React.ReactNode
  variant?: 'default' | 'accent' | 'success' | 'warning' | 'danger'
  className?: string
}

export function Badge({ children, variant = 'default', className }: BadgeProps) {
  return (
    <span
      className={cn(
        // .rw-badge is the shape; the semantic tints stay local because they
        // are this app's status vocabulary, not design-system colours.
        'rw-badge !rounded-full',
        variant === 'default' && '!bg-surface-3 !text-ink-2',
        variant === 'accent' && '!bg-accent/15 !text-accent',
        variant === 'success' && '!bg-success/15 !text-success',
        variant === 'warning' && '!bg-warning/15 !text-warning',
        variant === 'danger' && '!bg-danger/15 !text-danger',
        className,
      )}
    >
      {children}
    </span>
  )
}
