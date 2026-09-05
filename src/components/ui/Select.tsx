import { type SelectHTMLAttributes, forwardRef } from 'react'
import { cn } from '@/lib/utils'

interface SelectProps extends SelectHTMLAttributes<HTMLSelectElement> {
  label?: string
  options: { value: string; label: string }[]
}

export const Select = forwardRef<HTMLSelectElement, SelectProps>(
  ({ label, options, className, id, ...props }, ref) => {
    const selectId = id ?? label?.toLowerCase().replace(/\s+/g, '-')
    return (
      <div className="flex flex-col gap-1">
        {label && (
          <label htmlFor={selectId} className="rw-label">
            {label}
          </label>
        )}
        <select
          ref={ref}
          id={selectId}
          className={cn(
            // .rw-field plus the disclosure chevron; the native arrow is
            // removed so the control matches the other fields.
            'rw-field h-10 appearance-none pr-8',
            'bg-[url("data:image/svg+xml,%3Csvg%20xmlns%3D%27http%3A//www.w3.org/2000/svg%27%20width%3D%2712%27%20height%3D%2712%27%20fill%3D%27none%27%20stroke%3D%27%238e8e93%27%20stroke-width%3D%272%27%20stroke-linecap%3D%27round%27%3E%3Cpath%20d%3D%27M3%204.5l3%203%203-3%27/%3E%3C/svg%3E")]',
            'bg-no-repeat bg-[right_12px_center]',
            className,
          )}
          {...props}
        >
          {options.map((opt) => (
            <option key={opt.value} value={opt.value}>
              {opt.label}
            </option>
          ))}
        </select>
      </div>
    )
  },
)
Select.displayName = 'Select'
