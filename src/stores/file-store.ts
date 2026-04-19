import { create } from 'zustand'
import type { ExtractedPresentation } from '@/types'
import { parsePPTX, parsePPT } from '@/lib'

interface FileState {
  files: File[]
  presentations: ExtractedPresentation[]
  processingFile: string | null
  error: string | null

  addFiles: (files: File[]) => Promise<void>
  removeFile: (index: number) => void
  clearAll: () => void
  setError: (error: string | null) => void
  getPresentation: (id: string) => ExtractedPresentation | undefined
}

export const useFileStore = create<FileState>((set, get) => ({
  files: [],
  presentations: [],
  processingFile: null,
  error: null,

  addFiles: async (newFiles) => {
    const { files: existing, presentations } = get()
    const toProcess = newFiles.filter(
      (f) => !existing.find((e) => e.name === f.name),
    )
    if (toProcess.length === 0) return

    set({ files: [...existing, ...toProcess], error: null })

    for (const file of toProcess) {
      if (presentations.find((p) => p.fileName === file.name)) continue
      set({ processingFile: file.name })

      try {
        const isPPTX = file.name.toLowerCase().endsWith('.pptx')
        const data = isPPTX ? await parsePPTX(file) : await parsePPT(file)
        set((s) => ({ presentations: [...s.presentations, data] }))
      } catch (err) {
        const msg = err instanceof Error ? err.message : 'Unknown error'
        set({ error: `Error processing ${file.name}: ${msg}` })
      }
    }
    set({ processingFile: null })
  },

  removeFile: (index) => {
    const { files } = get()
    const file = files[index]
    if (!file) return
    set((s) => ({
      files: s.files.filter((_, i) => i !== index),
      presentations: s.presentations.filter((p) => p.fileName !== file.name),
    }))
  },

  clearAll: () => set({ files: [], presentations: [], error: null, processingFile: null }),

  setError: (error) => set({ error }),

  getPresentation: (id) => get().presentations.find((p) => p.id === id),
}))
