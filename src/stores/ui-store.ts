import { create } from 'zustand'
import type { ActiveView } from '@/types'

export type { ActiveView }

interface UIState {
  activeView: ActiveView
  isSidebarOpen: boolean
  isCommandPaletteOpen: boolean
  viewingPresentationId: string | null
  searchQuery: string
  toastMessage: string | null
  toastType: 'success' | 'error' | 'info'

  setActiveView: (view: ActiveView) => void
  toggleSidebar: () => void
  setSidebarOpen: (open: boolean) => void
  setCommandPaletteOpen: (open: boolean) => void
  toggleCommandPalette: () => void
  setViewingPresentation: (id: string | null) => void
  setSearchQuery: (query: string) => void
  showToast: (message: string, type?: 'success' | 'error' | 'info') => void
  clearToast: () => void
}

export const useUIStore = create<UIState>((set, get) => ({
  activeView: 'extractor',
  isSidebarOpen: true,
  isCommandPaletteOpen: false,
  viewingPresentationId: null,
  searchQuery: '',
  toastMessage: null,
  toastType: 'info',

  setActiveView: (view) => set({ activeView: view }),
  toggleSidebar: () => set({ isSidebarOpen: !get().isSidebarOpen }),
  setSidebarOpen: (open) => set({ isSidebarOpen: open }),
  setCommandPaletteOpen: (open) => set({ isCommandPaletteOpen: open }),
  toggleCommandPalette: () => set({ isCommandPaletteOpen: !get().isCommandPaletteOpen }),
  setViewingPresentation: (id) => set({ viewingPresentationId: id }),
  setSearchQuery: (query) => set({ searchQuery: query }),
  showToast: (message, type = 'info') => {
    set({ toastMessage: message, toastType: type })
    setTimeout(() => set({ toastMessage: null }), 3500)
  },
  clearToast: () => set({ toastMessage: null }),
}))
