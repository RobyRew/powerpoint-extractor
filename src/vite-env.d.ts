/// <reference types="vite/client" />

// Without this reference TypeScript does not know about import.meta.env and the
// build fails with TS2339: Property 'env' does not exist on type 'ImportMeta'.
// calendar-event-generator already had this file; this repo did not.
interface ImportMetaEnv {
  readonly VITE_UMAMI_SCRIPT_URL?: string;
  readonly VITE_UMAMI_WEBSITE_ID?: string;
  readonly VITE_UMAMI_DOMAINS?: string;
}

interface ImportMeta {
  readonly env: ImportMetaEnv;
}
