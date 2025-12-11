// Type declarations for packages without types

declare module 'pptx-parser' {
  interface SlideElement {
    type: string;
    content?: string;
    text?: string;
    shapeType?: string;
    table?: string[][];
  }

  interface SlideData {
    elements?: SlideElement[];
    notes?: string;
  }

  export default function parse(file: File): Promise<SlideData[]>;
}

declare module 'cfb' {
  interface CFBContainer {
    find(name: string): CFBEntry | null;
  }

  interface CFBEntry {
    content: Uint8Array & { l?: number; read_shift?: (bytes: number) => number };
  }

  export function read(data: Uint8Array | ArrayBuffer, options?: { type?: string }): CFBContainer;
}

declare module 'codepage' {
  export const utils: {
    decode(codepage: number, data: Uint8Array): string;
  };
}
