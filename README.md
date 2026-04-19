<p align="center">
  <img src="public/favicon.svg" width="80" height="80" alt="PowerPoint Extractor">
</p>

<h1 align="center">PowerPoint Extractor</h1>

<p align="center">
  Extract, view, and export data from PowerPoint files — entirely in your browser.
</p>

<p align="center">
  <a href="https://robyrew.github.io/powerpoint-extractor/"><strong>Live Demo</strong></a> ·
  <a href="#features">Features</a> ·
  <a href="#tech-stack">Tech Stack</a> ·
  <a href="#development">Development</a>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/version-2.0-orange?style=flat-square" alt="Version">
  <img src="https://img.shields.io/github/license/RobyRew/powerpoint-extractor?style=flat-square" alt="License">
  <img src="https://img.shields.io/badge/TypeScript-5.8-blue?style=flat-square&logo=typescript&logoColor=white" alt="TypeScript">
  <img src="https://img.shields.io/badge/React-19-blue?style=flat-square&logo=react&logoColor=white" alt="React">
  <img src="https://img.shields.io/badge/Tailwind_CSS-4-06B6D4?style=flat-square&logo=tailwindcss&logoColor=white" alt="Tailwind CSS">
  <img src="https://img.shields.io/badge/Vite-6-646CFF?style=flat-square&logo=vite&logoColor=white" alt="Vite">
</p>

---

## Features

**Extraction**
- Full data extraction from both `.pptx` (Office Open XML) and `.ppt` (MS-PPT binary) files
- Text, metadata, speaker notes, tables, shapes, themes, and color schemes
- Image and media extraction (JPEG, PNG, EMF, WMF)
- Multiple file processing — drag & drop or browse

**Export**
- Six export formats: **JSON**, **XML**, **CSV**, **TXT**, **HTML**, **PDF**
- Media download as ZIP archive
- Customizable filename patterns and export settings

**Interface**
- Three themes: Light, Dark, and OLED
- Four languages: English, Spanish, German, French
- Responsive design with mobile-first bottom navigation
- Command palette (⌘K) for quick actions
- Keyboard shortcuts for power users
- PWA — installable on any device

**Privacy**
- 100% client-side processing — no files leave your device
- No analytics, no tracking, no server uploads

## Supported Formats

| Format | Extension | Support |
|--------|-----------|---------|
| PowerPoint 2007+ | `.pptx` | Full (metadata, slides, notes, media, themes) |
| PowerPoint 97-2003 | `.ppt` | Full text extraction, metadata, images |

## Extracted Data

<details>
<summary><strong>PPTX files</strong></summary>

- **Metadata** — Title, creator, dates, revision, keywords, description, app version
- **Slides** — Title, text content, shapes, tables
- **Speaker Notes** — Full notes per slide
- **Themes** — Color schemes, font schemes
- **Media** — Images, videos, audio files
- **Custom Properties** — Any custom document properties
</details>

<details>
<summary><strong>PPT files (legacy)</strong></summary>

- **Full text extraction** using MS-PPT binary format specification
- **Unicode and ANSI** support (UTF-16LE and Windows-1252)
- **Metadata** from OLE property streams
- **Image extraction** (JPEG, PNG, EMF, WMF)
- **Slide organization** with automatic title detection
</details>

## Export Formats

| Format | Description | Use Case |
|--------|-------------|----------|
| JSON | Full structured data | Programming, APIs |
| XML | Structured markup | Data interchange |
| CSV | Spreadsheet format | Excel, data analysis |
| TXT | Plain text | Quick reading |
| HTML | Web page | Viewing in browser |
| PDF | Document | Printing, sharing |

## Tech Stack

| Category | Technology |
|----------|-----------|
| Framework | React 19 |
| Language | TypeScript 5.8 |
| Styling | Tailwind CSS 4 |
| Build | Vite 6 |
| State | Zustand 5 |
| PWA | vite-plugin-pwa |
| Parsers | JSZip, CFB, pptx-parser |
| PDF | jsPDF |
| Icons | Lucide React |
| Deploy | GitHub Pages |

## Development

### Prerequisites

- Node.js 20+
- npm

### Setup

```bash
git clone https://github.com/RobyRew/powerpoint-extractor.git
cd powerpoint-extractor
npm install
npm run dev
```

### Build

```bash
npm run build
```

### Preview

```bash
npm run preview
```

### Docker

```bash
docker build -t powerpoint-extractor .
docker run -p 80:80 powerpoint-extractor
```

## Deployment

Push to `main` and GitHub Actions will build and deploy to GitHub Pages automatically.

Manual deployment is also supported via Docker or any static hosting provider — just serve the `dist/` folder.

## Credits

- [SheetJS/js-cfb](https://github.com/SheetJS/js-cfb) — OLE Compound Document parsing
- [pptx-parser](https://www.npmjs.com/package/pptx-parser) — PPTX file parsing
- [js-ppt](https://github.com/nicwaller/js-ppt) — PPT binary format reference
- [MS-PPT Specification](https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-ppt/) — Microsoft binary format docs
- [Lucide](https://lucide.dev/) — Icon toolkit

## License

[MIT](LICENSE)

---

<p align="center">Made with ❤️ by <a href="https://github.com/RobyRew">RobyRew</a></p>
