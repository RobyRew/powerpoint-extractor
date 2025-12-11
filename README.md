# PowerPoint Extractor

A web application to extract and export data from PowerPoint files (PPT and PPTX).

![PowerPoint Extractor](https://img.shields.io/badge/PowerPoint-Extractor-orange)
![TypeScript](https://img.shields.io/badge/TypeScript-5.0-blue)
![React](https://img.shields.io/badge/React-18-blue)
![Tailwind CSS](https://img.shields.io/badge/Tailwind-CSS-cyan)

## Features

- 📊 **Full Data Extraction** - Extract text, metadata, themes, speaker notes, tables, shapes, and more
- 📁 **Multiple Export Formats** - Export to JSON, XML, CSV, TXT, HTML, and PDF
- 🖼️ **Media Extraction** - Extract and download images and media files separately
- 🎨 **Theme Support** - Light, Dark, OLED, and Neumorphic themes
- 🌐 **Multi-language Support** - English and Spanish (i18n)
- ⚙️ **Settings Panel** - Customize behavior with command palette support
- 📱 **Responsive Design** - Works on desktop and mobile devices
- 🔄 **Multiple Files** - Upload and process multiple files at once
- 👁️ **Data Viewer** - View extracted data in a beautiful modal interface

## Supported Formats

| Format | Extension | Support Level |
|--------|-----------|---------------|
| PowerPoint 2007+ | `.pptx` | Full support |
| PowerPoint 97-2003 | `.ppt` | Full text extraction (MS-PPT spec) |

## Extracted Data

### From PPTX files:
- **Metadata**: Title, creator, dates, revision, keywords, description, application version
- **Slides**: Title, text content, shapes, tables
- **Speaker Notes**: Full notes for each slide
- **Themes**: Color schemes, font schemes
- **Media**: Images, videos, audio files
- **Custom Properties**: Any custom document properties

### From PPT files (Legacy):
- **Full text extraction** using MS-PPT binary format specification
- **Unicode and ANSI text** support (UTF-16LE and Windows-1252)
- **Metadata** from OLE property streams
- **Image extraction** (JPEG, PNG, EMF, WMF)
- **Slide organization** with automatic title detection

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

- **React 18** - UI framework
- **TypeScript** - Type safety
- **Vite** - Build tool
- **Tailwind CSS** - Styling
- **JSZip** - ZIP file handling
- **CFB** - OLE Compound Document parsing (for .ppt files)
- **pptx-parser** - PPTX file parsing
- **jsPDF** - PDF generation
- **Lucide React** - Icons
- **i18next** - Internationalization

## Development

### Prerequisites

- Node.js 18+
- npm or yarn

### Installation

```bash
# Clone the repository
git clone https://github.com/RobyRew/powerpoint-extractor.git
cd powerpoint-extractor

# Install dependencies
npm install

# Start development server
npm run dev
```

### Build

```bash
npm run build
```

### Preview production build

```bash
npm run preview
```

## Deployment

### Docker

```bash
docker build -t powerpoint-extractor .
docker run -p 80:80 powerpoint-extractor
```

### Dokploy

This project is ready for deployment with Dokploy. Just connect your GitHub repository and deploy.

## Privacy

All processing happens locally in your browser. No files are uploaded to any server.

## Credits

- **[SheetJS/js-cfb](https://github.com/SheetJS/js-cfb)** - CFB (Compound File Binary) library for parsing OLE documents
- **[pptx-parser](https://www.npmjs.com/package/pptx-parser)** - PPTX file parsing library
- **[js-ppt](https://github.com/nicwaller/js-ppt)** - Reference implementation for PPT binary format parsing
- **[MS-PPT Specification](https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-ppt/)** - Microsoft PowerPoint Binary File Format documentation
- **[Lucide](https://lucide.dev/)** - Beautiful & consistent icon toolkit

## License

MIT License - feel free to use this project for any purpose.

## Author

Made with ❤️ by [RobyRew](https://github.com/RobyRew)
