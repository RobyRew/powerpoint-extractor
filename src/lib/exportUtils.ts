/**
 * Export utilities - Convert extracted data to various formats
 */

import JSZip from 'jszip';
import { saveAs } from 'file-saver';
import { jsPDF } from 'jspdf';
import type { ExtractedPresentation } from '../types';

/**
 * Export to JSON format
 */
export function exportToJSON(presentations: ExtractedPresentation[]): string {
  return JSON.stringify(presentations, null, 2);
}

/**
 * Export to XML format
 */
export function exportToXML(presentations: ExtractedPresentation[]): string {
  let xml = '<?xml version="1.0" encoding="UTF-8"?>\n';
  xml += '<presentations>\n';
  
  for (const pres of presentations) {
    xml += `  <presentation id="${escapeXml(pres.id)}">\n`;
    xml += `    <fileName>${escapeXml(pres.fileName)}</fileName>\n`;
    xml += `    <fileSize>${pres.fileSize}</fileSize>\n`;
    xml += `    <fileType>${escapeXml(pres.fileType)}</fileType>\n`;
    xml += `    <extractedAt>${escapeXml(pres.extractedAt)}</extractedAt>\n`;
    
    // Metadata
    xml += '    <metadata>\n';
    for (const [key, value] of Object.entries(pres.metadata)) {
      xml += `      <${key}>${escapeXml(String(value))}</${key}>\n`;
    }
    xml += '    </metadata>\n';
    
    // Slides
    xml += '    <slides>\n';
    for (const slide of pres.slides) {
      xml += `      <slide number="${slide.slideNumber}">\n`;
      xml += `        <title>${escapeXml(slide.title)}</title>\n`;
      xml += '        <textContent>\n';
      for (const text of slide.textContent) {
        xml += `          <text>${escapeXml(text)}</text>\n`;
      }
      xml += '        </textContent>\n';
      if (slide.notes) {
        xml += `        <notes>${escapeXml(slide.notes)}</notes>\n`;
      }
      
      // Shapes
      if (slide.shapes.length > 0) {
        xml += '        <shapes>\n';
        for (const shape of slide.shapes) {
          xml += `          <shape type="${escapeXml(shape.type)}">${escapeXml(shape.text)}</shape>\n`;
        }
        xml += '        </shapes>\n';
      }
      
      // Tables
      if (slide.tables.length > 0) {
        xml += '        <tables>\n';
        for (const table of slide.tables) {
          xml += `          <table rows="${table.rows}" columns="${table.columns}">\n`;
          for (const row of table.cells) {
            xml += '            <row>\n';
            for (const cell of row) {
              xml += `              <cell>${escapeXml(cell)}</cell>\n`;
            }
            xml += '            </row>\n';
          }
          xml += '          </table>\n';
        }
        xml += '        </tables>\n';
      }
      
      xml += '      </slide>\n';
    }
    xml += '    </slides>\n';
    
    // Themes
    if (pres.themes.length > 0) {
      xml += '    <themes>\n';
      for (const theme of pres.themes) {
        xml += `      <theme name="${escapeXml(theme.name)}">\n`;
        xml += '        <colors>\n';
        for (const color of theme.colors) {
          xml += `          <color>${escapeXml(color)}</color>\n`;
        }
        xml += '        </colors>\n';
        xml += '        <fonts>\n';
        for (const font of theme.fonts) {
          xml += `          <font>${escapeXml(font)}</font>\n`;
        }
        xml += '        </fonts>\n';
        xml += '      </theme>\n';
      }
      xml += '    </themes>\n';
    }
    
    // Custom Properties
    if (Object.keys(pres.customProperties).length > 0) {
      xml += '    <customProperties>\n';
      for (const [key, value] of Object.entries(pres.customProperties)) {
        xml += `      <property name="${escapeXml(key)}">${escapeXml(value)}</property>\n`;
      }
      xml += '    </customProperties>\n';
    }
    
    xml += '  </presentation>\n';
  }
  
  xml += '</presentations>';
  return xml;
}

/**
 * Export to CSV format (slide content focused)
 */
export function exportToCSV(presentations: ExtractedPresentation[]): string {
  const headers = [
    'File Name',
    'Slide Number',
    'Slide Title',
    'Text Content',
    'Notes',
    'Shape Count',
    'Table Count',
    'Creator',
    'Created Date',
    'Modified Date',
  ];
  
  let csv = headers.map(h => `"${h}"`).join(',') + '\n';
  
  for (const pres of presentations) {
    for (const slide of pres.slides) {
      const row = [
        pres.fileName,
        slide.slideNumber.toString(),
        slide.title,
        slide.textContent.join(' | '),
        slide.notes,
        slide.shapes.length.toString(),
        slide.tables.length.toString(),
        pres.metadata.creator,
        pres.metadata.created,
        pres.metadata.modified,
      ];
      csv += row.map(cell => `"${escapeCSV(cell)}"`).join(',') + '\n';
    }
  }
  
  return csv;
}

/**
 * Export to plain text format
 */
export function exportToText(presentations: ExtractedPresentation[]): string {
  let text = '';
  
  for (const pres of presentations) {
    text += '═'.repeat(80) + '\n';
    text += `FILE: ${pres.fileName}\n`;
    text += '═'.repeat(80) + '\n\n';
    
    // Metadata
    text += '--- METADATA ---\n';
    text += `Title: ${pres.metadata.title || 'N/A'}\n`;
    text += `Creator: ${pres.metadata.creator || 'N/A'}\n`;
    text += `Created: ${pres.metadata.created || 'N/A'}\n`;
    text += `Modified: ${pres.metadata.modified || 'N/A'}\n`;
    text += `Application: ${pres.metadata.application || 'N/A'}\n`;
    text += `Total Slides: ${pres.metadata.totalSlides}\n`;
    text += `Total Words: ${pres.metadata.totalWords}\n`;
    text += '\n';
    
    // Slides
    for (const slide of pres.slides) {
      text += '─'.repeat(60) + '\n';
      text += `SLIDE ${slide.slideNumber}: ${slide.title}\n`;
      text += '─'.repeat(60) + '\n';
      
      if (slide.textContent.length > 0) {
        text += '\nContent:\n';
        for (const content of slide.textContent) {
          text += `  • ${content}\n`;
        }
      }
      
      if (slide.notes) {
        text += `\nNotes:\n  ${slide.notes}\n`;
      }
      
      if (slide.tables.length > 0) {
        text += '\nTables:\n';
        for (const table of slide.tables) {
          text += `  [${table.rows}x${table.columns} table]\n`;
          for (const row of table.cells) {
            text += `    ${row.join(' | ')}\n`;
          }
        }
      }
      
      text += '\n';
    }
    
    // Themes
    if (pres.themes.length > 0) {
      text += '\n--- THEMES ---\n';
      for (const theme of pres.themes) {
        text += `Theme: ${theme.name}\n`;
        text += `  Fonts: ${theme.fonts.join(', ')}\n`;
      }
    }
    
    text += '\n\n';
  }
  
  return text;
}

/**
 * Export to HTML format
 */
export function exportToHTML(presentations: ExtractedPresentation[]): string {
  let html = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>PowerPoint Data Export</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background: #f5f5f5; color: #1a1a1a; line-height: 1.6; padding: 2rem; }
    .container { max-width: 1200px; margin: 0 auto; }
    h1 { font-size: 2rem; margin-bottom: 2rem; }
    .presentation { background: white; border-radius: 12px; padding: 2rem; margin-bottom: 2rem; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }
    .presentation-header { border-bottom: 2px solid #e5e5e5; padding-bottom: 1rem; margin-bottom: 1rem; }
    .file-name { font-size: 1.5rem; font-weight: 600; }
    .file-info { color: #666; font-size: 0.875rem; margin-top: 0.5rem; }
    .metadata { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 1rem; margin: 1rem 0; padding: 1rem; background: #f9f9f9; border-radius: 8px; }
    .meta-item { }
    .meta-label { font-size: 0.75rem; text-transform: uppercase; color: #888; }
    .meta-value { font-weight: 500; }
    .slides { margin-top: 2rem; }
    .slide { background: #fafafa; border: 1px solid #e5e5e5; border-radius: 8px; padding: 1.5rem; margin-bottom: 1rem; }
    .slide-header { display: flex; align-items: center; gap: 1rem; margin-bottom: 1rem; }
    .slide-number { background: #1a1a1a; color: white; width: 32px; height: 32px; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-size: 0.875rem; font-weight: 600; }
    .slide-title { font-size: 1.25rem; font-weight: 600; }
    .slide-content { padding-left: 2.5rem; }
    .content-item { margin: 0.5rem 0; padding: 0.5rem; background: white; border-radius: 4px; }
    .notes { margin-top: 1rem; padding: 1rem; background: #fff9e6; border-radius: 4px; border-left: 4px solid #f59e0b; }
    .notes-label { font-weight: 600; font-size: 0.875rem; color: #b45309; margin-bottom: 0.5rem; }
    table { width: 100%; border-collapse: collapse; margin: 1rem 0; }
    th, td { border: 1px solid #e5e5e5; padding: 0.5rem; text-align: left; }
    th { background: #f5f5f5; font-weight: 600; }
    .themes { margin-top: 2rem; }
    .theme { background: #f0f0f0; padding: 1rem; border-radius: 8px; margin-bottom: 1rem; }
    .theme-name { font-weight: 600; margin-bottom: 0.5rem; }
    .theme-detail { font-size: 0.875rem; color: #666; }
  </style>
</head>
<body>
  <div class="container">
    <h1>📊 PowerPoint Data Export</h1>
`;

  for (const pres of presentations) {
    html += `
    <div class="presentation">
      <div class="presentation-header">
        <div class="file-name">${escapeHtml(pres.fileName)}</div>
        <div class="file-info">
          ${formatFileSize(pres.fileSize)} • ${pres.fileType.toUpperCase()} • Extracted: ${new Date(pres.extractedAt).toLocaleString()}
        </div>
      </div>
      
      <div class="metadata">
        <div class="meta-item">
          <div class="meta-label">Title</div>
          <div class="meta-value">${escapeHtml(pres.metadata.title) || 'N/A'}</div>
        </div>
        <div class="meta-item">
          <div class="meta-label">Creator</div>
          <div class="meta-value">${escapeHtml(pres.metadata.creator) || 'N/A'}</div>
        </div>
        <div class="meta-item">
          <div class="meta-label">Created</div>
          <div class="meta-value">${pres.metadata.created ? new Date(pres.metadata.created).toLocaleDateString() : 'N/A'}</div>
        </div>
        <div class="meta-item">
          <div class="meta-label">Modified</div>
          <div class="meta-value">${pres.metadata.modified ? new Date(pres.metadata.modified).toLocaleDateString() : 'N/A'}</div>
        </div>
        <div class="meta-item">
          <div class="meta-label">Application</div>
          <div class="meta-value">${escapeHtml(pres.metadata.application) || 'N/A'}</div>
        </div>
        <div class="meta-item">
          <div class="meta-label">Total Slides</div>
          <div class="meta-value">${pres.metadata.totalSlides}</div>
        </div>
      </div>

      <div class="slides">
        <h3>Slides</h3>
`;

    for (const slide of pres.slides) {
      html += `
        <div class="slide">
          <div class="slide-header">
            <div class="slide-number">${slide.slideNumber}</div>
            <div class="slide-title">${escapeHtml(slide.title) || 'Untitled Slide'}</div>
          </div>
          <div class="slide-content">
`;
      
      for (const content of slide.textContent) {
        html += `            <div class="content-item">${escapeHtml(content)}</div>\n`;
      }
      
      if (slide.notes) {
        html += `
            <div class="notes">
              <div class="notes-label">📝 Speaker Notes</div>
              ${escapeHtml(slide.notes)}
            </div>
`;
      }
      
      for (const table of slide.tables) {
        html += `
            <table>
              <tbody>
`;
        for (let i = 0; i < table.cells.length; i++) {
          html += '                <tr>\n';
          for (const cell of table.cells[i]) {
            const tag = i === 0 ? 'th' : 'td';
            html += `                  <${tag}>${escapeHtml(cell)}</${tag}>\n`;
          }
          html += '                </tr>\n';
        }
        html += `
              </tbody>
            </table>
`;
      }
      
      html += `
          </div>
        </div>
`;
    }
    
    html += `      </div>
`;

    if (pres.themes.length > 0) {
      html += `
      <div class="themes">
        <h3>Themes</h3>
`;
      for (const theme of pres.themes) {
        html += `
        <div class="theme">
          <div class="theme-name">${escapeHtml(theme.name)}</div>
          <div class="theme-detail">Fonts: ${theme.fonts.map(f => escapeHtml(f)).join(', ')}</div>
        </div>
`;
      }
      html += `      </div>
`;
    }
    
    html += `    </div>
`;
  }

  html += `
  </div>
</body>
</html>`;

  return html;
}

/**
 * Export to PDF format
 */
export function exportToPDF(presentations: ExtractedPresentation[]): jsPDF {
  const doc = new jsPDF();
  let yPos = 20;
  const pageHeight = doc.internal.pageSize.height;
  const marginLeft = 20;
  const lineHeight = 7;
  
  const checkNewPage = (needed: number = lineHeight) => {
    if (yPos + needed > pageHeight - 20) {
      doc.addPage();
      yPos = 20;
    }
  };
  
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(20);
  doc.text('PowerPoint Data Export', marginLeft, yPos);
  yPos += 15;
  
  for (const pres of presentations) {
    checkNewPage(50);
    
    // File header
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(14);
    doc.text(pres.fileName, marginLeft, yPos);
    yPos += lineHeight;
    
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(10);
    doc.text(`${formatFileSize(pres.fileSize)} • ${pres.fileType.toUpperCase()}`, marginLeft, yPos);
    yPos += lineHeight * 2;
    
    // Metadata
    doc.setFont('helvetica', 'bold');
    doc.setFontSize(11);
    doc.text('Metadata', marginLeft, yPos);
    yPos += lineHeight;
    
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(9);
    const metaLines = [
      `Title: ${pres.metadata.title || 'N/A'}`,
      `Creator: ${pres.metadata.creator || 'N/A'}`,
      `Created: ${pres.metadata.created || 'N/A'}`,
      `Modified: ${pres.metadata.modified || 'N/A'}`,
      `Total Slides: ${pres.metadata.totalSlides}`,
    ];
    
    for (const line of metaLines) {
      checkNewPage();
      doc.text(line, marginLeft, yPos);
      yPos += lineHeight - 1;
    }
    yPos += lineHeight;
    
    // Slides
    for (const slide of pres.slides) {
      checkNewPage(30);
      
      doc.setFont('helvetica', 'bold');
      doc.setFontSize(11);
      doc.text(`Slide ${slide.slideNumber}: ${slide.title || 'Untitled'}`, marginLeft, yPos);
      yPos += lineHeight;
      
      doc.setFont('helvetica', 'normal');
      doc.setFontSize(9);
      
      for (const content of slide.textContent) {
        const lines = doc.splitTextToSize(`• ${content}`, 170);
        for (const line of lines) {
          checkNewPage();
          doc.text(line, marginLeft + 5, yPos);
          yPos += lineHeight - 2;
        }
      }
      
      if (slide.notes) {
        checkNewPage(15);
        doc.setFont('helvetica', 'italic');
        doc.text('Notes:', marginLeft + 5, yPos);
        yPos += lineHeight - 2;
        const noteLines = doc.splitTextToSize(slide.notes, 165);
        for (const line of noteLines) {
          checkNewPage();
          doc.text(line, marginLeft + 10, yPos);
          yPos += lineHeight - 2;
        }
        doc.setFont('helvetica', 'normal');
      }
      
      yPos += lineHeight;
    }
    
    yPos += lineHeight * 2;
  }
  
  return doc;
}

/**
 * Download content as file
 */
export function downloadFile(content: string | Blob, filename: string, type: string): void {
  const blob = content instanceof Blob ? content : new Blob([content], { type });
  saveAs(blob, filename);
}

/**
 * Download all media as ZIP
 */
export async function downloadMediaAsZip(presentations: ExtractedPresentation[]): Promise<void> {
  const zip = new JSZip();
  
  for (const pres of presentations) {
    if (pres.media.length === 0) continue;
    
    const folderName = pres.fileName.replace(/\.(pptx?|ppt)$/i, '');
    const folder = zip.folder(folderName);
    
    if (folder) {
      for (const media of pres.media) {
        if (media.data) {
          folder.file(media.name, media.data, { base64: true });
        }
      }
    }
  }
  
  const content = await zip.generateAsync({ type: 'blob' });
  saveAs(content, 'powerpoint-media.zip');
}

/**
 * Download all exports as ZIP
 */
export async function downloadAllAsZip(
  presentations: ExtractedPresentation[],
  formats: string[]
): Promise<void> {
  const zip = new JSZip();
  const timestamp = new Date().toISOString().split('T')[0];
  
  // Generate base filename from original file(s)
  const baseFilename = presentations.length === 1
    ? presentations[0].fileName.replace(/\.(pptx?|ppt)$/i, '')
    : 'presentations';
  
  if (formats.includes('json')) {
    zip.file(`${baseFilename}-export-${timestamp}.json`, exportToJSON(presentations));
  }
  if (formats.includes('xml')) {
    zip.file(`${baseFilename}-export-${timestamp}.xml`, exportToXML(presentations));
  }
  if (formats.includes('csv')) {
    zip.file(`${baseFilename}-export-${timestamp}.csv`, exportToCSV(presentations));
  }
  if (formats.includes('txt')) {
    zip.file(`${baseFilename}-export-${timestamp}.txt`, exportToText(presentations));
  }
  if (formats.includes('html')) {
    zip.file(`${baseFilename}-export-${timestamp}.html`, exportToHTML(presentations));
  }
  if (formats.includes('pdf')) {
    const pdf = exportToPDF(presentations);
    zip.file(`${baseFilename}-export-${timestamp}.pdf`, pdf.output('blob'));
  }
  
  // Add media folder
  const hasMedia = presentations.some(p => p.media.length > 0);
  if (hasMedia) {
    const mediaFolder = zip.folder('media');
    if (mediaFolder) {
      for (const pres of presentations) {
        if (pres.media.length === 0) continue;
        const presFolder = mediaFolder.folder(pres.fileName.replace(/\.(pptx?|ppt)$/i, ''));
        if (presFolder) {
          for (const media of pres.media) {
            if (media.data) {
              presFolder.file(media.name, media.data, { base64: true });
            }
          }
        }
      }
    }
  }
  
  const content = await zip.generateAsync({ type: 'blob' });
  saveAs(content, `${baseFilename}-export-${timestamp}.zip`);
}

// Helper functions
function escapeXml(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function escapeCSV(str: string): string {
  return str.replace(/"/g, '""');
}

function escapeHtml(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function formatFileSize(bytes: number): string {
  if (bytes === 0) return '0 B';
  const k = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return `${parseFloat((bytes / Math.pow(k, i)).toFixed(1))} ${sizes[i]}`;
}
