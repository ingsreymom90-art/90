import { saveAs } from 'file-saver';
import { jsPDF } from 'jspdf';
import { toPng } from 'html-to-image';

/**
 * Builds a Word-compatible header that closely matches the web design.
 * Uses VML-safe table structures with mso-* properties for background colors.
 */
function buildWordSafeHeader(
  schoolName: string,
  schoolAddress: string,
  studentLabel: string,
  classLabel: string,
  dateLabel: string,
  scoreLabel: string,
  teacherLabel: string,
  logoData: string | undefined,
  paperDesign: number,
  activeRulerColor: string,
  fontFamily: string,
  topic: string,
  globalLayout: number,
  baseLayout: number,
  withColor: boolean,
  isTopBottomLineEnabled: boolean = false,
  topBottomLineColor: string = '#0ea5e9'
): string {

  const logoHtml = logoData
    ? `<img src="${logoData}" style="width:60pt;height:auto;display:block;" />`
    : '';

  const accentColors: Record<number, string> = {
    5: '#1e293b', // Professional Navy
    8: '#166534', // Eco Green
    13: '#7f1d1d', // Bold Red (Maroon)
    14: '#92400e', // Royal Gold
    18: '#000000', // Academic Heavy
    19: '#4338ca', // Art Deco
    20: '#0ea5e9', // Futuristic
  };

  const accentColor = withColor ? (accentColors[paperDesign] || activeRulerColor || '#ea580c') : '#000000';
  
  // Derivative colors for frames (from globalLayout)
  let topBarColor = '';
  let leftBarColor = '';
  
  if (withColor) {
    if (globalLayout === 1) { // Orange Mix
      topBarColor = '#ea580c';
      leftBarColor = '#059669';
    } else if (globalLayout === 2) { // Emerald
      topBarColor = '#059669';
      leftBarColor = '#059669';
    } else if (globalLayout === 3) { // Lavender
      topBarColor = '#9333ea';
      leftBarColor = '#9333ea';
    } else if (globalLayout === 6) { // Sky
      topBarColor = '#0284c7';
      leftBarColor = '#0ea5e9';
    } else if (globalLayout === 15) { // Deep Ocean
      topBarColor = '#1e3a8a';
      leftBarColor = '#3b82f6';
    }
    
    // Override leftBarColor if paperDesign specifically requires it (e.g., Style 8 Eco Green)
    if (paperDesign === 8) leftBarColor = '#059669';
    if (paperDesign === 13) leftBarColor = '#b91c1c';
    if (paperDesign === 20) leftBarColor = '#0ea5e9';
  }

  // Custom Red Style (Style 13) specifically for the red horizontal bar
  const headerTopBarColor = paperDesign === 13 && withColor ? '#b91c1c' : topBarColor;
  
  const topBarHtml = headerTopBarColor ? `
    <table border="0" cellspacing="0" cellpadding="0" width="100%" style="width:100%; border-collapse:collapse; margin-bottom:0pt;">
      <tr><td style="background-color:${headerTopBarColor}; mso-shading:${headerTopBarColor}; height:12pt; font-size:1pt;">&nbsp;</td></tr>
    </table>` : '';

  // Header Style Mappings for left-bar professional headers
  if ([3, 7].includes(paperDesign)) {
    return `
    ${topBarHtml}
    <table border="0" cellspacing="0" cellpadding="0" width="100%"
      style="width:100%; border-collapse:collapse; margin-bottom:8pt; margin-top:12pt;">
      <tr>
        <td style="width:15%; vertical-align:middle; text-align:left;">
          ${logoHtml}
        </td>
        <td style="width:50%; vertical-align:middle; padding-left:10pt; border-left:${leftBarColor ? `15pt solid ${leftBarColor}` : '15pt solid #000000'};">
          <div style="font-size:22pt; font-weight:900; color:${accentColor}; 
                       font-family:'${fontFamily}'; text-transform:uppercase;
                       line-height:1.1; margin-bottom:4pt;">
            ${schoolName}
          </div>
          <div style="font-size:9pt; color:${accentColor}; font-weight:700;
                      text-transform:uppercase; letter-spacing:2pt;">
            ${topic ? topic.toUpperCase() : 'ACADEMIC EVALUATION'}
          </div>
        </td>
        <td style="width:35%; vertical-align:top; padding:8pt 0 4pt 10pt; text-align:right;">
          <table border="0" cellspacing="0" cellpadding="0" style="margin-left:auto;">
            <tr>
              <td style="font-size:9pt; font-style:italic; padding-bottom:4pt; text-align:right;">
                ${studentLabel}: ________________________
              </td>
            </tr>
            <tr>
              <td style="font-size:9pt; font-style:italic; padding-bottom:4pt; text-align:right;">
                ${classLabel}: ________________________
              </td>
            </tr>
            <tr>
              <td style="font-size:9pt; font-style:italic; padding-bottom:4pt; text-align:right;">
                ${dateLabel}: ________________________
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
    
    <!-- DIVIDER LINE -->
    <table border="0" cellspacing="0" cellpadding="0" width="100%"
      style="width:100%; border-collapse:collapse; margin-bottom:12pt; margin-top:8pt;">
      <tr>
        <td style="border-bottom:1.5pt solid ${accentColor}; 
                   mso-border-bottom-alt:1.5pt solid ${accentColor}; 
                   height:1pt; font-size:1pt;">&nbsp;</td>
      </tr>
    </table>`;
  }

  // ── Style 6 (paperDesign === 5): Green Nature (Boxed) ──
  if (paperDesign === 5) {
    return `
    <table border="0" cellspacing="0" cellpadding="10" width="100%"
      style="width:100%; border-collapse:collapse; border:3pt solid #16a34a; background-color:#f0fdf4; margin-bottom:12pt;">
      <tr>
        <td style="padding:10pt;">
          <table border="0" cellspacing="0" cellpadding="0" width="100%" style="border-bottom:2pt solid #16a34a; padding-bottom:6pt; margin-bottom:6pt;">
            <tr>
              <td style="font-size:18pt; font-weight:900; color:#065f46; text-transform:uppercase;">
                ${schoolName}
              </td>
              <td style="font-size:8pt; font-weight:700; color:#059669; text-align:right;">
                ${dateLabel}: ______/______/______<br/><br/>
                ${classLabel}: ______________________
              </td>
            </tr>
          </table>
          <div style="font-size:9pt; font-weight:700; color:#064e3b; margin-top:5pt;">
            ${studentLabel}: __________________________________________
          </div>
        </td>
      </tr>
    </table>`;
  }

  // ── Style 9 variants (paperDesign 8, 18, 19, 20, 21): Top-bar border ──
  if ([8, 18, 19, 20, 21].includes(paperDesign)) {
    const styleColors: Record<number, { text: string, sub: string, border: string }> = {
      8: { text: '#881337', sub: '#f43f5e', border: '#e11d48' },   // Modern Red (rose)
      18: { text: '#064e3b', sub: '#10b981', border: '#059669' },  // Modern Green (emerald)
      19: { text: '#1e3a8a', sub: '#3b82f6', border: '#2563eb' },  // Modern Blue (blue)
      20: { text: '#581c87', sub: '#a855f7', border: '#9333ea' },  // Modern Purple (purple)
      21: { text: '#7c2d12', sub: '#f97316', border: '#ea580c' }   // Modern Orange (orange)
    };
    const c = styleColors[paperDesign];
    return `
    <table border="0" cellspacing="0" cellpadding="0" width="100%"
      style="width:100%; border-collapse:collapse; margin-bottom:16pt; margin-top:8pt;">
      <tr>
        <td style="border-top: 4pt solid ${c.border}; padding-top: 12pt; width:60%; vertical-align:top; text-align:left;">
          ${logoHtml ? `<div style="margin-bottom:8pt;">${logoHtml}</div>` : ''}
          <div style="font-size:24pt; font-weight:900; color:${c.text}; 
                       font-family:'${fontFamily}'; text-transform:uppercase; line-height:1;">
            ${schoolName}
          </div>
          <div style="font-size:9pt; color:${c.sub}; font-weight:700;
                      text-transform:uppercase; letter-spacing:2pt; margin-top:6pt;">
            ${topic ? topic.toUpperCase() : 'ACADEMIC EVALUATION'}
          </div>
        </td>
        <td style="border-top: 4pt solid ${c.border}; padding-top: 12pt; width:40%; vertical-align:top; text-align:right;">
          <table border="0" cellspacing="0" cellpadding="0" style="margin-left:auto;">
            <tr>
              <td style="font-size:9pt; font-style:italic; padding-bottom:6pt; color:#64748b; text-align:right;">
                ${studentLabel}: __________________
              </td>
            </tr>
            <tr>
              <td style="font-size:9pt; font-style:italic; padding-bottom:6pt; color:#64748b; text-align:right;">
                ${classLabel}: __________________
              </td>
            </tr>
            <tr>
              <td style="font-size:9pt; font-style:italic; padding-bottom:6pt; color:#64748b; text-align:right;">
                ${dateLabel}: __________________
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </table>`;
  }

  // ── Style 1 (paperDesign === 1): Boxed header ──
  if (paperDesign === 1) {
    return `
    <table border="0" cellspacing="0" cellpadding="10" width="100%"
      style="width:100%; border-collapse:collapse; border:2pt solid #000000; margin-bottom:12pt;">
      <tr>
        <td colspan="2" style="text-align:center; padding:12pt; border-bottom:1pt solid #000000;">
          <div style="font-size:18pt; font-weight:900; text-transform:uppercase;">${schoolName}</div>
        </td>
      </tr>
      <tr>
        <td style="font-size:9pt; font-weight:700; padding:6pt 10pt; width:50%;">
          ${studentLabel}: ________________________
        </td>
        <td style="font-size:9pt; font-weight:700; padding:6pt 10pt; width:50%;">
          ${dateLabel}: ________________________
        </td>
      </tr>
      <tr>
        <td style="font-size:9pt; font-weight:700; padding:6pt 10pt;">
          ${classLabel}: _________________________
        </td>
        <td style="font-size:9pt; font-weight:700; padding:6pt 10pt;">
          ${scoreLabel}: ________ / ________
        </td>
      </tr>
    </table>`;
  }

  // ── Style 4 (paperDesign === 4): Dark header ──
  if (paperDesign === 4) {
    return `
    <table border="0" cellspacing="0" cellpadding="0" width="100%"
      style="width:100%; border-collapse:collapse; margin-bottom:12pt;">
      <tr>
        <td style="background-color:#1e293b; mso-shading:#1e293b; mso-pattern:solid;
                   padding:16pt; border-radius:8pt;">
          <div style="color:#ffffff; font-size:18pt; font-weight:900; 
                      text-transform:uppercase; margin-bottom:10pt;">${schoolName}</div>
          <table border="0" cellspacing="0" cellpadding="0" width="100%">
            <tr>
              <td style="color:#ffffff; font-size:9pt; font-weight:700; 
                         border-bottom:1pt solid rgba(255,255,255,0.3); padding-bottom:2pt; width:33%;">
                ${studentLabel}: ____________
              </td>
              <td style="color:#ffffff; font-size:9pt; font-weight:700;
                         border-bottom:1pt solid rgba(255,255,255,0.3); padding-bottom:2pt; width:33%;">
                ${classLabel}: ____________
              </td>
              <td style="color:#ffffff; font-size:9pt; font-weight:700;
                         border-bottom:1pt solid rgba(255,255,255,0.3); padding-bottom:2pt; width:34%;">
                ${scoreLabel}: ______
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </table>`;
  }

  // ── DEFAULT: Classic clean header (paperDesign 0, 2, 3, 5, 6, 7, 9, 10, 11...) ──
  return `
  <table border="0" cellspacing="0" cellpadding="0" width="100%"
    style="width:100%; border-collapse:collapse; margin-bottom:8pt;">
    <tr>
      <td style="border-bottom:2pt solid #000000; padding-bottom:10pt;">
        <table border="0" cellspacing="0" cellpadding="0" width="100%">
          <tr>
            <td style="vertical-align:middle; width:70%;">
              ${logoHtml}
              <div style="font-size:18pt; font-weight:900; text-transform:uppercase;
                          line-height:1.2;">${schoolName}</div>
              ${schoolAddress ? `<div style="font-size:9pt; color:#666666;">${schoolAddress}</div>` : ''}
            </td>
            <td style="vertical-align:top; width:30%; text-align:right;">
              <div style="font-size:9pt; font-weight:700; margin-bottom:4pt;">
                ${studentLabel}: _______________
              </div>
              <div style="font-size:9pt; font-weight:700; margin-bottom:4pt;">
                ${classLabel}: _______________
              </div>
              <div style="font-size:9pt; font-weight:700;">
                ${dateLabel}: _______________
              </div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
  </table>`;
}

export interface ExportMetadata {
  author?: string;
  date?: string;
  title?: string;
}

export const exportToWord = (
  htmlContent: string, 
  filename: string, 
  headerHtml: string = '', 
  marginValue: string = '0.4in 0.6in 0.4in 0.6in',
  fontFamily: string = 'Times New Roman',
  lineHeight: string = '1.15',
  metadata?: ExportMetadata,
  isFrameEnabled: boolean = false,
  activeDesign: string = '',
  paperStyles?: any,
  mcqStyle: number = 0,
  globalLayout: number = 0,
  baseLayout: number = 0,
  instructionRulerStyle: number = 0,
  instructionHeaderStyle: number = 0,
  instructionStyle: number = 0,
  isInstructionBackgroundEnabled: boolean = false,
  isColorExportEnabled: boolean = false,
  exportTheme: number = 1,
  isTopBottomLineEnabled: boolean = false,
  topBottomLineColor: string = '#0ea5e9',
  // ── NEW PARAMS ──
  brandSettings?: {
    schoolName?: string;
    schoolAddress?: string;
    studentLabel?: string;
    classLabel?: string;
    dateLabel?: string;
    scoreLabel?: string;
    teacherLabel?: string;
    logoData?: string;
  },
  paperDesignIndex?: number,
  topicText?: string
) => {
  const tempDiv = document.createElement('div');
  tempDiv.innerHTML = htmlContent;

  // UNWRAP WRAPPER DIVS IMMEDIATELY (Crucial to avoid removing the entire content during cleaning)
  while (tempDiv.children.length === 1 && (tempDiv.firstElementChild?.tagName === 'DIV' || tempDiv.firstElementChild?.classList.contains('prose'))) {
    const wrapper = tempDiv.firstElementChild as HTMLElement;
    tempDiv.innerHTML = wrapper.innerHTML;
  }

  const headerDiv = document.createElement('div');
  headerDiv.innerHTML = headerHtml;

  // Randomize Ruler Color if it's the middle ruler layout
  let activeRulerColor = '#334155'; // Use a neutral slate instead of purple
  if (globalLayout === 1 || activeDesign === 'design-playful') activeRulerColor = '#059669'; // Green for Orange Mix left bar
  else if (globalLayout === 2 || activeDesign === 'design-eco') activeRulerColor = '#059669';
  else if (activeDesign === 'design-modern-blue') activeRulerColor = '#2563eb';
  else if (globalLayout === 3) activeRulerColor = '#9333ea'; // Purple for Soft Lavender

  // ── WORD-SAFE HEADER OVERRIDE ──
  // The LLM-generated header uses CSS that Word ignores.
  // We replace it with a table-based Word-compatible version.
  const wordSafeHeader = buildWordSafeHeader(
    brandSettings?.schoolName || 'GLOBAL EDUCATION ACADEMY',
    brandSettings?.schoolAddress || '',
    brandSettings?.studentLabel || 'STUDENT NAME',
    brandSettings?.classLabel || 'CLASS',
    brandSettings?.dateLabel || 'DATE',
    brandSettings?.scoreLabel || 'SCORE',
    brandSettings?.teacherLabel || 'TEACHER',
    brandSettings?.logoData,
    paperDesignIndex || 0,
    activeRulerColor,
    fontFamily,
    topicText || '',
    globalLayout,
    baseLayout,
    isColorExportEnabled,
    isTopBottomLineEnabled,
    topBottomLineColor
  );

  // If we have brandSettings, aggressively remove redundant headers and student info from content
  if (brandSettings) {
    // 1. Remove explicit header elements
    tempDiv.querySelectorAll('.school-header, .worksheet-header, .academic-evaluation, h1').forEach(el => el.remove());
    
    // 2. Remove redundant school name and student board info from the TOP part of content only
    // To avoid removing the entire worksheet, we only target elements in the first 10-15 children.
    const schoolNamePart = (brandSettings.schoolName || '').split(' ')[0].toUpperCase();
    const children = Array.from(tempDiv.children);
    const topLimit = Math.min(children.length, 12);
    
    for (let i = 0; i < topLimit; i++) {
      const el = children[i];
      const text = el.textContent?.toUpperCase() || '';
      
      // Keywords that indicate header/student info
      const headerKeywords = ['NAME', 'CLASS', 'DATE', 'SCORE', 'TEACHER', 'ACADEMIC EVALUATION'];
      if (schoolNamePart.length > 2) headerKeywords.push(schoolNamePart);
      
      const matchCount = headerKeywords.filter(w => text.includes(w)).length;
      const hasContentMarkers = text.includes('PART') || text.includes('EXERCISE') || text.includes('QUESTION') || text.includes('I.') || text.includes('1.');
      
      // If it looks like a header (contains school name or multiple student fields) 
      // and doesn't look like actual content, remove it.
      if ((matchCount >= 2 || (schoolNamePart && text.includes(schoolNamePart) && text.length < 100)) && !hasContentMarkers) {
        el.remove();
      }
    }
  }

  // Dynamic Line Spacing Logic
  const spacingMap: Record<string, string> = {
    '1.0': '15pt',
    '1.15': '18pt',
    '1.5': '24pt',
    '2.0': '32pt'
  };
  const exactLineHeight = spacingMap[lineHeight] || `${Math.round(parseFloat(lineHeight) * 16)}pt`;

  // 1. Image Formatting (Synchronous to prevent Chrome's strict gesture-expiry download interruptions)
  const images = [...Array.from(tempDiv.querySelectorAll('img')), ...Array.from(headerDiv.querySelectorAll('img'))];
  for (const img of images) {
    const originalWidth = img.width || 550;
    const isLogo = img.style.maxHeight === '80pt' || img.classList.contains('logo') || headerDiv.contains(img);

    if (isLogo) {
      img.style.width = '1.25in';
      img.style.height = 'auto';
    } else if (originalWidth > 200) {
      img.style.width = '6.5in';
      img.style.height = 'auto';
    } else {
      img.style.width = `${(originalWidth / 96).toFixed(2)}in`;
      img.style.height = 'auto';
    }
    img.style.display = 'block';
    if (!isLogo) img.style.margin = '5px auto';
  }

  // 1.5. STRICT BACKGROUND STRIPPING (Instruction Mode)
  if (!isInstructionBackgroundEnabled) {
    const allHeaders = tempDiv.querySelectorAll('.header-row, .part-header, .instruction-header, h2, h3');
    allHeaders.forEach(el => {
      const header = el as HTMLElement;
      header.style.backgroundColor = 'transparent';
      header.style.color = '#000000';
      header.style.setProperty('mso-shading', 'transparent');
      header.style.border = 'none';
      header.style.setProperty('mso-border-alt', 'none');
    });
  }

  // Map activeDesign ID to design classes for logic
  const designClassMap: Record<string, string> = {
    '1': 'design-modern-blue',
    '2': 'design-classic',
    '3': 'design-minimalist',
    '8': 'design-eco'
  };
  const designClass = designClassMap[activeDesign] || '';

  // Comprehensive MCQ Styles
  const mcqElements = tempDiv.querySelectorAll('b, strong, span');
  mcqElements.forEach(el => {
    // PROTECT ANSWER KEY: Do not circle letters in the answer key section
    const isAnswerKey = el.closest('.answer-key-section') || el.closest('.answer-key');
    if (isAnswerKey) return;

    if (mcqStyle > 0) {
      let text = el.textContent?.trim().toUpperCase() || '';
      const match = text.match(/[A-D]/);
      if (match) text = match[0]; else return;

      if (['A', 'B', 'C', 'D'].includes(text)) {
        let borderColor = '#000000';
        let textColor = '#000000';
        let bgColor = '#ffffff';
        let isFilled = 'f';
        
        if (isColorExportEnabled) {
          if (activeDesign === 'design-eco' && (mcqStyle === 1 || mcqStyle === 15)) {
            borderColor = '#059669'; bgColor = '#ecfdf5'; textColor = '#065f46'; isFilled = 't';
          } else if ((activeDesign === 'design-modern-blue' || activeDesign === 'design-bold-red') && (mcqStyle === 1 || mcqStyle === 15)) {
            borderColor = '#2563eb'; bgColor = '#eff6ff'; textColor = '#171717'; isFilled = 't';
          }
        }

        if (mcqStyle === 1 || mcqStyle === 15) {
          const vmlFillAttr = isFilled === 't' ? `filled="t" fillcolor="${bgColor}"` : 'filled="f"';
          const htmlBgStyle = isFilled === 't' ? `background:${bgColor};` : 'background:transparent;';
          
          el.innerHTML = `<!--[if gte vml 1]><v:oval style="width:13pt;height:19pt;position:relative;top:2pt;" ${vmlFillAttr} strokecolor="${borderColor}" strokeweight="0.75pt" o:allowincell="t"><v:textbox inset="0,0,0,0" style="mso-fit-shape-to-text:f;mso-direction-alt:auto;v-text-anchor:top;"><div style="text-align:center;font-size:7pt;line-height:7pt;color:${textColor};font-weight:bold;font-family:'${fontFamily}';margin:0;padding:0;">${text}</div></v:textbox></v:oval><![endif]--><!--[if !mso]>--><span style="border:0.75pt solid ${borderColor}; width:13pt; height:19pt; border-radius:50%; ${htmlBgStyle} color:${textColor}; font-weight:bold; font-size:7pt; box-sizing:border-box; display:inline-flex; align-items:flex-start; justify-content:center; text-align:center; vertical-align:-2pt;">${text}</span><!--<![endif]-->&nbsp;&nbsp;&nbsp;&nbsp;`;
          (el as HTMLElement).style.border = 'none';
          (el as HTMLElement).style.setProperty('mso-border-alt', 'none');
          (el as HTMLElement).style.backgroundColor = 'transparent';
          (el as HTMLElement).style.setProperty('mso-shading', 'transparent');
          (el as HTMLElement).style.marginRight = '0pt';
          (el as HTMLElement).style.verticalAlign = 'baseline';
          (el as HTMLElement).classList.add('mcq-letter-vml');
        }
        else if (mcqStyle === 2) {
          el.innerHTML = `<!--[if gte vml 1]><v:rect style="width:13pt;height:19pt;position:relative;top:2pt;" filled="f" strokecolor="#000000" strokeweight="0.75pt" o:allowincell="t"><v:textbox inset="0,0,0,0" style="mso-fit-shape-to-text:f;v-text-anchor:top;"><div style="text-align:center;font-size:7pt;line-height:7pt;color:black;font-weight:bold;font-family:'${fontFamily}';margin:0;padding:0;">${text}</div></v:textbox></v:rect><![endif]--><!--[if !mso]>--><span style="border:0.75pt solid black; width:13pt; height:19pt; color:black; font-weight:bold; font-size:7pt; box-sizing:border-box; display:inline-flex; align-items:flex-start; justify-content:center; text-align:center; vertical-align:-2pt;">${text}</span><!--<![endif]-->&nbsp;&nbsp;&nbsp;&nbsp;`;
          (el as HTMLElement).style.border = 'none';
          (el as HTMLElement).style.setProperty('mso-border-alt', 'none');
          (el as HTMLElement).style.backgroundColor = 'transparent';
          (el as HTMLElement).style.setProperty('mso-shading', 'transparent');
          (el as HTMLElement).style.marginRight = '0pt';
          (el as HTMLElement).style.verticalAlign = 'baseline';
          (el as HTMLElement).classList.add('mcq-letter-vml');
        }
        else if (mcqStyle === 16) { // NEW: Pill / Tall Oval Style
          el.innerHTML = `<!--[if gte vml 1]><v:oval style="width:13pt;height:21pt;position:relative;top:0.5pt;" filled="f" strokecolor="#000000" strokeweight="1pt" o:allowincell="t"><v:textbox inset="0,0,0,0" style="mso-fit-shape-to-text:f;v-text-anchor:top;"><div style="text-align:center;font-size:7pt;line-height:7pt;color:black;font-weight:bold;font-family:'${fontFamily}';margin:0;padding:0;">${text}</div></v:textbox></v:oval><![endif]--><!--[if !mso]>--><span style="border:1pt solid black; width:13pt; height:21pt; border-radius:50% / 30%; color:black; font-weight:bold; font-size:7pt; box-sizing:border-box; display:inline-flex; align-items:flex-start; justify-content:center; text-align:center; vertical-align:-0.5pt;">${text}</span><!--<![endif]-->&nbsp;&nbsp;&nbsp;&nbsp;`;
          (el as HTMLElement).style.border = 'none';
          (el as HTMLElement).style.marginRight = '0pt';
          (el as HTMLElement).style.verticalAlign = 'baseline';
          (el as HTMLElement).classList.add('mcq-letter-vml');
        }
        else if (mcqStyle === 6) el.innerHTML = `◆${text}`;
        else if (mcqStyle === 8) {
          el.innerHTML = text === 'A' ? 'Ⓐ' : text === 'B' ? 'Ⓑ' : text === 'C' ? 'Ⓒ' : 'Ⓓ';
          (el as HTMLElement).style.border = 'none';
          (el as HTMLElement).style.setProperty('mso-border-alt', 'none');
          (el as HTMLElement).style.backgroundColor = 'transparent';
          (el as HTMLElement).style.setProperty('mso-shading', 'transparent');
        }
        else if (mcqStyle === 11 || mcqStyle === 12) {
          let borderType = mcqStyle === 11 ? 'dashstyle="solid" strokeweight="1pt"' : 'dashstyle="dot" strokeweight="1.2pt"';
          let htmlBorder = mcqStyle === 11 ? '1pt solid black' : '1.2pt dotted black';

          el.innerHTML = `<!--[if gte vml 1]><v:oval style="width:15pt;height:21pt;position:relative;top:2pt;" filled="f" strokecolor="black" ${borderType} o:allowincell="t"><v:textbox inset="0,0,0,0" style="mso-fit-shape-to-text:f;v-text-anchor:top;"><div style="text-align:center;font-size:7pt;line-height:7pt;color:black;font-weight:bold;font-family:'${fontFamily}';margin:0;padding:0;">${text}</div></v:textbox></v:oval><![endif]--><!--[if !mso]>--><span style="border:${htmlBorder}; width:15pt; height:21pt; border-radius:50%; background:transparent; color:black; font-weight:bold; font-size:7pt; box-sizing:border-box; display:inline-flex; align-items:flex-start; justify-content:center; text-align:center; vertical-align:-2pt;">${text}</span><!--<![endif]-->&nbsp;&nbsp;&nbsp;&nbsp;`;
          (el as HTMLElement).style.border = 'none';
          (el as HTMLElement).style.setProperty('mso-border-alt', 'none');
          (el as HTMLElement).style.backgroundColor = 'transparent';
          (el as HTMLElement).style.setProperty('mso-shading', 'transparent');
          (el as HTMLElement).style.marginRight = '0pt';
          (el as HTMLElement).style.verticalAlign = 'baseline';
        }
        else if (mcqStyle === 13 || mcqStyle === 14) {
          el.innerHTML = text === 'A' ? '🅐' : text === 'B' ? '🅑' : text === 'C' ? '🅒' : '🅓';
          (el as HTMLElement).style.border = 'none';
          (el as HTMLElement).style.setProperty('mso-border-alt', 'none');
          (el as HTMLElement).style.backgroundColor = 'transparent';
          (el as HTMLElement).style.setProperty('mso-shading', 'transparent');
          if (mcqStyle === 14) (el as HTMLElement).style.color = '#10b981';
          (el as HTMLElement).innerHTML += '&nbsp;&nbsp;&nbsp;&nbsp;';
          (el as HTMLElement).style.marginRight = '0pt';
          (el as HTMLElement).style.verticalAlign = 'baseline';
        }
      }
    }
  });

  // Instruction Rulers
  const existingRulers = tempDiv.querySelectorAll('[class*="instruction-ruler-"]');
  existingRulers.forEach(ruler => {
    const el = ruler as HTMLElement;
    const styleNum = parseInt(el.className.match(/instruction-ruler-(\d+)/)?.[1] || '0');
    el.style.width = '100%';
    el.style.margin = '5pt 0 10pt 0';
    el.innerHTML = '&nbsp;';
    if (styleNum === 1) el.style.borderBottom = `1.5pt solid ${activeRulerColor}`;
    else if (styleNum === 2) el.style.borderBottom = `2pt dashed ${activeRulerColor}`;
    else if (styleNum === 3) el.style.borderBottom = `4pt double ${activeRulerColor}`;
    else if (styleNum === 4) el.style.borderBottom = `4pt solid ${activeRulerColor}`;
    
    if (styleNum > 0) el.style.setProperty('mso-border-bottom-alt', el.style.borderBottom);
  });

  // 1. TRANSFORM MCQ OPTIONS INTO WORD-SAFE LAYOUT TABLES
  const optionsTables = tempDiv.querySelectorAll('.options-table, [data-type="mcq-options"]');
  optionsTables.forEach(table => {
    const parent = table.parentElement;
    if (!parent) return;

    // We use a clean table because Word treats nested inline-blocks as vertical piles.
    // This table is purely for structural layout (1-column, 2-column, or 4-column).
    const legacyRow = table.querySelector('tr');
    if (!legacyRow) return;
    
    const cells = Array.from(legacyRow.cells);
    const columnCount = cells.length || 1;
    const cellWidth = Math.floor(100 / columnCount);

    const layoutTable = document.createElement('table');
    layoutTable.setAttribute('border', '0');
    layoutTable.setAttribute('cellspacing', '0');
    layoutTable.setAttribute('cellpadding', '0');
    layoutTable.style.width = '100%';
    layoutTable.style.borderCollapse = 'collapse';
    layoutTable.style.border = 'none';
    layoutTable.style.setProperty('mso-border-alt', 'none');
    layoutTable.style.setProperty('mso-table-lspace', '0pt');
    layoutTable.style.setProperty('mso-table-rspace', '0pt');

    const tr = document.createElement('tr');
    cells.forEach(cell => {
      const td = document.createElement('td');
      td.style.width = `${cellWidth}%`;
      td.style.verticalAlign = 'top';
      td.style.padding = '4pt 2pt';
      td.style.border = 'none';

      // FIX ALIGNMENT: If the cell contains an MCQ letter, split it into a 2-col layout table
      const mcqLetter = cell.querySelector('.mcq-letter-vml');
      if (mcqLetter) {
        const letterHtml = mcqLetter.outerHTML;
        // Strip the letter from the original text
        const remainder = cell.innerHTML.replace(letterHtml, '').trim();
        
        td.innerHTML = `
          <table border="0" cellspacing="0" cellpadding="0" style="width:100%; border-collapse:collapse; border:none;">
            <tr>
              <td style="width:20pt; vertical-align:top; border:none; padding:0;">${letterHtml}</td>
              <td style="vertical-align:top; border:none; padding:0 0 0 4pt;">${remainder}</td>
            </tr>
          </table>
        `;
      } else {
        td.innerHTML = cell.innerHTML;
      }
      
      tr.appendChild(td);
    });
    layoutTable.appendChild(tr);

    table.replaceWith(layoutTable);
  });

  // Answer Key Section Cleanup
  const answerKeySections = tempDiv.querySelectorAll('.answer-key-section, .answer-key');
  answerKeySections.forEach(section => {
    const el = section as HTMLElement;
    el.style.marginTop = '20pt';
    el.style.padding = '15pt';
    el.style.border = `1.5pt solid ${activeRulerColor}`;
    el.style.borderRadius = '5pt';
    el.style.backgroundColor = '#f8fafc';
    el.style.setProperty('mso-shading', '#f8fafc');
    
    const title = el.querySelector('h2');
    if (title) {
        title.style.marginTop = '0';
        title.style.color = '#1e293b';
        title.style.fontSize = '14pt';
        title.style.borderBottom = `1pt solid ${activeRulerColor}`;
        title.style.paddingBottom = '5pt';
        title.style.marginBottom = '10pt';
    }

    const textEls = el.querySelectorAll('p, div, span');
    textEls.forEach(t => {
      (t as HTMLElement).style.fontSize = '11pt';
      (t as HTMLElement).style.lineHeight = '1.5';
      (t as HTMLElement).style.color = '#334155';
    });
  });

  const tables = tempDiv.querySelectorAll('table');
  tables.forEach(table => {
    const isNested = table.parentElement?.closest('table') !== null;
    if (isNested) {
      table.style.border = 'none';
      table.style.width = '100%';
      table.querySelectorAll('td').forEach(c => {
        (c as HTMLElement).style.border = 'none';
        (c as HTMLElement).style.padding = '2pt';
      });
    } else {
      const isRulerTable = table.classList.contains('ruler-table') || table.rows[0]?.cells.length === 2;
      if (isRulerTable) {
        table.style.border = 'none';
        table.style.borderCollapse = 'collapse';
        Array.from(table.rows).forEach(row => {
          Array.from(row.cells).forEach((c, idx) => {
            const cell = c as HTMLElement;
            cell.style.padding = '15pt';
            cell.style.border = 'none';
            if (idx === 0 && row.cells.length === 2) {
              cell.style.borderRight = `1.5pt solid ${activeRulerColor}`;
              cell.style.setProperty('mso-border-right-alt', `1.5pt solid ${activeRulerColor}`);
            }
          });
        });
      }
      
      // Header Styles - STRICTLY respect isInstructionBackgroundEnabled
      table.querySelectorAll('td').forEach(cell => {
        const isHeader = cell.classList.contains('header-row') || (cell.getAttribute('colspan') === '2');
        if (isHeader) {
          const c = cell as HTMLElement;
          // FORCE white background if disabled to prevent Word inheritance/defaults
          let bg = '#ffffff';
          let textColor = '#000000';
          let shading = '#ffffff';

          if (isInstructionBackgroundEnabled) {
            // Apply specific style mappings based on the chosen style
            const headerStyles: Record<number, { bg: string, color: string, border?: string }> = {
              0: { bg: '#facc15', color: '#000000', border: '3pt solid black' },
              1: { bg: '#f59e0b', color: '#ffffff' },
              3: { bg: '#1e293b', color: '#ffffff', border: '8pt solid #6366f1' },
              4: { bg: '#dcfce7', color: '#065f46', border: '2pt solid #10b981' },
              5: { bg: '#fde047', color: '#000000', border: '4pt solid black' },
              6: { bg: '#4f46e5', color: '#ffffff' },
              13: { bg: '#dcfce7', color: '#064e3b', border: '3pt solid #059669' },
              15: { bg: '#581c87', color: '#ffffff', border: '2pt solid #fbbf24' },
              19: { bg: '#ea580c', color: '#ffffff' }
            };

            const style = headerStyles[instructionHeaderStyle] || { bg: '#dcfce7', color: '#064e3b' };
            bg = style.bg;
            textColor = style.color;
            shading = style.bg;
            if (style.border) {
              c.style.border = style.border;
              c.style.setProperty('mso-border-alt', style.border);
            }
          }

          c.style.backgroundColor = bg;
          c.style.setProperty('mso-shading', shading);
          c.style.color = textColor;
          
          if (!isInstructionBackgroundEnabled) {
            c.style.border = 'none';
            c.style.borderBottom = `1.5pt solid ${activeRulerColor}`;
            c.style.setProperty('mso-border-bottom-alt', `1.5pt solid ${activeRulerColor}`);
          } else if (!c.style.border) {
            c.style.borderLeft = `6pt solid ${activeRulerColor}`;
            c.style.setProperty('mso-border-left-alt', `6pt solid ${activeRulerColor}`);
          }
          c.style.padding = '10pt';
          c.style.paddingLeft = '15pt';
          c.style.fontWeight = 'bold';
        }
      });
      
      // Zebra Striping detection
      const rows = Array.from(table.rows);
      if (table.classList.contains('zebra') || table.getAttribute('data-type') === 'zebra') {
        rows.forEach((row, idx) => {
          if (idx % 2 === 1) { // odd index = even row (1, 3, 5...)
            Array.from(row.cells).forEach(cell => {
              (cell as HTMLElement).style.backgroundColor = '#f8fafc';
              (cell as HTMLElement).style.setProperty('mso-shading', '#f8fafc');
            });
          }
        });
      }
    }
  });

  // Word Bank Box
  const wordBanks = tempDiv.querySelectorAll('.word-bank-box-alt, .word-bank');
  wordBanks.forEach(box => {
    const el = box as HTMLElement;
    el.style.border = '1.5pt solid #334155';
    el.style.padding = '10pt';
    el.style.margin = '10pt 0';
    el.style.backgroundColor = '#f1f5f9';
    el.style.setProperty('mso-shading', '#f1f5f9');
    el.style.textAlign = 'center';
    el.style.fontWeight = 'bold';
    el.style.borderRadius = '5pt';
  });

  let sections = Array.from(tempDiv.children);
  // Important: If everything is wrapped in ONE div (like .prose), unwrap it so we can loop over sections
  if (sections.length === 1 && (sections[0].classList.contains('prose') || sections[0].tagName === 'DIV')) {
    sections = Array.from(sections[0].children);
  }

  let finalHtml = "";
  sections.forEach(el => {
    // Only add non-empty elements
    if (el.textContent?.trim() || el.querySelector('img') || el.querySelector('table')) {
      finalHtml += `
      <table border="0" cellspacing="0" cellpadding="0" width="100%" style="width: 100%; border-collapse: collapse;">
        <tr>
          <td style="font-family: '${fontFamily}', serif; font-size: 12pt; padding-bottom: 8pt; line-height: ${exactLineHeight}; mso-line-height-rule: exactly;">
            ${(el as HTMLElement).outerHTML}
          </td>
        </tr>
      </table>`;
    }
  });

  // 5. Frame Style Simulation
  const frameStyle = '';
  const pageBorderStyle = isFrameEnabled ? `border: 3.0pt solid ${activeRulerColor}; padding: 24pt 24pt 24pt 24pt; mso-page-border-z-order: front; mso-page-border-surround-header:no; mso-page-border-surround-footer:no;` : '';
  const physicalFrameSyle = isFrameEnabled ? `border: 3.0pt solid ${activeRulerColor}; mso-border-alt: 3.5pt solid ${activeRulerColor}; padding: 10pt;` : '';

  // Paper Styles - Moved All Borders to TD level for better Word support
  let bodyBgColor = '#ffffff';
  let paperTdStyle = '';
  
  if (globalLayout === 1 && isTopBottomLineEnabled) { // Orange Mix + Top-bottom line enabled
    paperTdStyle = `border-top: 12pt solid #ea580c; mso-border-top-alt: 12pt solid #ea580c; border-top-left-radius: 40pt; padding-left: 20pt; padding-top: 20pt; background: #ffffff; mso-shading: windowtext 0% #ffffff;`;
  } else if (globalLayout === 1) { // Orange Mix, default (no extra borders if top-bottom line is off)
    paperTdStyle = `padding-left: 10pt; mso-shading: windowtext 0% #ffffff;`;
  } else if (globalLayout === 2) { // Modern Emerald
    paperTdStyle = `background-color: #f0fdf4; padding-left: 15pt; mso-shading: windowtext 0% #f0fdf4;`;
    if (isTopBottomLineEnabled) paperTdStyle += `border-left: 12pt solid #059669; mso-border-left-alt: 12pt solid #059669;`;
    bodyBgColor = '#f0fdf4';
  } else if (globalLayout === 17) {
    paperTdStyle = `background-color: #ffffff; border-left: 4.5pt double #ef4444; padding-left: 35pt; mso-shading: windowtext 0% #ffffff;`;
  } else if (globalLayout === 18) {
    paperTdStyle = `background-color: #fef3c7; border: 1pt solid #fde68a; mso-shading: windowtext 0% #fef3c7;`;
    bodyBgColor = '#fef3c7';
  } else {
    paperTdStyle = `mso-shading: windowtext 0% ${bodyBgColor};`;
  }

  const content = `
    <html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40">
    <head><meta charset='utf-8'>
      <style>
        @page Section1 { size: 8.5in 11.0in; margin: 0.5in; ${pageBorderStyle} }
        div.Section1 { page: Section1; }
        body { font-family: "${fontFamily}", serif; font-size: 12pt; line-height: ${exactLineHeight}; mso-line-height-rule: exactly; background-color: ${bodyBgColor}; }
        table { border-collapse: collapse; width: 100%; }
        td { padding: 0; vertical-align: top; }
        .options-table td { mso-line-height-rule: at-least; line-height: 24pt; height: 26pt; }
      </style>
    </head>
    <body>
      <div class="Section1">
        <!-- Master Table for Paper Design -->
    <table border="0" cellspacing="0" cellpadding="0" width="100%" style="width: 100%; border-collapse: collapse; ${physicalFrameSyle}">
      <tr>
        ${isTopBottomLineEnabled ? `
        <td style="width: 10pt; background-color: ${topBottomLineColor}; mso-shading: ${topBottomLineColor}; font-size: 1pt;">&nbsp;</td>
        ` : ''}
        <td style="padding: 30pt; ${paperTdStyle} ${frameStyle}">
            <div>
              ${brandSettings ? wordSafeHeader : headerDiv.innerHTML}
              ${finalHtml}
            </div>
          </td>
      </tr>
    </table>
      </div>
    </body>
    </html>`;

  const blob = new Blob(['\ufeff', content], { type: 'application/msword;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = `${filename.replace(/[^a-z0-9]/gi, '_')}.doc`;
  document.body.appendChild(link);
  link.click();
  
  // Delay cleanup to avoid interruption
  setTimeout(() => {
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  }, 100);
};

export const exportToHTML = (htmlContent: string, filename: string, headerHtml: string = '') => {
  const fullHtml = `<html><body><div class="header">${headerHtml}</div><div class="content">${htmlContent}</div></body></html>`;
  saveAs(new Blob([fullHtml], { type: 'text/html;charset=utf-8' }), `${filename}.html`);
};

export const exportToPDF = async (elementId: string, filename: string) => {
  const element = document.getElementById(elementId);
  if (!element) return;
  try {
    // High-resolution capture (300 DPI approx)
    const dataUrl = await toPng(element, { 
      quality: 1,
      pixelRatio: 2, // Double pixels for crispness
      skipFonts: false,
      cacheBust: true
    });
    
    const pdf = new jsPDF('p', 'mm', 'a4');
    const pdfWidth = pdf.internal.pageSize.getWidth();
    const pdfHeight = pdf.internal.pageSize.getHeight();
    
    // Calculate dimensions to fit the page
    const imgWidth = pdfWidth;
    const imgHeight = (element.offsetHeight * pdfWidth) / element.offsetWidth;
    
    // Handle multi-page if content is too long
    let heightLeft = imgHeight;
    let position = 0;
    
    pdf.addImage(dataUrl, 'PNG', 0, position, imgWidth, imgHeight, undefined, 'FAST');
    heightLeft -= pdfHeight;
    
    while (heightLeft >= 0) {
      position = heightLeft - imgHeight;
      pdf.addPage();
      pdf.addImage(dataUrl, 'PNG', 0, position, imgWidth, imgHeight, undefined, 'FAST');
      heightLeft -= pdfHeight;
    }
    
    pdf.save(`${filename}.pdf`);
  } catch (error) {
    console.error("PDF Export failed", error);
    window.print();
  }
};
