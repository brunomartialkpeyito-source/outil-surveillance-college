/* Minimal offline PDF generator - no external deps, bundled */
export interface PdfReportOpts {
  collegeName: string;
  logoDataUrl?: string;
  anneeActive: string;
  dateJour: string;
  periode?: string;
  pourquoi?: string;
  reportTitle: string;
  columns: string[];
  rows: string[][];
  fileName: string;
}

function pdfEscape(s: string): string {
  let out = "";
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    const code = s.charCodeAt(i);
    if (ch === "\\") out += "\\\\";
    else if (ch === "(") out += "\\(";
    else if (ch === ")") out += "\\)";
    else if (ch === "\r") out += "\\r";
    else if (ch === "\n") out += "\\n";
    else if (code < 32 || code > 126) {
      // WinAnsi octal
      if (code <= 255) {
        out += "\\" + code.toString(8).padStart(3, "0");
      } else {
        // replace accent with ascii approx for safety
        const map: Record<number, string> = {233:"e",232:"e",234:"e",224:"a",226:"a",231:"c",238:"i",244:"o",249:"u",251:"u",201:"E",200:"E",231:"c",244:"o"};
        const repl = map[code] || "?";
        out += repl;
      }
    } else out += ch;
  }
  return out;
}

function getTextWidth(text: string, fontSize: number): number {
  // approx Helvetica width factor
  return text.length * fontSize * 0.52;
}

function dataURLToUint8Array(dataURL: string): Uint8Array {
  const base64 = dataURL.split(",")[1] || "";
  const binary = atob(base64);
  const len = binary.length;
  const bytes = new Uint8Array(len);
  for (let i = 0; i < len; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

async function loadLogo(dataURL?: string): Promise<{width:number;height:number;data:Uint8Array}|null> {
  if (!dataURL || !dataURL.startsWith("data:")) return null;
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => {
      const w = img.width || 100;
      const h = img.height || 100;
      try {
        // Convert to JPEG small if png, to keep DCTDecode simple
        const canvas = document.createElement("canvas");
        // keep original size but limited to 100px for pdf display
        canvas.width = w;
        canvas.height = h;
        const ctx = canvas.getContext("2d");
        if (!ctx) {
          try { resolve({width:w,height:h,data:dataURLToUint8Array(dataURL)}); } catch { resolve(null); }
          return;
        }
        ctx.fillStyle = "#FFFFFF";
        ctx.fillRect(0,0,w,h);
        ctx.drawImage(img,0,0);
        const jpeg = canvas.toDataURL("image/jpeg", 0.6);
        const uint8 = dataURLToUint8Array(jpeg);
        resolve({width:w,height:h,data:uint8});
      } catch {
        try { resolve({width:w,height:h,data:dataURLToUint8Array(dataURL)}); } catch { resolve(null); }
      }
    };
    img.onerror = () => resolve(null);
    img.src = dataURL;
  });
}

const PAGE_W = 595.28;
const PAGE_H = 841.89;
const MARGIN = 40;
const FOOTER_RESERVE = 50;

function buildPages(
  opts: PdfReportOpts,
  logoInfo: {width:number;height:number;data:Uint8Array}|null
): string[] {
  const pages: string[][] = [];
  let current: string[] = [];
  let curY = PAGE_H - MARGIN;

  const addText = (x:number,y:number,text:string,font="F1",size=10) => {
    current.push(`BT /${font} ${size} Tf 1 0 0 1 ${x.toFixed(2)} ${y.toFixed(2)} Tm (${pdfEscape(text)}) Tj ET`);
  };
  const addLine = (x1:number,y1:number,x2:number,y2:number) => {
    current.push(`0.75 0.75 0.75 RG 0.5 w ${x1.toFixed(2)} ${y1.toFixed(2)} m ${x2.toFixed(2)} ${y2.toFixed(2)} l S 0 0 0 RG`);
  };
  const addRectFillStroke = (x:number,y:number,w:number,h:number,fill?:string,stroke?:string,lw?:number) => {
    if (fill) { current.push(`${fill} rg`); current.push(`${x.toFixed(2)} ${y.toFixed(2)} ${w.toFixed(2)} ${h.toFixed(2)} re f`); }
    if (stroke) {
      current.push(`${stroke} RG`);
      if (lw) current.push(`${lw} w`);
      current.push(`${x.toFixed(2)} ${y.toFixed(2)} ${w.toFixed(2)} ${h.toFixed(2)} re S`);
    }
    current.push(`0 0 0 rg 0 0 0 RG`);
  };

  const header = () => {
    curY = PAGE_H - MARGIN;
    if (logoInfo) {
      // draw 36x36 logo at top left
      const logoW = 36;
      const logoH = 36;
      current.push(`q ${logoW} 0 0 ${logoH} ${MARGIN.toFixed(2)} ${(curY - logoH).toFixed(2)} cm /Im0 Do Q`);
    }
    const college = opts.collegeName || "College Prive Saint Laurent";
    addText((PAGE_W - getTextWidth(college,14))/2, curY - 12, college, "F2", 14);
    curY -= 18;
    const subLines: string[] = [];
    subLines.push(`Annee scolaire: ${opts.anneeActive || "2026-2027"}`);
    subLines.push(`Date du jour: ${opts.dateJour}`);
    if (opts.periode) subLines.push(`Periode: ${opts.periode}`);
    if (opts.pourquoi) subLines.push(`Pourquoi: ${opts.pourquoi}`);
    for (const l of subLines) {
      addText((PAGE_W - getTextWidth(l,8))/2, curY, l, "F1", 8);
      curY -= 11;
    }
    curY -= 4;
    const title = opts.reportTitle || "Rapport";
    addText((PAGE_W - getTextWidth(title,12))/2, curY, title, "F2", 12);
    curY -= 16;
    addLine(MARGIN, curY, PAGE_W - MARGIN, curY);
    curY -= 12;
  };

  header();

  const availableW = PAGE_W - 2*MARGIN;
  const colCount = Math.max(1, opts.columns.length);
  const colW = availableW / colCount;
  const rowH = 16;

  // table header
  const drawTableHeader = () => {
    let x = MARGIN;
    for (let i=0;i<colCount;i++) {
      addRectFillStroke(x, curY - rowH, colW, rowH, "0.90 0.92 0.96", "0.80 0.82 0.88", 0.4);
      let txt = opts.columns[i] || "";
      const maxChars = Math.floor(colW / (8*0.52)) -1;
      if (txt.length > maxChars) txt = txt.slice(0, maxChars-1) + ".";
      addText(x+2, curY - rowH + 5, txt, "F2", 7.5);
      x += colW;
    }
    curY -= rowH;
  };

  drawTableHeader();

  for (let r=0;r<opts.rows.length;r++) {
    if (curY - rowH < MARGIN + FOOTER_RESERVE) {
      // close page
      pages.push(current);
      current = [];
      curY = PAGE_H - MARGIN;
      header();
      drawTableHeader();
    }
    let x = MARGIN;
    const row = opts.rows[r];
    for (let c=0;c<colCount;c++) {
      addRectFillStroke(x, curY - rowH, colW, rowH, undefined, "0.85 0.86 0.88", 0.25);
      let cell = (row[c] || "").toString();
      const maxChars = Math.floor(colW / (7*0.52)) -1;
      if (cell.length > maxChars) cell = cell.slice(0, maxChars-1) + ".";
      addText(x+2, curY - rowH + 4.5, cell, "F1", 7);
      x += colW;
    }
    curY -= rowH;
  }

  pages.push(current);
  // now add footers to each page string
  const finalPages: string[] = [];
  const genDate = new Date().toLocaleDateString("fr-FR");
  for (let i=0;i<pages.length;i++) {
    const pContent = pages[i];
    const footerArr: string[] = [...pContent];
    const footY = 28;
    const leftText = `Genere le ${genDate} par DualSync - Surveillant General`;
    footerArr.push(`BT /F1 7 Tf 1 0 0 1 ${MARGIN.toFixed(2)} ${footY.toFixed(2)} Tm (${pdfEscape(leftText)}) Tj ET`);
    const pageText = `Page ${i+1} / ${pages.length}`;
    const pw = getTextWidth(pageText,7);
    footerArr.push(`BT /F1 7 Tf 1 0 0 1 ${(PAGE_W - MARGIN - pw).toFixed(2)} ${footY.toFixed(2)} Tm (${pdfEscape(pageText)}) Tj ET`);
    // top border for footer
    footerArr.push(`0.80 0.82 0.88 RG 0.3 w ${MARGIN.toFixed(2)} ${(footY+10).toFixed(2)} m ${(PAGE_W-MARGIN).toFixed(2)} ${(footY+10).toFixed(2)} l S 0 0 0 RG`);
    finalPages.push(footerArr.join("\n"));
  }
  return finalPages;
}

function buildPdfBytes(pagesContents: string[], images: {width:number;height:number;data:Uint8Array}[]): Uint8Array {
  const enc = new TextEncoder();
  const chunks: Uint8Array[] = [];
  let offset = 0;
  const offsets: number[] = [];

  const pushChunk = (u: Uint8Array) => { chunks.push(u); offset += u.length; };
  const pushStr = (s: string) => { const b = enc.encode(s); pushChunk(b); };

  // header
  pushStr("%PDF-1.4\n%\xE2\xE3\xCF\xD3\n");

  const totalPages = pagesContents.length;
  const numImages = images.length;

  // object numbers:
  // 1 catalog, 2 pages, 3 helvetica, 4 helvetica-bold
  // pages: 5,6, 7,8, ... (page obj, content obj)
  // images after
  const fontHelvetica = 3;
  const fontBold = 4;
  const pageStart = 5;
  const imageStart = pageStart + totalPages*2;

  const totalObjs = imageStart + numImages -1;

  // helper to record offset and write obj
  const writeObj = (num: number, content: string | Uint8Array, isBinaryStream?: {header: string, data: Uint8Array, footer: string}) => {
    offsets[num] = offset;
    if (isBinaryStream) {
      // header
      pushStr(`${num} 0 obj\n${isBinaryStream.header}\nstream\n`);
      pushChunk(isBinaryStream.data);
      pushStr(`\nendstream\nendobj\n`);
    } else {
      const c = typeof content === "string" ? content : new TextDecoder().decode(content as any);
      pushStr(`${num} 0 obj\n${c}\nendobj\n`);
    }
  };

  // 1 catalog
  writeObj(1, "<< /Type /Catalog /Pages 2 0 R >>");
  // 2 pages
  const kids = [];
  for (let i=0;i<totalPages;i++) { const pageObj = pageStart + i*2; kids.push(`${pageObj} 0 R`); }
  writeObj(2, `<< /Type /Pages /Kids [${kids.join(" ")}] /Count ${totalPages} >>`);
  // 3,4 fonts
  writeObj(fontHelvetica, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>");
  writeObj(fontBold, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>");

  // page and content objs
  for (let i=0;i<totalPages;i++) {
    const pageObj = pageStart + i*2;
    const contentObj = pageStart + i*2 +1;
    const contentStr = pagesContents[i];
    const contentBytes = enc.encode(contentStr);
    // page obj
    let xobj = "";
    if (numImages>0) {
      // map Im0 -> imageStart
      xobj = ` /XObject << /Im0 ${imageStart} 0 R >>`;
    }
    const pageContent = `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 ${PAGE_W} ${PAGE_H}] /Resources << /Font << /F1 ${fontHelvetica} 0 R /F2 ${fontBold} 0 R >>${xobj} >> /Contents ${contentObj} 0 R >>`;
    writeObj(pageObj, pageContent);
    // content obj
    const header = `<< /Length ${contentBytes.length} >>`;
    offsets[contentObj] = offset;
    pushStr(`${contentObj} 0 obj\n${header}\nstream\n`);
    pushChunk(contentBytes);
    pushStr(`\nendstream\nendobj\n`);
  }

  // images
  for (let j=0;j<numImages;j++) {
    const img = images[j];
    const objNum = imageStart + j;
    const header = `<< /Type /XObject /Subtype /Image /Width ${img.width} /Height ${img.height} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${img.data.length} >>`;
    offsets[objNum] = offset;
    pushStr(`${objNum} 0 obj\n${header}\nstream\n`);
    pushChunk(img.data);
    pushStr(`\nendstream\nendobj\n`);
  }

  const xrefOffset = offset;
  // xref
  let xref = `xref\n0 ${totalObjs+1}\n0000000000 65535 f \n`;
  for (let i=1;i<=totalObjs;i++) {
    const off = offsets[i] || 0;
    xref += `${String(off).padStart(10,"0")} 00000 n \n`;
  }
  pushStr(xref);
  pushStr(`trailer\n<< /Size ${totalObjs+1} /Root 1 0 R >>\nstartxref\n${xrefOffset}\n%%EOF`);
  // concat
  const totalLen = chunks.reduce((a,c)=>a+c.length,0);
  const out = new Uint8Array(totalLen);
  let pos=0;
  for (const ch of chunks) { out.set(ch,pos); pos+=ch.length; }
  return out;
}

export async function generatePdfReport(opts: PdfReportOpts): Promise<void> {
  const logoInfo = await loadLogo(opts.logoDataUrl);
  const pages = buildPages(opts, logoInfo);
  const images = logoInfo ? [logoInfo] : [];
  const pdfBytes = buildPdfBytes(pages, images);
  const blob = new Blob([pdfBytes], {type:"application/pdf"});
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = opts.fileName.endsWith(".pdf") ? opts.fileName : `${opts.fileName}.pdf`;
  document.body.appendChild(a);
  a.click();
  setTimeout(()=>{ document.body.removeChild(a); URL.revokeObjectURL(url); }, 1000);
}
