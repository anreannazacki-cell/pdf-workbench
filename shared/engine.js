// PDF Workbench — shared engine
// Semua fungsi di sini murni (terima parameter eksplisit, tidak baca DOM langsung)
// supaya bisa dipakai ulang oleh semua halaman tool tanpa duplikasi kode.

import * as pdfjsLib from "https://cdn.jsdelivr.net/npm/pdfjs-dist@6.2.108/build/pdf.min.mjs";
pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdn.jsdelivr.net/npm/pdfjs-dist@6.2.108/build/pdf.worker.min.mjs";

const { PDFDocument } = window.PDFLib;

export const LIMIT = 50 * 1024 * 1024;
export const OCR_PAGE_LIMIT = 30;
export const SPLIT_OUTPUT_LIMIT = 30;
export const DEVICE_MEMORY_GB = navigator.deviceMemory || 4;
export const ORGANIZE_PAGE_LIMIT = DEVICE_MEMORY_GB<=4 ? 8 : (DEVICE_MEMORY_GB<=6 ? 20 : 40);
export const ORGANIZE_SIZE_LIMIT_MB = DEVICE_MEMORY_GB<=4 ? 8 : (DEVICE_MEMORY_GB<=6 ? 20 : 50);

export function fmtBytes(n){
  if(!Number.isFinite(n)) return "-";
  const u=["B","KB","MB","GB"]; let i=0;
  while(n>=1024&&i<u.length-1){n/=1024;i++}
  return `${n.toFixed(i?1:0)} ${u[i]}`;
}
export function safeBase(name){
  return (name||"document.pdf").replace(/\.pdf$/i,"").replace(/[\\/:*?"<>|]+/g,"-").trim()||"document";
}
export function sleep(ms){return new Promise(r=>setTimeout(r,ms))}
export function escapeHtml(s){return String(s).replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]))}
export function escapeAttr(s){return escapeHtml(s).replace(/`/g,"&#96;")}
export function failMessage(err){
  const s=String(err?.message||err||"");
  if(/password|encrypted|encryptedpdf|passwordexception/i.test(s))
    return "PDF ini terkunci password/enkripsi. PDF Workbench belum mendukung PDF terkunci.";
  if(/Invalid PDF|InvalidPDF|Missing PDF|format/i.test(s))
    return "File tidak dapat dibaca sebagai PDF yang valid.";
  return s || "Terjadi kesalahan saat memproses PDF.";
}

export async function loadPdf(file){
  const buf=await file.arrayBuffer();
  try{ return await pdfjsLib.getDocument({data:new Uint8Array(buf)}).promise; }
  catch(e){ throw new Error(failMessage(e)); }
}
export async function loadPdfLib(file){
  try{ return await PDFDocument.load(await file.arrayBuffer(),{ignoreEncryption:false,updateMetadata:false}); }
  catch(e){ throw new Error(failMessage(e)); }
}

let docxLib=null;
export async function getDocx(){
  if(docxLib) return docxLib;
  try{ docxLib=await import("https://cdn.jsdelivr.net/npm/docx@9.7.1/+esm"); return docxLib; }
  catch(e){ throw new Error("Engine Word gagal dimuat. Periksa koneksi internet lalu coba lagi."); }
}

async function renderPage(pdf,pageNo,scale,quality){
  const page=await pdf.getPage(pageNo);
  const viewport=page.getViewport({scale});
  const maxPx=1900;
  const ratio=Math.min(1,maxPx/Math.max(viewport.width,viewport.height));
  const canvas=document.createElement("canvas");
  canvas.width=Math.max(1,Math.round(viewport.width*ratio));
  canvas.height=Math.max(1,Math.round(viewport.height*ratio));
  const ctx=canvas.getContext("2d",{alpha:false});
  ctx.fillStyle="#fff";ctx.fillRect(0,0,canvas.width,canvas.height);
  const renderViewport=page.getViewport({scale:scale*ratio});
  await page.render({canvasContext:ctx,viewport:renderViewport}).promise;
  const blob=await new Promise(r=>canvas.toBlob(r,"image/jpeg",quality));
  page.cleanup();
  return {blob,width:canvas.width,height:canvas.height};
}

export const COMPRESS_MODES={
  light:{quality:.82,scale:1.35,label:"Light",desc:"Kualitas tinggi • ukuran lebih besar"},
  balanced:{quality:.68,scale:1.15,label:"Balanced",desc:"Kompresi seimbang • kualitas tetap baik"},
  strong:{quality:.52,scale:.95,label:"Strong",desc:"Ukuran lebih kecil • kualitas standar"}
};

export async function compressPdf(file,mode,onProgress,isCancelled){
  const pdf=await loadPdf(file), out=await PDFDocument.create();
  const cfg=COMPRESS_MODES[mode]||COMPRESS_MODES.balanced;
  for(let i=1;i<=pdf.numPages;i++){
    if(isCancelled?.()) throw new Error("CANCELLED");
    onProgress?.((i-1)/pdf.numPages*100,`Mengompres halaman ${i} dari ${pdf.numPages}...`);
    const img=await renderPage(pdf,i,cfg.scale,cfg.quality);
    const jpg=await out.embedJpg(await img.blob.arrayBuffer());
    const page=out.addPage([img.width,img.height]);
    page.drawImage(jpg,{x:0,y:0,width:img.width,height:img.height});
    await sleep(0);
  }
  const bytes=await out.save({useObjectStreams:true,addDefaultPage:false});
  onProgress?.(100,"Kompresi selesai.");
  return new Blob([bytes],{type:"application/pdf"});
}

export async function mergePdfs(files,onProgress,isCancelled){
  const out=await PDFDocument.create();
  for(let i=0;i<files.length;i++){
    if(isCancelled?.()) throw new Error("CANCELLED");
    onProgress?.(i/files.length*100,`Membaca file ${i+1} dari ${files.length}...`);
    const src=await loadPdfLib(files[i]);
    const pages=await out.copyPages(src,src.getPageIndices());
    pages.forEach(p=>out.addPage(p));
  }
  const bytes=await out.save({useObjectStreams:true});
  onProgress?.(100,"Penggabungan selesai.");
  return new Blob([bytes],{type:"application/pdf"});
}

export function parseGroups(input,max){
  if(!input.trim()) return [Array.from({length:max},(_,i)=>i+1)];
  const groups=[];
  for(const part of input.split(",")){
    const p=part.trim(); if(!p) continue;
    const m=p.match(/^(\d+)\s*-\s*(\d+)$/);
    if(m){
      let a=Number(m[1]),b=Number(m[2]); if(a>b)[a,b]=[b,a];
      if(a<1||b>max) throw new Error(`Rentang ${p} berada di luar halaman PDF (${max} halaman).`);
      const arr=[]; for(let i=a;i<=b;i++) arr.push(i);
      groups.push(arr);
    }else if(/^\d+$/.test(p)){
      const n=Number(p); if(n<1||n>max) throw new Error(`Halaman ${n} tidak tersedia. PDF hanya memiliki ${max} halaman.`);
      groups.push([n]);
    }else throw new Error(`Format rentang "${p}" tidak valid. Contoh yang benar: 1-3,5,8-10`);
  }
  if(!groups.length) throw new Error("Masukkan rentang halaman yang valid, atau kosongkan untuk semua halaman.");
  return groups;
}

export async function splitPdf(file,rangesInput,onProgress,isCancelled){
  const totalPages=(await loadPdf(file)).numPages;
  const groups=parseGroups(rangesInput||"",totalPages);
  if(groups.length>SPLIT_OUTPUT_LIMIT)
    throw new Error(`Permintaan ini menghasilkan ${groups.length} file terpisah, melebihi batas ${SPLIT_OUTPUT_LIMIT} file sekali proses. Kurangi jumlah rentang yang dipisah koma.`);
  const src=await loadPdfLib(file), results=[];
  for(let i=0;i<groups.length;i++){
    if(isCancelled?.()) throw new Error("CANCELLED");
    const group=groups[i];
    const out=await PDFDocument.create();
    const pages=await out.copyPages(src,group.map(n=>n-1));
    pages.forEach(p=>out.addPage(p));
    const bytes=await out.save({useObjectStreams:true});
    const label=group.length===1?`halaman-${group[0]}`:`halaman-${group[0]}-${group[group.length-1]}`;
    results.push({blob:new Blob([bytes],{type:"application/pdf"}),name:`${safeBase(file.name)}-${label}.pdf`});
    onProgress?.((i+1)/groups.length*100,`Membuat bagian ${i+1} dari ${groups.length}...`);
  }
  return results;
}

export async function extractText(file,onProgress,isCancelled){
  const pdf=await loadPdf(file), pages=[];
  for(let i=1;i<=pdf.numPages;i++){
    if(isCancelled?.()) throw new Error("CANCELLED");
    const page=await pdf.getPage(i), content=await page.getTextContent({disableCombineTextItems:true});
    const items=content.items||[];
    let text="", lastY=null, lastFontH=10;
    for(const item of items){
      const y=item.transform?.[5] ?? null;
      const isNewLine=lastY!==null && y!==null && Math.abs(y-lastY)>4;
      if(isNewLine && text && !text.endsWith("\n")) text+="\n";
      const isWhitespaceOnly=item.str && /^\s+$/.test(item.str);
      if(isWhitespaceOnly){
        const threshold=Math.max(lastFontH*0.9,8);
        if(item.width>threshold){
          if(!text.endsWith("\t") && !text.endsWith("\n")) text+="\t";
        } else if(!/[\s\t\n]$/.test(text)){
          text+=" ";
        }
      } else {
        text+=item.str||"";
        if(item.height) lastFontH=item.height;
      }
      if(item.hasEOL){ if(!text.endsWith("\n")) text+="\n"; }
      lastY=y;
    }
    pages.push(text.replace(/[ \t]+\n/g,"\n").replace(/\n{2,}/g,"\n").trim());
    page.cleanup();
    onProgress?.(i,pdf.numPages);
  }
  return pages;
}

export async function pdfToWord(file,onProgress,isCancelled){
  const docx=await getDocx();
  const pages=await extractText(file,(i,n)=>onProgress?.((i-1)/n*100,`Mengekstrak teks halaman ${i} dari ${n}...`),isCancelled);
  const children=[];
  pages.forEach((text,i)=>{
    if(i) children.push(new docx.Paragraph({children:[new docx.TextRun({text:`Halaman ${i+1}`,bold:true})],pageBreakBefore:true}));
    const lines=(text||"[Tidak ada text layer pada halaman ini]").split(/\n/);
    lines.forEach(line=>children.push(new docx.Paragraph({children:[new docx.TextRun({text:line})]})));
  });
  const document=new docx.Document({sections:[{properties:{},children}]});
  const blob=await docx.Packer.toBlob(document);
  onProgress?.(100,"Dokumen Word selesai.");
  return blob;
}

export function lineToCells(line){
  const cleaned=line.replace(/\u00a0/g," ").trim();
  if(!cleaned)return [""];
  if(/\t/.test(cleaned))return cleaned.split(/\t+/).map(s=>s.trim()).filter((s,idx,arr)=>s!==""||idx===arr.length-1);
  const cells=cleaned.split(/\s{2,}/).map(s=>s.trim()).filter(Boolean);
  return cells.length?cells:[cleaned];
}

export async function pdfToExcel(file,onProgress,isCancelled){
  const pages=await extractText(file,(i,n)=>onProgress?.((i-1)/n*100,`Mengekstrak data halaman ${i} dari ${n}...`),isCancelled);
  const rows=[["Halaman","Teks"]];
  pages.forEach((text,pi)=>{
    const lines=(text||"").split(/\n/).map(l=>l.trim()).filter(Boolean);
    if(!lines.length){ rows.push([pi+1,"[Tidak ada text layer pada halaman ini]"]); return; }
    lines.forEach(line=>{ const cells=lineToCells(line); rows.push([pi+1,...cells]); });
  });
  const wb=window.XLSX.utils.book_new(), ws=window.XLSX.utils.aoa_to_sheet(rows);
  ws["!cols"]=[{wch:10},{wch:90},{wch:40},{wch:40}];
  window.XLSX.utils.book_append_sheet(wb,ws,"PDF Data");
  const out=window.XLSX.write(wb,{bookType:"xlsx",type:"array"});
  onProgress?.(100,"Spreadsheet selesai.");
  return new Blob([out],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
}

export async function ocrPdf(file,lang,onProgress,isCancelled){
  const pdf=await loadPdf(file);
  if(pdf.numPages>OCR_PAGE_LIMIT) throw new Error(`PDF memiliki ${pdf.numPages} halaman. OCR dibatasi ${OCR_PAGE_LIMIT} halaman agar browser mobile tetap stabil.`);
  let worker=null, currentPage=1;
  try{
    worker=await window.Tesseract.createWorker(lang||"ind",1,{
      logger:m=>{
        if(isCancelled?.()) return;
        const inner=typeof m.progress==="number"?m.progress:0;
        onProgress?.(((currentPage-1)+inner)/pdf.numPages*100,`OCR halaman ${currentPage} dari ${pdf.numPages} • ${m.status||"memproses"}...`);
      }
    });
    let all=[];
    for(let i=1;i<=pdf.numPages;i++){
      if(isCancelled?.()) throw new Error("CANCELLED");
      currentPage=i;
      const rendered=await renderPage(pdf,i,1.55,.9);
      const result=await worker.recognize(await rendered.blob.arrayBuffer());
      all.push(`===== HALAMAN ${i} =====\n${result.data.text.trim()}`);
      onProgress?.(i/pdf.numPages*100,`OCR halaman ${i} dari ${pdf.numPages} selesai.`);
      await sleep(0);
    }
    return new Blob([all.join("\n\n")],{type:"text/plain;charset=utf-8"});
  } finally {
    if(worker){ try{ await worker.terminate(); }catch{} }
  }
}

export async function imagesToPdf(files,onProgress,isCancelled){
  const out=await PDFDocument.create();
  for(let i=0;i<files.length;i++){
    if(isCancelled?.()) throw new Error("CANCELLED");
    onProgress?.(i/files.length*100,`Menambahkan gambar ${i+1} dari ${files.length}...`);
    const f=files[i], bytes=await f.arrayBuffer();
    const isPng=f.type==="image/png"||/\.png$/i.test(f.name);
    let img;
    try{ img = isPng ? await out.embedPng(bytes) : await out.embedJpg(bytes); }
    catch(e){ throw new Error(`Gagal membaca gambar "${f.name}". Pastikan file JPG/PNG yang valid.`); }
    const page=out.addPage([img.width,img.height]);
    page.drawImage(img,{x:0,y:0,width:img.width,height:img.height});
  }
  const bytes=await out.save({useObjectStreams:true});
  onProgress?.(100,"PDF selesai dibuat.");
  return new Blob([bytes],{type:"application/pdf"});
}

export async function pdfToImages(file,format,scale,onProgress,isCancelled){
  const pdf=await loadPdf(file);
  const results=[];
  for(let i=1;i<=pdf.numPages;i++){
    if(isCancelled?.()) throw new Error("CANCELLED");
    onProgress?.((i-1)/pdf.numPages*100,`Merender halaman ${i} dari ${pdf.numPages}...`);
    const page=await pdf.getPage(i);
    const viewport=page.getViewport({scale});
    const canvas=document.createElement("canvas");
    canvas.width=Math.max(1,Math.round(viewport.width));
    canvas.height=Math.max(1,Math.round(viewport.height));
    const ctx=canvas.getContext("2d",{alpha:false});
    ctx.fillStyle="#fff";ctx.fillRect(0,0,canvas.width,canvas.height);
    await page.render({canvasContext:ctx,viewport}).promise;
    const mime=format==="png"?"image/png":"image/jpeg";
    const blob=await new Promise(r=>canvas.toBlob(r,mime,format==="png"?undefined:.92));
    results.push({blob,name:`${safeBase(file.name)}-halaman-${i}.${format}`});
    page.cleanup();
    canvas.width=1;canvas.height=1;
    onProgress?.(i/pdf.numPages*100,`Halaman ${i} dari ${pdf.numPages} selesai.`);
  }
  return results;
}

// ===== Organize PDF =====
export async function loadOrganizerPages(file){
  if(file.size>ORGANIZE_SIZE_LIMIT_MB*1024*1024)
    throw new Error(`File ini ${fmtBytes(file.size)}, melebihi batas ${ORGANIZE_SIZE_LIMIT_MB} MB untuk mode Organize di browser kamu. PDF berukuran besar biasanya berisi gambar/scan resolusi tinggi — coba compress dulu.`);
  const pdf=await loadPdf(file);
  if(pdf.numPages>ORGANIZE_PAGE_LIMIT){
    throw new Error(`PDF ini punya ${pdf.numPages} halaman, melebihi batas ${ORGANIZE_PAGE_LIMIT} halaman untuk mode Organize di browser kamu.`);
  }
  return pdf;
}
export async function renderThumbnail(pdf,pageNo,scale=0.28){
  const page=await pdf.getPage(pageNo);
  const viewport=page.getViewport({scale});
  const canvas=document.createElement("canvas");
  canvas.width=Math.max(1,Math.round(viewport.width));
  canvas.height=Math.max(1,Math.round(viewport.height));
  const ctx=canvas.getContext("2d",{alpha:false});
  ctx.fillStyle="#fff";ctx.fillRect(0,0,canvas.width,canvas.height);
  await page.render({canvasContext:ctx,viewport}).promise;
  page.cleanup();
  return canvas;
}
export async function buildOrganizedPdf(file,organizePages,onProgress,isCancelled){
  const kept=organizePages.filter(p=>!p.removed);
  if(!kept.length) throw new Error("Semua halaman ditandai hapus. Sisakan minimal satu halaman.");
  const src=await loadPdfLib(file), out=await PDFDocument.create();
  onProgress?.(5,"Menyalin halaman...");
  const copied=await out.copyPages(src,kept.map(e=>e.originalIndex));
  for(let i=0;i<copied.length;i++){
    if(isCancelled?.()) throw new Error("CANCELLED");
    const entry=kept[i], p=copied[i];
    if(entry.rotation) p.setRotation(window.PDFLib.degrees(entry.rotation));
    out.addPage(p);
    onProgress?.(5+(i+1)/copied.length*90,`Menyusun halaman ${i+1} dari ${copied.length}...`);
    await sleep(0);
  }
  const bytes=await out.save({useObjectStreams:true});
  onProgress?.(100,"Susunan PDF selesai.");
  return new Blob([bytes],{type:"application/pdf"});
}
