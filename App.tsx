import React, { useEffect, useState, useMemo, useRef, memo } from "react";
import { generatePdfReport } from "./pdf";

type Classe = { id: string; value: { name: string; createdAt: string } };
type Eleve = { id: string; value: { classeId: string; numero: number; nom: string; sexe: "M"|"F"; contactParent?: string } };
type AnyRec = { id: string; value: any };

const COLLECTIONS = {
  classes: "dualsync_classes",
  eleves: "dualsync_eleves",
  surveillances: "dualsync_surveillances",
  retards: "dualsync_retards",
  absTD: "dualsync_abs_td",
  absCours: "dualsync_abs_cours",
  deplacements: "dualsync_deplacements",
  permissions: "dualsync_permissions",
  ensAbs: "dualsync_ens_abs",
  settings: "dualsync_settings",
  parametres: "parametres",
  driveBackup: "dualsync_drive_backup_config",
};

const DEFAULT_MATIERES = [
  "MATHÉMATIQUES","FRANÇAIS","HISTOIRE-GÉOGRAPHIE","PHILOSOPHIE","ANGLAIS","SVT","PCT","EPS","ÉCONOMIE","ESPAGNOL","COMPTABILITÉ","DROITS","ÉTUDE DE CAS","MATHS FINANCIÈRES","MATHS GÉNÉRALES","ANGLAIS COMMERCIAL","ÉCONOMIE GÉNÉRALE","INFORMATIQUE","ÉDUCATION CIVIQUE","TD MATHS","TD PCT","TD SVT","TD FRANÇAIS","COMPTABILITÉ APPROFONDIE","FISCALITÉ","SECRETARIAT","DROIT COMMERCIAL","GESTION"
];
const DEFAULT_MOTIFS_SURV = ["Bruit","Bavardage","Insolence","Tenue incorrecte","Téléphone portable","Absence non justifiée","Retard répété","Violence","Triche","Refus de travail","Autre"];
const DEFAULT_MOTIFS_PERM = ["Maladie","Voyage","Décès familial","Convocation administrative","Rendez-vous médical","Problème familial","Autre"];
const THEMES = [
  {name:"Bleu", val:"#2563eb"},
  {name:"Vert", val:"#059669"},
  {name:"Violet", val:"#7c3aed"},
  {name:"Orange", val:"#ea580c"},
  {name:"Rouge", val:"#dc2626"},
  {name:"Sombre", val:"#1f2937"},
];

function normalizeNom(s:string): string {
  try {
    return s.trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\s+/g,' ').trim();
  } catch { return s.trim().toLowerCase().replace(/\s+/g,' ').trim(); }
}

/* ---------- COLLEGE PERSISTENCE - STABLE ---------- */
function readLocalCollege(){
  try{
    const raw = localStorage.getItem('college_info');
    if(!raw) return null;
    const j = JSON.parse(raw);
    if(!j) return null;
    return j as {nom?:string, logoBase64?:string, anneeActive?:string, couleur?:string, logo?:string};
  }catch{ return null; }
}
function writeLocalCollege(nom:string, logoBase64:string, anneeActive:string, couleur:string){
  try{
    localStorage.setItem('college_info', JSON.stringify({
      nom: nom || "Collège Privé Saint Laurent",
      logoBase64: logoBase64 || "",
      logo: logoBase64 || "",
      anneeActive: anneeActive || "2026-2027",
      couleur: couleur || "#2563eb",
      savedAt: new Date().toISOString()
    }));
  }catch{}
}
function readLocalDriveConfig(){
  try{
    const raw = localStorage.getItem('drive_backup_config');
    if(!raw) return null;
    return JSON.parse(raw);
  }catch{ return null; }
}
function writeLocalDriveConfig(val:any){
  try{ localStorage.setItem('drive_backup_config', JSON.stringify(val)); }catch{}
}
function writeLocalLatestBackup(payload:any){
  try{ localStorage.setItem('dualsync_latest_backup', JSON.stringify({savedAt:new Date().toISOString(), data:payload})); }catch{}
}

/* ---------- OFFLINE XLSX PARSER FAST ---------- */
async function decompressRaw(data: Uint8Array): Promise<Uint8Array> {
  // @ts-ignore
  const DS = (globalThis as any).DecompressionStream;
  if (!DS) throw new Error('Decompression non disponible');
  const ds = new DS('deflate-raw');
  const blob = new Blob([data as any]);
  // @ts-ignore
  const stream = blob.stream().pipeThrough(ds);
  const ab = await new Response(stream).arrayBuffer();
  return new Uint8Array(ab);
}
async function unzipXlsxFast(buffer: ArrayBuffer): Promise<Map<string,string>> {
  const view = new DataView(buffer);
  const u8 = new Uint8Array(buffer);
  let eocdOffset = -1;
  const maxSearch = Math.min(u8.length, 70000);
  const startSearch = Math.max(0, u8.length - maxSearch);
  for (let i = u8.length - 22; i >= startSearch; i--) {
    if (view.getUint32(i, true) === 0x06054b50) { eocdOffset = i; break; }
  }
  if (eocdOffset === -1) throw new Error('ZIP invalide');
  const cdCount = view.getUint16(eocdOffset + 8, true);
  const cdOffset = view.getUint32(eocdOffset + 16, true);
  const filesMeta = new Map<string, {compSize:number, offset:number, method:number}>();
  let offset = cdOffset;
  for (let i=0;i<cdCount;i++) {
    if (offset + 46 > u8.length) break;
    if (view.getUint32(offset, true) !== 0x02014b50) break;
    const method = view.getUint16(offset+10, true);
    const compSize = view.getUint32(offset+20, true);
    const fileNameLen = view.getUint16(offset+28, true);
    const extraLen = view.getUint16(offset+30, true);
    const commentLen = view.getUint16(offset+32, true);
    const localOffset = view.getUint32(offset+42, true);
    const nameStart = offset+46;
    const nameEnd = nameStart + fileNameLen;
    if (nameEnd > u8.length) break;
    const fileName = new TextDecoder().decode(u8.subarray(nameStart, nameEnd));
    filesMeta.set(fileName, {compSize, offset: localOffset, method});
    offset += 46 + fileNameLen + extraLen + commentLen;
  }
  const result = new Map<string,string>();
  const decoder = new TextDecoder();
  for (const [name, info] of filesMeta) {
    if (!name.startsWith('xl/')) continue;
    if (!(name.endsWith('.xml') || name.endsWith('.rels'))) continue;
    if (name !== 'xl/sharedStrings.xml' && name !== 'xl/worksheets/sheet1.xml' && !name.startsWith('xl/worksheets/sheet1')) {
      if (name.startsWith('xl/worksheets/') && name !== 'xl/worksheets/sheet1.xml') continue;
    }
    const localOff = info.offset;
    if (localOff + 30 > u8.length) continue;
    if (view.getUint32(localOff, true) !== 0x04034b50) continue;
    const fileNameLen = view.getUint16(localOff+26, true);
    const extraLen = view.getUint16(localOff+28, true);
    const dataStart = localOff + 30 + fileNameLen + extraLen;
    const dataEnd = dataStart + info.compSize;
    if (dataEnd > u8.length) continue;
    const compData = u8.subarray(dataStart, dataEnd);
    let decompressed: Uint8Array;
    if (info.method === 0) decompressed = compData;
    else if (info.method === 8) { try { decompressed = await decompressRaw(compData); } catch { continue; } }
    else continue;
    try { result.set(name, decoder.decode(decompressed)); } catch {}
  }
  return result;
}
async function parseXlsxOfflineFast(buffer: ArrayBuffer): Promise<any[][]> {
  const files = await unzipXlsxFast(buffer);
  let shared: string[] = [];
  const ssXml = files.get('xl/sharedStrings.xml');
  if (ssXml) {
    try {
      const doc = new DOMParser().parseFromString(ssXml, 'application/xml');
      shared = Array.from(doc.getElementsByTagName('si')).map(si => Array.from(si.getElementsByTagName('t')).map(t=>t.textContent||'').join(''));
    } catch {}
  }
  const sheetKeys = Array.from(files.keys()).filter(k=>k.startsWith('xl/worksheets/sheet')).sort();
  const sheetXml = sheetKeys.length? files.get(sheetKeys[0]) : undefined;
  if (!sheetXml) throw new Error('Feuille non trouvée');
  const doc = new DOMParser().parseFromString(sheetXml, 'application/xml');
  const rows: string[][] = [];
  const colLetterToIdx = (s:string) => { let n=0; for(let i=0;i<s.length;i++) n=n*26+(s.charCodeAt(i)-64); return n-1; };
  let rowCount=0;
  for (const rowEl of Array.from(doc.getElementsByTagName('row'))) {
    if (rowCount>=100) break;
    const cNodes = Array.from(rowEl.getElementsByTagName('c'));
    if (!cNodes.length) continue;
    const rowArr: string[] = [];
    for (const c of cNodes) {
      const rAttr = c.getAttribute('r') || '';
      const letters = rAttr.replace(/[0-9]/g,'');
      const colIdx = letters ? colLetterToIdx(letters) : rowArr.length;
      if (colIdx>2) continue;
      while (rowArr.length <= colIdx) rowArr.push('');
      const t = c.getAttribute('t');
      let val='';
      if (t==='s') { const vEl=c.getElementsByTagName('v')[0]; if(vEl){ const idx=parseInt(vEl.textContent||'0',10); val=shared[idx]||''; } }
      else if (t==='inlineStr') val=c.getElementsByTagName('t')[0]?.textContent||'';
      else { const vEl=c.getElementsByTagName('v')[0]; val=vEl?.textContent||c.getElementsByTagName('t')[0]?.textContent||''; }
      rowArr[colIdx]=val;
    }
    if (rowArr.some(x=> (x||'').trim())) { rows.push(rowArr.slice(0,3)); rowCount++; }
  }
  return rows;
}

/* ---------- API ---------- */
async function api(action:string, args:any){
  const res = await fetch("actions",{method:"POST",headers:{"content-type":"application/json"},body:JSON.stringify({action,args})});
  const j = await res.json();
  if(!res.ok || j.error) throw new Error(j.error || "api error");
  return j.data;
}
async function listAll(col:string):Promise<AnyRec[]>{
  try{
    const r = await api("records.list",{collection:col,limit:100});
    let all = r.records as AnyRec[];
    let cursor = r.next_cursor;
    while(cursor){
      const nr = await api("records.list",{collection:col,limit:100,cursor});
      all = all.concat(nr.records);
      cursor = nr.next_cursor;
    }
    return all;
  }catch{ return []; }
}

/* --- stable college_info put logic: records_put emulation --- */
async function putCollegeInfoStable(value: {nom:string, logoBase64:string, anneeActive:string, couleur:string}){
  writeLocalCollege(value.nom, value.logoBase64, value.anneeActive, value.couleur);
  const payload = {
    nom: value.nom,
    logoBase64: value.logoBase64,
    logo: value.logoBase64,
    anneeActive: value.anneeActive,
    couleur: value.couleur,
    _stableId: "college_info",
    updatedAt: new Date().toISOString()
  };
  try{
    const existingExact = await api("records.get",{collection:COLLECTIONS.parametres, id:"college_info"}).catch(()=>null);
    if(existingExact){
      await api("records.replace",{collection:COLLECTIONS.parametres, id:"college_info", value: payload});
      return;
    }
  }catch{}
  try{
    const all = await listAll(COLLECTIONS.parametres);
    const stable = all.find((r:any)=> r.id==="college_info" || r.value?._stableId==="college_info");
    if(stable){
      await api("records.replace",{collection:COLLECTIONS.parametres, id: stable.id, value: payload});
      return;
    }
    const first = all.find((r:any)=> r.value?.nom || r.value?.logoBase64);
    if(first && all.length===1){
      await api("records.replace",{collection:COLLECTIONS.parametres, id: first.id, value: payload});
      return;
    }
  }catch{}
  try{
    await api("records.create",{collection:COLLECTIONS.parametres, value: payload, id: "college_info"} as any);
    return;
  }catch{}
  try{
    await api("records.create",{collection:COLLECTIONS.parametres, value: payload});
  }catch{}
}

async function putDriveConfigStable(value: any){
  const payload = {
    ...value,
    _stableId: "drive_backup_config",
    updatedAt: new Date().toISOString()
  };
  writeLocalDriveConfig(payload);
  // try parametres collection first for unified stable id, then dedicated collection
  try{
    const existingExact = await api("records.get",{collection:COLLECTIONS.parametres, id:"drive_backup_config"}).catch(()=>null);
    if(existingExact){
      await api("records.replace",{collection:COLLECTIONS.parametres, id:"drive_backup_config", value: payload});
      return;
    }
  }catch{}
  try{
    const all = await listAll(COLLECTIONS.parametres);
    const stable = all.find((r:any)=> r.id==="drive_backup_config" || r.value?._stableId==="drive_backup_config");
    if(stable){
      await api("records.replace",{collection:COLLECTIONS.parametres, id: stable.id, value: payload});
      return;
    }
  }catch{}
  try{
    const allDrive = await listAll(COLLECTIONS.driveBackup);
    if(allDrive.length>0){
      await api("records.replace",{collection:COLLECTIONS.driveBackup, id: allDrive[0].id, value: payload});
      return;
    }
  }catch{}
  try{
    await api("records.create",{collection:COLLECTIONS.parametres, value: payload, id: "drive_backup_config"} as any);
    return;
  }catch{}
  try{
    await api("records.create",{collection:COLLECTIONS.driveBackup, value: payload});
  }catch{}
}

/* ---------- HOOKS ---------- */
function useDebounced<T>(value:T, ms=300):T{
  const [v,setV]=useState(value);
  useEffect(()=>{ const id=setTimeout(()=>setV(value),ms); return ()=>clearTimeout(id); },[value,ms]);
  return v;
}

/* ---------- SKELETON ---------- */
function Skeleton(){
  return (
    <div className="animate-pulse space-y-4">
      <div className="h-8 w-56 bg-slate-200 rounded-full"></div>
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
        {Array.from({length:8}).map((_,i)=><div key={i} className="h-[88px] bg-white border rounded-[20px] p-4"><div className="h-3 w-20 bg-slate-100 rounded mb-3"/><div className="h-7 w-16 bg-slate-200 rounded"/></div>)}
      </div>
      <div className="grid lg:grid-cols-2 gap-4">
        <div className="h-[300px] bg-white border rounded-[20px]"></div>
        <div className="h-[300px] bg-white border rounded-[20px]"></div>
      </div>
    </div>
  );
}

/* ================= MAIN APP ================= */
export default function App(){
  const [tab,setTab]=useState(0);
  const TABS = ["Tableau de bord","Classe","Surveillance","Retard","Absence TD","Absence Cours","Déplacement","Permission","Bilan","Élève","Enseignant","Paramètre"];
  const [visited,setVisited]=useState<Set<number>>(()=>new Set([0,1]));
  const markVisited = (i:number)=> setVisited(prev=>{ if(prev.has(i)) return prev; const n=new Set(prev); n.add(i); return n; });

  const [classes,setClasses]=useState<Classe[]>([]);
  const [eleves,setEleves]=useState<Eleve[]>([]);
  const [survs,setSurvs]=useState<AnyRec[]>([]);
  const [retards,setRetards]=useState<AnyRec[]>([]);
  const [absTD,setAbsTD]=useState<AnyRec[]>([]);
  const [absCours,setAbsCours]=useState<AnyRec[]>([]);
  const [depls,setDepls]=useState<AnyRec[]>([]);
  const [perms,setPerms]=useState<AnyRec[]>([]);
  const [ensAbs,setEnsAbs]=useState<AnyRec[]>([]);
  const [settings,setSettings]=useState<any>({
    collegeNom:"Collège Privé Saint Laurent",
    logo:"",
    anneeActive:"2026-2027",
    annees:["2026-2027"],
    theme:"#2563eb",
    matieres:DEFAULT_MATIERES,
    motifsSurv:DEFAULT_MOTIFS_SURV,
    motifsPerm:DEFAULT_MOTIFS_PERM
  });
  const [driveConfig,setDriveConfig]=useState<any>({
    connected:false,
    dossierId:"",
    activer:false,
    frequence:"quotidien",
    derniereSauvegarde:"",
    email:"",
  });

  const [online,setOnline]=useState(typeof navigator!=="undefined"?navigator.onLine:true);
  const [initialLoading,setInitialLoading]=useState(true);
  const [loadMs,setLoadMs]=useState<number|null>(null);
  const loadStartRef = useRef(performance.now());
  const cacheRef = useRef<{classes:boolean,eleves:boolean,settings:boolean,rest:boolean}>({classes:false,eleves:false,settings:false,rest:false});
  const backupTimerRef = useRef<any>(null);

  async function loadCore(){
    const local = readLocalCollege();
    if(local){
      setSettings((prev:any)=>({
        ...prev,
        collegeNom: local.nom || prev.collegeNom || "Collège Privé Saint Laurent",
        logo: local.logoBase64 || (local as any).logo || prev.logo || "",
        anneeActive: local.anneeActive || prev.anneeActive || "2026-2027",
        theme: local.couleur || prev.theme || "#2563eb",
      }));
    }
    const localDrive = readLocalDriveConfig();
    if(localDrive){
      setDriveConfig((prev:any)=>({ ...prev, ...localDrive }));
    }

    try{
      const [c,e,st,param, driveList] = await Promise.all([
        listAll(COLLECTIONS.classes),
        listAll(COLLECTIONS.eleves),
        listAll(COLLECTIONS.settings),
        listAll(COLLECTIONS.parametres),
        listAll(COLLECTIONS.driveBackup)
      ]);
      setClasses(c as any); setEleves(e as any);

      let collegeFromParams: any = null;
      if(param.length>0){
        const exact = param.find((r:any)=> r.id==="college_info");
        const stable = param.find((r:any)=> r.value?._stableId==="college_info");
        const firstWithNom = param.find((r:any)=> r.value?.nom || r.value?.logoBase64);
        collegeFromParams = exact || stable || firstWithNom || null;
      }

      let baseSettings: any = null;
      if(collegeFromParams){
        const v = collegeFromParams.value;
        baseSettings = {
          collegeNom: v.nom || v.collegeNom || v.name || "Collège Privé Saint Laurent",
          logo: v.logoBase64 || v.logo || "",
          anneeActive: v.anneeActive || "2026-2027",
          theme: v.couleur || v.theme || "#2563eb",
        };
        writeLocalCollege(baseSettings.collegeNom, baseSettings.logo, baseSettings.anneeActive, baseSettings.theme);
      }

      // drive config from parametres or dedicated collection
      let driveFromParams: any = null;
      if(param.length>0){
        const exactDrive = param.find((r:any)=> r.id==="drive_backup_config");
        const stableDrive = param.find((r:any)=> r.value?._stableId==="drive_backup_config");
        driveFromParams = exactDrive || stableDrive || null;
      }
      if(driveFromParams){
        setDriveConfig((prev:any)=>({ ...prev, ...driveFromParams.value }));
        writeLocalDriveConfig(driveFromParams.value);
      } else if(driveList.length>0){
        const dv = driveList[0].value;
        setDriveConfig((prev:any)=>({ ...prev, ...dv }));
        writeLocalDriveConfig(dv);
      }

      if(st.length>0){
        const merged = {...st[0].value};
        if(!merged.matieres || merged.matieres.length===0) merged.matieres = DEFAULT_MATIERES;
        if(!merged.motifsSurv || merged.motifsSurv.length===0) merged.motifsSurv = DEFAULT_MOTIFS_SURV;
        if(!merged.motifsPerm || merged.motifsPerm.length===0) merged.motifsPerm = DEFAULT_MOTIFS_PERM;
        if(!merged.collegeNom) merged.collegeNom = baseSettings?.collegeNom || "Collège Privé Saint Laurent";
        if(!merged.annees || merged.annees.length===0) merged.annees = ["2026-2027"];
        if(!merged.anneeActive) merged.anneeActive = baseSettings?.anneeActive || "2026-2027";
        if(!merged.theme) merged.theme = baseSettings?.theme || "#2563eb";
        if(baseSettings?.logo && !merged.logo) merged.logo = baseSettings.logo;
        if(baseSettings){
          merged.collegeNom = baseSettings.collegeNom;
          merged.logo = baseSettings.logo || merged.logo;
          merged.anneeActive = baseSettings.anneeActive;
          merged.theme = baseSettings.theme;
        }
        setSettings((prev:any)=>({...prev, ...merged, _id:st[0].id}));
        if(!collegeFromParams){
          await putCollegeInfoStable({
            nom: merged.collegeNom || "Collège Privé Saint Laurent",
            logoBase64: merged.logo || "",
            anneeActive: merged.anneeActive || "2026-2027",
            couleur: merged.theme || "#2563eb"
          });
        }
      } else if(baseSettings){
        setSettings((prev:any)=>({
          ...prev,
          collegeNom: baseSettings.collegeNom,
          logo: baseSettings.logo,
          anneeActive: baseSettings.anneeActive,
          theme: baseSettings.theme,
        }));
      } else if(local){
        await putCollegeInfoStable({
          nom: local.nom || "Collège Privé Saint Laurent",
          logoBase64: local.logoBase64 || (local as any).logo || "",
          anneeActive: local.anneeActive || "2026-2027",
          couleur: local.couleur || "#2563eb"
        });
      }

      cacheRef.current.classes=true; cacheRef.current.eleves=true; cacheRef.current.settings=true;
      const elapsed = Math.round(performance.now()-loadStartRef.current);
      setLoadMs(elapsed);
    }finally{
      setInitialLoading(false);
    }
    const idle = (window as any).requestIdleCallback || ((cb:any)=>setTimeout(cb,800));
    idle(async ()=>{
      const [sv,rt,at,ac,dp,pm,ea] = await Promise.all([
        listAll(COLLECTIONS.surveillances),
        listAll(COLLECTIONS.retards),
        listAll(COLLECTIONS.absTD),
        listAll(COLLECTIONS.absCours),
        listAll(COLLECTIONS.deplacements),
        listAll(COLLECTIONS.permissions),
        listAll(COLLECTIONS.ensAbs),
      ]);
      setSurvs(sv); setRetards(rt); setAbsTD(at); setAbsCours(ac); setDepls(dp); setPerms(pm); setEnsAbs(ea);
      cacheRef.current.rest=true;
    });
  }

  useEffect(()=>{ loadCore(); const on=()=>setOnline(true), off=()=>setOnline(false); window.addEventListener("online",on); window.addEventListener("offline",off); return ()=>{window.removeEventListener("online",on); window.removeEventListener("offline",off);}; },[]);

  async function saveSettings(newVal:any){
    const merged = {...settings, ...newVal}; 
    const id = merged._id;
    delete merged._id;
    if(!merged.matieres || merged.matieres.length===0) merged.matieres = DEFAULT_MATIERES;
    if(!merged.motifsSurv || merged.motifsSurv.length===0) merged.motifsSurv = DEFAULT_MOTIFS_SURV;
    if(!merged.motifsPerm || merged.motifsPerm.length===0) merged.motifsPerm = DEFAULT_MOTIFS_PERM;
    if(!merged.annees || merged.annees.length===0) merged.annees = ["2026-2027"];
    if(!merged.anneeActive) merged.anneeActive = "2026-2027";
    if(!merged.collegeNom || !String(merged.collegeNom).trim()) merged.collegeNom = "Collège Privé Saint Laurent";
    if(!merged.theme) merged.theme = "#2563eb";
    if (newVal.logo === undefined && settings.logo) {
      merged.logo = settings.logo;
    }
    if (merged.logo && typeof merged.logo === 'string' && merged.logo.length > 12000) {
      if (settings.logo && settings.logo.length < merged.logo.length) {
        merged.logo = settings.logo;
      } else {
        merged.logo = "";
      }
    }

    writeLocalCollege(merged.collegeNom, merged.logo || "", merged.anneeActive, merged.theme);
    putCollegeInfoStable({
      nom: merged.collegeNom,
      logoBase64: merged.logo || "",
      anneeActive: merged.anneeActive,
      couleur: merged.theme
    });

    try{
      if(id) await api("records.replace",{collection:COLLECTIONS.settings,id,value:merged});
      else { const r = await api("records.create",{collection:COLLECTIONS.settings,value:merged}); merged._id=r.id; }
      setSettings({...merged, _id: id || merged._id});
    }catch(e:any){ 
      setSettings({...merged, _id: id || merged._id});
    }
  }

  async function saveDriveConfig(newVal:any){
    const merged = {...driveConfig, ...newVal};
    setDriveConfig(merged);
    writeLocalDriveConfig(merged);
    await putDriveConfigStable(merged);
  }

  // Auto backup debounced 30s after mutations
  async function performBackup(){
    try{
      const allCollections = [COLLECTIONS.classes, COLLECTIONS.eleves, COLLECTIONS.surveillances, COLLECTIONS.retards, COLLECTIONS.absTD, COLLECTIONS.absCours, COLLECTIONS.deplacements, COLLECTIONS.permissions, COLLECTIONS.ensAbs, COLLECTIONS.settings, COLLECTIONS.parametres];
      const results: any = {};
      for(const col of allCollections){
        const recs = await listAll(col);
        results[col] = recs.map(r=>r.value);
      }
      const payload = {
        version:"dualsync_v12_final",
        exportedAt: new Date().toISOString(),
        collegeNom: settings.collegeNom,
        anneeActive: settings.anneeActive,
        data: results
      };
      writeLocalLatestBackup(payload);
      const now = new Date().toISOString();
      const updated = {...driveConfig, derniereSauvegarde: now};
      setDriveConfig(updated);
      writeLocalDriveConfig(updated);
      await putDriveConfigStable(updated);
    }catch{}
  }

  useEffect(()=>{
    if(!driveConfig.activer) return;
    if(backupTimerRef.current) clearTimeout(backupTimerRef.current);
    backupTimerRef.current = setTimeout(()=>{ performBackup(); }, 30000);
    return ()=>{ if(backupTimerRef.current) clearTimeout(backupTimerRef.current); };
  },[classes, eleves, survs, retards, absTD, absCours, depls, perms, ensAbs, driveConfig.activer]);

  const elevesByClasse = useMemo(()=>{
    const m:Record<string,Eleve[]>={};
    for(const el of eleves){ const k=el.value.classeId; if(!m[k]) m[k]=[]; m[k].push(el); }
    for(const k in m) m[k].sort((a,b)=>a.value.numero-b.value.numero);
    return m;
  },[eleves]);

  const today = useMemo(()=> new Date().toISOString().slice(0,10),[]);
  const stats = useMemo(()=>{
    const retardMatinToday = retards.filter(r=>r.value.date===today && r.value.period==="matin").reduce((acc,r)=>acc+(r.value.eleveIds?.length||0),0);
    const retardSoirToday = retards.filter(r=>r.value.date===today && r.value.period==="soir").reduce((acc,r)=>acc+(r.value.eleveIds?.length||0),0);
    const absTDToday = absTD.filter(r=>r.value.date===today).reduce((a,r)=>a+(r.value.eleveIds?.length||0),0);
    const absCoursToday = absCours.filter(r=>r.value.date===today).reduce((a,r)=>a+(r.value.absences?.length||0),0);
    const permAttente = perms.filter(p=>!p.value.traite).length;
    const ensAbsToday = ensAbs.filter(e=>e.value.date===today).length;
    const survToday = survs.filter(s=>s.value.date===today).length;
    return {totalEleves:eleves.length, retardMatinToday, retardSoirToday, absTDToday, absCoursToday, permAttente, ensAbsToday, survToday};
  },[eleves,retards,absTD,absCours,perms,ensAbs,survs,today]);

  useEffect(()=>{ 
    const theme = settings.theme||"#2563eb";
    document.documentElement.style.setProperty("--primary", theme);
    document.documentElement.style.setProperty("--primary-color", theme);
  },[settings.theme]);

  const handleTab = (i:number)=>{
    markVisited(i);
    setTab(i);
    if(!cacheRef.current.rest && i>1){
      if(i>=2){
        (async()=>{
          const [sv,rt,at,ac,dp,pm,ea] = await Promise.all([
            listAll(COLLECTIONS.surveillances), listAll(COLLECTIONS.retards), listAll(COLLECTIONS.absTD),
            listAll(COLLECTIONS.absCours), listAll(COLLECTIONS.deplacements), listAll(COLLECTIONS.permissions), listAll(COLLECTIONS.ensAbs)
          ]);
          setSurvs(sv); setRetards(rt); setAbsTD(at); setAbsCours(ac); setDepls(dp); setPerms(pm); setEnsAbs(ea);
          cacheRef.current.rest=true;
        })();
      }
    }
  };

  return (
    <div data-safe-area-inset-root className="min-h-screen bg-[#f8fafc] text-slate-800" style={{boxSizing:"border-box",paddingTop:"var(--safe-area-inset-top)", fontFamily:"ui-sans-system, -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif"}}>
      <style>{`
        :root{--primary:#2563eb; --primary-color:#2563eb; --primary-light:#eff6ff}
        .primary-bg{background:var(--primary-color)} .primary-text{color:var(--primary-color)} .primary-border{border-color:var(--primary-color)}
        .card{ background:white; border-radius:20px; box-shadow:0 6px 24px rgba(0,0,0,0.04); border:1px solid #f1f5f9 }
        .no-scrollbar::-webkit-scrollbar{display:none} .no-scrollbar{-ms-overflow-style:none; scrollbar-width:none}
        @keyframes spin { to { transform: rotate(360deg) } }
        .spinner { width:14px; height:14px; border:2px solid rgba(255,255,255,0.4); border-top-color:white; border-radius:50%; animation: spin 0.6s linear infinite; display:inline-block; }
        .spinner-dark { width:14px; height:14px; border:2px solid rgba(0,0,0,0.15); border-top-color:rgba(0,0,0,0.6); border-radius:50%; animation: spin 0.6s linear infinite; display:inline-block; }
        .shimmer{ background:linear-gradient(90deg,#f1f5f9 25%,#e2e8f0 37%,#f1f5f9 63%); background-size:400% 100%; animation:shimmer 1.2s ease-in-out infinite; }
        @keyframes shimmer{0%{background-position:100% 0}100%{background-position:-100% 0}}
      `}</style>

      <header data-safe-area-sticky-top className="sticky z-30 bg-white/90 backdrop-blur border-b border-slate-100" style={{top:"var(--safe-area-inset-top)"}}>
        <div className="max-w-[1400px] mx-auto px-4 lg:px-6 h-[60px] flex items-center justify-between">
          <div className="flex items-center gap-3">
            {settings.logo ? <img src={settings.logo} className="w-9 h-9 rounded-xl object-cover border" alt="logo collège" loading="lazy" decoding="async"/> : <div className="w-9 h-9 rounded-xl primary-bg text-white grid place-items-center font-bold">D</div>}
            <div>
              <div className="font-bold text-[17px] leading-none tracking-tight">{settings.collegeNom || "Collège Privé Saint Laurent"}</div>
              <div className="text-[11px] text-slate-500 font-semibold tracking-widest uppercase">Année {settings.anneeActive} • DualSync Vercel • PDF offline</div>
            </div>
          </div>
          <div className="flex items-center gap-2">
            <span className={`px-2.5 py-1 rounded-full text-[11px] font-bold border ${online?'bg-emerald-50 text-emerald-700 border-emerald-200':'bg-amber-50 text-amber-700 border-amber-200'}`}>{online?'En ligne':'Hors ligne'}</span>
            {loadMs!=null && <span className="hidden md:inline-flex px-2.5 py-1 rounded-full bg-slate-900 text-white text-[11px] font-bold">⚡ {loadMs}ms</span>}
            <span className="hidden md:inline-flex px-3 py-1 rounded-full bg-slate-900 text-white text-[11px] font-bold">{eleves.length} élèves</span>
          </div>
        </div>
        <div className="border-t border-slate-100">
          <div className="max-w-[1400px] mx-auto px-2 lg:px-4">
            <div className="flex gap-1 overflow-x-auto no-scrollbar py-2">
              {TABS.map((t,i)=>(
                <button key={t} onClick={()=>handleTab(i)} className={`whitespace-nowrap px-4 py-2.5 rounded-full text-[13px] font-semibold transition-all border ${tab===i?'primary-bg text-white border-transparent shadow':'bg-white text-slate-600 border-slate-200 hover:border-slate-300'}`}>{t}</button>
              ))}
            </div>
          </div>
        </div>
      </header>

      <main className="max-w-[1400px] mx-auto px-4 lg:px-6 py-6">
        {initialLoading ? <Skeleton/> : (
          <>
            <div style={{display: tab===0 ? 'block':'none'}}>
              <Dashboard stats={stats} classes={classes} elevesByClasse={elevesByClasse} survs={survs} retards={retards} perms={perms} today={today} online={online} setTab={handleTab} loadMs={loadMs}/>
            </div>
            <div style={{display: tab===1 ? 'block':'none'}}>
              <ClasseTabFast classes={classes} eleves={eleves} elevesByClasse={elevesByClasse} setClassesOpt={setClasses} setElevesOpt={setEleves} reloadCore={loadCore}/>
            </div>
            {visited.has(2) && <div style={{display: tab===2 ? 'block':'none'}}><SurveillanceTab classes={classes} elevesByClasse={elevesByClasse} survs={survs} settings={settings} setSurvs={setSurvs} eleves={eleves} setElevesOpt={setEleves}/></div>}
            {visited.has(3) && <div style={{display: tab===3 ? 'block':'none'}}><RetardTab classes={classes} elevesByClasse={elevesByClasse} retards={retards} setRetards={setRetards} settings={settings}/></div>}
            {visited.has(4) && <div style={{display: tab===4 ? 'block':'none'}}><AbsTDTab classes={classes} elevesByClasse={elevesByClasse} absTD={absTD} setAbsTD={setAbsTD} settings={settings}/></div>}
            {visited.has(5) && <div style={{display: tab===5 ? 'block':'none'}}><AbsCoursTab classes={classes} elevesByClasse={elevesByClasse} absCours={absCours} setAbsCours={setAbsCours} settings={settings}/></div>}
            {visited.has(6) && <div style={{display: tab===6 ? 'block':'none'}}><DeplacementTab classes={classes} elevesByClasse={elevesByClasse} setClassesOpt={setClasses} setElevesOpt={setEleves} depls={depls} settings={settings}/></div>}
            {visited.has(7) && <div style={{display: tab===7 ? 'block':'none'}}><PermissionTab classes={classes} elevesByClasse={elevesByClasse} perms={perms} setPerms={setPerms} settings={settings}/></div>}
            {visited.has(8) && <div style={{display: tab===8 ? 'block':'none'}}><BilanTab classes={classes} elevesByClasse={elevesByClasse} survs={survs} retards={retards} absTD={absTD} absCours={absCours} perms={perms} ensAbs={ensAbs} settings={settings} depls={depls}/></div>}
            {visited.has(9) && <div style={{display: tab===9 ? 'block':'none'}}><EleveTab eleves={eleves} classes={classes} survs={survs} retards={retards} absTD={absTD} absCours={absCours} perms={perms} depls={depls} settings={settings} ensAbs={ensAbs}/></div>}
            {visited.has(10) && <div style={{display: tab===10 ? 'block':'none'}}><EnseignantTab classes={classes} ensAbs={ensAbs} setEnsAbs={setEnsAbs} settings={settings}/></div>}
            {visited.has(11) && <div style={{display: tab===11 ? 'block':'none'}}><ParametreTab settings={settings} saveSettings={saveSettings} classes={classes} eleves={eleves} survs={survs} retards={retards} absTD={absTD} absCours={absCours} perms={perms} ensAbs={ensAbs} depls={depls} reloadCore={loadCore} driveConfig={driveConfig} saveDriveConfig={saveDriveConfig} performBackup={performBackup}/></div>}
          </>
        )}
      </main>
    </div>
  );
}

const Dashboard = memo(function Dashboard({stats,classes,elevesByClasse,survs,retards,perms,today,online,setTab,loadMs}:any){
  const recent = useMemo(()=> [...survs,...retards,...perms].sort((a:any,b:any)=> (b.value.date||"").localeCompare(a.value.date||"")).slice(0,8), [survs,retards,perms]);
  return (
    <div className="space-y-6">
      <div className="flex flex-wrap items-end justify-between gap-3">
        <div><h1 className="text-[32px] leading-none font-extrabold tracking-tight">Tableau de bord</h1>
          <p className="text-slate-500 mt-2 text-[13px]">Vue d'ensemble — aujourd'hui {today} {online?"• Connecté":"• Hors ligne"}</p>
          <div className="mt-2 flex gap-2">
            <span className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full bg-emerald-50 border border-emerald-200 text-emerald-800 text-[11px] font-bold">⚡ Fast Excel • PDF offline bundlé • college_info stable + drive_backup_config • {loadMs?`${loadMs}ms`:'optimisé'}</span>
          </div>
        </div>
        <button onClick={()=>setTab(1)} className="px-4 py-2 rounded-full primary-bg text-white text-sm font-bold">Gérer les classes</button>
      </div>
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
        <StatCard label="Total élèves" value={stats.totalEleves} sub={`${classes.length} classes`} accent/>
        <StatCard label="Retards Matin" value={stats.retardMatinToday} sub="7h-12h aujourd'hui"/>
        <StatCard label="Retards Soir" value={stats.retardSoirToday} sub="15h-18h aujourd'hui"/>
        <StatCard label="Absences TD" value={stats.absTDToday} sub="Aujourd'hui"/>
        <StatCard label="Absences Cours" value={stats.absCoursToday} sub="Heures aujourd'hui"/>
        <StatCard label="Surveillances" value={stats.survToday} sub="Aujourd'hui"/>
        <StatCard label="Permissions attente" value={stats.permAttente} sub="À traiter"/>
        <StatCard label="Enseignants absents" value={stats.ensAbsToday} sub="Aujourd'hui"/>
      </div>
      <div className="grid lg:grid-cols-[1.2fr_0.8fr] gap-4">
        <div className="card p-5">
          <div className="font-bold text-[15px] mb-4 flex items-center justify-between">Répartition par classe <span className="text-[11px] bg-slate-100 px-2 py-1 rounded-full">cache mémoire</span></div>
          <div className="space-y-3">
            {classes.length===0 ? <div className="text-sm text-slate-400 py-8 text-center">Aucune classe — base vide au démarrage Vercel, créez d'abord une classe.</div> : classes.map((c:any)=>{
              const count = elevesByClasse[c.id]?.length||0;
              const max = Math.max(1, ...classes.map((cc:any)=>elevesByClasse[cc.id]?.length||0));
              return (
                <div key={c.id} className="flex items-center gap-3">
                  <div className="w-[68px] text-[13px] font-semibold truncate">{c.value.name}</div>
                  <div className="flex-1 h-2 bg-slate-100 rounded-full overflow-hidden"><div className="h-full primary-bg" style={{width:`${(count/max)*100}%`}} /></div>
                  <div className="w-10 text-right text-[12px] font-bold">{count}</div>
                </div>
              );
            })}
          </div>
        </div>
        <div className="card p-5">
          <div className="font-bold text-[15px] mb-4">Activités récentes</div>
          <div className="space-y-2 max-h-[300px] overflow-auto">
            {recent.length===0 ? <div className="text-sm text-slate-400">Aucune activité</div> : recent.map((r:any,i:number)=>(
              <div key={r.id+i} className="flex gap-2 text-[12px] py-2 border-b border-slate-50 last:border-0">
                <div className="w-2 h-2 rounded-full primary-bg mt-1.5 shrink-0"/>
                <div className="flex-1"><span className="font-semibold">{r.value.date}</span> — {r.value.motif||r.value.period||r.value.type||"Événement"}</div>
              </div>
            ))}
          </div>
          <div className="mt-4 p-3 rounded-xl bg-slate-50 border text-[11px] leading-relaxed">
            <b>Tous les rapports en PDF offline:</b> • Entête collège + logo 36x36 • Année active • Date jour • Période • Pourquoi • Titre • Table auto • Footer Généré par DualSync • Pagination
          </div>
        </div>
      </div>
    </div>
  );
});
function StatCard({label,value,sub,accent}:any){
  return <div className={`card p-4 ${accent?'ring-2 ring-[var(--primary-color)]/10':''}`}><div className="text-[11px] font-bold tracking-widest uppercase text-slate-500">{label}</div><div className="text-[26px] font-extrabold mt-1">{value}</div><div className="text-[12px] text-slate-500">{sub}</div></div>;
}

/* ================= CLASSE FAST ================= */
function ClasseTabFast({classes,eleves,elevesByClasse,setClassesOpt,setElevesOpt,reloadCore}:any){
  const [selected,setSelected]=useState<string|null>(null);
  const [showCreateModal,setShowCreateModal]=useState(false);
  const [newClassName,setNewClassName]=useState("");
  const [classCreateError,setClassCreateError]=useState<string|null>(null);
  const [showRenameModal,setShowRenameModal]=useState(false);
  const [renameValue,setRenameValue]=useState("");
  const [renameError,setRenameError]=useState<string|null>(null);
  const [addTab,setAddTab]=useState<"manuel"|"excel">("manuel");

  const [bulkText,setBulkText]=useState("");
  const [bulkDuplicateError,setBulkDuplicateError]=useState<string|null>(null);
  const bulkDebounced = useDebounced(bulkText,300);
  const [searchRaw,setSearchRaw]=useState("");
  const searchEleves = useDebounced(searchRaw,300);

  const [excelParsed,setExcelParsed]=useState<{nom:string;sexe:"M"|"F";numero:number}[]>([]);
  const [excelDuplicates,setExcelDuplicates]=useState<{nom:string;sexe:"M"|"F";numero:number}[]>([]);
  const [excelToAdd,setExcelToAdd]=useState<{nom:string;sexe:"M"|"F";numero:number}[]>([]);
  const [showExcelPreview,setShowExcelPreview]=useState(false);
  const [excelError,setExcelError]=useState<string|null>(null);
  const [isDragOver,setIsDragOver]=useState(false);
  const [isParsing,setIsParsing]=useState(false);
  const [xlsxState,setXlsxState]=useState<{lib:any,status:"idle"|"loading"|"ready"|"offline"|"error"}>({lib:null,status:"idle"});
  const [excelProgress,setExcelProgress]=useState<null|{fileName:string,percent:number,label:string}>(null);
  const [importing,setImporting]=useState(false);
  const [homonymDialog,setHomonymDialog]=useState<null | {title:string,message:string,duplicates:any[],classeName:string,onConfirm:()=>void,onCancel:()=>void}>(null);

  const [showDeleteClasseModal,setShowDeleteClasseModal]=useState(false);
  const [deletingClasse,setDeletingClasse]=useState(false);
  const [deleteEleveTarget,setDeleteEleveTarget]=useState<{id:string; nom:string; classeName:string}|null>(null);
  const [deletingEleveId,setDeletingEleveId]=useState<string|null>(null);
  const [optimisticDeletedClasseIds,setOptimisticDeletedClasseIds]=useState<Set<string>>(()=>new Set());
  const [optimisticDeletedEleveIds,setOptimisticDeletedEleveIds]=useState<Set<string>>(()=>new Set());
  const [page,setPage]=useState(1);
  const PAGE_SIZE=20;

  const fileInputRef = useRef<HTMLInputElement>(null);
  const gridRef = useRef<HTMLDivElement>(null);
  const [gridReady,setGridReady]=useState(false);
  useEffect(()=>{
    setGridReady(false);
    const raf=requestAnimationFrame(()=>{ setGridReady(true); });
    return ()=>cancelAnimationFrame(raf);
  },[classes]);

  const visibleClasses = useMemo(()=> classes.filter((c:any)=> !optimisticDeletedClasseIds.has(c.id)), [classes, optimisticDeletedClasseIds]);

  useEffect(()=>{
    if(visibleClasses.length>0 && !selected) setSelected(visibleClasses[0].id);
    if(selected && !visibleClasses.find((c:any)=>c.id===selected)) setSelected(visibleClasses.length>0?visibleClasses[0].id:null);
  },[visibleClasses, selected]);

  useEffect(()=>{ if(selected){ const c=classes.find((x:any)=>x.id===selected); if(c) setRenameValue(c.value.name); } },[selected,classes]);

  const selectedClasse = classes.find((c:any)=>c.id===selected);
  const rawSelectedEleves = selected ? (elevesByClasse[selected]||[]) : [];
  const selectedEleves = useMemo(()=> rawSelectedEleves.filter((e:any)=> !optimisticDeletedEleveIds.has(e.id)), [rawSelectedEleves, optimisticDeletedEleveIds]);
  const garcons = useMemo(()=> selectedEleves.filter((e:any)=>e.value.sexe==="M").length, [selectedEleves]);
  const filles = useMemo(()=> selectedEleves.filter((e:any)=>e.value.sexe==="F").length, [selectedEleves]);

  const filteredEleves = useMemo(()=>{
    if(!searchEleves.trim()) return selectedEleves;
    const s=searchEleves.toLowerCase();
    return selectedEleves.filter((el:any)=> el.value.nom.toLowerCase().includes(s));
  },[selectedEleves,searchEleves]);

  const totalPages = Math.max(1, Math.ceil(filteredEleves.length / PAGE_SIZE));
  const paginatedEleves = useMemo(()=>{
    const start=(page-1)*PAGE_SIZE;
    return filteredEleves.slice(start,start+PAGE_SIZE);
  },[filteredEleves,page]);

  useEffect(()=>{ setPage(1); },[selected, searchEleves]);

  useEffect(()=>{
    if(!bulkDebounced.trim() || !selected){ setBulkDuplicateError(null); return; }
    const lines = parseBulkLines(bulkDebounced);
    if(!lines.length){ setBulkDuplicateError(null); return; }
    const existingNorms = new Set(selectedEleves.map((e:any)=>normalizeNom(e.value.nom)));
    const batchSeen = new Set<string>();
    const dups:string[]=[];
    for(const p of lines){ const n=normalizeNom(p.nom); if(existingNorms.has(n)||batchSeen.has(n)) dups.push(p.nom); else batchSeen.add(n); }
    setBulkDuplicateError(dups.length?`Doublon dans cette classe: ${dups.slice(0,3).join(", ")}${dups.length>3?` +${dups.length-3}`:''}`:null);
  },[bulkDebounced, selected, selectedEleves]);

  function getDuplicateSplitFast(parsed: {nom:string;sexe:"M"|"F";numero:number}[], existingSet: Set<string>){
    const seen = new Set<string>();
    const toAdd:typeof parsed=[]; const dups:typeof parsed=[];
    for(const p of parsed){
      const norm=normalizeNom(p.nom);
      if(existingSet.has(norm)||seen.has(norm)) dups.push(p);
      else { seen.add(norm); toAdd.push(p); }
    }
    return {toAdd,dups};
  }

  async function createClasse(){
    const name=newClassName.trim(); if(!name){ setClassCreateError("Nom requis"); return; }
    if(classes.some((c:any)=>c.value.name.trim().toLowerCase()===name.toLowerCase().trim())){ setClassCreateError("Cette classe existe déjà"); return; }
    setClassCreateError(null);
    try{
      const rec=await api("records.create",{collection:COLLECTIONS.classes,value:{name,createdAt:new Date().toISOString()}});
      setClassesOpt((prev:any)=>[...prev,{id:rec.id,value:{name,createdAt:new Date().toISOString()}}]);
      setNewClassName(""); setShowCreateModal(false); setSelected(rec.id);
    }catch(e:any){ }
  }
  async function renameClasse(){
    if(!selected || !renameValue.trim()){ setRenameError("Nom requis"); return; }
    if(classes.some((c:any)=> c.id!==selected && c.value.name.trim().toLowerCase()===renameValue.trim().toLowerCase())){ setRenameError("Cette classe existe déjà"); return; }
    setRenameError(null);
    const c=classes.find((x:any)=>x.id===selected); if(!c) return;
    try{
      await api("records.replace",{collection:COLLECTIONS.classes,id:selected,value:{...c.value,name:renameValue.trim()}});
      setClassesOpt((prev:any)=>prev.map((cc:any)=>cc.id===selected?{...cc,value:{...cc.value,name:renameValue.trim()}}:cc));
      setShowRenameModal(false);
    }catch(e:any){ }
  }

  function openDeleteClasseModal(){ if(!selected||!selectedClasse) return; setShowDeleteClasseModal(true); }
  async function confirmDeleteClasse(){
    if(!selected||!selectedClasse) return;
    const classeIdToDelete=selected;
    const toDeleteIds = eleves.filter((el:any)=> el.value.classeId===classeIdToDelete).map((e:any)=>e.id);
    const snapClasses=[...classes]; const snapEleves=[...eleves];
    setShowDeleteClasseModal(false);
    setOptimisticDeletedClasseIds(prev=>{const n=new Set(prev); n.add(classeIdToDelete); return n;});
    setOptimisticDeletedEleveIds(prev=>{const n=new Set(prev); toDeleteIds.forEach(id=>n.add(id)); return n;});
    setSelected(null);
    setClassesOpt((prev:any)=>prev.filter((c:any)=>c.id!==classeIdToDelete));
    setElevesOpt((prev:any)=>prev.filter((e:any)=>!toDeleteIds.includes(e.id)));
    setDeletingClasse(true);
    try{
      const ops=[api("records.delete",{collection:COLLECTIONS.classes,id:classeIdToDelete}), ...toDeleteIds.map(id=>api("records.delete",{collection:COLLECTIONS.eleves,id}))];
      const results=await Promise.allSettled(ops);
      if(results.every(r=>r.status==="rejected")) throw new Error("Échec total");
    }catch(e:any){
      setOptimisticDeletedClasseIds(prev=>{const n=new Set(prev); n.delete(classeIdToDelete); return n;});
      setOptimisticDeletedEleveIds(prev=>{const n=new Set(prev); toDeleteIds.forEach(id=>n.delete(id)); return n;});
      setClassesOpt(snapClasses); setElevesOpt(snapEleves); setSelected(classeIdToDelete);
    }finally{ setDeletingClasse(false); setTimeout(()=>{ setOptimisticDeletedClasseIds(prev=>{const n=new Set(prev); n.delete(classeIdToDelete); return n;}); setOptimisticDeletedEleveIds(prev=>{const n=new Set(prev); toDeleteIds.forEach(id=>n.delete(id)); return n;}); },800); }
  }

  function detectSexeFromText(txt:string):"M"|"F"{ const up=txt.toUpperCase().trim(); if(up.endsWith(" F")||up.endsWith(";F")||up.endsWith(",F")) return "F"; if(up.endsWith(" M")||up.endsWith(";M")||up.endsWith(",M")) return "M"; const tokens=up.split(/\s+/); const last=tokens[tokens.length-1]; if(last==="F"||last==="FILLE"||last==="FEMME"||last==="FÉMININ") return "F"; return "M"; }
  function cleanNom(raw:string){ let s=raw.trim(); s=s.replace(/\s+[MF]\s*$/i,"").replace(/;[MF]\s*$/i,"").trim(); s=s.replace(/\s+FILLE\s*$/i,"").replace(/\s+GARCON\s*$/i,"").replace(/\s+GARÇON\s*$/i,"").trim(); return s; }
  function parseBulkLines(txt:string){ const lines=txt.split("\n").map(l=>l.trim()).filter(l=>l.length>0); const out:{nom:string;sexe:"M"|"F";numero:number}[]=[]; lines.forEach((line,idx)=>{ let nom=""; let sexe:"M"|"F"="M"; let numero=idx+1; if(line.includes(";")){ const parts=line.split(";").map(p=>p.trim()).filter(p=>p.length>0); if(parts.length===1){ nom=parts[0]; sexe=detectSexeFromText(line); nom=cleanNom(nom); }else if(parts.length===2){ if(!isNaN(Number(parts[0]))){ numero=Number(parts[0]); nom=parts[1]; }else if(parts[1].toUpperCase()==="M"||parts[1].toUpperCase()==="F"){ nom=parts[0]; sexe=parts[1].toUpperCase() as any; }else nom=parts[1]||parts[0]; nom=cleanNom(nom); }else{ const firstIsNum=!isNaN(Number(parts[0])); if(firstIsNum){ numero=Number(parts[0]); nom=parts[1]; sexe=(parts[2]?.toUpperCase().startsWith("F")?"F":"M") as any; }else{ nom=parts[0]; sexe=(parts[1]?.toUpperCase().startsWith("F")?"F":"M") as any; } nom=cleanNom(nom); } }else{ const m=line.match(/^(\d+)[\s\.\-]+(.+)$/); if(m){ numero=Number(m[1]); let rest=m[2].trim(); sexe=detectSexeFromText(rest); nom=cleanNom(rest); }else{ sexe=detectSexeFromText(line); nom=cleanNom(line); } } if(nom.length>=2) out.push({nom,sexe,numero}); }); return out; }

  async function doAddEleves(list:{nom:string;sexe:"M"|"F";numero:number}[], forceSuffix=false, skipped:any[]=[]){
    if(!selected) return; try{ let count=0; const existingCount=selectedEleves.length; const existingNorms=new Set(selectedEleves.map((e:any)=>normalizeNom(e.value.nom))); const addedNorms=new Set<string>(); const newRecs:Eleve[]=[]; for(let i=0;i<list.length;i++){ const p=list[i]; let finalNom=p.nom; if(forceSuffix){ let norm=normalizeNom(finalNom); if(existingNorms.has(norm)||addedNorms.has(norm)){ let base=p.nom.replace(/\s*\(\d+\)\s*$/,'').trim(); let suffix=2; let candidate=`${base} (${suffix})`; while(existingNorms.has(normalizeNom(candidate))||addedNorms.has(normalizeNom(candidate))){ suffix++; candidate=`${base} (${suffix})`; } finalNom=candidate; } } const rec=await api("records.create",{collection:COLLECTIONS.eleves,value:{classeId:selected,numero:existingCount+count+1,nom:finalNom,sexe:p.sexe,contactParent:""}}); newRecs.push({id:rec.id,value:{classeId:selected,numero:existingCount+count+1,nom:finalNom,sexe:p.sexe}} as any); addedNorms.add(normalizeNom(finalNom)); count++; } setElevesOpt((prev:any)=>[...prev,...newRecs]); setBulkText(""); setBulkDuplicateError(null); }catch(e:any){ }
  }
  async function handleBulkAdd(){ if(!selected) return; const parsed=parseBulkLines(bulkText); if(!parsed.length) return; const existingSet=new Set(selectedEleves.map((e:any)=>normalizeNom(e.value.nom))); const {toAdd,dups}=getDuplicateSplitFast(parsed, existingSet); if(dups.length>0){ setHomonymDialog({title:"Doublon détecté",message:`Le nom ${dups[0].nom} existe déjà dans ${selectedClasse?.value.name}. Homonyme ?`,duplicates:dups,classeName:selectedClasse?.value.name||"",onConfirm:async()=>{setHomonymDialog(null); await doAddEleves(parsed,true,[]);},onCancel:async()=>{setHomonymDialog(null); await doAddEleves(toAdd,false,dups);}}); return; } await doAddEleves(toAdd,false,[]); }

  async function ensureXlsxLazy(){
    if(xlsxState.status==="ready"||xlsxState.status==="offline") return xlsxState.lib;
    setXlsxState({lib:null,status:"loading"});
    try{
      // @ts-ignore
      const mod = await import("xlsx");
      const lib = (mod as any).default ? (mod as any).default : mod;
      const normalized={read:(lib.read||(mod as any).read), utils:(lib.utils||(mod as any).utils), ...lib};
      if(normalized.read && normalized.utils){
        setXlsxState({lib:normalized,status:"ready"}); return normalized;
      }
      throw new Error("bundle fail");
    }catch{
      const offline={isOfflineParser:true, offline:true};
      setXlsxState({lib:offline,status:"offline"}); return offline;
    }
  }
  useEffect(()=>{ ensureXlsxLazy(); },[]);
  useEffect(()=>{ if(addTab==="excel") ensureXlsxLazy(); },[addTab]);

  function parseSheetRowsFast(rows:any[][]){
    if(!rows||!rows.length) return [];
    let headerIdx = rows.findIndex(r=> r && r.some((c:any)=> String(c||"").toUpperCase().includes("NOM")));
    if(headerIdx===-1) headerIdx=0;
    const dataRows = rows.slice(headerIdx+1, headerIdx+101);
    const out:{nom:string;sexe:"M"|"F";numero:number}[]=[];
    dataRows.forEach((r,i)=>{
      if(!r || r.length===0) return;
      const c0 = r[0];
      const c1 = r[1];
      const c2 = r[2];
      const nomRaw = String(c1!==undefined? c1 : r[0] || "").trim();
      if(nomRaw.length<2) return;
      if(nomRaw.toUpperCase().includes("NOM ET PRENOMS") || nomRaw.toUpperCase()==="NOM") return;
      let numero = i+1;
      if(c0!==undefined && c0!=="" && !isNaN(Number(c0))) numero = Number(c0);
      let sexeRaw = String(c2!==undefined? c2 : "M").trim().toUpperCase();
      const sexe = sexeRaw.startsWith("F") ? "F" as const : "M" as const;
      const finalNom = cleanNom(nomRaw);
      if(finalNom.length>=2) out.push({numero, nom:finalNom, sexe});
    });
    return out;
  }

  async function handleExcelFile(file:File){
    setExcelError(null);
    setIsParsing(true);
    setExcelProgress({fileName:file.name, percent:0, label:`Lecture de ${file.name}...`});
    setTimeout(async()=>{
      try{
        setExcelProgress(prev=> prev ? {...prev, percent:5, label:`${file.name} — fichier lu 0%`} : null);
        const buffer = await file.arrayBuffer();
        setExcelProgress({fileName:file.name, percent:15, label:`${file.name} — fichier lu 15%`});
        const lib = await ensureXlsxLazy();
        setExcelProgress({fileName:file.name, percent:30, label:`${file.name} — décodage 30%`});
        let rows:any[][]=[];
        if(lib && !lib.isOfflineParser && lib.read){
          const wb = lib.read(buffer, {type:'array', codepage:65001, cellDates:false, sheetRows:100} as any);
          setExcelProgress({fileName:file.name, percent:45, label:`${file.name} — feuille extraite 45%`});
          const firstSheetName = wb.SheetNames[0];
          const ws = wb.Sheets[firstSheetName];
          rows = lib.utils.sheet_to_json(ws, {header:1, defval:"", blankrows:false}) as any[][];
        }else{
          rows = await parseXlsxOfflineFast(buffer);
        }
        setExcelProgress({fileName:file.name, percent:55, label:`${file.name} — ${rows.length} lignes lues 50%`});
        const trimmedRows = rows.map(r=> r.slice(0,3));
        const parsed = parseSheetRowsFast(trimmedRows);
        if(!parsed.length) throw new Error("Aucune ligne valide - vérifiez colonnes N°, NOM, Sexe (seulement 3 colonnes lues)");
        const existingSet = new Set(selectedEleves.map((e:any)=>normalizeNom(e.value.nom)));
        const {toAdd,dups}=getDuplicateSplitFast(parsed, existingSet);
        setExcelParsed(parsed); setExcelToAdd(toAdd); setExcelDuplicates(dups);
        setExcelProgress({fileName:file.name, percent:70, label:`${parsed.length} élèves détectés — ${toAdd.length} nouveaux / ${dups.length} doublons 70%`});
        setShowExcelPreview(true);
      }catch(err:any){
        setExcelError(`Impossible de lire ${file.name}: ${err.message}`);
        setExcelProgress(null);
      }finally{
        setIsParsing(false);
      }
    },0);
  }

  async function doAddExcel(list:any[], skipped:any[], forceSuffix=false){
    if(!selected) return;
    setImporting(true);
    setExcelProgress(prev=> prev ? {...prev, percent:80, label:`${prev.fileName} — import ${list.length} élèves 80%`} : {fileName:"import", percent:80, label:`Import ${list.length} élèves 80%`});
    try{
      const existingCount=selectedEleves.length;
      const existingNorms=new Set(selectedEleves.map((e:any)=>normalizeNom(e.value.nom)));
      const addedNorms=new Set<string>();
      const finalItems:{nom:string;sexe:"M"|"F";numero:number}[]=[];
      for(const p of list){
        let finalNom=p.nom;
        if(forceSuffix){
          let norm=normalizeNom(finalNom);
          if(existingNorms.has(norm)||addedNorms.has(norm)){
            let base=p.nom.replace(/\s*\(\d+\)\s*$/,'').trim();
            let suffix=2;
            let cand=`${base} (${suffix})`;
            while(existingNorms.has(normalizeNom(cand))||addedNorms.has(normalizeNom(cand))){ suffix++; cand=`${base} (${suffix})`; }
            finalNom=cand;
          }
        }
        const normFinal=normalizeNom(finalNom);
        if(!existingNorms.has(normFinal) || forceSuffix){
          if(!addedNorms.has(normFinal)){
            addedNorms.add(normFinal);
            finalItems.push({nom:finalNom, sexe:p.sexe, numero: existingCount + finalItems.length + 1});
          }
        }
      }
      const promises = finalItems.map(item=>
        api("records.create",{collection:COLLECTIONS.eleves,value:{classeId:selected, numero:item.numero, nom:item.nom, sexe:item.sexe, contactParent:""}})
      );
      const results = await Promise.all(promises);
      const newRecs:Eleve[] = results.map((rec:any, idx:number)=>({
        id:rec.id,
        value:{classeId:selected, numero:finalItems[idx].numero, nom:finalItems[idx].nom, sexe:finalItems[idx].sexe}
      } as any));
      setElevesOpt((prev:any)=>[...prev,...newRecs]);
      setExcelParsed([]); setExcelDuplicates([]); setExcelToAdd([]); setShowExcelPreview(false);
      if(fileInputRef.current) fileInputRef.current.value="";
      setExcelProgress({fileName:`${finalItems.length} importés`, percent:100, label:`${finalItems.length} élèves importés — 100%`});
      setTimeout(()=>setExcelProgress(null), 1500);
    }catch(e:any){ setExcelProgress(null); }
    finally{ setImporting(false); }
  }

  async function confirmExcelImport(){
    if(!selected) return;
    const existingSet=new Set(selectedEleves.map((e:any)=>normalizeNom(e.value.nom)));
    const {toAdd,dups}=getDuplicateSplitFast(excelParsed, existingSet);
    if(dups.length>0){
      setHomonymDialog({
        title:"Doublons Excel",
        message:`Le nom ${dups[0].nom} existe déjà dans ${selectedClasse?.value.name}. Homonyme ?`,
        duplicates:dups,
        classeName:selectedClasse?.value.name||"",
        onConfirm:async()=>{setHomonymDialog(null); await doAddExcel(excelParsed,[],true);},
        onCancel:async()=>{setHomonymDialog(null); await doAddExcel(toAdd,dups,false);}
      });
      return;
    }
    await doAddExcel(toAdd,[],false);
  }

  function openDeleteEleveModal(el:any){ setDeleteEleveTarget({id:el.id,nom:el.value.nom,classeName:selectedClasse?.value.name||""}); }
  async function confirmDeleteEleve(){
    if(!deleteEleveTarget) return; const targetId=deleteEleveTarget.id; const snap=[...eleves];
    setDeleteEleveTarget(null); setOptimisticDeletedEleveIds(prev=>{const n=new Set(prev); n.add(targetId); return n;}); setElevesOpt((prev:any)=>prev.filter((e:any)=>e.id!==targetId)); setDeletingEleveId(targetId);
    try{ await api("records.delete",{collection:COLLECTIONS.eleves,id:targetId}); }catch(e:any){ setOptimisticDeletedEleveIds(prev=>{const n=new Set(prev); n.delete(targetId); return n;}); setElevesOpt(snap); }finally{ setDeletingEleveId(null); setTimeout(()=>{ setOptimisticDeletedEleveIds(prev=>{const n=new Set(prev); n.delete(targetId); return n;}); },600); }
  }

  return (
    <div className="space-y-4">
      <h2 className="text-[24px] font-extrabold tracking-tight">Classe <span className="text-[11px] font-bold bg-emerald-50 text-emerald-700 border border-emerald-200 px-2 py-0.5 rounded-full align-middle ml-2">⚡ Fast Excel • PDF offline • college_info stable + drive_backup_config</span></h2>
      <div className="grid lg:grid-cols-[40%_60%] gap-4 items-start">
        <div className="card p-4">
          <div className="flex items-center justify-between mb-4"><div className="font-bold text-[15px]">Mes classes</div><span className="px-2.5 py-1 rounded-full bg-slate-100 text-[12px] font-bold">{visibleClasses.length}</span></div>
          <button onClick={()=>{setClassCreateError(null); setNewClassName(""); setShowCreateModal(true);}} className="w-full py-3 rounded-full primary-bg text-white font-bold text-[14px] shadow hover:opacity-90 transition">➕ Créer une classe</button>
          <div className="mt-4" ref={gridRef}>
            {visibleClasses.length===0 ? <div className="text-sm text-slate-400 py-12 text-center border-2 border-dashed rounded-2xl">Aucune classe<br/>Base vide au démarrage Vercel — créez votre première classe.</div> : (
              !gridReady ? <div className="grid grid-cols-2 gap-3">{Array.from({length:4}).map((_,i)=><div key={i} className="h-[72px] shimmer rounded-2xl"></div>)}</div> :
              <div className="grid grid-cols-2 gap-3">
                {visibleClasses.map((c:any)=>{
                  const isSel=selected===c.id; const count=elevesByClasse[c.id]?.length||0; const optimisticRemoved=eleves.filter((e:any)=>e.value.classeId===c.id && optimisticDeletedEleveIds.has(e.id)).length; const displayCount=Math.max(0,count-optimisticRemoved);
                  return <button key={c.id} onClick={()=>setSelected(c.id)} className={`text-left p-4 rounded-2xl border-2 transition-all ${isSel?'bg-[var(--primary-light)] primary-border shadow-sm':'bg-slate-50 border-slate-200 hover:bg-white'}`}><div className={`font-bold text-[14px] ${isSel?'primary-text':''} truncate`}>{c.value.name}</div><div className={`text-[12px] mt-1 ${isSel?'text-slate-600':'text-slate-500'}`}>{displayCount} élèves</div></button>;
                })}
              </div>
            )}
          </div>
          {deletingClasse && <div className="mt-3 flex items-center gap-2 text-[12px] bg-amber-50 border border-amber-200 rounded-xl p-2"><span className="spinner-dark"></span><span className="font-bold">Suppression en cours...</span></div>}
        </div>

        <div className="card p-4 min-h-[560px]">
          {!selected || !selectedClasse ? (
            <div className="h-full min-h-[500px] grid place-items-center text-center px-6"><div><div className="text-[48px] mb-3">📚</div><div className="font-bold text-[16px]">Sélectionnez une classe à gauche</div><div className="text-sm text-slate-400 mt-2">Excel rapide: sheetRows 100, 3 colonnes, Promise.all, progress bar, preview 3 noms. PDF offline partout.</div></div></div>
          ) : (
            <div className="space-y-5">
              <div className="flex flex-wrap gap-2 items-center justify-between border-b border-slate-100 pb-3">
                <div className="font-bold text-[20px]">{selectedClasse.value.name}</div>
                <div className="flex gap-2"><button onClick={()=>{setRenameError(null); setShowRenameModal(true);}} className="px-3 py-1.5 rounded-full bg-slate-900 text-white text-[12px] font-bold">✏️ Renommer</button><button onClick={openDeleteClasseModal} disabled={deletingClasse} className="px-3 py-1.5 rounded-full bg-red-50 text-red-600 border border-red-200 text-[12px] font-bold hover:bg-red-100 disabled:opacity-50 flex items-center gap-1.5">{deletingClasse?<span className="spinner-dark"></span>:null}🗑️ Supprimer</button></div>
              </div>
              <div className="grid grid-cols-3 gap-2">
                <div className="border rounded-2xl p-3 bg-slate-50 text-center"><div className="text-[11px] font-bold uppercase text-slate-500">Total</div><div className="text-[18px] font-extrabold">{selectedEleves.length}</div></div>
                <div className="border rounded-2xl p-3 bg-blue-50 text-center"><div className="text-[11px] font-bold uppercase text-blue-600">Garçons</div><div className="text-[18px] font-bold text-blue-700">{garcons}</div></div>
                <div className="border rounded-2xl p-3 bg-pink-50 text-center"><div className="text-[11px] font-bold uppercase text-pink-600">Filles</div><div className="text-[18px] font-bold text-pink-700">{filles}</div></div>
              </div>

              <div>
                <div className="font-bold text-[14px] mb-3">Ajouter des élèves</div>
                <div className="flex gap-1 mb-3 p-1 bg-slate-100 rounded-full w-fit">
                  <button onClick={()=>setAddTab("manuel")} className={`px-4 py-1.5 rounded-full text-[13px] font-bold transition ${addTab==="manuel"?"bg-white shadow text-slate-900":"text-slate-500"}`}>Manuel</button>
                  <button onClick={async()=>{ setAddTab("excel"); await ensureXlsxLazy(); }} className={`px-4 py-1.5 rounded-full text-[13px] font-bold transition flex items-center gap-1.5 ${addTab==="excel"?"bg-white shadow text-slate-900":"text-slate-500"}`}>Excel {xlsxState.status==="loading"?<span className="w-3 h-3 border-2 border-slate-300 border-t-slate-900 rounded-full animate-spin inline-block"/>: xlsxState.status==="ready"?<span className="w-2 h-2 bg-emerald-500 rounded-full inline-block"/>: xlsxState.status==="offline"?<span className="w-2 h-2 bg-emerald-500 rounded-full inline-block"/>:null} {xlsxState.status==="loading"?"(chargement...)":xlsxState.status==="ready"?"✓ bundlé":""}</button>
                </div>

                {addTab==="manuel" && (
                  <div className="border rounded-2xl p-3 bg-slate-50 space-y-2">
                    <textarea value={bulkText} onChange={e=>setBulkText(e.target.value)} placeholder={"Un nom par ligne:\nADJIBADE Amir\nAGBOKOU Laïa"} className={`w-full h-[110px] rounded-xl border p-3 text-[13px] bg-white resize-y ${bulkDuplicateError?'border-amber-400 ring-2 ring-amber-100':''}`}></textarea>
                    {bulkDuplicateError && <div className="text-[12px] bg-amber-50 border border-amber-200 text-amber-800 rounded-xl p-2">⚠️ {bulkDuplicateError}</div>}
                    <button onClick={handleBulkAdd} className="w-full py-2.5 rounded-full primary-bg text-white font-bold text-sm">Ajouter {bulkText.trim()? parseBulkLines(bulkText).length : ""} élèves</button>
                  </div>
                )}
                {addTab==="excel" && (
                  <div className="space-y-2">
                    <div className="flex items-center gap-2 text-[11px] px-3 py-2 rounded-full border bg-slate-50">
                      {xlsxState.status==="idle" && <span className="font-bold text-slate-500">Préchargement XLSX bundlé...</span>}
                      {xlsxState.status==="loading" && <><span className="w-3 h-3 border-2 border-slate-300 border-t-slate-900 rounded-full animate-spin"></span><span className="font-bold">Chargement XLSX bundlé (1x)...</span></>}
                      {xlsxState.status==="ready" && <><span className="w-2 h-2 bg-emerald-500 rounded-full"></span><span className="font-bold text-emerald-700">XLSX bundlé prêt • sheetRows:100 • 3 colonnes seulement • Promise.all</span></>}
                      {xlsxState.status==="offline" && <><span className="w-2 h-2 bg-emerald-500 rounded-full"></span><span className="font-bold text-emerald-700">Parseur offline rapide prêt (100 lignes, 3 cols)</span></>}
                    </div>

                    {excelProgress && (
                      <div className="border rounded-2xl p-3 bg-white space-y-2">
                        <div className="flex items-center gap-2">
                          <span className="spinner-dark"></span>
                          <span className="text-[12px] font-bold truncate">{excelProgress.label}</span>
                          <span className="ml-auto text-[11px] font-bold bg-slate-900 text-white px-2 py-0.5 rounded-full">{excelProgress.percent}%</span>
                        </div>
                        <div className="w-full h-2 bg-slate-100 rounded-full overflow-hidden">
                          <div className="h-full primary-bg transition-all duration-300 ease-out" style={{width:`${excelProgress.percent}%`}} />
                        </div>
                      </div>
                    )}

                    <div onClick={()=>fileInputRef.current?.click()} onDragOver={e=>{e.preventDefault(); setIsDragOver(true);}} onDragLeave={()=>setIsDragOver(false)} onDrop={e=>{e.preventDefault(); setIsDragOver(false); const f=e.dataTransfer.files?.[0]; if(f) handleExcelFile(f);}} className={`border-2 border-dashed rounded-2xl p-6 text-center cursor-pointer transition select-none ${isDragOver?'bg-blue-50 border-blue-400':'bg-slate-50 border-slate-300 hover:bg-white'}`}>
                      <div className="text-[24px] mb-1">📄</div><div className="font-bold text-sm">Glissez Excel ici ou cliquez — toc toc ⚡</div><div className="text-[11px] text-slate-500 mt-1">.xlsx, .xls — N°, NOM, Sexe uniquement (3 cols) • 4B extra ignoré • max 100 lignes • {isParsing?"Analyse...":"<2s pour 34 élèves"}</div>
                    </div>
                    <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" onChange={e=>{const f=e.target.files?.[0]; if(f) handleExcelFile(f);}} className="hidden"/>
                    {excelError && <div className="text-[12px] bg-red-50 border border-red-200 text-red-700 rounded-xl p-3">{excelError}</div>}
                  </div>
                )}
              </div>

              <div>
                <div className="flex gap-2 items-center justify-between mb-2"><div className="font-bold text-[14px]">Liste {selectedClasse.value.name} <span className="text-[11px] text-slate-500 font-normal">({filteredEleves.length} • p.{page}/{totalPages})</span></div><input value={searchRaw} onChange={e=>setSearchRaw(e.target.value)} placeholder="Rechercher... (debounced 300ms)" className="px-3 py-1.5 rounded-full border text-[12px] w-[190px]"/></div>
                <div className="overflow-auto border rounded-2xl">
                  <table className="w-full text-sm"><thead className="bg-slate-50 sticky top-0"><tr><th className="text-left p-2.5 font-bold">N°</th><th className="text-left p-2.5 font-bold">NOM</th><th className="text-left p-2.5 font-bold">SEXE</th><th className="text-left p-2.5 font-bold"></th></tr></thead>
                    <tbody>{paginatedEleves.length===0?<tr><td colSpan={4} className="p-8 text-center text-slate-400 text-sm">Aucun élève</td></tr>:paginatedEleves.map((el:any,idx:number)=><tr key={el.id} className="border-t hover:bg-slate-50"><td className="p-2.5">{(page-1)*PAGE_SIZE+idx+1}</td><td className="p-2.5 font-medium">{el.value.nom}</td><td className="p-2.5"><span className={`px-2 py-0.5 rounded-full text-[11px] font-bold ${el.value.sexe==="F"?"bg-pink-50 text-pink-700 border border-pink-200":"bg-blue-50 text-blue-700 border border-blue-200"}`}>{el.value.sexe}</span></td><td className="p-2.5"><button onClick={()=>openDeleteEleveModal(el)} disabled={deletingEleveId===el.id} className="px-2 py-1 rounded-full bg-red-50 text-red-600 text-xs border border-red-100 hover:bg-red-100 disabled:opacity-50 flex items-center gap-1">{deletingEleveId===el.id?<span className="spinner-dark"></span>:null}🗑️</button></td></tr>)}</tbody>
                  </table>
                </div>
                <div className="flex items-center justify-between mt-2">
                  <div className="text-[11px] text-slate-400">20 lignes/page • 3 cols Excel • parallel • cache</div>
                  <div className="flex gap-1">
                    <button disabled={page<=1} onClick={()=>setPage(p=>Math.max(1,p-1))} className="px-3 py-1 rounded-full border text-[11px] font-bold disabled:opacity-40">←</button>
                    <span className="px-2 py-1 text-[11px] font-bold">{page}/{totalPages}</span>
                    <button disabled={page>=totalPages} onClick={()=>setPage(p=>Math.min(totalPages,p+1))} className="px-3 py-1 rounded-full border text-[11px] font-bold disabled:opacity-40">→</button>
                  </div>
                </div>
              </div>
            </div>
          )}
        </div>
      </div>

      {showCreateModal && (
        <div className="fixed inset-0 z-50 bg-black/40 grid place-items-center p-4">
          <div className="bg-white rounded-[20px] p-6 w-full max-w-[420px] shadow-2xl border"><div className="font-bold text-[18px] mb-1">Créer une classe</div><input autoFocus value={newClassName} onChange={e=>{setNewClassName(e.target.value); setClassCreateError(null);}} onKeyDown={e=>{if(e.key==="Enter") createClasse();}} placeholder="Nom de la classe" className={`w-full px-4 py-3 rounded-full border text-sm ${classCreateError?'border-red-300 ring-2 ring-red-100':''}`}/>{classCreateError && <div className="mt-2 text-[12px] text-red-600 bg-red-50 border border-red-200 rounded-xl p-2">⚠️ {classCreateError}</div>}<div className="flex gap-2 mt-4 justify-end"><button onClick={()=>{setShowCreateModal(false); setNewClassName(""); setClassCreateError(null);}} className="px-5 py-2.5 rounded-full bg-slate-100 text-slate-700 text-sm font-bold">Annuler</button><button onClick={createClasse} className="px-6 py-2.5 rounded-full primary-bg text-white text-sm font-bold">Créer</button></div></div>
        </div>
      )}
      {showRenameModal && (
        <div className="fixed inset-0 z-50 bg-black/40 grid place-items-center p-4"><div className="bg-white rounded-[20px] p-6 w-full max-w-[420px] shadow-2xl border"><div className="font-bold text-[18px] mb-2">Renommer</div><input autoFocus value={renameValue} onChange={e=>{setRenameValue(e.target.value); setRenameError(null);}} placeholder="Nouveau nom" className={`w-full px-4 py-3 rounded-full border text-sm ${renameError?'border-red-300':''}`}/>{renameError && <div className="mt-2 text-[12px] text-red-600 bg-red-50 border rounded-xl p-2">⚠️ {renameError}</div>}<div className="flex gap-2 mt-4 justify-end"><button onClick={()=>{setShowRenameModal(false); setRenameError(null);}} className="px-5 py-2.5 rounded-full bg-slate-100 text-sm font-bold">Annuler</button><button onClick={renameClasse} className="px-6 py-2.5 rounded-full bg-slate-900 text-white text-sm font-bold">Renommer</button></div></div></div>
      )}
      {showDeleteClasseModal && selectedClasse && (
        <div className="fixed inset-0 z-[70] bg-black/50 grid place-items-center p-4"><div className="bg-white rounded-[24px] p-6 w-full max-w-[460px] shadow-2xl border"><div className="w-12 h-12 rounded-full bg-red-100 text-red-600 grid place-items-center text-[22px] mb-3">⚠️</div><div className="font-bold text-[18px]">Supprimer définitivement ?</div><div className="text-[13px] text-slate-600 mt-2">Supprimer <b>{selectedClasse.value.name}</b> et ses <b>{selectedEleves.length} élèves</b> ?</div><div className="flex gap-2 mt-5 justify-end"><button disabled={deletingClasse} onClick={()=>setShowDeleteClasseModal(false)} className="px-5 py-2.5 rounded-full bg-slate-100 text-slate-700 text-sm font-bold">Annuler</button><button disabled={deletingClasse} onClick={confirmDeleteClasse} className="px-5 py-2.5 rounded-full bg-red-600 text-white text-sm font-bold flex items-center gap-2">{deletingClasse?<span className="spinner"></span>:null}Supprimer</button></div></div></div>
      )}
      {deleteEleveTarget && (
        <div className="fixed inset-0 z-[70] bg-black/50 grid place-items-center p-4"><div className="bg-white rounded-[22px] p-5 w-full max-w-[400px] shadow-2xl border"><div className="font-bold text-[16px]">Supprimer l'élève ?</div><div className="text-[13px] text-slate-600 mt-2">Supprimer <b>{deleteEleveTarget.nom}</b> ?</div><div className="flex gap-2 mt-5 justify-end"><button onClick={()=>setDeleteEleveTarget(null)} className="px-4 py-2 rounded-full bg-slate-100 text-sm font-bold">Annuler</button><button onClick={confirmDeleteEleve} className="px-4 py-2 rounded-full bg-red-600 text-white text-sm font-bold">Supprimer</button></div></div></div>
      )}
      {showExcelPreview && (
        <div className="fixed inset-0 z-50 bg-black/40 grid place-items-center p-4"><div className="bg-white rounded-[20px] p-6 w-full max-w-[560px] shadow-2xl border">
          <div className="font-bold text-[16px]">⚡ {excelParsed.length} détectés • {excelToAdd.length} nouveaux • {excelDuplicates.length} doublons</div>
          <div className="mt-1 text-[11px] text-slate-500">3 colonnes lues seulement (N°, NOM, Sexe) • 4B extra ignoré • 100 lignes max • &lt;2s</div>
          <div className="mt-3 border rounded-xl divide-y max-h-[180px] overflow-auto">
            {(excelToAdd.length>0?excelToAdd:excelParsed).slice(0,3).map((p,i)=><div key={i} className="p-2.5 flex justify-between text-sm"><span>{p.nom}</span><span className={`px-2 py-0.5 rounded-full text-[11px] ${p.sexe==="F"?"bg-pink-50 text-pink-700":"bg-blue-50 text-blue-700"}`}>{p.sexe}</span></div>)}
            {(excelToAdd.length>3) && <div className="p-2 text-center text-[11px] text-slate-400">+{excelToAdd.length-3} autres</div>}
          </div>
          {excelProgress && (
            <div className="mt-3">
              <div className="w-full h-1.5 bg-slate-100 rounded-full overflow-hidden"><div className="h-full primary-bg" style={{width:`${excelProgress.percent}%`}}/></div>
              <div className="text-[11px] mt-1 font-bold">{excelProgress.label}</div>
            </div>
          )}
          <div className="flex gap-2 mt-4 justify-end"><button onClick={()=>{setShowExcelPreview(false); setExcelProgress(null);}} className="px-5 py-2.5 rounded-full bg-slate-100 text-sm font-bold">Annuler</button><button disabled={importing} onClick={confirmExcelImport} className="px-6 py-2.5 rounded-full primary-bg text-white text-sm font-bold flex items-center gap-2">{importing?<span className="spinner"></span>:null}Confirmer ({excelToAdd.length}) ⚡</button></div>
        </div></div>
      )}
      {homonymDialog && (
        <div className="fixed inset-0 z-[60] bg-black/50 grid place-items-center p-4"><div className="bg-white rounded-[20px] p-6 w-full max-w-[500px] shadow-2xl border"><div className="font-bold">{homonymDialog.title}</div><div className="text-[13px] text-slate-600 mt-2">{homonymDialog.message}</div><div className="mt-3 bg-amber-50 border border-amber-200 rounded-xl p-3 text-[12px]">{homonymDialog.duplicates.map((d:any)=>d.nom).join(", ")}</div><div className="flex gap-2 mt-5 justify-end"><button onClick={homonymDialog.onCancel} className="px-5 py-2.5 rounded-full bg-slate-100 text-sm font-bold">Ignorer doublons</button><button onClick={homonymDialog.onConfirm} className="px-5 py-2.5 rounded-full bg-slate-900 text-white text-sm font-bold">Ajouter homonyme</button></div></div></div>
      )}
    </div>
  );
}

/* ---------- PHOTO OCR LAZY ---------- */
function PhotoOCR({elevesByClasse,onMatched}:any){
  const [preview,setPreview]=useState(""); const [ocrText,setOcrText]=useState(""); const [status,setStatus]=useState(""); const [matchedIds,setMatchedIds]=useState<string[]>([]); const [loadingTess,setLoadingTess]=useState(false); const [tessReady,setTessReady]=useState(false);

  async function ensureTesseract(){
    if(tessReady) return true;
    if((window as any).Tesseract) { setTessReady(true); return true; }
    setLoadingTess(true); setStatus("Chargement Tesseract (lazy)...");
    try{
      // @ts-ignore
      const mod = await import("tesseract.js");
      (window as any).Tesseract = mod.default || mod;
      setTessReady(true); setStatus("Tesseract prêt (lazy)"); return true;
    }catch(e:any){
      setStatus("Tesseract non disponible - mode manuel"); return false;
    }finally{ setLoadingTess(false); }
  }

  async function handleFile(f:File){
    const url = URL.createObjectURL(f); setPreview(url); setStatus("Préparation...");
    const hasTess = await ensureTesseract();
    if(hasTess && (window as any).Tesseract){
      try{
        setStatus("OCR en cours (Tesseract)...");
        const { data } = await (window as any).Tesseract.recognize(f, 'fra');
        setOcrText(data.text); fuzzyMatch(data.text); setStatus("OCR terminé");
      }catch(e:any){ setStatus("OCR erreur: "+e.message); }
    }else{
      setTimeout(()=>{ setOcrText("Saisissez noms manuellement"); setStatus("Mode manuel - éditez puis Valider"); },200);
    }
  }
  function fuzzyMatch(text:string){
    const allEleves = Object.values(elevesByClasse).flat() as any[]; const lines=text.split("\n").map(l=>l.trim()).filter(Boolean); const ids:string[]=[];
    for(const line of lines){ let best:any=null, bestScore=0; for(const el of allEleves){ const name=el.value.nom.toLowerCase(); const score=similarity(line.toLowerCase(),name); if(score>bestScore){ bestScore=score; best=el; } } if(bestScore>0.4 && best) ids.push(best.id); }
    setMatchedIds([...new Set(ids)]); onMatched([...new Set(ids)]);
  }
  function similarity(a:string,b:string){ const setA=new Set(a.split(/\s+/)); const setB=new Set(b.split(/\s+/)); let inter=0; setA.forEach(x=>{ if(setB.has(x)) inter++; }); return inter/Math.max(setA.size,setB.size,1); }

  return (
    <div className="border rounded-2xl p-3 bg-white space-y-2">
      <div className="font-bold text-sm flex items-center gap-2">Photo liste manuscrite <span className={`text-[10px] px-2 py-0.5 rounded-full border ${tessReady?'bg-emerald-50 text-emerald-700 border-emerald-200':'bg-slate-100 text-slate-600'}`}>{tessReady?'Tesseract lazy prêt':'Tesseract non chargé (lazy)'}</span></div>
      <div className="flex gap-2 flex-wrap">
        <label className="px-3 py-1.5 rounded-full bg-slate-900 text-white text-xs font-bold cursor-pointer">📸 Photo<input type="file" accept="image/*" capture="environment" onChange={e=>{const f=e.target.files?.[0]; if(f) handleFile(f);}} className="hidden"/></label>
        <label className="px-3 py-1.5 rounded-full bg-slate-100 text-slate-700 text-xs font-bold cursor-pointer">🖼️ Image<input type="file" accept="image/*" onChange={e=>{const f=e.target.files?.[0]; if(f) handleFile(f);}} className="hidden"/></label>
        <button disabled={loadingTess} onClick={ensureTesseract} className="px-3 py-1.5 rounded-full border text-xs font-bold disabled:opacity-40">{loadingTess?'Chargement...':'Charger Tesseract (lazy, ~2MB)'}</button>
      </div>
      {preview && <img src={preview} className="w-full max-h-[200px] object-contain border rounded-xl" alt="preview" loading="lazy"/>}
      <div className="text-[11px] text-slate-500">{status}</div>
      <textarea value={ocrText} onChange={e=>setOcrText(e.target.value)} placeholder="Texte OCR éditable" className="w-full h-[80px] border rounded-xl p-2 text-xs"/>
      <button onClick={()=>fuzzyMatch(ocrText)} className="w-full py-1.5 rounded-full bg-slate-900 text-white text-xs font-bold">Valider OCR → cocher auto</button>
      {matchedIds.length>0 && <div className="text-xs bg-emerald-50 border border-emerald-200 rounded-xl p-2">{matchedIds.length} correspondants détectés.</div>}
    </div>
  );
}

/* ---------- PDF HELPERS ---------- */
function todayStr(){ return new Date().toISOString().slice(0,10); }
function safeName(s:string){ return s.replace(/[^a-zA-Z0-9 _-]/g,"").trim() || "Rapport"; }

/* ---------- SURVEILLANCE WITH PDF ---------- */
function SurveillanceTab({classes,elevesByClasse,survs,settings,setSurvs,setElevesOpt}:any){
  const [classeId,setClasseId]=useState(""); const [searchRaw,setSearchRaw]=useState(""); const search=useDebounced(searchRaw,300);
  const [checked,setChecked]=useState<Record<string,{contact?:string}>>({}); const [date,setDate]=useState(new Date().toISOString().slice(0,10)); const [heure,setHeure]=useState("08:00"); const [motif,setMotif]=useState(settings.motifsSurv?.[0]||"Bruit"); const [motifFree,setMotifFree]=useState(""); const [matiere,setMatiere]=useState(settings.matieres?.[0]||"MATHS"); const [enseignant,setEnseignant]=useState(""); const [histSearchRaw,setHistSearchRaw]=useState(""); const histSearch=useDebounced(histSearchRaw,300);
  const [pdfBusy,setPdfBusy]=useState(false);
  useEffect(()=>{ if(classes.length>0 && !classeId) setClasseId(classes[0].id); },[classes]);
  const list = useMemo(()=> (classeId? (elevesByClasse[classeId]||[]) : []).filter((el:any)=> el.value.nom.toLowerCase().includes(search.toLowerCase())).slice(0,50), [classeId,elevesByClasse,search]);
  async function save(){
    const ids=Object.keys(checked); if(!ids.length) return;
    for(const id of ids){ const contact=checked[id].contact; if(contact){ const el=(Object.values(elevesByClasse).flat() as any).find((e:any)=>e.id===id); if(el) try{ await api("records.replace",{collection:COLLECTIONS.eleves,id, value:{...el.value, contactParent:contact}}); setElevesOpt((prev:any)=>prev.map((e:any)=>e.id===id?{...e,value:{...e.value,contactParent:contact}}:e)); }catch{} } }
    const rec=await api("records.create",{collection:COLLECTIONS.surveillances,value:{date,heure,classeId,eleveIds:ids,motif: motif==="Autre"?motifFree:motif, matiere, enseignant, contacts:checked}});
    setSurvs((prev:any)=>[...prev,{id:rec.id,value:{date,heure,classeId,eleveIds:ids,motif,matiere,enseignant,contacts:checked}}]); setChecked({});
  }
  const filteredHist = useMemo(()=> survs.filter((s:any)=> !histSearch || (s.value.eleveIds||[]).some((id:string)=>{ const el=(Object.values(elevesByClasse).flat() as any).find((e:any)=>e.id===id); return el?.value.nom.toLowerCase().includes(histSearch.toLowerCase()); })).sort((a:any,b:any)=>b.value.date.localeCompare(a.value.date)).slice(0,100), [survs,histSearch,elevesByClasse]);

  async function handlePdf(){
    setPdfBusy(true);
    try{
      const columns = ["Date","Heure","Classe","Eleve","Motif","Matiere","Enseignant","Contact Parent"];
      const rows: string[][] = [];
      const data = filteredHist.length ? filteredHist : survs;
      for(const s of data){
        const classeName = classes.find((c:any)=>c.id===s.value.classeId)?.value.name || "";
        for(const eid of (s.value.eleveIds||[])){
          const el = (Object.values(elevesByClasse).flat() as any).find((e:any)=>e.id===eid);
          const eleveName = el?.value.nom || eid;
          const contact = s.value.contacts?.[eid]?.contact || el?.value.contactParent || "";
          rows.push([s.value.date||"", s.value.heure||"", classeName, eleveName, s.value.motif||"", s.value.matiere||"", s.value.enseignant||"", contact]);
        }
      }
      if(rows.length===0) rows.push(["-","-","-","Aucune donnee","-","-","-","-"]);
      await generatePdfReport({
        collegeName: settings.collegeNom || "Collège Privé Saint Laurent",
        logoDataUrl: settings.logo,
        anneeActive: settings.anneeActive || "2026-2027",
        dateJour: todayStr(),
        reportTitle: `Surveillance - ${todayStr()}`,
        columns, rows,
        fileName: `Rapport_Surveillance_${todayStr()}.pdf`
      });
    }catch(e:any){ }
    finally{ setPdfBusy(false); }
  }

  return (
    <div className="space-y-4"><h2 className="text-[24px] font-extrabold flex items-center gap-2">Surveillance <span className="text-[11px] bg-emerald-50 border border-emerald-200 text-emerald-700 px-2 py-0.5 rounded-full">PDF offline</span></h2>
      <div className="grid lg:grid-cols-[1.1fr_0.9fr] gap-4">
        <div className="card p-4 space-y-3"><div className="flex gap-2 flex-wrap"><select value={classeId} onChange={e=>setClasseId(e.target.value)} className="px-3 py-2 rounded-full border text-sm"><option value="">Classe</option>{classes.map((c:any)=><option key={c.id} value={c.id}>{c.value.name}</option>)}</select><input type="date" value={date} onChange={e=>setDate(e.target.value)} className="px-3 py-2 rounded-full border text-sm"/><input type="time" value={heure} onChange={e=>setHeure(e.target.value)} className="px-3 py-2 rounded-full border text-sm"/><input value={searchRaw} onChange={e=>setSearchRaw(e.target.value)} placeholder="Rechercher (300ms debounce)" className="flex-1 min-w-[140px] px-3 py-2 rounded-full border text-sm"/></div>
          <div className="max-h-[380px] overflow-auto border rounded-2xl divide-y">{list.map((el:any)=>{const isChecked=!!checked[el.id]; return <div key={el.id} className={`p-2 flex flex-col ${isChecked?'bg-amber-50':''}`}><label className="flex items-center gap-2"><input type="checkbox" checked={isChecked} onChange={e=>{ if(e.target.checked) setChecked(prev=>({...prev,[el.id]:{contact:el.value.contactParent||""}})); else { const cp={...checked}; delete cp[el.id]; setChecked(cp);} }} className="w-4 h-4"/><span className="text-sm font-medium">{el.value.numero}. {el.value.nom}</span><span className="text-[11px] px-1.5 py-0.5 rounded-full bg-slate-100">{el.value.sexe}</span></label>{isChecked && <div className="ml-6 mt-2 flex gap-2"><input value={checked[el.id]?.contact||""} onChange={e=>setChecked(prev=>({...prev,[el.id]:{...prev[el.id],contact:e.target.value}}))} placeholder="Contact parent (sauvegardé cache)" className="flex-1 px-3 py-1.5 rounded-full border text-xs"/></div>}</div>;})}{list.length===0 && <div className="p-6 text-center text-sm text-slate-400">Aucun élève</div>}</div><button onClick={save} className="w-full py-2.5 rounded-full primary-bg text-white font-bold">Enregistrer ({Object.keys(checked).length})</button></div>
        <div className="card p-4 space-y-3"><div className="font-bold flex items-center justify-between">Détails + Historique <button disabled={pdfBusy} onClick={handlePdf} className="px-3 py-1.5 rounded-full bg-slate-900 text-white text-[11px] font-bold flex items-center gap-1">{pdfBusy?<span className="spinner"></span>:null}📄 PDF</button></div><div className="grid grid-cols-2 gap-2"><select value={motif} onChange={e=>setMotif(e.target.value)} className="px-3 py-2 rounded-full border text-sm col-span-2">{settings.motifsSurv.map((m:string)=><option key={m}>{m}</option>)}</select>{motif==="Autre" && <input value={motifFree} onChange={e=>setMotifFree(e.target.value)} placeholder="Motif libre" className="px-3 py-2 rounded-full border text-sm col-span-2"/>}<select value={matiere} onChange={e=>setMatiere(e.target.value)} className="px-3 py-2 rounded-full border text-sm col-span-2">{settings.matieres.map((m:string)=><option key={m}>{m}</option>)}</select><input value={enseignant} onChange={e=>setEnseignant(e.target.value)} placeholder="Enseignant libre" className="px-3 py-2 rounded-full border text-sm col-span-2"/></div>
          <div className="pt-2 border-t"><div className="flex gap-2 mb-2"><div className="font-bold text-sm">Historique</div><input value={histSearchRaw} onChange={e=>setHistSearchRaw(e.target.value)} placeholder="Rechercher élève (debounced)" className="ml-auto px-3 py-1 rounded-full border text-xs"/></div><div className="max-h-[300px] overflow-auto text-xs border rounded-xl"><table className="w-full"><thead className="bg-slate-50 sticky top-0"><tr><th className="p-1 text-left">Date</th><th className="p-1">Classe</th><th className="p-1">Nbre</th><th className="p-1">Motif</th></tr></thead><tbody>{filteredHist.map((s:any)=><tr key={s.id} className="border-t"><td className="p-1">{s.value.date}</td><td className="p-1">{classes.find((c:any)=>c.id===s.value.classeId)?.value.name}</td><td className="p-1">{s.value.eleveIds?.length}</td><td className="p-1">{s.value.motif}</td></tr>)}</tbody></table></div></div>
        </div>
      </div>
    </div>
  );
}
function RetardTab({classes,elevesByClasse,retards,setRetards,settings}:any){
  const [classeId,setClasseId]=useState(""); const [period,setPeriod]=useState<"matin"|"soir">("matin"); const [date,setDate]=useState(new Date().toISOString().slice(0,10)); const [checked,setChecked]=useState<Record<string,boolean>>({}); const [filter,setFilter]=useState<"all"|"matin"|"soir">("all"); const [searchRaw,setSearchRaw]=useState(""); const search=useDebounced(searchRaw,300);
  const [pdfBusy,setPdfBusy]=useState(false);
  useEffect(()=>{ if(classes.length>0&&!classeId) setClasseId(classes[0].id); },[classes]);
  const list = useMemo(()=> (classeId? elevesByClasse[classeId]||[] : []).filter((el:any)=>el.value.nom.toLowerCase().includes(search.toLowerCase())).slice(0,50), [classeId,elevesByClasse,search]);
  async function save(){ const ids=Object.keys(checked); if(!ids.length) return; const rec=await api("records.create",{collection:COLLECTIONS.retards,value:{date,period,classeId,eleveIds:ids}}); setRetards((prev:any)=>[...prev,{id:rec.id,value:{date,period,classeId,eleveIds:ids}}]); setChecked({}); }
  const hist = useMemo(()=> retards.filter((r:any)=> filter==="all"||r.value.period===filter).sort((a:any,b:any)=>b.value.date.localeCompare(a.value.date)).slice(0,100), [retards,filter]);

  async function handlePdf(){
    setPdfBusy(true);
    try{
      const columns=["Date","Classe","Eleve","Periode","Matin/Soir"];
      const rows: string[][]=[];
      for(const r of hist){
        const classeName = classes.find((c:any)=>c.id===r.value.classeId)?.value.name||"";
        for(const eid of (r.value.eleveIds||[])){
          const el = (Object.values(elevesByClasse).flat() as any).find((e:any)=>e.id===eid);
          rows.push([r.value.date||"", classeName, el?.value.nom||eid, r.value.period||"", r.value.period==="matin"?"Matin":"Soir"]);
        }
      }
      if(rows.length===0) rows.push([todayStr(),"-","Aucune donnee","-","-"]);
      await generatePdfReport({
        collegeName: settings?.collegeNom || "Collège Privé Saint Laurent",
        logoDataUrl: settings?.logo,
        anneeActive: settings?.anneeActive || "2026-2027",
        dateJour: todayStr(),
        reportTitle: `Retard - ${filter} - ${todayStr()}`,
        columns, rows,
        fileName: `Rapport_Retard_${filter}_${todayStr()}.pdf`
      });
    }catch(e:any){ }
    finally{ setPdfBusy(false); }
  }

  return (
    <div className="space-y-4"><h2 className="text-[24px] font-extrabold flex items-center gap-2">Retard <span className="text-[11px] bg-emerald-50 border border-emerald-200 text-emerald-700 px-2 py-0.5 rounded-full">PDF offline</span></h2>
      <div className="grid lg:grid-cols-[1.1fr_0.9fr] gap-4">
        <div className="card p-4 space-y-3"><div className="flex flex-wrap gap-2"><select value={classeId} onChange={e=>setClasseId(e.target.value)} className="px-3 py-2 rounded-full border text-sm"><option value="">Classe</option>{classes.map((c:any)=><option key={c.id} value={c.id}>{c.value.name}</option>)}</select><div className="flex rounded-full border overflow-hidden"><button onClick={()=>setPeriod("matin")} className={`px-3 py-2 text-sm font-bold ${period==="matin"?"primary-bg text-white":"bg-white"}`}>Matin</button><button onClick={()=>setPeriod("soir")} className={`px-3 py-2 text-sm font-bold ${period==="soir"?"primary-bg text-white":"bg-white"}`}>Soir</button></div><input type="date" value={date} onChange={e=>setDate(e.target.value)} className="px-3 py-2 rounded-full border text-sm"/><input value={searchRaw} onChange={e=>setSearchRaw(e.target.value)} placeholder="Rechercher (debounced)" className="flex-1 px-3 py-2 rounded-full border text-sm"/></div>
          <PhotoOCR elevesByClasse={elevesByClasse} onMatched={(ids:string[])=>{ const m:Record<string,boolean>={}; ids.forEach(id=>m[id]=true); setChecked(m); }}/>
          <div className="max-h-[320px] overflow-auto border rounded-2xl divide-y">{list.map((el:any)=><label key={el.id} className={`flex items-center gap-2 p-2 ${checked[el.id]?'bg-amber-50':''}`}><input type="checkbox" checked={!!checked[el.id]} onChange={()=>setChecked(p=>{const cp={...p}; if(cp[el.id]) delete cp[el.id]; else cp[el.id]=true; return cp;})} className="w-4 h-4"/><span className="text-sm">{el.value.numero}. {el.value.nom}</span></label>)}</div><button onClick={save} className="w-full py-2.5 rounded-full primary-bg text-white font-bold">Valider ({Object.keys(checked).length})</button></div>
        <div className="card p-4"><div className="flex gap-2 mb-3"><div className="font-bold">Historique</div><div className="ml-auto flex gap-1"><button onClick={()=>setFilter("all")} className={`px-3 py-1 rounded-full text-xs border ${filter==="all"?"bg-slate-900 text-white":"bg-white"}`}>Tous</button><button onClick={()=>setFilter("matin")} className={`px-3 py-1 rounded-full text-xs border ${filter==="matin"?"bg-slate-900 text-white":"bg-white"}`}>Matin</button><button onClick={()=>setFilter("soir")} className={`px-3 py-1 rounded-full text-xs border ${filter==="soir"?"bg-slate-900 text-white":"bg-white"}`}>Soir</button></div><button disabled={pdfBusy} onClick={handlePdf} className="px-3 py-1 rounded-full bg-slate-900 text-white text-[11px] font-bold flex items-center gap-1">{pdfBusy?<span className="spinner"></span>:null}📄 PDF</button></div><div className="overflow-auto max-h-[500px] border rounded-xl text-xs"><table className="w-full"><thead className="bg-slate-50 sticky top-0"><tr><th className="p-2 text-left">Date</th><th className="p-2">Période</th><th className="p-2">Classe</th><th className="p-2">Nbre</th></tr></thead><tbody>{hist.map((r:any)=><tr key={r.id} className="border-t"><td className="p-2">{r.value.date}</td><td className="p-2">{r.value.period}</td><td className="p-2">{classes.find((c:any)=>c.id===r.value.classeId)?.value.name}</td><td className="p-2">{r.value.eleveIds?.length}</td></tr>)}</tbody></table></div></div>
      </div>
    </div>
  );
}
function AbsTDTab({classes,elevesByClasse,absTD,setAbsTD,settings}:any){
  const [classeId,setClasseId]=useState(""); const [date,setDate]=useState(new Date().toISOString().slice(0,10)); const [matiere,setMatiere]=useState("TD"); const [checked,setChecked]=useState<Record<string,boolean>>({}); const [searchRaw,setSearchRaw]=useState(""); const search=useDebounced(searchRaw,300);
  const [pdfBusy,setPdfBusy]=useState(false);
  useEffect(()=>{ if(classes.length>0&&!classeId) setClasseId(classes[0].id); },[classes]);
  const list = useMemo(()=> (classeId?elevesByClasse[classeId]||[]:[]).filter((el:any)=>el.value.nom.toLowerCase().includes(search.toLowerCase())).slice(0,50), [classeId,elevesByClasse,search]);
  async function save(){ const ids=Object.keys(checked); if(!ids.length) return; const rec=await api("records.create",{collection:COLLECTIONS.absTD,value:{date,matiere,classeId,eleveIds:ids}}); setAbsTD((p:any)=>[...p,{id:rec.id,value:{date,matiere,classeId,eleveIds:ids}}]); setChecked({}); }

  async function handlePdf(){
    setPdfBusy(true);
    try{
      const columns=["Date","Classe","Eleve","Matiere"];
      const rows: string[][]=[];
      for(const r of absTD){
        const classeName = classes.find((c:any)=>c.id===r.value.classeId)?.value.name||"";
        for(const eid of (r.value.eleveIds||[])){
          const el = (Object.values(elevesByClasse).flat() as any).find((e:any)=>e.id===eid);
          rows.push([r.value.date||"", classeName, el?.value.nom||eid, r.value.matiere||""]);
        }
      }
      if(rows.length===0) rows.push([todayStr(),"-","Aucune donnee","-"]);
      await generatePdfReport({
        collegeName: settings?.collegeNom || "Collège Privé Saint Laurent",
        logoDataUrl: settings?.logo,
        anneeActive: settings?.anneeActive || "2026-2027",
        dateJour: todayStr(),
        reportTitle: `Absence TD - ${todayStr()}`,
        columns, rows,
        fileName: `Rapport_Absence_TD_${todayStr()}.pdf`
      });
    }catch(e:any){ }
    finally{ setPdfBusy(false); }
  }

  return (
    <div className="space-y-4"><h2 className="text-[24px] font-extrabold flex items-center gap-2">Absence TD <span className="text-[11px] bg-emerald-50 border border-emerald-200 text-emerald-700 px-2 py-0.5 rounded-full">PDF offline</span></h2>
      <div className="grid lg:grid-cols-[1.1fr_0.9fr] gap-4">
        <div className="card p-4 space-y-3"><div className="flex flex-wrap gap-2"><select value={classeId} onChange={e=>setClasseId(e.target.value)} className="px-3 py-2 rounded-full border text-sm"><option value="">Classe</option>{classes.map((c:any)=><option key={c.id} value={c.id}>{c.value.name}</option>)}</select><input type="date" value={date} onChange={e=>setDate(e.target.value)} className="px-3 py-2 rounded-full border text-sm"/><input value={matiere} onChange={e=>setMatiere(e.target.value)} className="px-3 py-2 rounded-full border text-sm"/><input value={searchRaw} onChange={e=>setSearchRaw(e.target.value)} placeholder="Rechercher debounced" className="flex-1 px-3 py-2 rounded-full border text-sm"/></div>
          <PhotoOCR elevesByClasse={elevesByClasse} onMatched={(ids:string[])=>{ const m:Record<string,boolean>={}; ids.forEach(id=>m[id]=true); setChecked(m); }}/>
          <div className="max-h-[320px] overflow-auto border rounded-2xl divide-y">{list.map((el:any)=><label key={el.id} className={`flex gap-2 p-2 ${checked[el.id]?'bg-amber-50':''}`}><input type="checkbox" checked={!!checked[el.id]} onChange={()=>setChecked(p=>{const cp={...p}; if(cp[el.id]) delete cp[el.id]; else cp[el.id]=true; return cp;})} className="w-4 h-4"/><span className="text-sm">{el.value.numero}. {el.value.nom}</span></label>)}</div><button onClick={save} className="w-full py-2.5 rounded-full primary-bg text-white font-bold">Valider ({Object.keys(checked).length})</button></div>
        <div className="card p-4"><div className="font-bold mb-2 flex items-center justify-between">Historique <button disabled={pdfBusy} onClick={handlePdf} className="px-3 py-1 rounded-full bg-slate-900 text-white text-[11px] font-bold flex items-center gap-1">{pdfBusy?<span className="spinner"></span>:null}📄 PDF</button></div><div className="overflow-auto max-h-[500px] border rounded-xl text-xs"><table className="w-full"><thead className="bg-slate-50 sticky top-0"><tr><th className="p-2 text-left">Date</th><th className="p-2">Matière</th><th className="p-2">Classe</th><th className="p-2">Nbre</th></tr></thead><tbody>{absTD.slice(0,100).sort((a:any,b:any)=>b.value.date.localeCompare(a.value.date)).map((r:any)=><tr key={r.id} className="border-t"><td className="p-2">{r.value.date}</td><td className="p-2">{r.value.matiere}</td><td className="p-2">{classes.find((c:any)=>c.id===r.value.classeId)?.value.name}</td><td className="p-2">{r.value.eleveIds?.length}</td></tr>)}</tbody></table></div></div>
      </div>
    </div>
  );
}
function AbsCoursTab({classes,elevesByClasse,absCours,setAbsCours,settings}:any){
  const [classeId,setClasseId]=useState(""); const [date,setDate]=useState(new Date().toISOString().slice(0,10)); const [checked,setChecked]=useState<Record<string,{heures:Record<string,boolean>}>>({}); const [searchRaw,setSearchRaw]=useState(""); const search=useDebounced(searchRaw,300);
  const [pdfBusy,setPdfBusy]=useState(false);
  const HEURES=["7H","8H","9H","10H","11H","15H","16H"];
  useEffect(()=>{ if(classes.length>0&&!classeId) setClasseId(classes[0].id); },[classes]);
  const list=useMemo(()=> (classeId?elevesByClasse[classeId]||[]:[]).filter((el:any)=>el.value.nom.toLowerCase().includes(search.toLowerCase())).slice(0,30), [classeId,elevesByClasse,search]);
  function toggleHeure(eleveId:string, h:string){ setChecked(prev=>{ const cur=prev[eleveId]||{heures:{}}; const nh={...cur.heures, [h]: !cur.heures[h]}; if(!nh[h]) delete nh[h]; if(Object.keys(nh).length===0){ const cp={...prev}; delete cp[eleveId]; return cp; } return {...prev,[eleveId]:{heures:nh}}; }); }
  async function save(){ const absences=Object.entries(checked).map(([eleveId, v]:any)=>({eleveId, heures:Object.keys(v.heures)})); if(!absences.length) return; const rec=await api("records.create",{collection:COLLECTIONS.absCours,value:{date,classeId,absences}}); setAbsCours((p:any)=>[...p,{id:rec.id,value:{date,classeId,absences}}]); setChecked({}); }

  async function handlePdf(){
    setPdfBusy(true);
    try{
      const columns=["Date","Classe","Eleve","Heures manq.","Matiere"];
      const rows: string[][]=[];
      for(const r of absCours){
        const classeName = classes.find((c:any)=>c.id===r.value.classeId)?.value.name||"";
        for(const ab of (r.value.absences||[])){
          const el = (Object.values(elevesByClasse).flat() as any).find((e:any)=>e.id===ab.eleveId);
          rows.push([r.value.date||"", classeName, el?.value.nom||ab.eleveId, (ab.heures||[]).join(", "), "Cours"]);
        }
      }
      if(rows.length===0) rows.push([todayStr(),"-","Aucune donnee","-","-"]);
      await generatePdfReport({
        collegeName: settings?.collegeNom || "Collège Privé Saint Laurent",
        logoDataUrl: settings?.logo,
        anneeActive: settings?.anneeActive || "2026-2027",
        dateJour: todayStr(),
        reportTitle: `Absence Cours - ${todayStr()}`,
        columns, rows,
        fileName: `Rapport_Absence_Cours_${todayStr()}.pdf`
      });
    }catch(e:any){ }
    finally{ setPdfBusy(false); }
  }

  return (
    <div className="space-y-4"><h2 className="text-[24px] font-extrabold flex items-center gap-2">Absence Cours <span className="text-[11px] bg-emerald-50 border border-emerald-200 text-emerald-700 px-2 py-0.5 rounded-full">PDF offline</span></h2>
      <div className="grid lg:grid-cols-[1.2fr_0.8fr] gap-4">
        <div className="card p-4 space-y-3"><div className="flex flex-wrap gap-2"><select value={classeId} onChange={e=>setClasseId(e.target.value)} className="px-3 py-2 rounded-full border text-sm"><option value="">Classe</option>{classes.map((c:any)=><option key={c.id} value={c.id}>{c.value.name}</option>)}</select><input type="date" value={date} onChange={e=>setDate(e.target.value)} className="px-3 py-2 rounded-full border text-sm"/><input value={searchRaw} onChange={e=>setSearchRaw(e.target.value)} placeholder="Rechercher debounced" className="flex-1 px-3 py-2 rounded-full border text-sm"/></div>
          <div className="max-h-[460px] overflow-auto border rounded-2xl divide-y">{list.map((el:any)=>{ const cur=checked[el.id]; return <div key={el.id} className={`p-2 ${cur?'bg-amber-50':''}`}><div className="flex items-center gap-2"><span className="text-sm font-medium flex-1">{el.value.numero}. {el.value.nom}</span>{cur && <span className="text-[11px] bg-slate-900 text-white px-2 py-0.5 rounded-full">{Object.keys(cur.heures).length}h</span>}</div><div className="flex flex-wrap gap-1 mt-2 ml-2">{HEURES.map(h=><button key={h} onClick={()=>toggleHeure(el.id,h)} className={`px-2.5 py-1 rounded-full text-[11px] font-bold border ${cur?.heures[h]?'primary-bg text-white border-transparent':'bg-white'}`}>{h}</button>)}</div></div>; })}</div><button onClick={save} className="w-full py-2.5 rounded-full primary-bg text-white font-bold">Valider ({Object.keys(checked).length})</button></div>
        <div className="card p-4"><div className="font-bold mb-2 flex items-center justify-between">Historique <button disabled={pdfBusy} onClick={handlePdf} className="px-3 py-1 rounded-full bg-slate-900 text-white text-[11px] font-bold flex items-center gap-1">{pdfBusy?<span className="spinner"></span>:null}📄 PDF</button></div><div className="overflow-auto max-h-[500px] border rounded-xl text-xs"><table className="w-full"><thead className="bg-slate-50 sticky top-0"><tr><th className="p-2 text-left">Date</th><th className="p-2">Classe</th><th className="p-2">Détails</th></tr></thead><tbody>{absCours.slice(0,100).sort((a:any,b:any)=>b.value.date.localeCompare(a.value.date)).map((r:any)=><tr key={r.id} className="border-t"><td className="p-2">{r.value.date}</td><td className="p-2">{classes.find((c:any)=>c.id===r.value.classeId)?.value.name}</td><td className="p-2">{r.value.absences?.length} élèves</td></tr>)}</tbody></table></div></div>
      </div>
    </div>
  );
}
function DeplacementTab({classes,elevesByClasse,depls,settings}:any){
  const [src,setSrc]=useState(""); const [dest,setDest]=useState(""); const [checked,setChecked]=useState<Record<string,boolean>>({}); const [searchRaw,setSearchRaw]=useState(""); const search=useDebounced(searchRaw,300);
  const [pdfBusy,setPdfBusy]=useState(false);
  useEffect(()=>{ if(classes.length>0&&!src) setSrc(classes[0].id); },[classes]);
  const srcList = useMemo(()=> (src? elevesByClasse[src]||[]:[]).filter((el:any)=>el.value.nom.toLowerCase().includes(search.toLowerCase())).slice(0,50), [src,elevesByClasse,search]);
  const destOptions = useMemo(()=> classes.filter((c:any)=>c.id!==src), [classes,src]);
  async function move(){
    const ids=Object.keys(checked); if(!ids.length||!dest) return;
    for(const id of ids){ const el=(elevesByClasse[src]||[]).find((e:any)=>e.id===id) as any; if(el){ await api("records.replace",{collection:COLLECTIONS.eleves,id, value:{...el.value, classeId:dest}}); } }
    await api("records.create",{collection:COLLECTIONS.deplacements,value:{date:new Date().toISOString().slice(0,10),sourceClasseId:src,destClasseId:dest,eleveIds:ids}}); setChecked({});
  }

  async function handlePdf(){
    setPdfBusy(true);
    try{
      const columns=["Date","Eleve","Classe source","->","Classe accueil"];
      const rows: string[][]=[];
      for(const d of (depls||[])){
        const srcName = classes.find((c:any)=>c.id===d.value.sourceClasseId)?.value.name||d.value.sourceClasseId||"";
        const destName = classes.find((c:any)=>c.id===d.value.destClasseId)?.value.name||d.value.destClasseId||"";
        for(const eid of (d.value.eleveIds||[])){
          const el = (Object.values(elevesByClasse).flat() as any).find((e:any)=>e.id===eid);
          rows.push([d.value.date||"", el?.value.nom||eid, srcName, "->", destName]);
        }
      }
      if(rows.length===0) rows.push([todayStr(),"Aucune donnee","-","-","-"]);
      await generatePdfReport({
        collegeName: settings?.collegeNom || "Collège Privé Saint Laurent",
        logoDataUrl: settings?.logo,
        anneeActive: settings?.anneeActive || "2026-2027",
        dateJour: todayStr(),
        reportTitle: `Deplacement - ${todayStr()}`,
        columns, rows,
        fileName: `Rapport_Deplacement_${todayStr()}.pdf`
      });
    }catch(e:any){ }
    finally{ setPdfBusy(false); }
  }

  return (
    <div className="space-y-4"><h2 className="text-[24px] font-extrabold flex items-center gap-2">Déplacement <span className="text-[11px] bg-emerald-50 border border-emerald-200 text-emerald-700 px-2 py-0.5 rounded-full">PDF offline</span></h2>
      <div className="card p-4 space-y-3"><div className="flex flex-wrap gap-2"><select value={src} onChange={e=>{setSrc(e.target.value); setChecked({});}} className="px-3 py-2 rounded-full border text-sm"><option value="">Source</option>{classes.map((c:any)=><option key={c.id} value={c.id}>{c.value.name}</option>)}</select><span className="self-center">→</span><select value={dest} onChange={e=>setDest(e.target.value)} className="px-3 py-2 rounded-full border text-sm"><option value="">Accueil</option>{destOptions.map((c:any)=><option key={c.id} value={c.id}>{c.value.name}</option>)}</select><input value={searchRaw} onChange={e=>setSearchRaw(e.target.value)} placeholder="Rechercher debounced" className="flex-1 px-3 py-2 rounded-full border text-sm"/><button disabled={pdfBusy} onClick={handlePdf} className="ml-auto px-3 py-2 rounded-full bg-slate-900 text-white text-xs font-bold flex items-center gap-1">{pdfBusy?<span className="spinner"></span>:null}📄 PDF Deplacements</button></div><div className="max-h-[360px] overflow-auto border rounded-2xl divide-y">{srcList.map((el:any)=><label key={el.id} className={`flex gap-2 p-2 ${checked[el.id]?'bg-amber-50':''}`}><input type="checkbox" checked={!!checked[el.id]} onChange={()=>setChecked(p=>{const cp={...p}; if(cp[el.id]) delete cp[el.id]; else cp[el.id]=true; return cp;})} className="w-4 h-4"/><span className="text-sm">{el.value.numero}. {el.value.nom}</span></label>)}</div><button onClick={move} className="w-full py-2.5 rounded-full primary-bg text-white font-bold">Valider déplacement ({Object.keys(checked).length})</button></div>
    </div>
  );
}
function PermissionTab({classes,elevesByClasse,perms,setPerms,settings}:any){
  const [type,setType]=useState<"Apprenant"|"Parent">("Apprenant"); const [lien,setLien]=useState("père"); const [num,setNum]=useState(""); const [classeId,setClasseId]=useState(""); const [checked,setChecked]=useState<string>(""); const [datePerm,setDatePerm]=useState(new Date().toISOString().slice(0,10)); const [dateRetour,setDateRetour]=useState(""); const [motif,setMotif]=useState(settings.motifsPerm?.[0]||"Maladie"); const [searchRaw,setSearchRaw]=useState(""); const search=useDebounced(searchRaw,300);
  const [pdfBusy,setPdfBusy]=useState(false);
  useEffect(()=>{ if(classes.length>0&&!classeId) setClasseId(classes[0].id); },[classes]);
  const list=useMemo(()=> (classeId?elevesByClasse[classeId]||[]:[]).filter((el:any)=>el.value.nom.toLowerCase().includes(search.toLowerCase())).slice(0,30), [classeId,elevesByClasse,search]);
  async function save(){ if(!checked) return; const rec=await api("records.create",{collection:COLLECTIONS.permissions,value:{type,lien: type==="Parent"?lien:"",numeroParent: type==="Parent"?num:"",classeId,eleveId:checked,datePerm,dateRetour,motif}}); setPerms((p:any)=>[...p,{id:rec.id,value:{type,lien,numeroParent:num,classeId,eleveId:checked,datePerm,dateRetour,motif}}]); setChecked(""); }

  async function handlePdf(){
    setPdfBusy(true);
    try{
      const columns=["Date perm","Retour","Classe","Eleve","Motif","Demandeur","Lien parente"];
      const rows: string[][]=[];
      for(const p of perms){
        const classeName = classes.find((c:any)=>c.id===p.value.classeId)?.value.name||"";
        const el = (Object.values(elevesByClasse).flat() as any).find((e:any)=>e.id===p.value.eleveId);
        rows.push([p.value.datePerm||"", p.value.dateRetour||"-", classeName, el?.value.nom||p.value.eleveId||"", p.value.motif||"", p.value.type||"", p.value.lien||"-"]);
      }
      if(rows.length===0) rows.push([todayStr(),"-","-","Aucune donnee","-","-","-"]);
      await generatePdfReport({
        collegeName: settings.collegeNom || "Collège Privé Saint Laurent",
        logoDataUrl: settings.logo,
        anneeActive: settings.anneeActive || "2026-2027",
        dateJour: todayStr(),
        reportTitle: `Permission - ${todayStr()}`,
        columns, rows,
        fileName: `Rapport_Permission_${todayStr()}.pdf`
      });
    }catch(e:any){ }
    finally{ setPdfBusy(false); }
  }

  return (
    <div className="space-y-4"><h2 className="text-[24px] font-extrabold flex items-center gap-2">Permission <span className="text-[11px] bg-emerald-50 border border-emerald-200 text-emerald-700 px-2 py-0.5 rounded-full">PDF offline</span></h2>
      <div className="grid lg:grid-cols-2 gap-4"><div className="card p-4 space-y-3"><div className="flex gap-2"><button onClick={()=>setType("Apprenant")} className={`px-4 py-2 rounded-full text-sm font-bold border ${type==="Apprenant"?"primary-bg text-white":"bg-white"}`}>Apprenant</button><button onClick={()=>setType("Parent")} className={`px-4 py-2 rounded-full text-sm font-bold border ${type==="Parent"?"primary-bg text-white":"bg-white"}`}>Parent</button></div>{type==="Parent" && <div className="flex gap-2"><select value={lien} onChange={e=>setLien(e.target.value)} className="px-3 py-2 rounded-full border text-sm"><option>père</option><option>mère</option><option>frère/soeur</option><option>tuteur</option></select><input value={num} onChange={e=>setNum(e.target.value)} placeholder="Numéro parent" className="flex-1 px-3 py-2 rounded-full border text-sm"/></div>}<div className="flex flex-wrap gap-2"><select value={classeId} onChange={e=>setClasseId(e.target.value)} className="px-3 py-2 rounded-full border text-sm"><option value="">Classe</option>{classes.map((c:any)=><option key={c.id} value={c.id}>{c.value.name}</option>)}</select><input value={searchRaw} onChange={e=>setSearchRaw(e.target.value)} placeholder="Rechercher debounced" className="flex-1 px-3 py-2 rounded-full border text-sm"/><input type="date" value={datePerm} onChange={e=>setDatePerm(e.target.value)} className="px-3 py-2 rounded-full border text-sm"/><input type="date" value={dateRetour} onChange={e=>setDateRetour(e.target.value)} className="px-3 py-2 rounded-full border text-sm"/><select value={motif} onChange={e=>setMotif(e.target.value)} className="px-3 py-2 rounded-full border text-sm w-full">{settings.motifsPerm.map((m:string)=><option key={m}>{m}</option>)}</select></div><div className="max-h-[300px] overflow-auto border rounded-2xl divide-y">{list.map((el:any)=><label key={el.id} className={`flex gap-2 p-2 ${checked===el.id?'bg-amber-50':''}`}><input type="radio" name="perm" checked={checked===el.id} onChange={()=>setChecked(el.id)} className="w-4 h-4"/><span className="text-sm">{el.value.numero}. {el.value.nom}</span></label>)}</div><button onClick={save} className="w-full py-2.5 rounded-full primary-bg text-white font-bold">Enregistrer permission</button></div><div className="card p-4"><div className="font-bold mb-2 flex items-center justify-between">Historique <button disabled={pdfBusy} onClick={handlePdf} className="px-3 py-1.5 rounded-full bg-slate-900 text-white text-[11px] font-bold flex items-center gap-1">{pdfBusy?<span className="spinner"></span>:null}📄 PDF</button></div><div className="overflow-auto max-h-[500px] border rounded-xl text-xs"><table className="w-full"><thead className="bg-slate-50 sticky top-0"><tr><th className="p-2 text-left">Date perm</th><th className="p-2">Retour</th><th className="p-2">Type</th><th className="p-2">Classe</th><th className="p-2">Motif</th></tr></thead><tbody>{perms.slice(0,100).sort((a:any,b:any)=>b.value.datePerm.localeCompare(a.value.datePerm)).map((p:any)=><tr key={p.id} className="border-t"><td className="p-2">{p.value.datePerm}</td><td className="p-2">{p.value.dateRetour||"—"}</td><td className="p-2">{p.value.type}</td><td className="p-2">{classes.find((c:any)=>c.id===p.value.classeId)?.value.name}</td><td className="p-2">{p.value.motif}</td></tr>)}</tbody></table></div></div></div>
    </div>
  );
}
function BilanTab({classes,elevesByClasse,survs,retards,absTD,absCours,perms,ensAbs,settings,depls}:any){
  const [bilanTab,setBilanTab]=useState(0); const B_TABS=["Classe","Retard","Permission","Absence Cours","Absence TD","Surveillance","Enseignant"];
  const [debut,setDebut]=useState(new Date().toISOString().slice(0,10)); const [fin,setFin]=useState(new Date().toISOString().slice(0,10)); const [classeId,setClasseId]=useState(""); const [searchRaw,setSearchRaw]=useState(""); const search=useDebounced(searchRaw,300); const [pourquoi,setPourquoi]=useState("");
  const [pdfBusy,setPdfBusy]=useState(false);
  function inPeriod(d:string){ return d>=debut && d<=fin; }
  const data = useMemo(()=>{
    switch(B_TABS[bilanTab]){
      case "Retard": return retards.filter((r:any)=>inPeriod(r.value.date) && (!classeId||r.value.classeId===classeId));
      case "Permission": return perms.filter((r:any)=>inPeriod(r.value.datePerm) && (!classeId||r.value.classeId===classeId));
      case "Absence Cours": return absCours.filter((r:any)=>inPeriod(r.value.date) && (!classeId||r.value.classeId===classeId));
      case "Absence TD": return absTD.filter((r:any)=>inPeriod(r.value.date) && (!classeId||r.value.classeId===classeId));
      case "Surveillance": return survs.filter((r:any)=>inPeriod(r.value.date) && (!classeId||r.value.classeId===classeId));
      case "Enseignant": return ensAbs.filter((r:any)=>inPeriod(r.value.date) && (!classeId||r.value.classeId===classeId));
      default: return classes.filter((c:any)=>!classeId||c.id===classeId);
    }
  },[bilanTab,retards,perms,absCours,absTD,survs,ensAbs,classes,classeId,debut,fin]);

  async function downloadPdf(){
    setPdfBusy(true);
    try{
      const periode = `${debut} au ${fin}`;
      const titleBase = `Bilan ${B_TABS[bilanTab]} - ${periode}`;
      let columns: string[] = [];
      let rows: string[][] = [];
      const allElevesFlat = Object.values(elevesByClasse).flat() as any[];
      if(B_TABS[bilanTab]==="Classe"){
        columns = ["Classe","Effectif","Garcons","Filles"];
        (data as any[]).forEach((c:any)=>{
          const list = elevesByClasse[c.id]||[];
          const g = list.filter((e:any)=>e.value.sexe==="M").length;
          const f = list.filter((e:any)=>e.value.sexe==="F").length;
          if(!search || c.value.name.toLowerCase().includes(search.toLowerCase()) || String(list.length).includes(search)){
            rows.push([c.value.name, String(list.length), String(g), String(f)]);
          }
        });
      } else if(B_TABS[bilanTab]==="Retard"){
        columns=["Date","Periode","Classe","Eleve","Matin/Soir"];
        (data as any[]).forEach((r:any)=>{
          const classeName = classes.find((c:any)=>c.id===r.value.classeId)?.value.name||"";
          (r.value.eleveIds||[]).forEach((eid:string)=>{
            const el = allElevesFlat.find((e:any)=>e.id===eid);
            if(!search || (el?.value.nom||"").toLowerCase().includes(search.toLowerCase())){
              rows.push([r.value.date||"", r.value.period||"", classeName, el?.value.nom||eid, r.value.period||""]);
            }
          });
        });
      } else if(B_TABS[bilanTab]==="Permission"){
        columns=["Date perm","Retour","Classe","Eleve","Motif","Demandeur","Lien"];
        (data as any[]).forEach((p:any)=>{
          const classeName = classes.find((c:any)=>c.id===p.value.classeId)?.value.name||"";
          const el = allElevesFlat.find((e:any)=>e.id===p.value.eleveId);
          if(!search || JSON.stringify(p.value).toLowerCase().includes(search.toLowerCase())){
            rows.push([p.value.datePerm||"", p.value.dateRetour||"-", classeName, el?.value.nom||"", p.value.motif||"", p.value.type||"", p.value.lien||"-"]);
          }
        });
      } else if(B_TABS[bilanTab]==="Absence Cours"){
        columns=["Date","Classe","Eleve","Heures","Total h"];
        (data as any[]).forEach((r:any)=>{
          const classeName = classes.find((c:any)=>c.id===r.value.classeId)?.value.name||"";
          (r.value.absences||[]).forEach((ab:any)=>{
            const el = allElevesFlat.find((e:any)=>e.id===ab.eleveId);
            rows.push([r.value.date||"", classeName, el?.value.nom||ab.eleveId, (ab.heures||[]).join(", "), String((ab.heures||[]).length)]);
          });
        });
      } else if(B_TABS[bilanTab]==="Absence TD"){
        columns=["Date","Classe","Eleve","Matiere"];
        (data as any[]).forEach((r:any)=>{
          const classeName = classes.find((c:any)=>c.id===r.value.classeId)?.value.name||"";
          (r.value.eleveIds||[]).forEach((eid:string)=>{
            const el = allElevesFlat.find((e:any)=>e.id===eid);
            rows.push([r.value.date||"", classeName, el?.value.nom||eid, r.value.matiere||""]);
          });
        });
      } else if(B_TABS[bilanTab]==="Surveillance"){
        columns=["Date","Heure","Classe","Eleve","Motif","Matiere","Enseignant","Contact"];
        (data as any[]).forEach((s:any)=>{
          const classeName = classes.find((c:any)=>c.id===s.value.classeId)?.value.name||"";
          (s.value.eleveIds||[]).forEach((eid:string)=>{
            const el = allElevesFlat.find((e:any)=>e.id===eid);
            rows.push([s.value.date||"", s.value.heure||"", classeName, el?.value.nom||eid, s.value.motif||"", s.value.matiere||"", s.value.enseignant||"", s.value.contacts?.[eid]?.contact||el?.value.contactParent||""]);
          });
        });
      } else if(B_TABS[bilanTab]==="Enseignant"){
        columns=["Date","Classe","Matiere","Heure","Enseignant"];
        (data as any[]).forEach((r:any)=>{
          const classeName = classes.find((c:any)=>c.id===r.value.classeId)?.value.name||"";
          rows.push([r.value.date||"", classeName, r.value.matiere||"", r.value.heure||"", r.value.enseignant||""]);
        });
      }
      if(rows.length===0) rows.push(["-","Aucune donnee pour cette periode","","",""]);
      if(B_TABS[bilanTab]==="Classe"){
        const total = rows.reduce((a,r)=>a+ (parseInt(r[1])||0),0);
        rows.push(["TOTAL", String(total), "", ""]);
      }
      await generatePdfReport({
        collegeName: settings.collegeNom || "Collège Privé Saint Laurent",
        logoDataUrl: settings.logo,
        anneeActive: settings.anneeActive || "2026-2027",
        dateJour: todayStr(),
        periode,
        pourquoi: pourquoi || "",
        reportTitle: titleBase + (classeId ? ` - ${classes.find((c:any)=>c.id===classeId)?.value.name||""}` : ""),
        columns, rows,
        fileName: `Rapport_Bilan_${B_TABS[bilanTab]}_${debut}_${fin}.pdf`
      });
    }catch(e:any){
      console.error(e);
    }finally{ setPdfBusy(false); }
  }

  return (
    <div className="space-y-4"><h2 className="text-[24px] font-extrabold flex items-center gap-2">Bilan <span className="text-[11px] bg-emerald-50 border border-emerald-200 text-emerald-700 px-2 py-0.5 rounded-full">PDF offline partout</span></h2>
      <div className="card p-4"><div className="flex gap-1 overflow-x-auto no-scrollbar mb-3">{B_TABS.map((t,i)=><button key={t} onClick={()=>setBilanTab(i)} className={`whitespace-nowrap px-3 py-1.5 rounded-full text-xs font-bold border ${bilanTab===i?'primary-bg text-white border-transparent':'bg-white'}`}>{t}</button>)}</div>
        <div className="flex flex-wrap gap-2 items-end"><div><label className="text-[11px] font-bold uppercase text-slate-500">Début</label><input type="date" value={debut} onChange={e=>setDebut(e.target.value)} className="block px-3 py-2 rounded-full border text-sm"/></div><div><label className="text-[11px] font-bold uppercase text-slate-500">Fin</label><input type="date" value={fin} onChange={e=>setFin(e.target.value)} className="block px-3 py-2 rounded-full border text-sm"/></div><div><label className="text-[11px] font-bold uppercase text-slate-500">Classe</label><select value={classeId} onChange={e=>setClasseId(e.target.value)} className="block px-3 py-2 rounded-full border text-sm"><option value="">Toutes</option>{classes.map((c:any)=><option key={c.id} value={c.id}>{c.value.name}</option>)}</select></div><div className="flex-1 min-w-[140px]"><label className="text-[11px] font-bold uppercase text-slate-500">Recherche debounced</label><input value={searchRaw} onChange={e=>setSearchRaw(e.target.value)} placeholder="Filtrer..." className="block w-full px-3 py-2 rounded-full border text-sm"/></div><div className="flex-1 min-w-[200px]"><label className="text-[11px] font-bold uppercase text-slate-500">Pourquoi</label><input value={pourquoi} onChange={e=>setPourquoi(e.target.value)} placeholder="Ex: Conseil..." className="block w-full px-3 py-2 rounded-full border text-sm"/></div><button disabled={pdfBusy} onClick={downloadPdf} className="px-4 py-2.5 rounded-full primary-bg text-white text-sm font-bold flex items-center gap-2">{pdfBusy?<span className="spinner"></span>:null}📄 Télécharger PDF</button></div>
        <div className="mt-4 border rounded-xl overflow-auto max-h-[420px]"><div className="p-3 bg-slate-50 border-b flex items-center gap-3">{settings.logo && <img src={settings.logo} className="w-8 h-8 rounded object-cover border" alt="logo"/>}<div className="font-bold text-sm">{settings.collegeNom || "Collège Privé Saint Laurent"} — {B_TABS[bilanTab]} • {debut} → {fin} • {data.length} lignes • Année {settings.anneeActive}</div></div><div className="p-2 text-xs">{B_TABS[bilanTab]==="Classe" ? <table className="w-full"><thead><tr><th className="text-left p-1">Classe</th><th className="p-1">Effectif</th></tr></thead><tbody>{(data as any[]).map((c:any)=><tr key={c.id} className="border-t"><td className="p-1">{c.value.name}</td><td className="p-1">{elevesByClasse[c.id]?.length||0}</td></tr>)}</tbody></table> : <table className="w-full"><tbody>{(data as any[]).slice(0,100).filter((r:any)=> !search || JSON.stringify(r.value).toLowerCase().includes(search.toLowerCase())).map((r:any)=><tr key={r.id} className="border-t"><td className="p-1">{JSON.stringify(r.value).slice(0,200)}</td></tr>)}</tbody></table>}</div></div>
      </div>
    </div>
  );
}
function EleveTab({eleves,classes,survs,retards,absTD,absCours,perms,depls,settings,ensAbs}:any){
  const [qRaw,setQRaw]=useState(""); const q=useDebounced(qRaw,300); const [selected,setSelected]=useState<string|null>(null);
  const [pdfBusy,setPdfBusy]=useState(false);
  const matched = useMemo(()=>{ if(!q.trim()) return eleves.slice(0,20); const s=q.toLowerCase(); return eleves.filter((e:any)=> e.value.nom.toLowerCase().includes(s)).slice(0,30); },[q,eleves]);
  const sel = useMemo(()=> eleves.find((e:any)=>e.id===selected) as any, [eleves,selected]);

  async function downloadRapport(){
    if(!sel) return;
    setPdfBusy(true);
    try{
      const classeName = classes.find((c:any)=>c.id===sel.value.classeId)?.value.name||"";
      const retardCount = retards.filter((r:any)=>r.value.eleveIds?.includes(sel.id)).length;
      const survCount = survs.filter((s:any)=>s.value.eleveIds?.includes(sel.id)).length;
      const permCount = perms.filter((p:any)=>p.value.eleveId===sel.id).length;
      const absCoursCount = absCours.filter((a:any)=>a.value.absences?.some((ab:any)=>ab.eleveId===sel.id)).length;
      const absTDCount = absTD.filter((a:any)=>a.value.eleveIds?.includes(sel.id)).length;
      const deplCount = depls.filter((d:any)=>d.value.eleveIds?.includes(sel.id)).length;
      const columns=["Categorie","Details","Nombre"];
      const rows: string[][]=[
        ["Eleve", sel.value.nom, ""],
        ["Classe", classeName, ""],
        ["Sexe", sel.value.sexe||"", ""],
        ["Contact Parent", sel.value.contactParent||"-", ""],
        ["Retards", `Total fiches retards contenant l'eleve`, String(retardCount)],
        ["Surveillances", `Total surveillances`, String(survCount)],
        ["Absences Cours", `Nombre de jours avec absence`, String(absCoursCount)],
        ["Absences TD", `Nombre de fiches TD`, String(absTDCount)],
        ["Permissions", `Permissions demandees`, String(permCount)],
        ["Deplacements", `Deplacements de classe`, String(deplCount)],
      ];
      const details: string[][]=[];
      retards.filter((r:any)=>r.value.eleveIds?.includes(sel.id)).slice(0,10).forEach((r:any)=> details.push(["Retard", `${r.value.date} ${r.value.period}`, ""]));
      survs.filter((s:any)=>s.value.eleveIds?.includes(sel.id)).slice(0,10).forEach((s:any)=> details.push(["Surveillance", `${s.value.date} ${s.value.motif} ${s.value.matiere}`, ""]));
      perms.filter((p:any)=>p.value.eleveId===sel.id).slice(0,10).forEach((p:any)=> details.push(["Permission", `${p.value.datePerm} -> ${p.value.dateRetour} ${p.value.motif}`, ""]));

      const allRows = [...rows, ["---","---","---"], ...details];

      await generatePdfReport({
        collegeName: settings.collegeNom || "Collège Privé Saint Laurent",
        logoDataUrl: settings.logo,
        anneeActive: settings.anneeActive || "2026-2027",
        dateJour: todayStr(),
        reportTitle: `Fiche Eleve - ${sel.value.nom}`,
        columns, rows: allRows,
        fileName: `Rapport_Fiche_${safeName(sel.value.nom)}.pdf`
      });
    }catch(e){ console.error(e); }
    finally{ setPdfBusy(false); }
  }

  return (
    <div className="space-y-4"><h2 className="text-[24px] font-extrabold flex items-center gap-2">Élève <span className="text-[11px] bg-emerald-50 border border-emerald-200 text-emerald-700 px-2 py-0.5 rounded-full">PDF offline</span></h2>
      <div className="grid lg:grid-cols-[0.9fr_1.1fr] gap-4"><div className="card p-4"><input value={qRaw} onChange={e=>setQRaw(e.target.value)} placeholder="Rechercher fuzzy (debounced 300ms)" className="w-full px-4 py-2.5 rounded-full border text-sm"/><div className="mt-3 max-h-[500px] overflow-auto divide-y border rounded-2xl">{matched.map((el:any)=><button key={el.id} onClick={()=>setSelected(el.id)} className={`w-full text-left p-3 flex justify-between ${selected===el.id?'bg-slate-900 text-white':''}`}><span className="text-sm font-medium">{el.value.nom}</span><span className="text-xs opacity-70">{classes.find((c:any)=>c.id===el.value.classeId)?.value.name}</span></button>)}</div></div><div className="card p-4 min-h-[400px]">{!sel ? <div className="py-20 text-center text-sm text-slate-400">Sélectionnez un élève</div> : <div className="space-y-3"><div className="flex justify-between"><div><div className="text-[20px] font-extrabold">{sel.value.nom}</div><div className="text-sm text-slate-500">Classe {classes.find((c:any)=>c.id===sel.value.classeId)?.value.name} • {sel.value.sexe}</div></div><button disabled={pdfBusy} onClick={downloadRapport} className="h-fit px-4 py-2 rounded-full primary-bg text-white text-sm font-bold flex items-center gap-2">{pdfBusy?<span className="spinner"></span>:null}📄 Fiche PDF</button></div><div className="grid grid-cols-2 gap-2 text-xs"><div className="border rounded-xl p-2"><div className="font-bold">Retards</div><div>{retards.filter((r:any)=>r.value.eleveIds?.includes(sel.id)).length}</div></div><div className="border rounded-xl p-2"><div className="font-bold">Surveillances</div><div>{survs.filter((s:any)=>s.value.eleveIds?.includes(sel.id)).length}</div></div><div className="border rounded-xl p-2"><div className="font-bold">Permissions</div><div>{perms.filter((p:any)=>p.value.eleveId===sel.id).length}</div></div><div className="border rounded-xl p-2"><div className="font-bold">Abs Cours</div><div>{absCours.filter((a:any)=>a.value.absences?.some((ab:any)=>ab.eleveId===sel.id)).length}</div></div></div></div>}</div></div>
    </div>
  );
}
function EnseignantTab({classes,ensAbs,setEnsAbs,settings}:any){
  const [classeId,setClasseId]=useState(""); const [matiere,setMatiere]=useState(settings.matieres?.[0]||"MATHS"); const [heure,setHeure]=useState("7H"); const [enseignant,setEnseignant]=useState(""); const [date,setDate]=useState(new Date().toISOString().slice(0,10));
  const [pdfBusy,setPdfBusy]=useState(false);
  useEffect(()=>{ if(classes.length>0&&!classeId) setClasseId(classes[0].id); },[classes]);
  async function save(){ if(!enseignant.trim()) return; const rec=await api("records.create",{collection:COLLECTIONS.ensAbs,value:{date,classeId,matiere,heure,enseignant}}); setEnsAbs((p:any)=>[...p,{id:rec.id,value:{date,classeId,matiere,heure,enseignant}}]); setEnseignant(""); }

  async function handlePdf(){
    setPdfBusy(true);
    try{
      const columns=["Date","Classe","Matiere","Heure","Nom enseignant"];
      const rows: string[][]=[];
      for(const r of ensAbs){
        const classeName = classes.find((c:any)=>c.id===r.value.classeId)?.value.name||"";
        rows.push([r.value.date||"", classeName, r.value.matiere||"", r.value.heure||"", r.value.enseignant||""]);
      }
      if(rows.length===0) rows.push([todayStr(),"-","-","-","Aucune donnee"]);
      await generatePdfReport({
        collegeName: settings.collegeNom || "Collège Privé Saint Laurent",
        logoDataUrl: settings.logo,
        anneeActive: settings.anneeActive || "2026-2027",
        dateJour: todayStr(),
        reportTitle: `Enseignants Absents - ${todayStr()}`,
        columns, rows,
        fileName: `Rapport_Enseignants_Abs_${todayStr()}.pdf`
      });
    }catch(e:any){ }
    finally{ setPdfBusy(false); }
  }

  return (
    <div className="space-y-4"><h2 className="text-[24px] font-extrabold flex items-center gap-2">Enseignant <span className="text-[11px] bg-emerald-50 border border-emerald-200 text-emerald-700 px-2 py-0.5 rounded-full">PDF offline</span></h2>
      <div className="grid lg:grid-cols-2 gap-4"><div className="card p-4 space-y-3"><div className="grid grid-cols-2 gap-2"><select value={classeId} onChange={e=>setClasseId(e.target.value)} className="px-3 py-2 rounded-full border text-sm"><option value="">Classe</option>{classes.map((c:any)=><option key={c.id} value={c.id}>{c.value.name}</option>)}</select><select value={matiere} onChange={e=>setMatiere(e.target.value)} className="px-3 py-2 rounded-full border text-sm">{settings.matieres.map((m:string)=><option key={m}>{m}</option>)}</select><select value={heure} onChange={e=>setHeure(e.target.value)} className="px-3 py-2 rounded-full border text-sm">{["7H","8H","9H","10H","11H","15H","16H"].map(h=><option key={h}>{h}</option>)}</select><input type="date" value={date} onChange={e=>setDate(e.target.value)} className="px-3 py-2 rounded-full border text-sm"/><input value={enseignant} onChange={e=>setEnseignant(e.target.value)} placeholder="Nom enseignant libre" className="col-span-2 px-3 py-2 rounded-full border text-sm"/></div><button onClick={save} className="w-full py-2.5 rounded-full primary-bg text-white font-bold">Enregistrer</button></div><div className="card p-4"><div className="font-bold mb-2 flex items-center justify-between">Historique <button disabled={pdfBusy} onClick={handlePdf} className="px-3 py-1.5 rounded-full bg-slate-900 text-white text-[11px] font-bold flex items-center gap-1">{pdfBusy?<span className="spinner"></span>:null}📄 PDF</button></div><div className="overflow-auto max-h-[400px] border rounded-xl text-xs"><table className="w-full"><thead className="bg-slate-50 sticky top-0"><tr><th className="p-2 text-left">Date</th><th className="p-2">Classe</th><th className="p-2">Matière</th><th className="p-2">Heure</th><th className="p-2">Nom</th></tr></thead><tbody>{ensAbs.slice(0,100).sort((a:any,b:any)=>b.value.date.localeCompare(a.value.date)).map((r:any)=><tr key={r.id} className="border-t"><td className="p-2">{r.value.date}</td><td className="p-2">{classes.find((c:any)=>c.id===r.value.classeId)?.value.name}</td><td className="p-2">{r.value.matiere}</td><td className="p-2">{r.value.heure}</td><td className="p-2">{r.value.enseignant}</td></tr>)}</tbody></table></div></div></div>
    </div>
  );
}

/* ================= PARAMÈTRE COMPLET - 12th SECTION + DRIVE BACKUP ================= */
function ParametreTab({settings,saveSettings,classes,eleves,survs,retards,absTD,absCours,perms,ensAbs,depls,reloadCore,driveConfig,saveDriveConfig,performBackup}:any){
  const [anneeInput,setAnneeInput]=useState("");
  const [anneeError,setAnneeError]=useState<string|null>(null);
  const [anneeDeleteConfirm,setAnneeDeleteConfirm]=useState<string|null>(null);

  const [collegeNom,setCollegeNom]=useState(settings.collegeNom || "Collège Privé Saint Laurent");
  const [logoPreview,setLogoPreview]=useState(settings.logo || "");
  const [collegeSaving,setCollegeSaving]=useState(false);

  const [newMatiere,setNewMatiere]=useState("");
  const [editingMatiereIdx,setEditingMatiereIdx]=useState<number|null>(null);
  const [editingMatiereVal,setEditingMatiereVal]=useState("");
  const [deleteMatiereConfirm,setDeleteMatiereConfirm]=useState<number|null>(null);

  const [newMotifS,setNewMotifS]=useState("");
  const [editingMotifSIdx,setEditingMotifSIdx]=useState<number|null>(null);
  const [editingMotifSVal,setEditingMotifSVal]=useState("");
  const [deleteMotifSConfirm,setDeleteMotifSConfirm]=useState<number|null>(null);

  const [newMotifP,setNewMotifP]=useState("");
  const [editingMotifPIdx,setEditingMotifPIdx]=useState<number|null>(null);
  const [editingMotifPVal,setEditingMotifPVal]=useState("");
  const [deleteMotifPConfirm,setDeleteMotifPConfirm]=useState<number|null>(null);

  const [confirmDeleteStep,setConfirmDeleteStep]=useState(0);
  const [importFile,setImportFile]=useState<File|null>(null);
  const [importConfirmData,setImportConfirmData]=useState<any|null>(null);
  const [importingAll,setImportingAll]=useState(false);
  const [pdfBusyAll,setPdfBusyAll]=useState(false);
  const fileImportRef = useRef<HTMLInputElement>(null);

  // Drive backup UI states
  const [showDriveInstructions,setShowDriveInstructions]=useState(false);
  const [driveDossierId,setDriveDossierId]=useState(driveConfig?.dossierId||"");
  const [driveFrequence,setDriveFrequence]=useState(driveConfig?.frequence||"quotidien");
  const [driveActiver,setDriveActiver]=useState(!!driveConfig?.activer);
  const [driveSaving,setDriveSaving]=useState(false);
  const [driveBackupNowBusy,setDriveBackupNowBusy]=useState(false);

  useEffect(()=>{ setCollegeNom(settings.collegeNom || "Collège Privé Saint Laurent"); setLogoPreview(settings.logo || ""); },[settings.collegeNom, settings.logo]);
  useEffect(()=>{ setDriveDossierId(driveConfig?.dossierId||""); setDriveFrequence(driveConfig?.frequence||"quotidien"); setDriveActiver(!!driveConfig?.activer); },[driveConfig]);

  async function handleCreateAnnee(){
    const v = anneeInput.trim();
    if(!v){ setAnneeError("Saisissez une année ex: 2027-2028"); return; }
    if(!/^\d{4}-\d{4}$/.test(v)){ setAnneeError("Format attendu: 2027-2028"); return; }
    const existing = settings.annees || [];
    if(existing.includes(v)){ setAnneeError("Cette année existe déjà"); return; }
    setAnneeError(null);
    const newList = [...existing, v];
    await saveSettings({annees:newList, anneeActive: v});
    setAnneeInput("");
  }
  async function handleActivateAnnee(a:string){
    await saveSettings({anneeActive:a});
  }
  async function handleDeleteAnnee(a:string){
    if(settings.anneeActive===a){ return; }
    if(anneeDeleteConfirm!==a){ setAnneeDeleteConfirm(a); return; }
    const newList = (settings.annees||[]).filter((x:string)=>x!==a);
    await saveSettings({annees:newList});
    setAnneeDeleteConfirm(null);
  }

  function handleLogoUpload(e:React.ChangeEvent<HTMLInputElement>){
    const f = e.target.files?.[0]; if(!f) return;
    if(f.size>2*1024*1024){ return; }
    const reader = new FileReader();
    reader.onload = () => {
      const result = reader.result as string;
      const img = new Image();
      img.onload = () => {
        const canvas = document.createElement("canvas");
        const max=128;
        let w=img.width, h=img.height;
        if(w>max||h>max){ const ratio=Math.min(max/w, max/h); w=Math.round(w*ratio); h=Math.round(h*ratio); }
        canvas.width=w; canvas.height=h;
        const ctx=canvas.getContext("2d");
        if(!ctx){ setLogoPreview(result); return; }
        ctx.fillStyle="#FFFFFF";
        ctx.fillRect(0,0,w,h);
        ctx.drawImage(img,0,0,w,h);
        let compressed = canvas.toDataURL("image/jpeg",0.4);
        if(compressed.length>11000){
          canvas.width=96; canvas.height=96;
          ctx.fillStyle="#FFFFFF";
          ctx.fillRect(0,0,96,96);
          ctx.drawImage(img,0,0,96,96);
          compressed = canvas.toDataURL("image/jpeg",0.3);
        }
        setLogoPreview(compressed);
      };
      img.onerror = () => setLogoPreview(result);
      img.src = result;
    };
    reader.readAsDataURL(f);
  }
  async function handleSaveCollege(){
    if(!collegeNom.trim()){ return; }
    setCollegeSaving(true);
    let logoToSave = logoPreview;
    if(logoToSave && logoToSave.length>11000){
      logoToSave = "";
    }
    await saveSettings({collegeNom: collegeNom.trim(), logo: logoToSave});
    setCollegeSaving(false);
  }

  async function addMatiere(){
    const v = newMatiere.trim().toUpperCase(); if(!v) return;
    if(settings.matieres.includes(v)){ return; }
    await saveSettings({matieres:[...settings.matieres, v]});
    setNewMatiere("");
  }
  async function saveEditMatiere(){
    if(editingMatiereIdx===null) return;
    const v = editingMatiereVal.trim().toUpperCase(); if(!v){ return; }
    const list = [...settings.matieres];
    list[editingMatiereIdx]=v;
    await saveSettings({matieres:list});
    setEditingMatiereIdx(null); setEditingMatiereVal("");
  }
  async function deleteMatiere(idx:number){
    if(deleteMatiereConfirm!==idx){ setDeleteMatiereConfirm(idx); return; }
    const list = settings.matieres.filter((_:any,i:number)=>i!==idx);
    await saveSettings({matieres:list});
    setDeleteMatiereConfirm(null);
  }

  async function addMotifS(){
    const v = newMotifS.trim(); if(!v) return;
    if(settings.motifsSurv.includes(v)){ return; }
    await saveSettings({motifsSurv:[...settings.motifsSurv, v]});
    setNewMotifS("");
  }
  async function saveEditMotifS(){
    if(editingMotifSIdx===null) return;
    const v = editingMotifSVal.trim(); if(!v) return;
    const list = [...settings.motifsSurv];
    list[editingMotifSIdx]=v;
    await saveSettings({motifsSurv:list});
    setEditingMotifSIdx(null); setEditingMotifSVal("");
  }
  async function deleteMotifS(idx:number){
    if(deleteMotifSConfirm!==idx){ setDeleteMotifSConfirm(idx); return; }
    const list = settings.motifsSurv.filter((_:any,i:number)=>i!==idx);
    await saveSettings({motifsSurv:list});
    setDeleteMotifSConfirm(null);
  }

  async function addMotifP(){
    const v = newMotifP.trim(); if(!v) return;
    if(settings.motifsPerm.includes(v)){ return; }
    await saveSettings({motifsPerm:[...settings.motifsPerm, v]});
    setNewMotifP("");
  }
  async function saveEditMotifP(){
    if(editingMotifPIdx===null) return;
    const v = editingMotifPVal.trim(); if(!v) return;
    const list = [...settings.motifsPerm];
    list[editingMotifPIdx]=v;
    await saveSettings({motifsPerm:list});
    setEditingMotifPIdx(null); setEditingMotifPVal("");
  }
  async function deleteMotifP(idx:number){
    if(deleteMotifPConfirm!==idx){ setDeleteMotifPConfirm(idx); return; }
    const list = settings.motifsPerm.filter((_:any,i:number)=>i!==idx);
    await saveSettings({motifsPerm:list});
    setDeleteMotifPConfirm(null);
  }

  async function exportTotal(){
    const payload = {
      version:"dualsync_v12_final",
      exportedAt:new Date().toISOString(),
      collegeNom: settings.collegeNom,
      anneeActive: settings.anneeActive,
      logo: settings.logo ? "[base64 présent]" : "",
      classes: classes.map((c:any)=>c.value),
      eleves: eleves.map((e:any)=>e.value),
      surveillances: survs.map((s:any)=>s.value),
      retards: retards.map((r:any)=>r.value),
      absences_td: absTD.map((r:any)=>r.value),
      absences_cours: absCours.map((r:any)=>r.value),
      deplacements: depls.map((r:any)=>r.value),
      permissions: perms.map((r:any)=>r.value),
      enseignants_abs: ensAbs.map((r:any)=>r.value),
      parametres: {
        collegeNom: settings.collegeNom,
        logo: settings.logo,
        anneeActive: settings.anneeActive,
        annees: settings.annees,
        theme: settings.theme,
        matieres: settings.matieres,
        motifsSurv: settings.motifsSurv,
        motifsPerm: settings.motifsPerm
      },
      drive_backup_config: driveConfig
    };
    const fullRaw = {
      ...payload,
      _raw: {
        classes, eleves, survs, retards, absTD, absCours, depls, perms, ensAbs,
        settings,
        driveConfig
      }
    };
    const blob = new Blob([JSON.stringify(fullRaw, null, 2)], {type:"application/json"});
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href=url;
    a.download=`dualsync_export_total_${settings.anneeActive||"2026-2027"}_${new Date().toISOString().slice(0,10)}.json`;
    a.click();
    URL.revokeObjectURL(url);
  }

  async function exportAllPdf(){
    setPdfBusyAll(true);
    try{
      const columns=["Collection","Nombre","Details"];
      const rows: string[][]=[
        ["Classes", String(classes.length), classes.map((c:any)=>c.value.name).slice(0,5).join(", ")],
        ["Eleves", String(eleves.length), `${eleves.filter((e:any)=>e.value.sexe==="M").length} garcons, ${eleves.filter((e:any)=>e.value.sexe==="F").length} filles`],
        ["Surveillances", String(survs.length), `${survs.reduce((a:any,r:any)=>a+(r.value.eleveIds?.length||0),0)} eleves concernes`],
        ["Retards", String(retards.length), `${retards.reduce((a:any,r:any)=>a+(r.value.eleveIds?.length||0),0)} retards`],
        ["Abs TD", String(absTD.length), ""],
        ["Abs Cours", String(absCours.length), ""],
        ["Deplacements", String(depls.length), ""],
        ["Permissions", String(perms.length), ""],
        ["Enseignants abs", String(ensAbs.length), ""],
        ["Annee active", settings.anneeActive||"", ""],
        ["College", settings.collegeNom||"", ""],
        ["Drive Backup", driveConfig.activer?"Actif":"Inactif", driveConfig.dossierId||"-"],
      ];
      await generatePdfReport({
        collegeName: settings.collegeNom || "Collège Privé Saint Laurent",
        logoDataUrl: settings.logo,
        anneeActive: settings.anneeActive || "2026-2027",
        dateJour: todayStr(),
        reportTitle: `Export Total - ${settings.anneeActive} - ${todayStr()}`,
        columns, rows,
        fileName: `Rapport_Export_Total_${settings.anneeActive}_${todayStr()}.pdf`
      });
    }catch(e:any){ console.error(e); }
    finally{ setPdfBusyAll(false); }
  }

  async function handleImportFileSelect(e:React.ChangeEvent<HTMLInputElement>){
    const f = e.target.files?.[0]; if(!f) return;
    try{
      const text = await f.text();
      const data = JSON.parse(text);
      setImportFile(f);
      setImportConfirmData(data);
    }catch(err:any){
    }
    if(fileImportRef.current) fileImportRef.current.value="";
  }

  async function confirmImportRestore(){
    if(!importConfirmData) return;
    setImportingAll(true);
    try{
      const data = importConfirmData;
      const raw = data._raw || data;

      const getValues = (keyOptions:string[]):any[] => {
        for(const k of keyOptions){
          if(raw[k] && Array.isArray(raw[k]) && raw[k].length>0){
            if(raw[k][0]?.value) return raw[k].map((r:any)=>r.value || r);
            return raw[k];
          }
          if(data[k] && Array.isArray(data[k])){
            if(data[k][0]?.value) return data[k].map((r:any)=>r.value||r);
            return data[k];
          }
        }
        return [];
      };

      const allCols = [COLLECTIONS.classes, COLLECTIONS.eleves, COLLECTIONS.surveillances, COLLECTIONS.retards, COLLECTIONS.absTD, COLLECTIONS.absCours, COLLECTIONS.deplacements, COLLECTIONS.permissions, COLLECTIONS.ensAbs];
      for(const col of allCols){
        const recs = await listAll(col);
        for(const r of recs){
          try{ await api("records.delete",{collection:col,id:r.id}); }catch{}
        }
      }

      const parametresToRestore = raw.parametres || raw.settings || data.parametres || data.settings || null;
      if(parametresToRestore){
        const settingsVal = parametresToRestore.matieres ? parametresToRestore : (parametresToRestore.value || parametresToRestore);
        try{
          const valToSave = {
            collegeNom: settingsVal.collegeNom || "Collège Privé Saint Laurent",
            logo: settingsVal.logo || "",
            anneeActive: settingsVal.anneeActive || "2026-2027",
            annees: settingsVal.annees || ["2026-2027"],
            theme: settingsVal.theme || "#2563eb",
            matieres: settingsVal.matieres || DEFAULT_MATIERES,
            motifsSurv: settingsVal.motifsSurv || settingsVal.motifsSurveillance || DEFAULT_MOTIFS_SURV,
            motifsPerm: settingsVal.motifsPerm || settingsVal.motifsPermission || DEFAULT_MOTIFS_PERM
          };
          await api("records.create",{collection:COLLECTIONS.settings,value:valToSave});
          await putCollegeInfoStable({
            nom: valToSave.collegeNom,
            logoBase64: valToSave.logo,
            anneeActive: valToSave.anneeActive,
            couleur: valToSave.theme
          });
        }catch{}
      }

      const classesVals = getValues(["classes"]);
      for(const v of classesVals){
        const val = v.value ? v.value : v;
        await api("records.create",{collection:COLLECTIONS.classes,value:{name: val.name || val.nom || "Classe", createdAt: val.createdAt || new Date().toISOString()}});
      }
      const newClasses = await listAll(COLLECTIONS.classes);
      const classNameToId:Record<string,string>={};
      newClasses.forEach((c:any)=>{ classNameToId[c.value.name]=c.id; });

      const elevesVals = getValues(["eleves"]);
      for(const v of elevesVals){
        const val = v.value ? v.value : v;
        let classeId = val.classeId;
        if(val.classeName && classNameToId[val.classeName]) classeId = classNameToId[val.classeName];
        if(!classeId || !newClasses.find((c:any)=>c.id===classeId)){
          classeId = newClasses[0]?.id || val.classeId;
        }
        if(!classeId) continue;
        await api("records.create",{collection:COLLECTIONS.eleves,value:{classeId, numero: val.numero || 1, nom: val.nom || val.name || "Élève", sexe: val.sexe || "M", contactParent: val.contactParent || ""}});
      }

      const survVals = getValues(["surveillances","survs"]);
      for(const v of survVals){ const val = v.value ? v.value : v; await api("records.create",{collection:COLLECTIONS.surveillances,value:val}); }
      const retardVals = getValues(["retards"]);
      for(const v of retardVals){ const val = v.value ? v.value : v; await api("records.create",{collection:COLLECTIONS.retards,value:val}); }
      const absTDVals = getValues(["absences_td","absTD","abs_td"]);
      for(const v of absTDVals){ const val = v.value ? v.value : v; await api("records.create",{collection:COLLECTIONS.absTD,value:val}); }
      const absCoursVals = getValues(["absences_cours","absCours","abs_cours"]);
      for(const v of absCoursVals){ const val = v.value ? v.value : v; await api("records.create",{collection:COLLECTIONS.absCours,value:val}); }
      const deplVals = getValues(["deplacements","depls"]);
      for(const v of deplVals){ const val = v.value ? v.value : v; await api("records.create",{collection:COLLECTIONS.deplacements,value:val}); }
      const permVals = getValues(["permissions","perms"]);
      for(const v of permVals){ const val = v.value ? v.value : v; await api("records.create",{collection:COLLECTIONS.permissions,value:val}); }
      const ensVals = getValues(["enseignants_abs","enseignants","ensAbs","ens_abs"]);
      for(const v of ensVals){ const val = v.value ? v.value : v; await api("records.create",{collection:COLLECTIONS.ensAbs,value:val}); }

      // drive config restore
      const driveRestore = data.drive_backup_config || raw.drive_backup_config || data.driveConfig || raw.driveConfig;
      if(driveRestore){
        await putDriveConfigStable(driveRestore.value || driveRestore);
      }

      setImportConfirmData(null);
      setImportFile(null);
      await reloadCore();
    }catch(err:any){
    }finally{
      setImportingAll(false);
    }
  }

  async function deleteAllData(){
    if(confirmDeleteStep===0){ setConfirmDeleteStep(1); return; }
    if(confirmDeleteStep===1){ setConfirmDeleteStep(2); return; }
    const allCols=[COLLECTIONS.classes,COLLECTIONS.eleves,COLLECTIONS.surveillances,COLLECTIONS.retards,COLLECTIONS.absTD,COLLECTIONS.absCours,COLLECTIONS.deplacements,COLLECTIONS.permissions,COLLECTIONS.ensAbs,COLLECTIONS.settings, COLLECTIONS.parametres, COLLECTIONS.driveBackup];
    for(const col of allCols){
      const recs=await listAll(col);
      for(const r of recs){ try{ await api("records.delete",{collection:col,id:r.id}); }catch{} }
    }
    try{ localStorage.removeItem('college_info'); localStorage.removeItem('drive_backup_config'); localStorage.removeItem('dualsync_latest_backup'); }catch{}
    setConfirmDeleteStep(0);
    await reloadCore();
  }

  async function handleSaveDrive(){
    setDriveSaving(true);
    try{
      await saveDriveConfig({
        dossierId: driveDossierId,
        frequence: driveFrequence,
        activer: driveActiver,
        connected: true,
        email: driveConfig.email || "",
        derniereSauvegarde: driveConfig.derniereSauvegarde || ""
      });
    }catch{}
    setDriveSaving(false);
  }

  async function handleBackupNow(){
    setDriveBackupNowBusy(true);
    try{
      await performBackup();
    }catch{}
    setDriveBackupNowBusy(false);
  }

  return (
    <div className="space-y-6">
      <div className="flex items-center justify-between flex-wrap gap-2">
        <h2 className="text-[26px] font-extrabold tracking-tight">Paramètre <span className="ml-2 text-[11px] bg-slate-900 text-white px-2.5 py-1 rounded-full align-middle">12e section • PDF offline bundlé • college_info + drive_backup_config • Vercel ready</span></h2>
        <div className="text-[11px] text-slate-500">Base vide classes/eleves au démarrage • parametres pré-chargés • 12 sections horizontales</div>
      </div>

      <div className="grid lg:grid-cols-2 gap-5 items-start">
        <div className="space-y-5">
          <div className="card p-5">
            <div className="flex items-center gap-2 mb-3">
              <div className="w-8 h-8 rounded-full primary-bg text-white grid place-items-center font-bold text-sm">1</div>
              <div>
                <div className="font-bold text-[15px]">Année scolaire</div>
                <div className="text-[11px] text-slate-500">Active: <span className="font-bold text-slate-900">{settings.anneeActive || "2026-2027"}</span> • Stockée en collection parametres + utilisée partout</div>
              </div>
            </div>

            <div className="bg-slate-50 border rounded-2xl p-3 mb-3 flex items-center justify-between">
              <div className="text-[13px]"><span className="text-slate-500">Année active actuelle:</span> <span className="font-extrabold text-[16px] primary-text">{settings.anneeActive}</span></div>
              <span className="px-2.5 py-1 rounded-full bg-emerald-50 text-emerald-700 border border-emerald-200 text-[11px] font-bold">Utilisée dans header & bilan PDF</span>
            </div>

            <div className="flex gap-2 mb-4">
              <input value={anneeInput} onChange={e=>{setAnneeInput(e.target.value); setAnneeError(null);}} placeholder="Ex: 2027-2028" className={`flex-1 px-4 py-2.5 rounded-full border text-sm ${anneeError?'border-red-300 ring-2 ring-red-100':''}`}/>
              <button onClick={handleCreateAnnee} className="px-5 py-2.5 rounded-full primary-bg text-white text-sm font-bold shadow">Créer année</button>
            </div>
            {anneeError && <div className="mb-3 text-[12px] text-red-600 bg-red-50 border border-red-200 rounded-xl p-2">⚠️ {anneeError}</div>}

            <div className="space-y-2 max-h-[200px] overflow-auto">
              {(settings.annees||[]).map((a:string)=>(
                <div key={a} className={`flex items-center justify-between p-3 rounded-xl border ${settings.anneeActive===a?'bg-[var(--primary-light)] primary-border border-2':'bg-white'}`}>
                  <div className="flex items-center gap-2">
                    <span className={`text-sm font-bold ${settings.anneeActive===a?'primary-text':''}`}>{a}</span>
                    {settings.anneeActive===a && <span className="text-[10px] px-2 py-0.5 rounded-full primary-bg text-white font-bold">Active</span>}
                  </div>
                  <div className="flex gap-1">
                    {settings.anneeActive!==a && <button onClick={()=>handleActivateAnnee(a)} className="px-3 py-1 rounded-full bg-slate-900 text-white text-[11px] font-bold">Activer</button>}
                    <button onClick={()=>handleDeleteAnnee(a)} className={`px-3 py-1 rounded-full text-[11px] font-bold border ${anneeDeleteConfirm===a?'bg-red-600 text-white border-red-600':'bg-white text-red-600 border-red-200'}`}>{anneeDeleteConfirm===a?'Confirmer?':'Supprimer'}</button>
                  </div>
                </div>
              ))}
            </div>
          </div>

          <div className="card p-5">
            <div className="flex items-center gap-2 mb-3">
              <div className="w-8 h-8 rounded-full primary-bg text-white grid place-items-center font-bold text-sm">2</div>
              <div>
                <div className="font-bold text-[15px]">Infos collège - PERSISTANCE FIXÉE</div>
                <div className="text-[11px] text-slate-500">Nom et logo persistés via parametres id=college_info + localStorage + dualsync_settings • PDF header</div>
              </div>
            </div>

            <div className="space-y-3">
              <div>
                <label className="text-[11px] font-bold uppercase text-slate-500">Nom du collège</label>
                <input value={collegeNom} onChange={e=>setCollegeNom(e.target.value)} placeholder="Collège Privé Saint Laurent" className="w-full mt-1 px-4 py-2.5 rounded-full border text-sm"/>
              </div>

              <div>
                <label className="text-[11px] font-bold uppercase text-slate-500">Logo (FileReader DataURL → preview → save DB stable)</label>
                <div className="mt-1 flex gap-3 items-center">
                  <label className="px-4 py-2 rounded-full bg-slate-900 text-white text-xs font-bold cursor-pointer">📁 Choisir logo<input type="file" accept="image/*" onChange={handleLogoUpload} className="hidden"/></label>
                  {logoPreview ? <img src={logoPreview} className="w-12 h-12 rounded-xl object-cover border" alt="logo preview"/> : <div className="w-12 h-12 rounded-xl bg-slate-100 border grid place-items-center text-[10px] text-slate-400">Aucun</div>}
                  {logoPreview && <button onClick={()=>setLogoPreview("")} className="px-3 py-1 rounded-full bg-red-50 text-red-600 border border-red-200 text-[11px] font-bold">Retirer</button>}
                </div>
                <div className="text-[11px] text-slate-400 mt-1">Compressé 128px JPEG 0.4, &lt;11KB, sauvegardé via records_put collection parametres id college_info + localStorage backup.</div>
                {logoPreview && <div className="text-[10px] mt-1 text-slate-500">Taille logo: {(logoPreview.length/1024).toFixed(1)} KB — {logoPreview.length>11000 ? "⚠️ trop grand, sera réduit" : "✅ OK pour sauvegarde stable"}</div>}
              </div>

              <button disabled={collegeSaving} onClick={handleSaveCollege} className="w-full py-2.5 rounded-full primary-bg text-white font-bold text-sm flex items-center justify-center gap-2">{collegeSaving?<span className="spinner"></span>:null}Enregistrer infos collège (fix persistance college_info)</button>

              <div className="p-3 rounded-xl bg-slate-50 border text-[11px]">
                <div className="font-bold mb-1">Aperçu header + PDF entête (logo + nom + année active):</div>
                <div className="flex items-center gap-2 bg-white border rounded-full px-3 py-2">
                  {logoPreview ? <img src={logoPreview} className="w-7 h-7 rounded-lg object-cover border" alt="logo"/> : <div className="w-7 h-7 rounded-lg primary-bg text-white grid place-items-center text-[12px] font-bold">D</div>}
                  <span className="font-bold text-[13px]">{collegeNom || "Collège Privé Saint Laurent"}</span>
                  <span className="ml-auto text-[10px] text-slate-500">Année {settings.anneeActive}</span>
                </div>
              </div>
            </div>
          </div>

          {/* DRIVE BACKUP CARD */}
          <div className="card p-5 border-2 border-blue-100 bg-gradient-to-br from-white to-blue-50/50">
            <div className="flex items-center gap-2 mb-3">
              <div className="w-8 h-8 rounded-full bg-blue-600 text-white grid place-items-center font-bold text-sm">☁️</div>
              <div className="flex-1">
                <div className="font-bold text-[15px]">Sauvegarde auto Google Drive</div>
                <div className="text-[11px] text-slate-500">Via actions endpoint, pas de fetch direct googleapis • records_put id drive_backup_config • debounce 30s</div>
              </div>
              <span className={`px-2.5 py-1 rounded-full text-[10px] font-bold border ${driveConfig.connected?'bg-emerald-50 text-emerald-700 border-emerald-200':'bg-slate-100 text-slate-600'}`}>{driveConfig.connected?'Connecté':'Non connecté'}</span>
            </div>

            <div className="space-y-3">
              <div className="flex gap-2">
                <button onClick={()=>setShowDriveInstructions(true)} className="flex-1 py-2.5 rounded-full bg-slate-900 text-white text-[12px] font-bold">🔗 Connecter — instructions Vercel env vars</button>
                <button onClick={()=>saveDriveConfig({connected:!driveConfig.connected})} className={`px-4 py-2.5 rounded-full border text-[12px] font-bold ${driveConfig.connected?'bg-emerald-50 border-emerald-200 text-emerald-700':'bg-white'}`}>{driveConfig.connected?'Déconnecter':'Activer'}</button>
              </div>

              <div>
                <label className="text-[11px] font-bold uppercase text-slate-500">Dossier ID Google Drive (optionnel)</label>
                <input value={driveDossierId} onChange={e=>setDriveDossierId(e.target.value)} placeholder="Ex: 1a2b3c... (ID du dossier Drive)" className="w-full mt-1 px-4 py-2.5 rounded-full border text-sm bg-white"/>
                <div className="text-[10px] text-slate-400 mt-1">Utilisé côté serveur Vercel via actions endpoint — pas de fetch direct https://www.googleapis.com dans le client.</div>
              </div>

              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="text-[11px] font-bold uppercase text-slate-500">Fréquence</label>
                  <select value={driveFrequence} onChange={e=>setDriveFrequence(e.target.value)} className="w-full mt-1 px-3 py-2.5 rounded-full border text-sm bg-white">
                    <option value="15min">Toutes les 15 min</option>
                    <option value="heure">Chaque heure</option>
                    <option value="quotidien">Quotidien</option>
                    <option value="manuel">Manuel seulement</option>
                  </select>
                </div>
                <div className="flex flex-col">
                  <label className="text-[11px] font-bold uppercase text-slate-500">Activer auto-backup</label>
                  <button onClick={()=>setDriveActiver(!driveActiver)} className={`mt-1 flex-1 px-3 py-2.5 rounded-full border text-sm font-bold transition ${driveActiver?'bg-emerald-600 text-white border-emerald-600':'bg-white border-slate-200'}`}>{driveActiver?'✅ Activé (debounce 30s)':'⭕ Désactivé'}</button>
                </div>
              </div>

              <div className="bg-white border rounded-2xl p-3 text-[12px]">
                <div className="flex justify-between">
                  <span className="text-slate-500 font-bold uppercase text-[10px]">Dernière sauvegarde</span>
                  <span className="font-mono text-[11px]">{driveConfig.derniereSauvegarde ? new Date(driveConfig.derniereSauvegarde).toLocaleString('fr-FR') : "Jamais"}</span>
                </div>
                <div className="mt-2 text-[11px] text-slate-600">
                  Sauvegarde locale: <span className="font-bold">{(() => { try{ const r=localStorage.getItem('dualsync_latest_backup'); if(!r) return "aucune"; const j=JSON.parse(r); return new Date(j.savedAt).toLocaleString('fr-FR'); }catch{ return "aucune"; } })()}</span>
                </div>
              </div>

              <div className="flex gap-2">
                <button disabled={driveSaving} onClick={handleSaveDrive} className="flex-1 py-2.5 rounded-full bg-blue-600 text-white font-bold text-sm flex items-center justify-center gap-2">{driveSaving?<span className="spinner"></span>:null}💾 Enregistrer config Drive</button>
                <button disabled={driveBackupNowBusy} onClick={handleBackupNow} className="px-5 py-2.5 rounded-full bg-slate-900 text-white font-bold text-sm flex items-center justify-center gap-2">{driveBackupNowBusy?<span className="spinner"></span>:null}Sauvegarder maintenant</button>
              </div>

              <div className="text-[10px] text-slate-500 bg-slate-50 border rounded-xl p-2">
                <b>Logique Vercel-safe:</b> putDriveConfigStable() → parametres id drive_backup_config + localStorage drive_backup_config. performBackup() → records.list toutes collections via actions, save localStorage dualsync_latest_backup. Debounce 30s après mutations si activer. Aucun fetch direct googleapis côté client — serveur Vercel lira l'env var GOOGLE_DRIVE_FOLDER_ID et utilisera actions endpoint.
              </div>
            </div>
          </div>

          <div className="card p-5">
            <div className="flex items-center gap-2 mb-3">
              <div className="w-8 h-8 rounded-full primary-bg text-white grid place-items-center font-bold text-sm">3</div>
              <div>
                <div className="font-bold text-[15px]">Thème couleur</div>
                <div className="text-[11px] text-slate-500">6 couleurs • --primary-color • persistant via college_info</div>
              </div>
            </div>

            <div className="grid grid-cols-3 gap-3">
              {THEMES.map(t=>{
                const isSelected = settings.theme===t.val;
                return (
                  <button key={t.val} onClick={()=>saveSettings({theme:t.val})} className={`relative p-3 rounded-2xl border-2 flex flex-col items-center gap-2 transition-all ${isSelected?'border-slate-900 shadow bg-slate-50':'border-slate-200 bg-white hover:border-slate-300'}`}>
                    <div className="w-12 h-12 rounded-full shadow-inner border-2 border-white relative" style={{background:t.val}}>
                      {isSelected && <span className="absolute inset-0 grid place-items-center text-white font-bold text-[16px]">✓</span>}
                    </div>
                    <div className="text-center">
                      <div className="font-bold text-[12px]">{t.name}</div>
                      <div className="text-[10px] text-slate-500 font-mono">{t.val}</div>
                    </div>
                    {isSelected && <div className="absolute -top-1 -right-1 w-5 h-5 bg-slate-900 text-white rounded-full grid place-items-center text-[10px]">✓</div>}
                  </button>
                );
              })}
            </div>
          </div>
        </div>

        <div className="space-y-5">
          <div className="card p-5">
            <div className="flex items-center gap-2 mb-3">
              <div className="w-8 h-8 rounded-full primary-bg text-white grid place-items-center font-bold text-sm">4a</div>
              <div className="font-bold text-[15px]">Matières</div>
              <span className="ml-auto px-2 py-0.5 rounded-full bg-slate-100 text-[11px] font-bold">{settings.matieres?.length||0} pré-chargées</span>
            </div>
            <div className="flex gap-2 mb-3">
              <input value={newMatiere} onChange={e=>setNewMatiere(e.target.value)} placeholder="Nouvelle matière ex: INFORMATIQUE" className="flex-1 px-4 py-2 rounded-full border text-sm"/>
              <button onClick={addMatiere} className="px-4 py-2 rounded-full bg-slate-900 text-white text-xs font-bold">Ajouter</button>
            </div>
            <div className="space-y-1 max-h-[260px] overflow-auto">
              {settings.matieres?.map((m:string,idx:number)=>(
                <div key={`${m}-${idx}`} className="flex items-center gap-2 p-2 rounded-xl border bg-slate-50">
                  {editingMatiereIdx===idx ? (
                    <>
                      <input autoFocus value={editingMatiereVal} onChange={e=>setEditingMatiereVal(e.target.value)} className="flex-1 px-3 py-1 rounded-full border text-sm"/>
                      <button onClick={saveEditMatiere} className="px-3 py-1 rounded-full primary-bg text-white text-[11px] font-bold">Sauver</button>
                      <button onClick={()=>{setEditingMatiereIdx(null); setEditingMatiereVal("");}} className="px-3 py-1 rounded-full bg-slate-200 text-[11px] font-bold">Annuler</button>
                    </>
                  ) : (
                    <>
                      <span className="flex-1 text-[13px] font-medium">{m}</span>
                      <button onClick={()=>{setEditingMatiereIdx(idx); setEditingMatiereVal(m);}} className="px-2.5 py-1 rounded-full bg-white border text-[11px] font-bold">✏️ Edit</button>
                      <button onClick={()=>deleteMatiere(idx)} className={`px-2.5 py-1 rounded-full border text-[11px] font-bold ${deleteMatiereConfirm===idx?'bg-red-600 text-white border-red-600':'bg-white text-red-600 border-red-200'}`}>{deleteMatiereConfirm===idx?'Confirmer?':'🗑️'}</button>
                    </>
                  )}
                </div>
              ))}
            </div>
          </div>

          <div className="card p-5">
            <div className="flex items-center gap-2 mb-3">
              <div className="w-8 h-8 rounded-full primary-bg text-white grid place-items-center font-bold text-sm">4b</div>
              <div className="font-bold text-[15px]">Motifs surveillance</div>
              <span className="ml-auto px-2 py-0.5 rounded-full bg-slate-100 text-[11px] font-bold">{settings.motifsSurv?.length||0} pré-chargés</span>
            </div>
            <div className="flex gap-2 mb-3">
              <input value={newMotifS} onChange={e=>setNewMotifS(e.target.value)} placeholder="Nouveau motif surveillance" className="flex-1 px-4 py-2 rounded-full border text-sm"/>
              <button onClick={addMotifS} className="px-4 py-2 rounded-full bg-slate-900 text-white text-xs font-bold">Ajouter</button>
            </div>
            <div className="space-y-1 max-h-[180px] overflow-auto">
              {settings.motifsSurv?.map((m:string,idx:number)=>(
                <div key={`${m}-${idx}`} className="flex items-center gap-2 p-2 rounded-xl border bg-slate-50">
                  {editingMotifSIdx===idx ? (
                    <>
                      <input autoFocus value={editingMotifSVal} onChange={e=>setEditingMotifSVal(e.target.value)} className="flex-1 px-3 py-1 rounded-full border text-sm"/>
                      <button onClick={saveEditMotifS} className="px-3 py-1 rounded-full primary-bg text-white text-[11px] font-bold">Sauver</button>
                      <button onClick={()=>{setEditingMotifSIdx(null); setEditingMotifSVal("");}} className="px-3 py-1 rounded-full bg-slate-200 text-[11px] font-bold">Annuler</button>
                    </>
                  ) : (
                    <>
                      <span className="flex-1 text-[13px]">{m}</span>
                      <button onClick={()=>{setEditingMotifSIdx(idx); setEditingMotifSVal(m);}} className="px-2.5 py-1 rounded-full bg-white border text-[11px] font-bold">✏️</button>
                      <button onClick={()=>deleteMotifS(idx)} className={`px-2.5 py-1 rounded-full border text-[11px] font-bold ${deleteMotifSConfirm===idx?'bg-red-600 text-white border-red-600':'bg-white text-red-600 border-red-200'}`}>{deleteMotifSConfirm===idx?'Confirmer?':'🗑️'}</button>
                    </>
                  )}
                </div>
              ))}
            </div>
          </div>

          <div className="card p-5">
            <div className="flex items-center gap-2 mb-3">
              <div className="w-8 h-8 rounded-full primary-bg text-white grid place-items-center font-bold text-sm">4c</div>
              <div className="font-bold text-[15px]">Motifs permission</div>
              <span className="ml-auto px-2 py-0.5 rounded-full bg-slate-100 text-[11px] font-bold">{settings.motifsPerm?.length||0} pré-chargés</span>
            </div>
            <div className="flex gap-2 mb-3">
              <input value={newMotifP} onChange={e=>setNewMotifP(e.target.value)} placeholder="Nouveau motif permission" className="flex-1 px-4 py-2 rounded-full border text-sm"/>
              <button onClick={addMotifP} className="px-4 py-2 rounded-full bg-slate-900 text-white text-xs font-bold">Ajouter</button>
            </div>
            <div className="space-y-1 max-h-[180px] overflow-auto">
              {settings.motifsPerm?.map((m:string,idx:number)=>(
                <div key={`${m}-${idx}`} className="flex items-center gap-2 p-2 rounded-xl border bg-slate-50">
                  {editingMotifPIdx===idx ? (
                    <>
                      <input autoFocus value={editingMotifPVal} onChange={e=>setEditingMotifPVal(e.target.value)} className="flex-1 px-3 py-1 rounded-full border text-sm"/>
                      <button onClick={saveEditMotifP} className="px-3 py-1 rounded-full primary-bg text-white text-[11px] font-bold">Sauver</button>
                      <button onClick={()=>{setEditingMotifPIdx(null); setEditingMotifPVal("");}} className="px-3 py-1 rounded-full bg-slate-200 text-[11px] font-bold">Annuler</button>
                    </>
                  ) : (
                    <>
                      <span className="flex-1 text-[13px]">{m}</span>
                      <button onClick={()=>{setEditingMotifPIdx(idx); setEditingMotifPVal(m);}} className="px-2.5 py-1 rounded-full bg-white border text-[11px] font-bold">✏️</button>
                      <button onClick={()=>deleteMotifP(idx)} className={`px-2.5 py-1 rounded-full border text-[11px] font-bold ${deleteMotifPConfirm===idx?'bg-red-600 text-white border-red-600':'bg-white text-red-600 border-red-200'}`}>{deleteMotifPConfirm===idx?'Confirmer?':'🗑️'}</button>
                    </>
                  )}
                </div>
              ))}
            </div>
          </div>

          <div className="card p-5 border-2 border-slate-200">
            <div className="flex items-center gap-2 mb-3">
              <div className="w-8 h-8 rounded-full bg-slate-900 text-white grid place-items-center font-bold text-sm">5</div>
              <div>
                <div className="font-bold text-[15px]">Sécurité données + PDF offline</div>
                <div className="text-[11px] text-slate-500">Export JSON sauvegarde + Export tout en PDF + college_info stable + drive_backup_config</div>
              </div>
            </div>

            <div className="space-y-3">
              <div className="p-3 rounded-xl bg-emerald-50 border border-emerald-200 text-[12px] text-emerald-900">
                <div className="font-bold">💾 Vercel deployable — base vide classes/eleves, parametres pré-chargés</div>
                <div className="mt-1 leading-relaxed">Collection parametres id college_info = nom, logoBase64, anneeActive 2026-2027, couleur #2563eb. Collection parametres id drive_backup_config = dossierId, activer, frequence, derniereSauvegarde. Backup localStorage college_info + drive_backup_config + dualsync_latest_backup. Aucune notification/toast en bas. Tous PDF offline via pdf.ts bundlé (pas de CDN googleapis). Header affiche logo (si existe) + nom + année active. PDFs utilisent ce nom/logo.</div>
              </div>

              <div className="grid grid-cols-1 gap-2">
                <button onClick={exportTotal} className="w-full py-3 rounded-full bg-emerald-600 text-white font-bold text-sm shadow flex items-center justify-center gap-2">📤 Exporter toutes les données (JSON sauvegarde)</button>
                <button disabled={pdfBusyAll} onClick={exportAllPdf} className="w-full py-3 rounded-full bg-slate-900 text-white font-bold text-sm shadow flex items-center justify-center gap-2">{pdfBusyAll?<span className="spinner"></span>:null}📄 Exporter tout en PDF offline (résumé)</button>

                <div className="border-t pt-3">
                  <label className="text-[11px] font-bold uppercase text-slate-500">Importer backup total (JSON)</label>
                  <div className="mt-2 flex gap-2">
                    <input ref={fileImportRef} type="file" accept=".json" onChange={handleImportFileSelect} className="flex-1 text-xs file:mr-2 file:px-3 file:py-1.5 file:rounded-full file:border-0 file:bg-slate-900 file:text-white file:text-xs file:font-bold"/>
                  </div>
                  {importFile && <div className="mt-2 text-[11px] bg-amber-50 border border-amber-200 rounded-xl p-2">Fichier: <b>{importFile.name}</b> prêt à restaurer.</div>}
                  {importConfirmData && (
                    <div className="mt-2 p-3 rounded-xl bg-amber-50 border border-amber-200">
                      <div className="font-bold text-[13px]">Confirmer restauration totale?</div>
                      <div className="text-[11px] mt-1">Cela supprimera les données actuelles et restaurera depuis le backup.</div>
                      <div className="flex gap-2 mt-3">
                        <button onClick={()=>{setImportConfirmData(null); setImportFile(null);}} className="px-4 py-2 rounded-full bg-white border text-xs font-bold">Annuler</button>
                        <button disabled={importingAll} onClick={confirmImportRestore} className="px-4 py-2 rounded-full bg-amber-600 text-white text-xs font-bold flex items-center gap-2">{importingAll?<span className="spinner"></span>:null}Confirmer import total</button>
                      </div>
                    </div>
                  )}
                </div>

                <div className="border-t pt-3">
                  <div className="font-bold text-[13px] mb-2">Zone dangereuse — double confirmation</div>
                  <div className="flex gap-2">
                    <button onClick={deleteAllData} className={`px-4 py-2 rounded-full text-xs font-bold border ${confirmDeleteStep===0?'bg-red-50 text-red-600 border-red-200':confirmDeleteStep===1?'bg-amber-100 text-amber-800 border-amber-300':'bg-red-600 text-white border-red-600'}`}>{confirmDeleteStep===0?'Supprimer toutes les données':confirmDeleteStep===1?'Confirmez encore':"SUPPRIMER DÉFINITIVEMENT"}</button>
                    {confirmDeleteStep>0 && <button onClick={()=>setConfirmDeleteStep(0)} className="px-3 py-2 rounded-full bg-slate-100 text-xs font-bold">Annuler</button>}
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>

      {showDriveInstructions && (
        <div className="fixed inset-0 z-[80] bg-black/50 grid place-items-center p-4">
          <div className="bg-white rounded-[24px] p-6 w-full max-w-[560px] shadow-2xl border max-h-[90vh] overflow-auto">
            <div className="font-bold text-[18px] mb-2">Sauvegarde auto Google Drive — Vercel env vars</div>
            <div className="text-[13px] text-slate-600 space-y-3 leading-relaxed">
              <p><b>Objectif:</b> Connecter Drive sans fetch direct https://www.googleapis.com côté client. Le client utilise uniquement <code>actions</code> endpoint (records_put id drive_backup_config + localStorage). Le serveur Vercel fait l'upload Drive.</p>
              <p><b>Étapes Vercel:</b></p>
              <ol className="list-decimal pl-5 space-y-1">
                <li>Créer projet Google Cloud, activer Drive API</li>
                <li>Créer Service Account, télécharger JSON key</li>
                <li>Partager votre dossier Drive avec l'email du service account</li>
                <li>Dans Vercel → Settings → Environment Variables, ajouter:
                  <div className="mt-2 bg-slate-900 text-white rounded-xl p-3 font-mono text-[11px] space-y-1">
                    <div>GOOGLE_DRIVE_CLIENT_EMAIL=xxx@xxx.iam.gserviceaccount.com</div>
                    <div>GOOGLE_DRIVE_PRIVATE_KEY="-----BEGIN PRIVATE KEY-----\n..."</div>
                    <div>GOOGLE_DRIVE_FOLDER_ID={driveDossierId||"votre dossier ID"}</div>
                    <div>DRIVE_BACKUP_ENABLED=true</div>
                  </div>
                </li>
                <li>Dans votre API Vercel (ex: /api/actions ou /api/drive-backup), lire les collections via actions endpoint, puis upload vers Drive avec googleapis côté serveur uniquement.</li>
                <li>Côté client, ce module sauvegarde déjà en localStorage dualsync_latest_backup toutes les 30s (debounce) si Activer coché, et persiste config en parametres id drive_backup_config.</li>
              </ol>
              <p className="text-[11px] bg-blue-50 border border-blue-200 rounded-xl p-2"><b>Sécurité:</b> Aucun fetch direct googleapis dans le navigateur. Tout passe par <code>fetch("actions")</code> ou votre endpoint Vercel serveur. La carte ci-dessous enregistre dossierId, frequence, activer, derniereSauvegarde via records_put id drive_backup_config.</p>
            </div>
            <div className="flex gap-2 justify-end mt-5">
              <button onClick={()=>setShowDriveInstructions(false)} className="px-5 py-2.5 rounded-full bg-slate-900 text-white text-sm font-bold">Fermer</button>
            </div>
          </div>
        </div>
      )}

      <div className="card p-4 bg-slate-50">
        <div className="text-[12px] leading-relaxed">
          <b>Final Vercel version:</b> Aucune notification/toast en bas. Matières full list pré-chargée (MATHÉMATIQUES, FRANÇAIS, HISTOIRE-GÉOGRAPHIE, PHILOSOPHIE, ANGLAIS, SVT, PCT, EPS, ÉCONOMIE, ESPAGNOL, COMPTABILITÉ, DROITS, ÉTUDE DE CAS, MATHS FINANCIÈRES, MATHS GÉNÉRALES, ANGLAIS COMMERCIAL, ÉCONOMIE GÉNÉRALE, INFORMATIQUE, ÉDUCATION CIVIQUE, TD MATHS, TD PCT, TD SVT, TD FRANÇAIS...). Motifs surveillance & permission pré-chargés. Tous les rapports en PDF offline bundlé via pdf.ts (pas de CDN). Header nom collège + logo 36x36 + année scolaire + date + title + table + footer. College persistence fixed: stable id college_info via records_put + localStorage fallback. Drive auto backup UI: card, Connecter (instructions Vercel env vars), input Dossier ID, toggle Activer, fréquence select, dernière sauvegarde, button Sauvegarder maintenant. Logic: records_put id drive_backup_config, gathers all collections via records_list, saves to localStorage latest-backup, debounced 30s. 12 sections horizontales, Paramètre last. Database empty for classes/eleves at start, parametres pre-loaded. Lazy XLSX + pagination.
        </div>
      </div>
    </div>
  );
}
