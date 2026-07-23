<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Convertisseur Planning FR → MA</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.4.0/exceljs.min.js"></script>
<style>
  :root{
    --navy:#132A45;
    --navy-2:#1C3D63;
    --blue:#2E5395;
    --blue-light:#DCE6F5;
    --paper:#F6F5F1;
    --panel:#FFFFFF;
    --ink:#1B2430;
    --muted:#667085;
    --line:#E1E4EA;
    --fr:#2255A4;
    --ma:#C1272D;
    --ok:#3E8E5A;
    --warn:#B9862F;
    --radius:10px;
  }
  *{box-sizing:border-box;}
  body{
    margin:0; background:var(--paper); color:var(--ink);
    font-family:"Segoe UI", Inter, Arial, sans-serif;
    -webkit-font-smoothing:antialiased;
  }
  header.app{
    background:linear-gradient(120deg, var(--navy) 0%, var(--navy-2) 100%);
    color:#fff; padding:22px 28px 26px;
  }
  header.app h1{
    margin:0; font-size:22px; letter-spacing:.2px; font-weight:700;
  }
  header.app p{margin:6px 0 0; color:#C6D3E8; font-size:13.5px; max-width:760px; line-height:1.5;}
  .tz-badge{
    display:inline-flex; align-items:center; gap:8px; margin-top:14px;
    background:rgba(255,255,255,.08); border:1px solid rgba(255,255,255,.18);
    padding:6px 12px; border-radius:999px; font-size:12.5px; font-weight:600;
  }
  .tz-badge .fr{color:#9FC1FF;}
  .tz-badge .ma{color:#FF9E9E;}
  .tz-badge svg{width:14px;height:14px;}

  main{max-width:1180px; margin:0 auto; padding:26px 20px 80px;}
  .card{
    background:var(--panel); border:1px solid var(--line); border-radius:var(--radius);
    padding:20px 22px; margin-bottom:20px; box-shadow:0 1px 2px rgba(20,30,50,.04);
  }
  .card h2{
    margin:0 0 4px; font-size:15px; text-transform:uppercase; letter-spacing:.6px;
    color:var(--navy-2); display:flex; align-items:center; gap:8px;
  }
  .card h2 .num{
    background:var(--blue); color:#fff; width:22px;height:22px; border-radius:6px;
    display:inline-flex; align-items:center; justify-content:center; font-size:12px;
  }
  .card > .sub{color:var(--muted); font-size:13px; margin:0 0 14px;}

  #dropzone{
    border:2px dashed #B9C6DC; border-radius:var(--radius); padding:26px;
    text-align:center; cursor:pointer; transition:.15s; background:#FAFBFD;
  }
  #dropzone.drag{border-color:var(--blue); background:var(--blue-light);}
  #dropzone p{margin:4px 0; color:var(--muted); font-size:13.5px;}
  #dropzone strong{color:var(--navy-2);}
  #fileInput{display:none;}

  #fileList{margin-top:14px; display:flex; flex-direction:column; gap:6px;}
  .file-chip{
    display:flex; align-items:center; justify-content:space-between;
    background:#F1F3F7; border:1px solid var(--line); border-radius:8px;
    padding:8px 12px; font-size:13px;
  }
  .file-chip .name{font-weight:600; color:var(--navy-2);}
  .file-chip .meta{color:var(--muted); font-size:12px;}
  .file-chip button{
    background:none;border:none;color:#B23A48;cursor:pointer;font-size:12px;font-weight:600;
  }

  .btn{
    display:inline-flex; align-items:center; gap:8px; border:none; cursor:pointer;
    padding:11px 18px; border-radius:8px; font-size:13.5px; font-weight:600;
    background:var(--blue); color:#fff; transition:.15s;
  }
  .btn:hover{background:var(--navy-2);}
  .btn:disabled{background:#B7C1D1; cursor:not-allowed;}
  .btn.secondary{background:#fff; color:var(--blue); border:1.5px solid var(--blue);}
  .btn.secondary:hover{background:var(--blue-light);}
  .btn.export{background:var(--ok);}
  .btn.export:hover{background:#2F6E44;}
  .actions-row{display:flex; gap:10px; margin-top:14px; flex-wrap:wrap;}

  .status-line{font-size:13px; color:var(--muted); margin-top:10px; min-height:18px;}
  .status-line.error{color:#B23A48; font-weight:600;}
  .status-line.ok{color:var(--ok); font-weight:600;}

  table.people{width:100%; border-collapse:collapse; font-size:12.8px;}
  table.people th, table.people td{
    border:1px solid var(--line); padding:6px 8px; text-align:left; vertical-align:middle;
  }
  table.people thead th{
    background:var(--navy-2); color:#fff; font-weight:600; font-size:11.5px;
    text-transform:uppercase; letter-spacing:.3px; position:sticky; top:0;
  }
  table.people tbody tr:nth-child(even){background:#FAFBFD;}
  table.people input[type=text], table.people select{
    width:100%; border:1px solid var(--line); border-radius:5px; padding:4px 6px; font-size:12.5px;
    font-family:inherit; background:#fff;
  }
  table.people .week-head td{
    background:var(--blue-light); font-weight:700; color:var(--navy-2); padding:8px;
  }
  .day-cell{white-space:nowrap; font-variant-numeric:tabular-nums;}
  .day-cell .fr{color:var(--muted); font-size:11px;}
  .day-cell .ma{color:var(--ma); font-weight:700;}
  .day-cell .arrow{color:#B9C1CE; margin:0 2px;}
  .tag-off{color:var(--muted); font-style:italic;}
  .tag-conge{color:var(--ok); font-weight:600;}
  .tag-ferie{color:var(--warn); font-weight:600;}
  .tag-mission{color:#8A7000; font-weight:600;}
  .tag-presse{color:#8A7000; font-weight:600;}
  .role-manager{background:#DCEFD9 !important;}
  .table-scroll{overflow-x:auto; border:1px solid var(--line); border-radius:8px;}

  .legend{display:flex; gap:18px; flex-wrap:wrap; margin-top:12px; font-size:12px; color:var(--muted);}
  .legend span.dot{display:inline-block; width:10px;height:10px;border-radius:3px;margin-right:5px;vertical-align:-1px;}

  footer.note{max-width:1180px;margin:0 auto;padding:0 20px 40px;color:var(--muted);font-size:12px;line-height:1.6;}
  .empty-hint{color:var(--muted); font-size:13px; padding:10px 2px;}
</style>
</head>
<body>

<header class="app">
  <h1>Convertisseur de plannings — Timesquare (FR) → Excel (Maroc)</h1>
  <p>Dépose un ou plusieurs PDF Timesquare — anciens exports « Plannings individuels » ou nouveaux exports « Plannings périodiques » (tableau par tâche). L'outil lit les créneaux, les convertit à l'heure marocaine (HM = HF − 1h) et génère un Excel semaine par semaine.</p>
  <div class="tz-badge">
    <span class="fr">France (UTC+2 été)</span>
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M5 12h14M13 6l6 6-6 6"/></svg>
    <span class="ma">Maroc (UTC+1) · −1h</span>
  </div>
</header>

<main>

  <div class="card">
    <h2><span class="num">1</span>Importer les PDF</h2>
    <p class="sub">Un PDF par personne et par semaine (export Timesquare classique), ou un PDF multi-pages regroupant plusieurs personnes — chaque page est analysée indépendamment.</p>
    <div id="dropzone">
      <p><strong>Cliquer pour choisir</strong> ou glisser-déposer un ou plusieurs fichiers .pdf ici</p>
      <p>Le nom du projet lu dans le PDF (Accessibilité, ou tout autre) n'est jamais utilisé pour filtrer — tous les projets sont supportés.</p>
    </div>
    <input type="file" id="fileInput" accept="application/pdf" multiple>
    <div id="fileList"></div>
    <div class="actions-row">
      <button class="btn" id="analyzeBtn" disabled>Analyser les PDF</button>
      <button class="btn secondary" id="clearBtn">Tout effacer</button>
    </div>
    <div class="status-line" id="status1"></div>
  </div>

  <div class="card" id="reviewCard" style="display:none;">
    <h2><span class="num">2</span>Vérifier &amp; compléter</h2>
    <p class="sub">Les horaires marocains sont calculés automatiquement. Complète la Zone, le Rôle (Collaborateur / Manager), le code TT et le commentaire pour chaque ligne avant l'export — comme dans ton fichier Excel de suivi.</p>
    <div class="table-scroll">
      <table class="people" id="peopleTable">
        <thead>
          <tr>
            <th>Semaine</th>
            <th>Nom</th>
            <th>Zone</th>
            <th>Rôle</th>
            <th>Lun</th><th>Mar</th><th>Mer</th><th>Jeu</th><th>Ven</th><th>Sam</th><th>Dim</th>
            <th>Total</th>
            <th>TT</th>
            <th>Commentaire</th>
          </tr>
        </thead>
        <tbody id="peopleBody"></tbody>
      </table>
    </div>
    <div class="legend">
      <span><span class="dot" style="background:#DCEFD9"></span>Manager</span>
      <span><span class="dot" style="background:#F2DCA6"></span>OFF / Repos</span>
      <span><span class="dot" style="background:#D9EAD3"></span>Congé</span>
      <span><span class="dot" style="background:#D9D9D9"></span>Férié</span>
      <span><span class="dot" style="background:#FFF200"></span>Missionnée / Presse</span>
      <span style="color:var(--muted)">gris = heure France · <span style="color:var(--ma);font-weight:600">rouge = heure Maroc (−1h)</span></span>
    </div>
  </div>

  <div class="card" id="exportCard" style="display:none;">
    <h2><span class="num">3</span>Générer le fichier Excel</h2>
    <p class="sub">Le fichier généré contient deux feuilles : « Plannings (heure Maroc) » (shifts complets) et « Planning pause déjeuner » — dans la mise en page Zones / Collaborateur / D P / FP par jour, comme ton fichier de suivi actuel.</p>
    <div class="actions-row">
      <button class="btn export" id="exportBtn">⭳ Télécharger le fichier Excel (heure Maroc)</button>
    </div>
    <div class="status-line" id="status2"></div>
  </div>

</main>

<footer class="note">
  Fonctionnement : tout le traitement (lecture PDF, conversion, export) se fait dans ton navigateur — aucun fichier n'est envoyé à un serveur.
  Ce convertisseur reconnaît la structure standard des exports « Plannings individuels » (Timesquare v2) : nom, semaine « Du … au … (n°) », lignes de jour « lun. jj/mm/aaaa », statuts Repos / Congé / Férié, et le total « BC : ». Il ignore volontairement le libellé du projet (Accessibilité, etc.) afin de fonctionner pour n'importe quel projet.
</footer>

<script>
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

// ---------- Time helpers ----------
function timeToMinutes(t){ const [h,m]=t.split(':').map(Number); return h*60+m; }
function minutesToTime(m){ m=((m%1440)+1440)%1440; const h=Math.floor(m/60); const mm=m%60; return String(h).padStart(2,'0')+':'+String(mm).padStart(2,'0'); }
function shiftMinus1h(t){ return t ? minutesToTime(timeToMinutes(t)-60) : null; }

const DAY_ABBR = ['lun','mar','mer','jeu','ven','sam','dim'];
const DAY_FULL = {lun:'Lundi',mar:'Mardi',mer:'Mercredi',jeu:'Jeudi',ven:'Vendredi',sam:'Samedi',dim:'Dimanche'};

function frDateToJs(s){
  const [d,m,y] = s.split('/').map(Number);
  return new Date(y, m-1, d);
}
function jsDateToFr(d){
  return String(d.getDate()).padStart(2,'0')+'/'+String(d.getMonth()+1).padStart(2,'0')+'/'+d.getFullYear();
}

// ---------- Core parser ----------
// Note : selon les exports Timesquare, l'en-tête « Du ... au ... (n°) » peut apparaître
// soit deux fois de façon complète pour une même personne (en-tête + ligne répétée),
// soit une seule fois complète suivie d'une répétition TRONQUÉE (ex: « Du lundi 20/07/2026
// (30) au dimanche » sans date de fin ni numéro). Le parseur gère les deux cas.
function parsePersonBlocksFromPage(rawText){
  const text = rawText.replace(/\s+/g,' ').trim();
  const weekRe = /Du\s+([A-Za-zéûèêîôàç]+)\s+(\d{2}\/\d{2}\/\d{4})\s*\((\d{1,2})\)\s*au\s+([A-Za-zéûèêîôàç]+)\s+(\d{2}\/\d{2}\/\d{4})\s*\((\d{1,2})\)/gi;
  const weekMatches = [...text.matchAll(weekRe)];
  if(weekMatches.length===0) return [];

  const dayReGlobal = /(lun|mar|mer|jeu|ven|sam|dim)\.\s*(\d{2}\/\d{2}\/\d{4})/gi;

  const results = [];
  let i = 0;
  while(i < weekMatches.length){
    const first = weekMatches[i];

    // Cherche le premier marqueur de jour après ce match d'en-tête
    dayReGlobal.lastIndex = first.index + first[0].length;
    const dayProbe = dayReGlobal.exec(text);
    const dayStart = dayProbe ? dayProbe.index : text.length;

    // Si un DEUXIÈME match complet de semaine apparaît avant le premier jour,
    // c'est le style "en-tête + ligne répétée complète" : on utilise le 2e comme référence.
    let header = first;
    if(i+1 < weekMatches.length && weekMatches[i+1].index < dayStart){
      header = weekMatches[i+1];
      i++;
    }

    const weekNumber = header[3];
    const weekStart = header[2];
    const weekEnd = header[5];

    // Le nom se trouve entre la fin du 1er match et soit le 2e match complet (style ancien),
    // soit le début des jours (style tronqué) — dans ce dernier cas on retire le fragment
    // tronqué "Du ... " qui traîne dans le nom.
    let namePart = text.substring(first.index+first[0].length, header===first ? dayStart : header.index);
    namePart = namePart.replace(/Plannings individuels/gi,'').trim();
    const duFragIdx = namePart.search(/\bDu\s+[A-Za-zà-ÿ]+\s+\d{2}\/\d{2}\/\d{4}/i);
    if(duFragIdx > -1) namePart = namePart.substring(0, duFragIdx).trim();
    let name = namePart.replace(/^[:\-]\s*/,'').trim();
    if(!name){
      const before = text.substring(0, first.index).replace(/Plannings individuels/gi,'').trim();
      name = before.split(/(?<=[a-zà-ÿ])(?=[A-ZÀ-Ý])/).pop() || before;
    }

    const sectionEnd = (i+1<weekMatches.length) ? weekMatches[i+1].index : text.length;
    const rest = text.substring(header.index+header[0].length, sectionEnd);

    const dayRe = /(lun|mar|mer|jeu|ven|sam|dim)\.\s*(\d{2}\/\d{2}\/\d{4})/g;
    const dayMatches = [...rest.matchAll(dayRe)];
    if(dayMatches.length===0){ i++; continue; }

    const totalRe = /(\d{1,3}:\d{2})\s*BC\s*:/i;
    const days = [];
    for(let i=0;i<dayMatches.length;i++){
      const cur = dayMatches[i];
      const nextIdx = (i+1<dayMatches.length) ? dayMatches[i+1].index : rest.length;
      let content = rest.substring(cur.index+cur[0].length, nextIdx);
      const totalMatch = content.match(totalRe);
      if(totalMatch){ content = content.substring(0, totalMatch.index); }

      const timeRanges = [...content.matchAll(/(\d{2}:\d{2})\s*-\s*(\d{2}:\d{2})/g)];
      let status='inconnu', startFR=null, endFR=null, pauseFR=null;
      if(timeRanges.length>0){
        status='travail';
        const starts = timeRanges.map(r=>timeToMinutes(r[1]));
        const ends = timeRanges.map(r=>timeToMinutes(r[2]));
        startFR = minutesToTime(Math.min(...starts));
        endFR = minutesToTime(Math.max(...ends));
        if(timeRanges.length>=2){ pauseFR = [timeRanges[0][2], timeRanges[1][1]]; }
      } else if(/repos/i.test(content)){ status='repos'; }
        else if(/cong[ée]/i.test(content) || /\bCP\b/.test(content)){ status='conge'; }
        else if(/f[ée]ri[ée]/i.test(content)){ status='ferie'; }
        else if(/mission/i.test(content)){ status='mission'; }
        else if(/presse/i.test(content)){ status='presse'; }
        else if(/off/i.test(content)){ status='repos'; }

      days.push({
        dayAbbr: cur[1].toLowerCase(), date: cur[2], status,
        startFR, endFR, pauseFR,
        startMA: shiftMinus1h(startFR), endMA: shiftMinus1h(endFR),
        pauseMA: pauseFR ? [shiftMinus1h(pauseFR[0]), shiftMinus1h(pauseFR[1])] : null
      });
    }
    const totalM = rest.match(/(\d{1,3}:\d{2})\s*BC\s*:\s*(\d{1,3}:\d{2})/i);
    const totalHours = totalM ? totalM[1] : '';

    if(name && days.length){
      results.push({ id: name+'_'+weekNumber+'_'+weekStart, name, weekNumber, weekStart, weekEnd, days, totalHours, zone:'', role:'Collaborateur', tt:'', comment:'RAS' });
    }
    i++;
  }
  return results;
}

// ---------- Parser for new "Plannings périodiques" table format ----------
// PDF has a table per Tâche, with 7 day columns and rows per personne.
// Uses positional items from pdf.js to reconstruct cells (multi-line shifts).
function parsePeriodicPageFromItems(items, rawText){
  if(!/Plannings\s+périodiques/i.test(rawText)) return [];
  if(/R[ée]capitulatif\s+des\s+(absences|repos)/i.test(rawText)) return [];

  const weekMatch = rawText.match(/Du\s+\S+\s+(\d{2}\/\d{2}\/\d{4})\s*\((\d{1,2})\)\s*au\s+\S+\s+(\d{2}\/\d{2}\/\d{4})/i);
  if(!weekMatch) return [];
  const weekStart = weekMatch[1], weekNumber = weekMatch[2], weekEnd = weekMatch[3];

  // Try to grab task name; it may be split across items so use rawText.
  let taskName = '';
  const taskMatch = rawText.match(/T[âa]che\s*:\s*([^\n]+?)(?=\s{2,}|lun\.|Total|$)/i);
  if(taskMatch) taskName = taskMatch[1].replace(/\s+/g,' ').trim().slice(0,40);

  const positioned = items
    .map(it => ({ str: it.str, x: it.transform[4], y: it.transform[5] }))
    .filter(it => it.str && it.str.trim().length>0);

  // Find items that look like day-column headers (may combine abbr+date in same item, or split).
  // Detect the header row by scanning for the largest y that has >=5 day tokens on it.
  // Combine adjacent items on same y first.
  const byY = {};
  positioned.forEach(it=>{
    const yk = Math.round(it.y);
    (byY[yk] = byY[yk]||[]).push(it);
  });
  const yKeys = Object.keys(byY).map(Number).sort((a,b)=>b-a);

  let headerY = null, dayCols = null;
  for(const yk of yKeys){
    const row = byY[yk].slice().sort((a,b)=>a.x-b.x);
    const rowText = row.map(it=>it.str).join(' ');
    const dayHits = [...rowText.matchAll(/(lun|mar|mer|jeu|ven|sam|dim)\.\s*(\d{2}\/\d{2}\/\d{4})/gi)];
    if(dayHits.length >= 5){
      // Locate the x of each day header by finding the item whose str starts the abbr.
      const cols = [];
      const seenAbbrs = new Set();
      for(const it of row){
        const m = it.str.match(/(lun|mar|mer|jeu|ven|sam|dim)\.?/i);
        if(m && !seenAbbrs.has(m[1].toLowerCase())){
          // Find date near this item
          const combined = row.filter(r=>Math.abs(r.x-it.x)<80 && r.x>=it.x).map(r=>r.str).join(' ');
          const dm = combined.match(/(\d{2}\/\d{2}\/\d{4})/);
          if(dm){
            cols.push({ abbr: m[1].toLowerCase(), date: dm[1], x: it.x });
            seenAbbrs.add(m[1].toLowerCase());
          }
        }
      }
      if(cols.length >= 5){
        headerY = yk;
        dayCols = cols.sort((a,b)=>a.x-b.x);
        break;
      }
    }
  }
  if(!dayCols || dayCols.length < 5) return [];

  // Fill missing days from expected 7-day sequence
  const orderAbbr = ['lun','mar','mer','jeu','ven','sam','dim'];
  if(dayCols.length < 7){
    // Extrapolate x-spacing from existing columns
    const spans = [];
    for(let i=1;i<dayCols.length;i++) spans.push((dayCols[i].x-dayCols[i-1].x)/1);
    const step = spans.reduce((a,b)=>a+b,0)/spans.length;
    const firstIdx = orderAbbr.indexOf(dayCols[0].abbr);
    const start = frDateToJs(dayCols[0].date);
    const rebuilt = [];
    for(let i=0;i<7;i++){
      const d = new Date(start); d.setDate(start.getDate() + (i - firstIdx));
      const existing = dayCols.find(c=>c.abbr===orderAbbr[i]);
      rebuilt.push(existing || { abbr: orderAbbr[i], date: jsDateToFr(d), x: dayCols[0].x + (i-firstIdx)*step });
    }
    dayCols = rebuilt;
  }
  dayCols = dayCols.slice(0,7);

  // Total column x = to the right of last day
  const totalColX = dayCols[6].x + (dayCols[6].x - dayCols[5].x);
  const colCentersX = dayCols.map(c=>c.x);
  const nameColMaxX = dayCols[0].x - 8;

  // Group rows below header by y
  const belowY = yKeys.filter(y=>y < headerY - 2);
  const rowLabels = belowY.map(y=>{
    const leftItems = byY[y].filter(it=>it.x < nameColMaxX);
    const text = leftItems.map(it=>it.str).join(' ').replace(/\s+/g,' ').trim();
    return { y, text };
  });

  const persons = [];
  for(let i=0; i<rowLabels.length; i++){
    const lbl = rowLabels[i];
    if(!lbl.text) continue;
    if(/^Total$/i.test(lbl.text)) continue;
    if(!/[A-Za-zÀ-ÿ]{2,}/.test(lbl.text)) continue;
    // Skip page footers / titles that landed in left col
    if(/Plannings|T[âa]che|Page\s*\d|Timesquare|Holy-Dis|R[ée]capitulatif/i.test(lbl.text)) continue;

    // Find matching Total row
    let totalIdx = -1;
    for(let j=i+1; j<rowLabels.length; j++){
      if(/^Total$/i.test(rowLabels[j].text)){ totalIdx = j; break; }
      if(rowLabels[j].text && /[A-Za-zÀ-ÿ]{2,}/.test(rowLabels[j].text) &&
         !/Plannings|T[âa]che|Page\s*\d|Timesquare|Holy-Dis/i.test(rowLabels[j].text)){
        break;
      }
    }
    if(totalIdx < 0) continue;

    const yTop = lbl.y, yBot = rowLabels[totalIdx].y;
    const blockItems = [];
    for(const yk of belowY){
      if(yk > yBot + 1 && yk <= yTop + 1){
        for(const it of byY[yk]) blockItems.push(it);
      }
    }

    // Assign each item to a day column (or skip if it's in name col or total col)
    const cells = Array.from({length:7}, ()=>[]);
    blockItems.forEach(it=>{
      if(it.x < nameColMaxX) return;
      let best=-1, bestD=Infinity;
      for(let c=0;c<7;c++){
        const d = Math.abs(it.x - colCentersX[c]);
        if(d<bestD){ bestD=d; best=c; }
      }
      const dTotal = Math.abs(it.x - totalColX);
      if(dTotal < bestD) return; // belongs to total column
      cells[best].push(it);
    });

    const days = dayCols.map((dc, idx)=>{
      cells[idx].sort((a,b)=> b.y - a.y);
      const cellText = cells[idx].map(it=>it.str).join(' ').replace(/\s+/g,' ').trim();
      const timeRanges = [...cellText.matchAll(/(\d{2}:\d{2})\s*-\s*(\d{2}:\d{2})/g)];
      let status='inconnu', startFR=null, endFR=null, pauseFR=null;
      if(timeRanges.length>0){
        status='travail';
        const starts = timeRanges.map(r=>timeToMinutes(r[1]));
        const ends = timeRanges.map(r=>timeToMinutes(r[2]));
        startFR = minutesToTime(Math.min(...starts));
        endFR = minutesToTime(Math.max(...ends));
        if(timeRanges.length>=2){ pauseFR = [timeRanges[0][2], timeRanges[1][1]]; }
      } else if(/repos/i.test(cellText)) status='repos';
        else if(/cong[ée]/i.test(cellText) || /\bCP\b/.test(cellText)) status='conge';
        else if(/f[ée]ri[ée]/i.test(cellText)) status='ferie';
        else if(/mission/i.test(cellText)) status='mission';
        else if(/presse/i.test(cellText)) status='presse';
        else if(/\bHS\b/.test(cellText)) status='repos';
        else if(!cellText) status='repos';
      return {
        dayAbbr: dc.abbr, date: dc.date, status,
        startFR, endFR, pauseFR,
        startMA: shiftMinus1h(startFR), endMA: shiftMinus1h(endFR),
        pauseMA: pauseFR ? [shiftMinus1h(pauseFR[0]), shiftMinus1h(pauseFR[1])] : null
      };
    });

    // Weekly total from rightmost item on Total row
    const totalRowItems = (byY[rowLabels[totalIdx].y]||[]).slice().sort((a,b)=>b.x-a.x);
    let totalHours='';
    for(const it of totalRowItems){
      const m = it.str.match(/(\d{1,3}:\d{2})/);
      if(m){ totalHours = m[1]; break; }
    }

    persons.push({
      id: lbl.text + '_' + weekNumber + '_' + weekStart + '_' + taskName,
      name: lbl.text,
      weekNumber, weekStart, weekEnd,
      days, totalHours,
      zone:'', role:'Collaborateur', tt:'',
      comment: taskName || 'RAS'
    });

    i = totalIdx;
  }

  return persons;
}

// Merge duplicate person/week entries coming from multiple task pages of the
// "Plannings périodiques" export: prefer 'travail' days over 'repos' and
// aggregate task names in the comment column.
function mergePeople(list){
  const RANK = { travail:5, mission:4, presse:4, conge:3, ferie:3, repos:2, inconnu:1 };
  const map = new Map();
  for(const p of list){
    const key = p.name + '|' + p.weekStart;
    if(!map.has(key)){ map.set(key, JSON.parse(JSON.stringify(p))); continue; }
    const cur = map.get(key);
    p.days.forEach(d=>{
      const ex = cur.days.find(x=>x.dayAbbr===d.dayAbbr);
      if(!ex){ cur.days.push(d); return; }
      if((RANK[d.status]||0) > (RANK[ex.status]||0)) Object.assign(ex, d);
    });
    if(p.comment && p.comment!=='RAS' && !cur.comment.split(/\s*\+\s*/).includes(p.comment)){
      cur.comment = (cur.comment==='RAS' || !cur.comment) ? p.comment : cur.comment+' + '+p.comment;
    }
    // Keep the larger totalHours (usually identical anyway)
    if(!cur.totalHours && p.totalHours) cur.totalHours = p.totalHours;
  }
  return [...map.values()];
}

// ---------- State ----------
let loadedFiles = []; // {file, pages:[]}
let people = []; // parsed records

const fileInput = document.getElementById('fileInput');
const dropzone = document.getElementById('dropzone');
const fileListEl = document.getElementById('fileList');
const analyzeBtn = document.getElementById('analyzeBtn');
const clearBtn = document.getElementById('clearBtn');
const status1 = document.getElementById('status1');
const status2 = document.getElementById('status2');
const reviewCard = document.getElementById('reviewCard');
const exportCard = document.getElementById('exportCard');
const peopleBody = document.getElementById('peopleBody');

dropzone.addEventListener('click', ()=>fileInput.click());
dropzone.addEventListener('dragover', e=>{e.preventDefault(); dropzone.classList.add('drag');});
dropzone.addEventListener('dragleave', ()=>dropzone.classList.remove('drag'));
dropzone.addEventListener('drop', e=>{
  e.preventDefault(); dropzone.classList.remove('drag');
  handleFiles(e.dataTransfer.files);
});
fileInput.addEventListener('change', e=>handleFiles(e.target.files));

function handleFiles(fileList){
  for(const f of fileList){
    if(f.type==='application/pdf' || f.name.toLowerCase().endsWith('.pdf')){
      loadedFiles.push({file:f});
    }
  }
  renderFileList();
  analyzeBtn.disabled = loadedFiles.length===0;
}

function renderFileList(){
  fileListEl.innerHTML='';
  loadedFiles.forEach((lf, idx)=>{
    const div = document.createElement('div');
    div.className='file-chip';
    div.innerHTML = `<span class="name">📄 ${lf.file.name}</span><span class="meta">${(lf.file.size/1024).toFixed(0)} Ko</span>`;
    const btn = document.createElement('button');
    btn.textContent = 'Retirer';
    btn.onclick = ()=>{ loadedFiles.splice(idx,1); renderFileList(); analyzeBtn.disabled = loadedFiles.length===0; };
    div.appendChild(btn);
    fileListEl.appendChild(div);
  });
}

clearBtn.addEventListener('click', ()=>{
  loadedFiles=[]; people=[];
  fileInput.value='';
  renderFileList();
  analyzeBtn.disabled=true;
  reviewCard.style.display='none';
  exportCard.style.display='none';
  status1.textContent=''; status1.className='status-line';
});

analyzeBtn.addEventListener('click', async ()=>{
  status1.textContent = 'Lecture des PDF en cours…';
  status1.className='status-line';
  analyzeBtn.disabled = true;
  people = [];
  try{
    for(const lf of loadedFiles){
      const buf = await lf.file.arrayBuffer();
      const pdf = await pdfjsLib.getDocument({data:buf}).promise;
      for(let p=1;p<=pdf.numPages;p++){
        const page = await pdf.getPage(p);
        const content = await page.getTextContent();
        const text = content.items.map(it=>it.str).join(' ');
        let parsed = [];
        if(/Plannings\s+périodiques/i.test(text)){
          parsed = parsePeriodicPageFromItems(content.items, text);
        }
        if(!parsed || parsed.length===0){
          parsed = parsePersonBlocksFromPage(text);
        }
        people.push(...parsed);
      }
    }
    // Merge duplicates (same person / same week across multiple task pages)
    people = mergePeople(people);

    if(people.length===0){
      status1.textContent = "Aucun planning reconnu dans ces PDF. Vérifie qu'il s'agit bien d'un export Timesquare (Plannings individuels ou Plannings périodiques).";
      status1.className='status-line error';
    } else {
      status1.textContent = `${people.length} planning(s) individuel(s) détecté(s) et converti(s) à l'heure marocaine.`;
      status1.className='status-line ok';
      renderPeopleTable();
      reviewCard.style.display='block';
      exportCard.style.display='block';
    }
  } catch(err){
    console.error(err);
    status1.textContent = 'Erreur pendant la lecture des PDF : '+err.message;
    status1.className='status-line error';
  }
  analyzeBtn.disabled = false;
});

function dayCellHTML(day){
  if(!day) return '<span class="empty-hint">—</span>';
  if(day.status==='travail'){
    return `<div class="day-cell">
      <div class="fr">${day.startFR}–${day.endFR}</div>
      <div class="ma">${day.startMA}–${day.endMA}</div>
    </div>`;
  }
  if(day.status==='repos') return '<span class="tag-off">OFF</span>';
  if(day.status==='conge') return '<span class="tag-conge">Congé</span>';
  if(day.status==='ferie') return '<span class="tag-ferie">Férié</span>';
  if(day.status==='mission') return '<span class="tag-mission">Missionnée</span>';
  if(day.status==='presse') return '<span class="tag-presse">Presse</span>';
  return '<span class="empty-hint">?</span>';
}

function renderPeopleTable(){
  peopleBody.innerHTML='';
  // group by week for readability
  const weeks = {};
  people.forEach(p=>{
    const key = p.weekNumber+'_'+p.weekStart;
    (weeks[key] = weeks[key]||[]).push(p);
  });
  const weekKeys = Object.keys(weeks).sort((a,b)=>{
    const da = frDateToJs(weeks[a][0].weekStart), db = frDateToJs(weeks[b][0].weekStart);
    return da-db;
  });

  weekKeys.forEach(wk=>{
    const rows = weeks[wk];
    const headTr = document.createElement('tr');
    headTr.className='week-head';
    headTr.innerHTML = `<td colspan="14">Semaine S${rows[0].weekNumber} — du ${rows[0].weekStart} au ${rows[0].weekEnd}</td>`;
    peopleBody.appendChild(headTr);

    rows.forEach(person=>{
      const tr = document.createElement('tr');
      tr.dataset.id = person.id;

      const dayByAbbr = {};
      person.days.forEach(d=>dayByAbbr[d.dayAbbr]=d);

      tr.innerHTML = `
        <td>S${person.weekNumber}</td>
        <td><strong>${person.name}</strong></td>
        <td><input type="text" class="zone-input" placeholder="ex: ZONE 6" value="${person.zone}"></td>
        <td>
          <select class="role-input">
            <option value="Collaborateur" ${person.role==='Collaborateur'?'selected':''}>Collaborateur</option>
            <option value="Manager" ${person.role==='Manager'?'selected':''}>Manager</option>
          </select>
        </td>
        ${DAY_ABBR.map(a=>`<td>${dayCellHTML(dayByAbbr[a])}</td>`).join('')}
        <td>${person.totalHours||'—'}</td>
        <td><input type="text" class="tt-input" style="width:34px" maxlength="2" value="${person.tt}"></td>
        <td><input type="text" class="comment-input" value="${person.comment}"></td>
      `;
      peopleBody.appendChild(tr);

      tr.querySelector('.zone-input').addEventListener('input', e=>person.zone=e.target.value);
      tr.querySelector('.role-input').addEventListener('change', e=>{
        person.role=e.target.value;
        tr.classList.toggle('role-manager', person.role==='Manager');
      });
      tr.querySelector('.tt-input').addEventListener('input', e=>person.tt=e.target.value);
      tr.querySelector('.comment-input').addEventListener('input', e=>person.comment=e.target.value);
      if(person.role==='Manager') tr.classList.add('role-manager');
    });
  });
}

// ---------- Excel export ----------
const COLORS = {
  band: 'FF2E5395',
  bandText: 'FFFFFFFF',
  subHeader: 'FF1C3D63',
  subHeaderText: 'FFFFFFFF',
  weekendHeader: 'FFEDEFF3',
  off: 'FFF2DCA6',
  conge: 'FFD9EAD3',
  ferie: 'FFD9D9D9',
  managerTag: 'FFA9D18E',
  zebra: 'FFF7F8FA',
  border: 'FFB9C1CE',
  pauseTime: 'FFDCE6F5',
  missionPresse: 'FFFFF200',
  trsptKo: 'FFED9B33'
};
function thinBorder(){
  return { top:{style:'thin',color:{argb:COLORS.border}}, left:{style:'thin',color:{argb:COLORS.border}},
           bottom:{style:'thin',color:{argb:COLORS.border}}, right:{style:'thin',color:{argb:COLORS.border}} };
}
function fill(argb){ return { type:'pattern', pattern:'solid', fgColor:{argb} }; }

function mostFrequent(arr){
  if(arr.length===0) return null;
  const counts = {};
  arr.forEach(v=>{ const k=JSON.stringify(v); counts[k]=(counts[k]||0)+1; });
  let best=null,bestCount=-1;
  Object.keys(counts).forEach(k=>{ if(counts[k]>bestCount){bestCount=counts[k]; best=JSON.parse(k);} });
  return best;
}

async function buildWorkbook(){
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Convertisseur Planning FR-MA';
  const ws = wb.addWorksheet('Plannings (heure Maroc)');

  ws.getColumn(1).width = 11;  // role tag
  ws.getColumn(2).width = 10;  // zone
  ws.getColumn(3).width = 22;  // nom
  for(let c=4;c<=17;c++) ws.getColumn(c).width = 9.5; // 7 days x2
  ws.getColumn(18).width = 7;   // OFF count
  ws.getColumn(19).width = 13;  // NB heures
  ws.getColumn(20).width = 6;   // TT
  ws.getColumn(21).width = 16;  // Commentaires

  const weeks = {};
  people.forEach(p=>{
    const key = p.weekNumber+'_'+p.weekStart;
    (weeks[key]=weeks[key]||[]).push(p);
  });
  const weekKeys = Object.keys(weeks).sort((a,b)=>{
    const da=frDateToJs(weeks[a][0].weekStart), db=frDateToJs(weeks[b][0].weekStart);
    return da-db;
  });

  let row = 1;

  function writeGroupTable(members, roleLabel, weekInfo){
    if(members.length===0) return;
    const startRow = row;
    // Row 1: band
    ws.mergeCells(startRow,2,startRow,3);
    const bandLabelCell = ws.getCell(startRow,2);
    bandLabelCell.value = 'S'+weekInfo.weekNumber;
    bandLabelCell.font = {bold:true, color:{argb:COLORS.bandText}, size:12};
    bandLabelCell.fill = fill(COLORS.band);
    bandLabelCell.alignment = {vertical:'middle', horizontal:'center'};

    const refDays = weekInfo.refDays; // array of 7 {dayAbbr,date}
    for(let d=0; d<7; d++){
      const col = 4 + d*2;
      ws.mergeCells(startRow, col, startRow, col+1);
      const cell = ws.getCell(startRow, col);
      const info = refDays[d];
      cell.value = info ? (DAY_FULL[info.dayAbbr]+' '+info.date) : DAY_FULL[DAY_ABBR[d]];
      cell.alignment = {vertical:'middle', horizontal:'center'};
      cell.font = {bold:true, color:{argb: d<5?COLORS.bandText:'FF1B2430'}, size:10.5};
      cell.fill = fill(d<5?COLORS.band:COLORS.weekendHeader);
    }
    ws.mergeCells(startRow,18,startRow+1,18);
    ws.mergeCells(startRow,19,startRow+1,19);
    ws.mergeCells(startRow,20,startRow+1,20);
    ws.mergeCells(startRow,21,startRow+1,21);
    ['OFF','NB heures','TT','Commentaires'].forEach((label, i)=>{
      const cell = ws.getCell(startRow, 18+i);
      cell.value = label;
      cell.font = {bold:true, color:{argb:COLORS.bandText}, size:10};
      cell.fill = fill(COLORS.subHeader);
      cell.alignment = {vertical:'middle', horizontal:'center', wrapText:true};
    });

    // Row 2: sub-header
    const subRow = startRow+1;
    ws.getCell(subRow,2).value = 'Zones';
    ws.getCell(subRow,3).value = roleLabel;
    [2,3].forEach(c=>{
      const cell = ws.getCell(subRow,c);
      cell.font = {bold:true, color:{argb:COLORS.subHeaderText}, size:10.5};
      cell.fill = fill(COLORS.subHeader);
      cell.alignment = {vertical:'middle', horizontal:'center'};
    });
    for(let d=0; d<7; d++){
      const col = 4+d*2;
      ws.getCell(subRow,col).value = 'Début de shift';
      ws.getCell(subRow,col+1).value = 'Fin de shift';
      [col,col+1].forEach(c=>{
        const cell = ws.getCell(subRow,c);
        cell.font = {bold:true, color:{argb:d<5?COLORS.subHeaderText:'FF1B2430'}, size:9};
        cell.fill = fill(d<5?COLORS.subHeader:COLORS.weekendHeader);
        cell.alignment = {vertical:'middle', horizontal:'center', wrapText:true};
      });
    }
    for(let c=1;c<=21;c++){ ws.getCell(startRow,c).border = thinBorder(); ws.getCell(subRow,c).border = thinBorder(); }

    row = subRow+1;

    // Data rows
    members.forEach((person, idx)=>{
      const r = row;
      const byAbbr = {};
      person.days.forEach(d=>byAbbr[d.dayAbbr]=d);

      const roleCell = ws.getCell(r,1);
      if(person.role==='Manager'){
        roleCell.value = 'Manager';
        roleCell.fill = fill(COLORS.managerTag);
        roleCell.font = {bold:true, size:9.5};
        roleCell.alignment = {vertical:'middle', horizontal:'center'};
      }

      ws.getCell(r,2).value = person.zone || '';
      ws.getCell(r,3).value = person.name;
      ws.getCell(r,3).font = {bold:true, size:10};

      let offCount = 0;
      for(let d=0; d<7; d++){
        const col = 4+d*2;
        const day = byAbbr[DAY_ABBR[d]];
        const c1 = ws.getCell(r,col), c2 = ws.getCell(r,col+1);
        if(!day){
          c1.value=''; c2.value='';
        } else if(day.status==='travail'){
          c1.value = day.startMA; c2.value = day.endMA;
        } else if(day.status==='repos'){
          c1.value='OFF'; c2.value='OFF';
          c1.fill=fill(COLORS.off); c2.fill=fill(COLORS.off);
          offCount++;
        } else if(day.status==='conge'){
          ws.mergeCells(r,col,r,col+1);
          c1.value='Congé'; c1.fill=fill(COLORS.conge); c1.alignment={horizontal:'center'};
        } else if(day.status==='ferie'){
          ws.mergeCells(r,col,r,col+1);
          c1.value='Férié Français'; c1.fill=fill(COLORS.ferie); c1.alignment={horizontal:'center'};
        }
        c1.font = {size:10}; c2.font = {size:10};
      }

      ws.getCell(r,18).value = offCount;
      ws.getCell(r,18).alignment = {horizontal:'center'};
      ws.getCell(r,19).value = person.totalHours ? person.totalHours+':00' : '';
      ws.getCell(r,19).alignment = {horizontal:'center'};
      ws.getCell(r,20).value = person.tt || '';
      ws.getCell(r,20).alignment = {horizontal:'center'};
      ws.getCell(r,21).value = person.comment || '';

      if(idx % 2 === 1){
        for(let c=1;c<=21;c++){
          const cell = ws.getCell(r,c);
          if(!cell.fill || cell.fill.fgColor === undefined) cell.fill = fill(COLORS.zebra);
        }
      }
      for(let c=1;c<=21;c++){ ws.getCell(r,c).border = thinBorder(); }
      row++;
    });

    row += 1; // spacer
    return startRow;
  }

  weekKeys.forEach(wk=>{
    const members = weeks[wk];
    const refPerson = members.slice().sort((a,b)=>b.days.length-a.days.length)[0];
    const weekInfo = { weekNumber: refPerson.weekNumber, refDays: refPerson.days };

    const collabs = members.filter(m=>m.role!=='Manager');
    const managers = members.filter(m=>m.role==='Manager');

    writeGroupTable(collabs, 'Collaborateur', weekInfo);

    // Shift / Pause déjeuner reference block, derived from the group's own converted times
    const shiftStarts = [], shiftEnds = [], pauses = [];
    collabs.concat(managers).forEach(p=>p.days.forEach(d=>{
      if(d.status==='travail'){ shiftStarts.push(d.startMA); shiftEnds.push(d.endMA); }
      if(d.pauseMA) pauses.push(d.pauseMA);
    }));
    const commonStart = mostFrequent(shiftStarts);
    const commonEnd = mostFrequent(shiftEnds);
    const commonPause = mostFrequent(pauses);

    if(commonStart || commonPause){
      const r1 = row, r2 = row+1;
      ws.mergeCells(r1,2,r1,3); ws.mergeCells(r1,4,r1,5);
      ws.getCell(r1,2).value='Shift'; ws.getCell(r1,4).value='Pause déjeuner';
      [2,4].forEach(c=>{
        const cell = ws.getCell(r1,c);
        cell.font={bold:true,color:{argb:COLORS.subHeaderText}}; cell.fill=fill(COLORS.subHeader);
        cell.alignment={horizontal:'center'};
      });
      ws.getCell(r2,2).value = commonStart||'';
      ws.getCell(r2,3).value = commonEnd||'';
      ws.getCell(r2,4).value = commonPause? commonPause[0]:'';
      ws.getCell(r2,5).value = commonPause? commonPause[1]:'';
      [2,3,4,5].forEach(c=>{ ws.getCell(r2,c).alignment={horizontal:'center'}; ws.getCell(r2,c).font={size:10}; });
      for(let c=1;c<=5;c++){ ws.getCell(r1,c).border=thinBorder(); ws.getCell(r2,c).border=thinBorder(); }
      row = r2+2;
    }

    writeGroupTable(managers, 'Manager', weekInfo);
  });

  ws.views = [{state:'frozen', ySplit:0}];

  buildPauseSheet(wb, weeks, weekKeys);

  return wb;
}

// ---------- Feuille "Planning pause déjeuner" ----------
function pauseStatusForDay(day){
  if(!day) return {type:'none'};
  if(day.status==='travail'){
    if(day.pauseMA) return {type:'time', start:day.pauseMA[0], end:day.pauseMA[1]};
    return {type:'none'};
  }
  if(day.status==='repos') return {type:'off'};
  if(day.status==='conge') return {type:'conge'};
  if(day.status==='ferie') return {type:'ferie'};
  if(day.status==='mission') return {type:'mission'};
  if(day.status==='presse') return {type:'presse'};
  return {type:'none'};
}

function buildPauseSheet(wb, weeks, weekKeys){
  const ws = wb.addWorksheet('Planning pause déjeuner');

  ws.getColumn(1).width = 11;  // Zones
  ws.getColumn(2).width = 24;  // Collaborateur
  for(let c=3;c<=16;c++) ws.getColumn(c).width = 8.5; // 7 jours x 2 (D P / FP)

  let row = 1;

  weekKeys.forEach(wk=>{
    const members = weeks[wk];
    const refPerson = members.slice().sort((a,b)=>b.days.length-a.days.length)[0];
    const refDays = refPerson.days;

    // Bandeau titre
    const titleRow = row;
    ws.mergeCells(titleRow,1,titleRow,16);
    const titleCell = ws.getCell(titleRow,1);
    titleCell.value = 'Planning pause déjeuner';
    titleCell.font = {bold:true, color:{argb:COLORS.bandText}, size:12};
    titleCell.fill = fill(COLORS.band);
    titleCell.alignment = {vertical:'middle', horizontal:'left', indent:1};

    // Ligne "Zones" / "Collaborateur" + jours (fusion verticale sur 2 lignes pour les 2 premières colonnes)
    const headRow = titleRow+1, subRow = titleRow+2;
    ws.mergeCells(headRow,1,subRow,1);
    ws.mergeCells(headRow,2,subRow,2);
    ws.getCell(headRow,1).value = 'Zones';
    ws.getCell(headRow,2).value = 'Collaborateur';
    [1,2].forEach(c=>{
      const cell = ws.getCell(headRow,c);
      cell.font = {bold:true, color:{argb:COLORS.bandText}, size:10.5};
      cell.fill = fill(COLORS.subHeader);
      cell.alignment = {vertical:'middle', horizontal:'center'};
    });

    for(let d=0; d<7; d++){
      const col = 3 + d*2;
      ws.mergeCells(headRow, col, headRow, col+1);
      const info = refDays[d];
      const headCell = ws.getCell(headRow, col);
      headCell.value = info ? (DAY_FULL[info.dayAbbr]+' '+info.date) : DAY_FULL[DAY_ABBR[d]];
      headCell.font = {bold:true, color:{argb:COLORS.bandText}, size:10};
      headCell.fill = fill(COLORS.band);
      headCell.alignment = {vertical:'middle', horizontal:'center'};

      ws.getCell(subRow,col).value = 'D P';
      ws.getCell(subRow,col+1).value = 'FP';
      [col,col+1].forEach(c=>{
        const cell = ws.getCell(subRow,c);
        cell.font = {bold:true, color:{argb:COLORS.bandText}, size:9.5};
        cell.fill = fill(COLORS.subHeader);
        cell.alignment = {vertical:'middle', horizontal:'center'};
      });
    }
    for(let c=1;c<=16;c++){
      ws.getCell(headRow,c).border = thinBorder();
      ws.getCell(subRow,c).border = thinBorder();
    }
    ws.autoFilter = { from:{row:headRow, column:1}, to:{row:headRow, column:16} };

    row = subRow+1;

    members.forEach(person=>{
      const r = row;
      const byAbbr = {};
      person.days.forEach(d=>byAbbr[d.dayAbbr]=d);

      const zoneCell = ws.getCell(r,1);
      zoneCell.value = person.zone || '';
      zoneCell.alignment = {vertical:'middle', horizontal:'center'};
      zoneCell.font = {bold:true, size:9.5, color:{argb: (person.zone||'').toUpperCase()==='TRSPT KO' ? 'FFFFFFFF' : 'FFFFFFFF'}};
      zoneCell.fill = fill((person.zone||'').toUpperCase()==='TRSPT KO' ? COLORS.trsptKo : COLORS.subHeader);

      const nameCell = ws.getCell(r,2);
      nameCell.value = person.name;
      nameCell.font = {bold:true, size:10, color:{argb:'FFFFFFFF'}};
      nameCell.fill = fill(COLORS.subHeader);
      nameCell.alignment = {vertical:'middle', horizontal:'left', indent:1};

      for(let d=0; d<7; d++){
        const col = 3 + d*2;
        const day = byAbbr[DAY_ABBR[d]];
        const info = pauseStatusForDay(day);
        const c1 = ws.getCell(r,col), c2 = ws.getCell(r,col+1);

        if(info.type==='time'){
          c1.value = info.start; c2.value = info.end;
          c1.fill = fill(COLORS.pauseTime); c2.fill = fill(COLORS.pauseTime);
          c1.font = {size:10}; c2.font = {size:10};
        } else if(info.type==='off'){
          ws.mergeCells(r,col,r,col+1);
          c1.value = 'OFF'; c1.fill = fill(COLORS.off);
          c1.font = {size:10, color:{argb:'FF8A5A00'}};
          c1.alignment = {horizontal:'center'};
        } else if(info.type==='conge'){
          ws.mergeCells(r,col,r,col+1);
          c1.value = 'Congé'; c1.fill = fill(COLORS.conge);
          c1.font = {size:10, bold:true, color:{argb:'FF2F6E44'}};
          c1.alignment = {horizontal:'center'};
        } else if(info.type==='ferie'){
          ws.mergeCells(r,col,r,col+1);
          c1.value = 'Férié'; c1.fill = fill(COLORS.ferie);
          c1.font = {size:10};
          c1.alignment = {horizontal:'center'};
        } else if(info.type==='mission'){
          ws.mergeCells(r,col,r,col+1);
          c1.value = 'Missionnée'; c1.fill = fill(COLORS.missionPresse);
          c1.font = {size:10, bold:true};
          c1.alignment = {horizontal:'center'};
        } else if(info.type==='presse'){
          ws.mergeCells(r,col,r,col+1);
          c1.value = 'Presse'; c1.fill = fill(COLORS.missionPresse);
          c1.font = {size:10, bold:true};
          c1.alignment = {horizontal:'center'};
        } else {
          c1.value=''; c2.value='';
        }
      }
      for(let c=1;c<=16;c++){ ws.getCell(r,c).border = thinBorder(); }
      row++;
    });

    row += 2; // espace avant la semaine suivante
  });

  ws.views = [{state:'frozen', xSplit:2, ySplit:3}];
}

document.getElementById('exportBtn').addEventListener('click', async ()=>{
  status2.textContent = 'Génération du fichier Excel…';
  status2.className='status-line';
  try{
    const wb = await buildWorkbook();
    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'Planning_Maroc_'+new Date().toISOString().slice(0,10)+'.xlsx';
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    status2.textContent = 'Fichier Excel généré et téléchargé ✓';
    status2.className='status-line ok';
  } catch(err){
    console.error(err);
    status2.textContent = 'Erreur pendant la génération : '+err.message;
    status2.className='status-line error';
  }
});
</script>
</body>
</html>
