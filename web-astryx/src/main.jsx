import React, {useEffect, useMemo, useState} from 'react';
import {createRoot} from 'react-dom/client';
import '@astryxdesign/core/reset.css';
import '@astryxdesign/core/astryx.css';
import '@astryxdesign/theme-neutral/theme.css';
import {Card} from '@astryxdesign/core/Card';
import {Badge} from '@astryxdesign/core/Badge';
import {Button} from '@astryxdesign/core/Button';
import './styles.css';

const tabs = [
  ['dashboard', 'Dashboard'],
  ['ingest', 'Upload / Folders'],
  ['library', 'Library'],
  ['ask', 'Ask AI'],
  ['community', 'Community Install'],
  ['settings', 'Settings'],
];

const communityRepoUrl = 'https://github.com/Buafra/docwise-community.git';

const communityFreshInstall = `git clone ${communityRepoUrl}
cd docwise-community
.\\start.bat`;

const communityUpdate = `cd docwise-community
git pull
.\\start.bat`;

async function api(path, opts = {}) {
  const res = await fetch(path, opts);
  if (!res.ok) {
    let msg = res.statusText;
    try { msg = (await res.json()).detail || msg; } catch {}
    throw new Error(msg);
  }
  return res.json();
}
function isArabic(s='') { return /[\u0600-\u06FF]/.test(s); }
function money(v, c='USD') { return `${c} ${Number(v || 0).toFixed(2)}`; }

function Shell() {
  const [tab, setTab] = useState('dashboard');
  const [status, setStatus] = useState(null);
  const [docs, setDocs] = useState([]);
  const [selectedDoc, setSelectedDoc] = useState(null);
  const [toast, setToast] = useState('');
  const [mode, setMode] = useState(() => localStorage.getItem('docwise-mode') || 'light');

  async function refresh() {
    const s = await api('/api/status');
    setStatus(s);
    const d = await api('/api/documents?limit=80');
    setDocs(d.documents || []);
  }
  useEffect(() => { refresh().catch(e => setToast(e.message)); }, []);
  useEffect(() => { const t = setInterval(() => api('/api/status').then(setStatus).catch(()=>{}), 5000); return () => clearInterval(t); }, []);
  useEffect(() => { document.documentElement.dataset.mode = mode; localStorage.setItem('docwise-mode', mode); }, [mode]);

  return <div className="appShell">
    <aside className="sideNav">
      <div className="brand"><div className="logo">د</div><div><h1>DocWise</h1><p>Astryx OCR + RAG</p></div></div>
      <nav>{tabs.map(([id,label]) => <button key={id} className={tab===id?'active':''} onClick={() => setTab(id)}>{label}</button>)}</nav>
      <Card variant="muted" padding={4} className="statusCard">
        <b>System</b>
        <p>{status ? `${status.documents} docs · ${status.chunks || 0} chunks · ${status.embeddings || 0} vectors` : 'Loading...'}</p>
        <p>{status?.ocr?.arabic ? 'Arabic OCR ready' : 'Arabic OCR missing'} · {status?.ocr?.openai_vision ? 'OpenAI on' : status?.ocr?.ollama?.available ? 'Ollama on' : 'Local mode'}</p>
        {status?.ocr_progress?.running && <p className="live">OCR: page {status.ocr_progress.pages_done}/{status.ocr_progress.pages_total} · {status.ocr_progress.file}</p>}
        {status?.scan?.running && <p className="live">{status.scan.message}</p>}
        {status?.ocr?.embedding?.downloading && <p className="live">{status.ocr.embedding.progress}</p>}
        {status?.ocr?.embedding?.backfill_total > 0 && status?.ocr?.embedding?.backfill_done < status?.ocr?.embedding?.backfill_total && <p className="live">Upgrading search vectors: {status.ocr.embedding.backfill_done}/{status.ocr.embedding.backfill_total}</p>}
      </Card>
    </aside>
    <main className="main">
      <header className="topbar"><div><Badge variant="purple" label="Meta Astryx UI" /><h2>{tabs.find(t=>t[0]===tab)?.[1]}</h2></div><div className="topActions"><Button label={mode === 'dark' ? 'Light mode' : 'Dark mode'} variant="secondary" onClick={() => setMode(mode === 'dark' ? 'light' : 'dark')} /><Button label="Refresh" variant="secondary" onClick={refresh} /></div></header>
      {toast && <Card variant="yellow" padding={3} className="toast"><span>{toast}</span><button onClick={()=>setToast('')}>×</button></Card>}
      {tab === 'dashboard' && <Dashboard status={status} docs={docs} openDoc={setSelectedDoc} />}
      {tab === 'ingest' && <Ingest onChange={refresh} setToast={setToast} />}
      {tab === 'library' && <Library docs={docs} setDocs={setDocs} openDoc={setSelectedDoc} refresh={refresh} setToast={setToast} />}
      {tab === 'ask' && <Ask openDoc={setSelectedDoc} />}
      {tab === 'community' && <CommunityInstall setToast={setToast} />}
      {tab === 'settings' && <Settings status={status} refresh={refresh} setToast={setToast} />}
    </main>
    {selectedDoc && <DocDialog docId={selectedDoc.id || selectedDoc} initialPage={selectedDoc.page || 1} highlight={selectedDoc.q || ''} onClose={()=>setSelectedDoc(null)} refresh={refresh} setToast={setToast} />}
  </div>;
}

function Metric({value,label,variant='cyan'}) { return <Card variant={variant} padding={5} elevation="low" className="metric"><strong>{value}</strong><span>{label}</span></Card>; }
function Dashboard({status, docs, openDoc}) {
  return <section>
    <div className="metrics"><Metric value={status?.documents ?? 0} label="Indexed docs" /><Metric value={status?.needs_review ?? 0} label="Need review" variant="yellow" /><Metric value={status?.folders ?? 0} label="Watched folders" variant="green" /></div>
    <div className="grid2"><Card padding={5}><h3>Recent documents</h3><DocList docs={docs.slice(0,8)} openDoc={openDoc} /></Card><Card padding={5}><h3>RAG health</h3><div className="setupRows"><Row label="FTS rows" value={status?.fts_rows ?? 0}/><Row label="Embeddings" value={status?.embeddings ?? 0}/><Row label="OCR quality" value={Object.entries(status?.ocr_quality_counts || {}).map(([k,v])=>`${k}:${v}`).join(' · ') || 'none'}/></div></Card></div>
  </section>;
}

function Ingest({onChange,setToast}) {
  const [folder,setFolder] = useState('');
  const [busy,setBusy] = useState(false);
  async function upload(e) {
    const files = e.target.files;
    if (!files?.length) return;
    const fd = new FormData();
    [...files].forEach(f => fd.append('files', f));
    setToast('Uploading and indexing...');
    const r = await api('/api/upload', {method:'POST', body: fd});
    const dups = (r.results||[]).filter(x=>x.status==='duplicate').length;
    setToast(dups ? `Upload complete - ${dups} duplicate${dups>1?'s':''} skipped (already in archive)` : 'Upload/index complete');
    onChange();
  }
  async function browse() {
    setBusy(true);
    setToast('Choose a folder in the window that opened (it may be behind the browser)...');
    try {
      const r = await api('/api/pick-folder', {method:'POST'});
      if (r.cancelled) { setToast('Folder choice cancelled'); return; }
      setFolder(r.path);
      setToast('Folder chosen - click "Add folder + scan"');
    } catch (e) { setToast(e.message); } finally { setBusy(false); }
  }
  async function addFolder() {
    if (!folder.trim()) return setToast('Choose or type a folder path');
    setToast('Scanning folder...');
    const r = await api('/api/folders', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({path:folder, recursive:true, watch:true, scan_now:true})});
    const s = r.scan || {};
    setToast(`Folder added: ${s.indexed||0} indexed, ${s.unchanged||0} unchanged, ${s.duplicates||0} duplicates skipped. It is watched - new files index automatically.`);
    onChange();
  }
  return <div className="grid2"><Card padding={5}><h3>Upload files</h3><p>PDF, images, Word, Excel, PowerPoint, text and RTF. Duplicates are detected and skipped automatically.</p><input className="fileInput" type="file" multiple onChange={upload}/></Card><Card padding={5}><h3>Folder access</h3><p>Watched folders rescan automatically when new files arrive.</p><div className="folderRow"><Button label={busy?'Choosing...':'Browse...'} variant="primary" onClick={browse}/><input value={folder} onChange={e=>setFolder(e.target.value)} placeholder="...or type a path like C:\\Users\\You\\Documents\\Scans"/></div><Button label="Add folder + scan" variant="secondary" onClick={addFolder}/></Card></div>;
}

function Library({docs,setDocs,openDoc,refresh,setToast}) {
  const [q,setQ] = useState(''); const [type,setType] = useState('all'); const [checked,setChecked] = useState({});
  async function search() {
    if (!q.trim() && type==='all') return refresh();
    const res = await api('/api/search',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({q,doc_type:type,limit:100})});
    setDocs(res.results || []);
  }
  const ids = Object.keys(checked).filter(k=>checked[k]).map(Number);
  async function bulkIndex() { if(!ids.length) return; await api('/api/documents/bulk-delete',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({ids,delete_files:false})}); setChecked({}); refresh(); }
  return <Card padding={5}><div className="toolbar"><input value={q} onChange={e=>setQ(e.target.value)} onKeyDown={e=>e.key==='Enter'&&search()} placeholder="Search Arabic/English..."/><select value={type} onChange={e=>setType(e.target.value)}><option value="all">All</option><option value="invoice">Invoice</option><option value="contract">Contract</option><option value="news">News</option><option value="id">ID</option><option value="general">General</option></select><Button label="Search" variant="primary" onClick={search}/><Button label="Export Excel" variant="secondary" onClick={()=>window.open('/api/export/xlsx','_blank')}/><Button label="Remove selected index" variant="destructive" onClick={bulkIndex}/></div><div className="docGrid">{docs.map(d=><DocCard key={d.id} d={d} checked={!!checked[d.id]} onCheck={v=>setChecked({...checked,[d.id]:v})} openDoc={openDoc}/>)}</div></Card>;
}
function DocCard({d,checked,onCheck,openDoc}) { return <Card variant="default" padding={4} elevation="low" className="docCard"><input type="checkbox" checked={checked} onChange={e=>onCheck(e.target.checked)} /><h3 dir={isArabic(d.title)?'rtl':'ltr'}>{d.title || d.original_name}</h3><p className="path">{d.path}</p><div className="badges"><Badge variant="cyan" label={d.doc_type || 'general'} /><Badge variant={d.status==='indexed'?'green':'yellow'} label={d.status || 'new'} /><Badge variant="purple" label={d.ocr_quality || 'ocr'} /></div><p dir={isArabic(d.snippet || d.summary)?'rtl':'ltr'}>{d.snippet || d.summary || 'No summary'}</p><Button label="Open" variant="secondary" onClick={()=>openDoc(d)} /></Card>; }
function DocList({docs,openDoc}) { return <div className="list">{docs.map(d => <button key={d.id} className="listItem" onClick={()=>openDoc(d)}><b>{d.title || d.original_name}</b><span>{d.doc_type} · {d.status}</span></button>)}</div>; }

function Ask({openDoc}) {
  const [question,setQuestion]=useState(''); const [answer,setAnswer]=useState(null); const [loading,setLoading]=useState(false);
  async function ask(){ if(!question.trim()) return; setLoading(true); try{setAnswer(await api('/api/ask',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({question,use_ai:true,limit:8})}));} finally{setLoading(false);} }
  return <Card padding={5}><h3>Ask your archive</h3><textarea rows="4" value={question} onChange={e=>setQuestion(e.target.value)} placeholder="مثال: كم قيمة آخر فاتورة؟"/><Button label={loading?'Thinking...':'Ask'} variant="primary" onClick={ask}/>{answer && <div className="answer"><Card variant="muted" padding={4}><p dir={isArabic(answer.answer)?'rtl':'ltr'}>{answer.answer}</p></Card><h3>Sources — click to open the exact page</h3>{(answer.sources||[]).map((s,i)=><Card key={i} padding={3} className="source sourceClick" onClick={()=>openDoc({id:s.document_id, page:s.page, q:question})}><b>[{i+1}] {s.title}</b><small>Page {s.page} · {s.retrieval} · score {s.score}</small><p dir={isArabic(s.snippet)?'rtl':'ltr'}>{s.snippet}</p></Card>)}</div>}</Card>;
}

function CommandBlock({title, code, setToast}) {
  async function copy() {
    try {
      await navigator.clipboard.writeText(code);
      setToast('Command copied');
    } catch {
      setToast('Select and copy the command manually');
    }
  }
  return <Card padding={5} className="commandCard"><div className="commandHead"><h3>{title}</h3><Button label="Copy" variant="secondary" onClick={copy}/></div><pre className="commandBlock"><code>{code}</code></pre></Card>;
}

function CommunityInstall({setToast}) {
  return <section>
    <Card padding={5} variant="purple" className="communityHero"><Badge variant="cyan" label="Community"/><h3>Pull the files and install DocWise Community</h3><p>Share these command lines so users can clone, install dependencies, and run the community version.</p><p><b>Repo:</b> <code>{communityRepoUrl}</code></p></Card>
    <div className="grid2">
      <CommandBlock title="Fresh Windows install" code={communityFreshInstall} setToast={setToast}/>
      <CommandBlock title="Update an existing install" code={communityUpdate} setToast={setToast}/>
    </div>
    <Card padding={5}><h3>Requirements</h3><div className="setupRows"><Row label="Python" value="Python 3.10+"/><Row label="Git" value="Required for clone / pull"/><Row label="Start" value="Open http://127.0.0.1:8120 after start.bat"/></div></Card>
  </section>;
}

function Settings({status,refresh,setToast}) { async function rebuild(){setToast('Rebuilding RAG indexes...'); const r=await api('/api/rebuild-rag',{method:'POST'}); setToast(`Rebuilt ${r.rebuilt_chunks} chunk indexes`); refresh();} return <div className="grid2"><Card padding={5}><h3>OCR / RAG status</h3><div className="setupRows"><Row label="Tesseract" value={status?.ocr?.tesseract?'Available':'Missing'}/><Row label="Arabic OCR" value={status?.ocr?.arabic?'Available':'Missing'}/><Row label="OpenAI Vision" value={status?.ocr?.openai_vision?'Available':'Off'}/><Row label="Ollama local vision" value={status?.ocr?.ollama?.available?`On (${status.ocr.ollama.model})`:'Not detected'}/><Row label="Azure Document AI" value={status?.ocr?.azure?'Configured':'Off'}/><Row label="Embedding model" value={status?.ocr?.embedding_model || ''}/><Row label="Semantic search" value={status?.ocr?.embedding?.ready?'Active (local, offline)':status?.ocr?.embedding?.downloading?'Downloading model...':status?.ocr?.embedding?.runtime?'Preparing':'Lexical only'}/></div><Button label="Rebuild RAG indexes" variant="primary" onClick={rebuild}/></Card><Card padding={5}><h3>Reference docs</h3><p><a href="manual.html" target="_blank">User manual</a></p><p><a href="delete-help.html" target="_blank">Delete help</a></p></Card></div>; }
function Row({label,value}) { return <div className="row"><span>{label}</span><b>{value}</b></div>; }

function PageViewer({doc, initialPage, highlight}) {
  const maxPage = useMemo(()=>Math.max(1, ...(doc.chunks||[]).map(c=>c.page||1)), [doc]);
  const [page,setPage]=useState(Math.min(initialPage||1, maxPage));
  const [q,setQ]=useState(highlight||'');
  const [boxes,setBoxes]=useState({words:[],matched:[]});
  const [imgOk,setImgOk]=useState(true);
  useEffect(()=>{
    let live = true;
    api(`/api/documents/${doc.id}/page/${page}/words?q=${encodeURIComponent(q)}`)
      .then(r=>{ if(live) setBoxes(r); }).catch(()=>{ if(live) setBoxes({words:[],matched:[]}); });
    return ()=>{ live=false; };
  },[doc.id,page,q]);
  const matchedSet = new Set(boxes.matched||[]);
  return <div className="viewerWrap">
    <div className="viewerBar">
      <Button label="◀" variant="secondary" onClick={()=>setPage(p=>Math.max(1,p-1))}/>
      <span className="pageNum">Page {page}/{maxPage}</span>
      <Button label="▶" variant="secondary" onClick={()=>setPage(p=>Math.min(maxPage,p+1))}/>
      <input className="hlInput" value={q} onChange={e=>setQ(e.target.value)} placeholder="Highlight words on page... / ظلّل كلمة"/>
      {q && boxes.matched?.length ? <Badge variant="green" label={`${boxes.matched.length} hits`}/> : null}
    </div>
    <div className="viewerBox">
      {imgOk
        ? <div className="pageStack"><img src={`/api/documents/${doc.id}/page/${page}/image`} onError={()=>setImgOk(false)} alt={`page ${page}`}/>
            {(boxes.words||[]).map((w,i)=> matchedSet.has(i)
              ? <span key={i} className="hlBox" style={{left:`${w.x0*100}%`, top:`${w.y0*100}%`, width:`${(w.x1-w.x0)*100}%`, height:`${(w.y1-w.y0)*100}%`}}/>
              : null)}
          </div>
        : <iframe src={`/api/file/${doc.id}`} title="file preview"/>}
    </div>
  </div>;
}

function DocDialog({docId,initialPage,highlight,onClose,refresh,setToast}) {
  const [doc,setDoc]=useState(null); const [text,setText]=useState('');
  async function load(){ const d=await api(`/api/documents/${docId}`); setDoc(d); setText(d.text||''); }
  useEffect(()=>{load()},[docId]);
  if(!doc) return <div className="modal"><Card padding={5}>Loading...</Card></div>;
  async function save(){ await api(`/api/documents/${doc.id}/text`,{method:'PUT',headers:{'Content-Type':'application/json'},body:JSON.stringify({text})}); setToast('OCR text saved'); refresh(); load(); }
  async function vision(){ if(!confirm('Use OpenAI Vision OCR?'))return; const r=await api(`/api/documents/${doc.id}/vision-ocr`,{method:'POST'}); setToast(`Vision OCR: ${r.ocr_quality}`); refresh(); load(); }
  async function del(){ if(!confirm('Remove from index only?'))return; await api(`/api/documents/${doc.id}`,{method:'DELETE'}); onClose(); refresh(); }
  const ext=(doc.file_ext||'').toLowerCase();
  const viewable = ext==='.pdf' || ['.png','.jpg','.jpeg','.webp','.bmp','.tif','.tiff'].includes(ext);
  return <div className="modal"><Card padding={0} className="dialog"><div className="dialogHead"><div><h3>{doc.title}</h3><p>{doc.doc_type} · {doc.status} · OCR {doc.ocr_quality} {doc.ocr_score}</p></div><Button label="Close" variant="secondary" onClick={onClose}/></div><div className="dialogBody"><div className="preview">{viewable?<PageViewer doc={doc} initialPage={initialPage} highlight={highlight}/>:<p>No preview. Use Open file.</p>}</div><div className="docPane"><div className="actions"><Button label="Open file" onClick={()=>api(`/api/open/${doc.id}`,{method:'POST'})}/>{viewable && <Button label="Searchable PDF" variant="secondary" onClick={()=>window.open(`/api/documents/${doc.id}/searchable-pdf`,'_blank')}/>}<Button label="Save OCR Text" variant="primary" onClick={save}/><Button label="OpenAI Vision OCR" variant="secondary" onClick={vision}/><Button label="Remove index" variant="destructive" onClick={del}/></div><h3>Summary</h3><p dir={isArabic(doc.summary)?'rtl':'ltr'}>{doc.summary}</p><h3>Structured fields</h3><pre>{JSON.stringify(doc.fields||{},null,2)}</pre><h3>Extracted text</h3><textarea value={text} onChange={e=>setText(e.target.value)} dir={isArabic(text)?'rtl':'ltr'} /></div></div></Card></div>;
}

createRoot(document.getElementById('root')).render(<Shell />);
