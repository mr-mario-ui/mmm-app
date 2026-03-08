// src/pages/Editor.jsx
import { useState, useEffect, useRef, useCallback } from 'react'
import { useParams, useNavigate } from 'react-router-dom'
import { supabase } from '../lib/supabase'
import { Document, Packer, Paragraph, TextRun, HeadingLevel } from 'docx'
import { saveAs } from 'file-saver'

const COLORS = ['#3b82f6','#ef4444','#f59e0b','#10b981','#8b5cf6','#06b6d4','#f97316','#ec4899']
const COL_W  = 200
const ROW_H  = 38

function getDesc(id, all) {
  const r = [id]
  all.forEach(n => { if (n.parent_id === id) r.push(...getDesc(n.id, all)) })
  return r
}
function getDepth(node, all) {
  if (!node.parent_id) return 0
  const p = all.find(n => n.id === node.parent_id)
  return p ? getDepth(p, all) + 1 : 1
}
function getColor(node, all) {
  if (!node.parent_id) return '#1e293b'
  let cur = node
  while (cur.parent_id) {
    const p = all.find(n => n.id === cur.parent_id)
    if (!p || !p.parent_id) break
    cur = p
  }
  const root = all.find(n => !n.parent_id)
  if (!root) return COLORS[0]
  const idx = all.filter(n => n.parent_id === root.id).findIndex(n => n.id === cur.id)
  return COLORS[Math.abs(idx) % COLORS.length]
}
function getOutlineNumber(node, all) {
  if (!node.parent_id) return ''
  const parent = all.find(n => n.id === node.parent_id)
  if (!parent) return '?'
  const pNum = getOutlineNumber(parent, all)
  const idx  = all.filter(n => n.parent_id === node.parent_id).findIndex(n => n.id === node.id)
  return pNum ? `${pNum}.${idx+1}` : `${idx+1}`
}
function getOrderedNodes(all) {
  const result = []
  function walk(pid) { all.filter(n => n.parent_id===pid).forEach(c => { result.push(c); walk(c.id) }) }
  all.filter(n => !n.parent_id).forEach(r => { result.push(r); walk(r.id) })
  return result
}

// Horizontaler Baum: Root links, Kinder rechts
function treeLayout(nodes, collapsed, ox=80, oy=60) {
  if (!nodes.length) return nodes
  const pos = {}
  function leafCount(id) {
    const ch = nodes.filter(n => n.parent_id === id)
    if (!ch.length || collapsed[id]) return 1
    return ch.reduce((s,c) => s + leafCount(c.id), 0)
  }
  function place(id, depth, yTop) {
    const lc   = leafCount(id)
    pos[id]    = { x: ox + depth * COL_W, y: yTop + (lc * ROW_H) / 2 }
    if (!collapsed[id]) {
      let cy = yTop
      nodes.filter(n => n.parent_id === id).forEach(child => {
        const clc = leafCount(child.id)
        place(child.id, depth+1, cy)
        cy += clc * ROW_H
      })
    }
  }
  let rootY = oy
  nodes.filter(n => !n.parent_id).forEach(r => {
    const lc = leafCount(r.id)
    place(r.id, 0, rootY)
    rootY += lc * ROW_H + ROW_H
  })
  return nodes.map(n => pos[n.id] ? { ...n, x: pos[n.id].x, y: pos[n.id].y } : n)
}

// ── Quill Rich Text Editor ────────────────────────────────────
function RichTextEditor({ value, onChange, accentColor }) {
  const quillRef   = useRef(null)
  const isInternal = useRef(false)
  const mounted    = useRef(false)

  const cbRef = useCallback((el) => {
    if (!el || mounted.current) return
    mounted.current = true
    const init = () => {
      if (!el || quillRef.current) return
      const Q = window.Quill; if (!Q) return
      quillRef.current = new Q(el, {
        theme: 'snow', placeholder: 'Kommentar...',
        modules: { toolbar: [
          ['bold','italic','underline','strike'],
          [{ color:[] }, { background:[] }],
          [{ list:'ordered' }, { list:'bullet' }],
          ['clean']
        ]}
      })
      quillRef.current.on('text-change', () => {
        isInternal.current = true
        const h = quillRef.current.root.innerHTML
        onChange(h === '<p><br></p>' ? '' : h)
        setTimeout(() => { isInternal.current = false }, 0)
      })
    }
    if (!document.getElementById('quill-css')) {
      const l = document.createElement('link')
      l.id='quill-css'; l.rel='stylesheet'
      l.href='https://cdn.quilljs.com/1.3.7/quill.snow.css'
      document.head.appendChild(l)
    }
    if (window.Quill) { init() }
    else if (!document.getElementById('quill-script')) {
      const s = document.createElement('script')
      s.id='quill-script'; s.src='https://cdn.quilljs.com/1.3.7/quill.min.js'
      s.onload=init; document.head.appendChild(s)
    } else {
      const t = setInterval(() => { if (window.Quill) { clearInterval(t); init() } }, 50)
    }
  }, [])

  useEffect(() => {
    if (!quillRef.current || isInternal.current) return
    const cur = quillRef.current.root.innerHTML
    if ((cur==='<p><br></p>'?'':cur) !== (value||''))
      quillRef.current.clipboard.dangerouslyPasteHTML(value||'')
  }, [value])

  return (
    <div style={{ flex:1, display:'flex', flexDirection:'column', overflow:'hidden', minHeight:0 }}>
      <style>{`
        .ql-toolbar{border-radius:6px 6px 0 0!important;border-color:#e2e8f0!important;background:#fafafa;padding:3px 6px!important;flex-shrink:0}
        .ql-container{border-radius:0 0 6px 6px!important;border-color:#e2e8f0!important;flex:1;overflow-y:auto;font-size:12.5px!important}
        .ql-editor{min-height:50px;line-height:1.6;color:#334155;padding:8px 10px}
        .ql-editor.ql-blank::before{color:#94a3b8;font-style:normal}
        .ql-toolbar button:hover .ql-stroke,.ql-toolbar button.ql-active .ql-stroke{stroke:${accentColor}!important}
        .ql-toolbar button:hover .ql-fill,.ql-toolbar button.ql-active .ql-fill{fill:${accentColor}!important}
      `}</style>
      <div style={{ flex:1, display:'flex', flexDirection:'column', overflow:'hidden',
        border:'1px solid #e2e8f0', borderRadius:6 }}>
        <div ref={cbRef} style={{ flex:1, overflow:'hidden', display:'flex', flexDirection:'column' }}/>
      </div>
    </div>
  )
}

function Btn({ icon, tip, onClick, color='#64748b', disabled, hi }) {
  return (
    <button onClick={onClick} disabled={disabled} title={tip}
      style={{ width:32, height:32, borderRadius:6, border:'none', cursor:disabled?'default':'pointer',
        background:'transparent', color: hi?'#ef4444': disabled?'#d1d5db':color,
        fontSize:15, display:'flex', alignItems:'center', justifyContent:'center',
        opacity:disabled?0.35:1, flexShrink:0 }}
      onMouseEnter={e => { if (!disabled) e.currentTarget.style.background=hi?'#fee2e2':'#f1f5f9' }}
      onMouseLeave={e => { e.currentTarget.style.background='transparent' }}>
      {icon}
    </button>
  )
}

const MAX_UNDO = 8

export default function Editor() {
  const { mapId } = useParams()
  const navigate  = useNavigate()

  const [nodes, setNodes]           = useState([])
  const [mapTitle, setMapTitle]     = useState('')
  const [selectedId, setSelectedId] = useState(null)
  const [collapsed, setCollapsed]   = useState({})
  const [comment, setComment]       = useState('')
  const [saved, setSaved]           = useState(true)
  const [pan, setPan]               = useState({ x:0, y:0 })
  const [zoom, setZoom]             = useState(1)
  const [listW, setListW]           = useState(280)
  const [cmtH, setCmtH]             = useState(180)

  const [renameId, setRenameId]     = useState(null)
  const [renameText, setRenameText] = useState('')
  const renameRef   = useRef(null)
  const undoStack   = useRef([])
  const cTimer      = useRef(null)
  const svgRef      = useRef()
  const isPan       = useRef(false)
  const panStart    = useRef({x:0,y:0})
  const isResL      = useRef(false); const resLx = useRef(0); const resLw = useRef(0)
  const isResC      = useRef(false); const resCy = useRef(0); const resCh = useRef(0)

  const pushUndo = useCallback(s => {
    undoStack.current = [...undoStack.current.slice(-MAX_UNDO+1), JSON.parse(JSON.stringify(s))]
  }, [])

  useEffect(() => {
    const load = async () => {
      const { data: map } = await supabase.from('maps').select('title').eq('id', mapId).single()
      if (map) setMapTitle(map.title)
      const { data } = await supabase.from('nodes').select('*').eq('map_id', mapId)
      if (data?.length) {
        setNodes(treeLayout(data, {}))
      } else {
        const root = { id: crypto.randomUUID(), map_id: mapId, label: map?.title||'Start',
          x:80, y:300, parent_id:null, comment:'' }
        setNodes([root]); await supabase.from('nodes').insert(root)
      }
    }
    load()
  }, [mapId])

  useEffect(() => {
    if (renameId && renameRef.current)
      setTimeout(() => { renameRef.current?.focus(); renameRef.current?.select() }, 20)
  }, [renameId])

  useEffect(() => {
    setComment(nodes.find(n => n.id === selectedId)?.comment || '')
  }, [selectedId])

  useEffect(() => {
    const onKey = e => {
      if ((e.ctrlKey||e.metaKey) && e.key==='z' && !renameId) {
        e.preventDefault()
        if (!undoStack.current.length) return
        const prev = undoStack.current.pop()
        setNodes(prev)
        supabase.from('nodes').delete().eq('map_id', mapId).then(() => supabase.from('nodes').insert(prev))
        setSaved(true)
      }
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [renameId, mapId])

  const relayout = async (ns, coll) => {
    const c   = coll !== undefined ? coll : collapsed
    const laid = treeLayout(ns, c)
    setNodes(laid)
    for (const n of laid) await supabase.from('nodes').update({ x:n.x, y:n.y }).eq('id', n.id)
    return laid
  }

  const openRename = node => { setRenameId(node.id); setRenameText(node.label) }

  const confirmRename = async () => {
    if (!renameId) return
    const t = renameText.trim() || 'Punkt'
    pushUndo(nodes)
    const next = nodes.map(n => n.id===renameId ? {...n, label:t} : n)
    setRenameId(null); setSaved(false)
    await supabase.from('nodes').update({ label:t }).eq('id', renameId)
    await relayout(next); setSaved(true)
  }

  const addChild = async parentId => {
    pushUndo(nodes)
    const nn = { id:crypto.randomUUID(), map_id:mapId, label:'Neuer Punkt',
      x:0, y:0, parent_id:parentId, comment:'' }
    await supabase.from('nodes').insert(nn)
    const laid = await relayout([...nodes, nn])
    setSelectedId(nn.id)
    setTimeout(() => openRename(laid.find(n=>n.id===nn.id)||nn), 60)
  }

  const deleteNode = async id => {
    pushUndo(nodes)
    const del = getDesc(id, nodes)
    if (del.includes(selectedId)) setSelectedId(null)
    await supabase.from('nodes').delete().in('id', del)
    await relayout(nodes.filter(n => !del.includes(n.id)))
  }

  const toggleCollapse = id => {
    const next = { ...collapsed, [id]: !collapsed[id] }
    setCollapsed(next); relayout(nodes, next)
  }

  const handleComment = val => {
    setComment(val); if (!selectedId) return
    clearTimeout(cTimer.current); setSaved(false)
    cTimer.current = setTimeout(async () => {
      setNodes(prev => prev.map(n => n.id===selectedId ? {...n, comment:val} : n))
      await supabase.from('nodes').update({ comment:val }).eq('id', selectedId)
      setSaved(true)
    }, 800)
  }

  // Word Export
  const htmlToRuns = html => {
    if (!html) return []
    const runs = []; const div = document.createElement('div'); div.innerHTML = html
    const hex = css => {
      if (!css) return undefined
      const m = css.match(/rgb\((\d+),\s*(\d+),\s*(\d+)\)/)
      if (m) return [m[1],m[2],m[3]].map(x=>parseInt(x).toString(16).padStart(2,'0')).join('').toUpperCase()
      if (css.startsWith('#')) return css.slice(1).toUpperCase()
    }
    const walk = (el, f) => {
      if (el.nodeType===Node.TEXT_NODE) { if (el.textContent) runs.push(new TextRun({text:el.textContent,...f})); return }
      if (el.nodeType!==Node.ELEMENT_NODE) return
      const tag=el.tagName.toLowerCase(), st=el.getAttribute('style')||''
      const fg=st.match(/(?<!background-)color:\s*([^;]+)/)?.[1]
      const bg=st.match(/background-color:\s*([^;]+)/)?.[1]
      const nf={...f, bold:f.bold||tag==='strong'||tag==='b', italics:f.italics||tag==='em'||tag==='i',
        underline:(f.underline||tag==='u')?{}:f.underline, strike:f.strike||tag==='s'||tag==='del',
        color:hex(fg?.trim())||f.color, shading:bg?{fill:hex(bg.trim())}:f.shading }
      if (tag==='br'){runs.push(new TextRun({text:'',break:1}));return}
      if (tag==='p'||tag==='div'){if(runs.length)runs.push(new TextRun({text:'',break:1}));el.childNodes.forEach(c=>walk(c,nf));runs.push(new TextRun({text:'',break:1}));return}
      if (tag==='li'){runs.push(new TextRun({text:'• ',...nf}));el.childNodes.forEach(c=>walk(c,nf));runs.push(new TextRun({text:'',break:1}));return}
      el.childNodes.forEach(c=>walk(c,nf))
    }
    div.childNodes.forEach(c=>walk(c,{size:22}))
    while(runs.length&&!runs[0].text)runs.shift()
    while(runs.length&&!runs[runs.length-1].text)runs.pop()
    return runs.length?runs:[new TextRun({text:''})]
  }

  const exportToWord = async () => {
    const ordered = getOrderedNodes(nodes)
    const hMap = [HeadingLevel.HEADING_1, HeadingLevel.HEADING_2, HeadingLevel.HEADING_3]
    const ch = [new Paragraph({text:mapTitle,heading:HeadingLevel.TITLE,spacing:{after:400}})]
    ordered.forEach(n => {
      const d=getDepth(n,nodes), num=getOutlineNumber(n,nodes)
      ch.push(new Paragraph({children:[
        new TextRun({text:num?num+'  ':'',bold:true,color:'2563eb',size:Math.max(24-d*2,18)}),
        new TextRun({text:n.label,bold:d===0,size:Math.max(24-d*2,18)})
      ],heading:hMap[Math.min(d,2)],indent:{left:d*360},spacing:{before:d===0?300:100,after:n.comment?40:80}}))
      if (n.comment?.trim()) {
        const runs = htmlToRuns(n.comment)
        if (runs.length) ch.push(new Paragraph({children:runs,indent:{left:d*360+360},spacing:{after:80},
          border:{left:{style:'single',size:4,color:'CBD5E0',space:8}}}))
      }
    })
    const blob = await Packer.toBlob(new Document({sections:[{properties:{},children:ch}]}))
    saveAs(blob, `${mapTitle||'MindMap'}.docx`)
  }

  // Sichtbare Knoten
  const visibleNodes = (() => {
    const hidden = new Set()
    nodes.forEach(n => {
      if (!n.parent_id) return
      let cur = n
      while (cur.parent_id) {
        if (collapsed[cur.parent_id]) { hidden.add(n.id); break }
        cur = nodes.find(x => x.id===cur.parent_id) || {}
      }
    })
    return nodes.filter(n => !hidden.has(n.id))
  })()

  const selectedNode = nodes.find(n => n.id===selectedId)
  const orderedNodes = getOrderedNodes(nodes)
  const canUndo      = undoStack.current.length > 0

  // ── SVG ──────────────────────────────────────────────────
  const edges = visibleNodes.map(node => {
    if (!node.parent_id) return null
    const p = visibleNodes.find(n => n.id===node.parent_id)
    if (!p) return null
    const color = getColor(node, nodes)
    const midX  = p.x + COL_W * 0.5
    return (
      <path key={`e-${node.id}`}
        d={`M ${p.x} ${p.y} H ${midX} V ${node.y} H ${node.x - 16}`}
        fill="none" stroke={color} strokeWidth="1.4" strokeOpacity="0.45"
        strokeLinejoin="round" strokeLinecap="round"/>
    )
  })

  const nodeEls = visibleNodes.map(node => {
    const depth   = getDepth(node, nodes)
    const color   = getColor(node, nodes)
    const isSel   = node.id === selectedId
    const isColl  = !!collapsed[node.id]
    const hasKids = nodes.some(n => n.parent_id===node.id)
    const kids    = nodes.filter(n => n.parent_id===node.id).length

    if (depth === 0) {
      const R = 36
      return (
        <g key={node.id} onClick={() => setSelectedId(node.id)}
          onDoubleClick={e => { e.stopPropagation(); openRename(node) }} style={{ cursor:'pointer' }}>
          <circle cx={node.x} cy={node.y} r={R} fill="white"
            stroke={isSel?'#3b82f6':'#d1d5db'} strokeWidth={isSel?2.5:1.5} filter="url(#sh)"/>
          {isSel && <circle cx={node.x} cy={node.y} r={R+5} fill="none"
            stroke="#3b82f6" strokeWidth="1" strokeOpacity="0.3" strokeDasharray="3 3"/>}
          <text x={node.x} y={node.y+5} textAnchor="middle"
            fill="#111827" fontSize="13" fontWeight="800"
            style={{ pointerEvents:'none', userSelect:'none' }}>
            {node.label.length>11 ? node.label.slice(0,10)+'…' : node.label}
          </text>
        </g>
      )
    }

    const fs  = depth===1 ? 12 : depth===2 ? 11 : 10.5
    const lbl = node.label.length>24 ? node.label.slice(0,23)+'…' : node.label
    const W   = Math.min(Math.max(lbl.length * fs * 0.57 + 20, 56), 175)
    const H   = depth===1 ? 28 : 22
    const rx  = H/2

    return (
      <g key={node.id} onClick={() => setSelectedId(node.id)}
        onDoubleClick={e => { e.stopPropagation(); openRename(node) }} style={{ cursor:'pointer' }}>
        <rect x={node.x-W/2} y={node.y-H/2} width={W} height={H} rx={rx}
          fill={isSel ? color : depth===1 ? color+'20' : '#f8fafc'}
          stroke={isSel ? color : depth===1 ? color+'80' : '#e5e7eb'}
          strokeWidth={isSel?1.5:1}
          filter={depth===1?'url(#sh)':'none'}/>
        <text x={node.x} y={node.y+fs*0.38} textAnchor="middle"
          fill={isSel?'white': depth===1 ? color : '#374151'}
          fontSize={fs} fontWeight={depth===1?'700':'400'}
          style={{ pointerEvents:'none', userSelect:'none' }}>{lbl}</text>
        {/* Collapse */}
        {hasKids && (
          <g onClick={e=>{e.stopPropagation();toggleCollapse(node.id)}} style={{cursor:'pointer'}}>
            <circle cx={node.x+W/2+11} cy={node.y} r={9}
              fill={isColl?color:'white'} stroke={color} strokeWidth="1.2"/>
            <text x={node.x+W/2+11} y={node.y+4} textAnchor="middle"
              fill={isColl?'white':color} fontSize="9" fontWeight="700"
              style={{pointerEvents:'none'}}>{isColl ? kids : '−'}</text>
          </g>
        )}
        {/* Kommentar-Punkt */}
        {node.comment && (
          <circle cx={node.x-W/2-7} cy={node.y} r={3} fill={color} fillOpacity="0.6"/>
        )}
      </g>
    )
  })

  return (
    <div style={{ width:'100vw', height:'100vh', display:'flex', flexDirection:'column',
      background:'white', fontFamily:'"Inter",system-ui,sans-serif', userSelect:'none' }}
      onMouseMove={e => {
        if (isPan.current) setPan({x:e.clientX-panStart.current.x, y:e.clientY-panStart.current.y})
        if (isResL.current) setListW(Math.max(180,Math.min(500,resLw.current+(e.clientX-resLx.current))))
        if (isResC.current) setCmtH(Math.max(80,Math.min(400,resCh.current-(e.clientY-resCy.current))))
      }}
      onMouseUp={() => { isPan.current=false; isResL.current=false; isResC.current=false; document.body.style.cursor='' }}>

      {/* Rename */}
      {renameId && (
        <div style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.25)', display:'flex',
          alignItems:'center', justifyContent:'center', zIndex:1000 }}>
          <div style={{ background:'white', borderRadius:12, padding:22, width:310,
            boxShadow:'0 8px 32px rgba(0,0,0,0.12)' }} onMouseDown={e=>e.stopPropagation()}>
            <div style={{ fontSize:13, fontWeight:700, color:'#111827', marginBottom:10 }}>Umbenennen</div>
            <input ref={renameRef} value={renameText} onChange={e=>setRenameText(e.target.value)}
              onKeyDown={e=>{if(e.key==='Enter')confirmRename();if(e.key==='Escape')setRenameId(null)}}
              style={{ width:'100%', padding:'8px 11px', fontSize:13, border:'2px solid #3b82f6',
                borderRadius:7, outline:'none', boxSizing:'border-box' }}/>
            <div style={{ display:'flex', gap:8, marginTop:12, justifyContent:'flex-end' }}>
              <button onClick={()=>setRenameId(null)} style={{ padding:'6px 14px', border:'1px solid #e5e7eb',
                borderRadius:6, background:'white', cursor:'pointer', fontSize:12, color:'#6b7280' }}>Abbrechen</button>
              <button onClick={confirmRename} style={{ padding:'6px 16px', border:'none', borderRadius:6,
                background:'#3b82f6', color:'white', cursor:'pointer', fontSize:12, fontWeight:600 }}>OK</button>
            </div>
          </div>
        </div>
      )}

      {/* Header */}
      <div style={{ height:46, borderBottom:'1px solid #e5e7eb', display:'flex', alignItems:'center',
        gap:2, padding:'0 8px', flexShrink:0, background:'white', zIndex:10 }}>
        <Btn icon="←" tip="Dashboard" onClick={()=>navigate('/dashboard')}/>
        <span style={{ fontWeight:700, fontSize:13, color:'#111827', flex:1, marginLeft:4,
          overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{mapTitle}</span>

        <div style={{ width:1, height:18, background:'#e5e7eb', margin:'0 4px' }}/>
        <Btn icon="⊞" tip="Layout neu" onClick={()=>relayout(nodes)}/>
        <Btn icon="↩" tip="Rückgängig" onClick={()=>{
          if(!canUndo)return
          const prev=undoStack.current.pop(); setNodes(prev)
          supabase.from('nodes').delete().eq('map_id',mapId).then(()=>supabase.from('nodes').insert(prev))
          setSaved(true)
        }} disabled={!canUndo}/>
        <Btn icon="📄" tip="Word Export" onClick={exportToWord} color="#3b82f6"/>

        {selectedNode && (<>
          <div style={{ width:1, height:18, background:'#e5e7eb', margin:'0 4px' }}/>
          <span style={{ fontSize:10, color:'#9ca3af', fontFamily:'monospace' }}>
            {getOutlineNumber(selectedNode,nodes)||'●'}
          </span>
          <span style={{ fontSize:12, fontWeight:600, color:'#374151', maxWidth:110,
            overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap', marginLeft:4 }}>
            {selectedNode.label}
          </span>
          <Btn icon="✏️" tip="Umbenennen" onClick={()=>openRename(selectedNode)}/>
          <Btn icon="＋" tip="Unterknoten" onClick={()=>addChild(selectedNode.id)} color="#10b981"/>
          {nodes.some(n=>n.parent_id===selectedNode.id) && (
            <Btn icon={collapsed[selectedNode.id]?'▶':'▼'}
              tip={collapsed[selectedNode.id]?'Aufklappen':'Zuklappen'}
              onClick={()=>toggleCollapse(selectedNode.id)} color="#8b5cf6"/>
          )}
          {selectedNode.parent_id && (
            <Btn icon="🗑" tip="Löschen" onClick={()=>deleteNode(selectedNode.id)} hi/>
          )}
        </>)}

        <div style={{ marginLeft:8, fontSize:11, color:saved?'#22c55e':'#f59e0b', fontWeight:600, flexShrink:0 }}>
          {saved?'✓':'⟳'}
        </div>
      </div>

      {/* Body */}
      <div style={{ flex:1, display:'flex', overflow:'hidden', minHeight:0 }}>

        {/* Canvas */}
        <div style={{ flex:1, position:'relative', overflow:'hidden' }}>
          <svg ref={svgRef} style={{ width:'100%', height:'100%', background:'#f9fafb', display:'block' }}
            onMouseDown={e => {
              const tag = e.target.tagName
              if (tag==='svg'||tag==='g') {
                isPan.current=true; panStart.current={x:e.clientX-pan.x, y:e.clientY-pan.y}
                setSelectedId(null)
              }
            }}
            onWheel={e => { e.preventDefault(); setZoom(z=>Math.min(2.5,Math.max(0.25,z-e.deltaY*0.001))) }}>
            <defs>
              <filter id="sh"><feDropShadow dx="0" dy="1" stdDeviation="3" floodOpacity="0.1"/></filter>
            </defs>
            <g transform={`translate(${pan.x},${pan.y}) scale(${zoom})`}>
              {edges}
              {nodeEls}
            </g>
          </svg>
          <div style={{ position:'absolute', bottom:8, left:8, fontSize:10, color:'#9ca3af',
            background:'rgba(255,255,255,0.85)', padding:'2px 8px', borderRadius:20, pointerEvents:'none' }}>
            Scroll = Zoom · Drag = Verschieben · Doppelklick = Umbenennen
          </div>
        </div>

        {/* Resize handle */}
        <div onMouseDown={e=>{isResL.current=true;resLx.current=e.clientX;resLw.current=listW;document.body.style.cursor='col-resize'}}
          style={{ width:4, background:'#f3f4f6', cursor:'col-resize', flexShrink:0 }}
          onMouseEnter={e=>e.currentTarget.style.background='#bfdbfe'}
          onMouseLeave={e=>e.currentTarget.style.background='#f3f4f6'}/>

        {/* List + Comment */}
        <div style={{ width:listW, flexShrink:0, display:'flex', flexDirection:'column',
          overflow:'hidden', borderLeft:'1px solid #e5e7eb' }}>

          {/* Gliederung */}
          <div style={{ flex:1, overflowY:'auto', padding:'8px 6px' }}>
            <div style={{ fontSize:9, fontWeight:700, color:'#9ca3af', letterSpacing:1.2,
              textTransform:'uppercase', padding:'4px 8px 8px' }}>Gliederung</div>
            {orderedNodes.map(node => {
              const d=getDepth(node,nodes), color=getColor(node,nodes), num=getOutlineNumber(node,nodes)
              const isSel = node.id===selectedId
              return (
                <div key={node.id}
                  onClick={()=>setSelectedId(node.id)}
                  onDoubleClick={()=>openRename(node)}
                  style={{ display:'flex', alignItems:'center', gap:4, padding:'4px 8px',
                    marginLeft:d*10, borderRadius:5, marginBottom:1,
                    background:isSel?color+'14':'transparent',
                    border:`1px solid ${isSel?color+'50':'transparent'}`,
                    cursor:'pointer' }}>
                  <span style={{ width:5, height:5, borderRadius:'50%', background:color, flexShrink:0 }}/>
                  <span style={{ fontSize:9, fontWeight:700, color, fontFamily:'monospace',
                    minWidth:22, flexShrink:0 }}>{num}</span>
                  <span style={{ flex:1, fontSize:11, color:isSel?color:'#374151',
                    fontWeight:d<=1?600:400, whiteSpace:'nowrap', overflow:'hidden', textOverflow:'ellipsis' }}>
                    {node.label}
                  </span>
                  {node.comment && <span style={{ fontSize:8, color:'#9ca3af' }}>●</span>}
                  <button onClick={e=>{e.stopPropagation();addChild(node.id)}}
                    style={{ background:'none',border:'none',color,cursor:'pointer',fontSize:14,padding:'0 1px',lineHeight:1,opacity:0.5 }}>+</button>
                  {node.parent_id && (
                    <button onClick={e=>{e.stopPropagation();deleteNode(node.id)}}
                      style={{ background:'none',border:'none',color:'#ef4444',cursor:'pointer',fontSize:12,padding:'0 1px',lineHeight:1,opacity:0.4 }}>×</button>
                  )}
                </div>
              )
            })}
          </div>

          {/* Kommentar */}
          {selectedNode && (<>
            <div onMouseDown={e=>{isResC.current=true;resCy.current=e.clientY;resCh.current=cmtH;document.body.style.cursor='row-resize'}}
              style={{ height:5, cursor:'row-resize', background:'#f9fafb',
                borderTop:'1px solid #e5e7eb', flexShrink:0, display:'flex', alignItems:'center', justifyContent:'center' }}
              onMouseEnter={e=>e.currentTarget.style.background='#bfdbfe'}
              onMouseLeave={e=>e.currentTarget.style.background='#f9fafb'}>
              <div style={{ width:24, height:2, background:'#d1d5db', borderRadius:2 }}/>
            </div>
            <div style={{ height:cmtH, flexShrink:0, display:'flex', flexDirection:'column',
              padding:'8px 10px', background:'#fafafa' }}>
              <div style={{ display:'flex', alignItems:'center', gap:5, marginBottom:6 }}>
                <span style={{ width:6, height:6, borderRadius:'50%', background:getColor(selectedNode,nodes) }}/>
                <span style={{ fontSize:11, fontWeight:600, color:'#4b5563', flex:1,
                  overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>
                  {getOutlineNumber(selectedNode,nodes)} {selectedNode.label}
                </span>
                <span style={{ fontSize:10, color:saved?'#22c55e':'#f59e0b' }}>{saved?'✓':'⟳'}</span>
              </div>
              <RichTextEditor value={comment} onChange={handleComment}
                accentColor={getColor(selectedNode,nodes)}/>
            </div>
          </>)}
        </div>
      </div>
    </div>
  )
}
