// src/pages/Editor.jsx
import { useState, useEffect, useRef, useCallback } from 'react'
import { useParams, useNavigate } from 'react-router-dom'
import { supabase } from '../lib/supabase'
import { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType } from 'docx'
import { saveAs } from 'file-saver'

const DEPTH_COLORS = ['#2563eb','#16a34a','#dc2626','#d97706','#7c3aed','#0891b2','#be185d']

// ── Rich Text Editor (Quill + eigene Farb-Toolbar) ───────────
const BG_COLORS  = ['#fef08a','#bbf7d0','#bfdbfe','#fecaca','#e9d5ff','#fed7aa']
const TXT_COLORS = ['#1e293b','#dc2626','#16a34a','#2563eb','#7c3aed','#d97706','#0891b2','#be185d']

function RichTextEditor({ value, onChange, accentColor }) {
  const containerRef  = useRef(null)
  const quillRef      = useRef(null)
  const isInternalChange = useRef(false)
  const [bgOpen,  setBgOpen]  = useState(false)
  const [txtOpen, setTxtOpen] = useState(false)

  useEffect(() => {
    if (!document.getElementById('quill-css')) {
      const link = document.createElement('link')
      link.id = 'quill-css'; link.rel = 'stylesheet'
      link.href = 'https://cdn.quilljs.com/1.3.7/quill.snow.css'
      document.head.appendChild(link)
    }
    const initQuill = () => {
      if (!containerRef.current || quillRef.current) return
      const Quill = window.Quill
      if (!Quill) return
      quillRef.current = new Quill(containerRef.current, {
        theme: 'snow',
        placeholder: 'Kommentar schreiben...',
        modules: {
          toolbar: [
            ['bold', 'italic', 'underline', 'strike'],
            [{ list: 'ordered' }, { list: 'bullet' }],
            ['clean']
          ]
        }
      })
      if (value) quillRef.current.clipboard.dangerouslyPasteHTML(value)
      quillRef.current.on('text-change', () => {
        isInternalChange.current = true
        const html = quillRef.current.root.innerHTML
        onChange(html === '<p><br></p>' ? '' : html)
        setTimeout(() => { isInternalChange.current = false }, 0)
      })
    }
    if (window.Quill) initQuill()
    else {
      const script = document.createElement('script')
      script.src = 'https://cdn.quilljs.com/1.3.7/quill.min.js'
      script.onload = initQuill
      document.head.appendChild(script)
    }
    return () => { quillRef.current = null }
  }, [])

  useEffect(() => {
    if (!quillRef.current || isInternalChange.current) return
    const cur = quillRef.current.root.innerHTML
    const norm = cur === '<p><br></p>' ? '' : cur
    if (norm !== value) quillRef.current.clipboard.dangerouslyPasteHTML(value || '')
  }, [value])

  const applyBg  = (color) => { quillRef.current?.format('background', color || false); setBgOpen(false) }
  const applyTxt = (color) => { quillRef.current?.format('color', color); setTxtOpen(false) }

  // Schließe Picker bei Klick außerhalb
  useEffect(() => {
    if (!bgOpen && !txtOpen) return
    const close = () => { setBgOpen(false); setTxtOpen(false) }
    document.addEventListener('mousedown', close)
    return () => document.removeEventListener('mousedown', close)
  }, [bgOpen, txtOpen])

  const pickerStyle = {
    position:'absolute', top:'100%', left:0, zIndex:200,
    background:'white', border:'1px solid #e2e8f0', borderRadius:8,
    padding:6, display:'flex', flexWrap:'wrap', gap:4, width:148,
    boxShadow:'0 4px 16px rgba(0,0,0,0.12)'
  }

  return (
    <div style={{ flex:1, display:'flex', flexDirection:'column', overflow:'hidden', minHeight:0 }}>
      <style>{`
        .ql-toolbar { border-radius: 8px 8px 0 0 !important; border-color: #e2e8f0 !important; background: white; padding: 4px 8px !important; flex-shrink: 0; }
        .ql-container { border-radius: 0 0 8px 8px !important; border-color: #e2e8f0 !important; font-family: inherit !important; font-size: 13px !important; }
        .ql-editor { min-height: 60px; line-height: 1.7; color: #334155; padding: 10px 12px; }
        .ql-editor.ql-blank::before { color: #94a3b8; font-style: normal; }
        .ql-toolbar button:hover .ql-stroke, .ql-toolbar button.ql-active .ql-stroke { stroke: ${accentColor} !important; }
        .ql-toolbar button:hover .ql-fill, .ql-toolbar button.ql-active .ql-fill { fill: ${accentColor} !important; }
      `}</style>

      {/* Eigene Farb-Buttons rechts neben der Quill-Toolbar */}
      <div style={{ display:'flex', alignItems:'center', gap:4, padding:'3px 8px',
        background:'white', borderRadius:'8px 8px 0 0', borderBottom:'1px solid #e2e8f0',
        borderLeft:'1.5px solid #e2e8f0', borderRight:'1.5px solid #e2e8f0', borderTop:'1.5px solid #e2e8f0',
        flexShrink:0 }}>
        <span style={{ fontSize:11, color:'#94a3b8', marginRight:2 }}>Formatierung:</span>

        {/* Hintergrundfarbe */}
        <div style={{ position:'relative' }} onMouseDown={e => e.stopPropagation()}>
          <button onMouseDown={e => { e.preventDefault(); setBgOpen(o => !o); setTxtOpen(false) }}
            title="Hintergrundfarbe"
            style={{ display:'flex', alignItems:'center', gap:3, padding:'3px 7px', borderRadius:5,
              border:'1px solid #e2e8f0', background:'#f8fafc', cursor:'pointer', fontSize:11, color:'#475569' }}>
            🖍 <span style={{ fontSize:10 }}>Hintergrund</span>
          </button>
          {bgOpen && (
            <div style={pickerStyle} onMouseDown={e => e.stopPropagation()}>
              {BG_COLORS.map(c => (
                <div key={c} onMouseDown={() => applyBg(c)}
                  style={{ width:22, height:22, borderRadius:4, background:c,
                    border:'1.5px solid #e2e8f0', cursor:'pointer' }}/>
              ))}
              {/* Entfernen */}
              <div onMouseDown={() => applyBg(false)} title="Entfernen"
                style={{ width:22, height:22, borderRadius:4, cursor:'pointer',
                  border:'1.5px solid #e2e8f0', fontSize:13, display:'flex', alignItems:'center', justifyContent:'center' }}>✕</div>
            </div>
          )}
        </div>

        {/* Schriftfarbe */}
        <div style={{ position:'relative' }} onMouseDown={e => e.stopPropagation()}>
          <button onMouseDown={e => { e.preventDefault(); setTxtOpen(o => !o); setBgOpen(false) }}
            title="Schriftfarbe"
            style={{ display:'flex', alignItems:'center', gap:3, padding:'3px 7px', borderRadius:5,
              border:'1px solid #e2e8f0', background:'#f8fafc', cursor:'pointer', fontSize:11, color:'#475569' }}>
            A <span style={{ fontSize:10 }}>Farbe</span>
          </button>
          {txtOpen && (
            <div style={pickerStyle} onMouseDown={e => e.stopPropagation()}>
              {TXT_COLORS.map(c => (
                <div key={c} onMouseDown={() => applyTxt(c)}
                  style={{ width:22, height:22, borderRadius:4, background:c,
                    border:'1.5px solid #e2e8f0', cursor:'pointer' }}/>
              ))}
              {/* Zurücksetzen auf Standard */}
              <div onMouseDown={() => applyTxt(false)} title="Standard"
                style={{ width:22, height:22, borderRadius:4, cursor:'pointer',
                  border:'1.5px solid #e2e8f0', fontSize:13, display:'flex', alignItems:'center', justifyContent:'center', color:'#475569' }}>✕</div>
            </div>
          )}
        </div>
      </div>

      <div ref={containerRef}
        style={{ flex:1, border:'1.5px solid #e2e8f0', borderTop:'none',
          borderRadius:'0 0 8px 8px', overflow:'hidden', minHeight:0 }}/>
    </div>
  )
}

function getDepth(node, all) {
  if (!node.parent_id) return 0
  const p = all.find(n => n.id === node.parent_id)
  return p ? getDepth(p, all) + 1 : 1
}
function getColor(node, all) {
  return DEPTH_COLORS[getDepth(node, all) % DEPTH_COLORS.length]
}
function getDesc(id, all) {
  const r = [id]
  all.forEach(n => { if (n.parent_id === id) r.push(...getDesc(n.id, all)) })
  return r
}
function getOutlineNumber(node, all) {
  if (!node.parent_id) return ''   // Root: keine Nummer
  const parent = all.find(n => n.id === node.parent_id)
  if (!parent) return '?'
  const siblings = all.filter(n => n.parent_id === node.parent_id)
  const idx = siblings.findIndex(n => n.id === node.id)
  if (!parent.parent_id) {
    // Direkte Kinder des Root: 1, 2, 3 …
    return `${idx + 1}`
  }
  const parentNum = getOutlineNumber(parent, all)
  return `${parentNum}.${idx + 1}`
}
function getOrderedNodes(all) {
  const result = []
  function walk(parentId) {
    all.filter(n => n.parent_id === parentId).forEach(child => { result.push(child); walk(child.id) })
  }
  all.filter(n => !n.parent_id).forEach(root => { result.push(root); walk(root.id) })
  return result
}

// ── Auto-Layout: Baum neu ausrichten ─────────────────────────
// Gibt jedem Knoten eine saubere x/y Position basierend auf Baumstruktur
function autoLayout(nodes) {
  if (nodes.length === 0) return nodes

  const NODE_W  = 230   // horizontaler Abstand
  const BOX_H   = 28    // tatsächliche Knotenhöhe (vgl. Rendering H=28/32)
  const GAP     = Math.round(BOX_H / 3)   // ≈ 9px zwischen Knoten
  const NODE_H  = BOX_H + GAP             // ≈ 37px Schritt pro Blatt
  const START_X = 100

  // Berechne Subtree-Höhe (Anzahl Blätter)
  function subtreeHeight(id) {
    const children = nodes.filter(n => n.parent_id === id)
    if (children.length === 0) return 1
    return children.reduce((sum, c) => sum + subtreeHeight(c.id), 0)
  }

  const positioned = {}

  function place(id, depth, yStart) {
    const children = nodes.filter(n => n.parent_id === id)
    const h = subtreeHeight(id)
    const yCenter = yStart + (h * NODE_H) / 2

    positioned[id] = {
      x: START_X + depth * NODE_W,
      y: yCenter
    }

    let childY = yStart
    children.forEach(child => {
      const ch = subtreeHeight(child.id)
      place(child.id, depth + 1, childY)
      childY += ch * NODE_H
    })
  }

  // Alle Wurzeln
  const roots = nodes.filter(n => !n.parent_id)
  let rootY = 60
  roots.forEach(root => {
    const h = subtreeHeight(root.id)
    place(root.id, 0, rootY)
    rootY += h * NODE_H + 80
  })

  return nodes.map(n => positioned[n.id]
    ? { ...n, x: positioned[n.id].x, y: positioned[n.id].y }
    : n
  )
}

const MAX_UNDO = 5

// ── Toolbar Button Helper ────────────────────────────────────
function ToolBtn({ icon, label, onClick, color, disabled, primary }) {
  return (
    <button
      onClick={onClick}
      disabled={disabled}
      title={label}
      style={{
        display:'flex', alignItems:'center', gap:4,
        padding:'4px 10px', borderRadius:6, border: primary ? 'none' : '1px solid #e2e8f0',
        background: primary ? color : disabled ? 'transparent' : '#f8fafc',
        color: primary ? 'white' : disabled ? '#cbd5e1' : color,
        cursor: disabled ? 'default' : 'pointer',
        fontSize:12, fontWeight: primary ? 600 : 500,
        whiteSpace:'nowrap', transition:'all 0.12s',
        opacity: disabled ? 0.5 : 1
      }}
      onMouseEnter={e => { if (!disabled && !primary) e.currentTarget.style.background = color+'18' }}
      onMouseLeave={e => { if (!disabled && !primary) e.currentTarget.style.background = '#f8fafc' }}>
      <span style={{ fontSize:13 }}>{icon}</span>
      <span style={{ fontSize:11 }}>{label}</span>
    </button>
  )
}

export default function Editor() {
  const { mapId } = useParams()
  const navigate  = useNavigate()

  const [nodes, setNodes]           = useState([])
  const [mapTitle, setMapTitle]     = useState('')
  const [selectedId, setSelectedId] = useState(null)
  const [comment, setComment]       = useState('')
  const [saved, setSaved]           = useState(true)
  const [pan, setPan]               = useState({ x: 0, y: 0 })
  const [zoom, setZoom]             = useState(1)
  const zoomRef   = useRef(1)
  const panRef    = useRef({ x: 0, y: 0 })
  const setPanSync = (np) => { panRef.current = np; setPan(np) }
  const [isMobile, setIsMobile]     = useState(window.innerWidth < 768)
  const [activeTab, setActiveTab]   = useState('map')

  // Rename Popup
  const [renameId, setRenameId]     = useState(null)
  const [renameText, setRenameText] = useState('')
  const renameInputRef              = useRef(null)

  // Undo
  const undoStack = useRef([])
  const [collapsed, setCollapsed] = useState({}) // nodeId -> true wenn zugeklappt

  const toggleCollapse = (nodeId) => {
    setCollapsed(prev => ({ ...prev, [nodeId]: !prev[nodeId] }))
  }
  // Gibt alle sichtbaren nodes zurück (collapsed berücksichtigt)
  const getVisibleNodes = (all) => {
    const hidden = new Set()
    all.forEach(n => {
      if (!n.parent_id) return
      // Prüfe ob ein Vorfahre collapsed ist
      let cur = n
      while (cur.parent_id) {
        if (collapsed[cur.parent_id]) { hidden.add(n.id); break }
        cur = all.find(x => x.id === cur.parent_id) || {}
      }
    })
    return all.filter(n => !hidden.has(n.id))
  }
  const pushUndo  = useCallback((snap) => {
    undoStack.current = [...undoStack.current.slice(-MAX_UNDO + 1), JSON.parse(JSON.stringify(snap))]
  }, [])

  // Multi-Select
  const [selectedIds, setSelectedIds]   = useState(new Set())
  const [selRect, setSelRect]           = useState(null) // { x1,y1,x2,y2 } in screen px
  const selRectStart                    = useRef(null)
  const isSelecting                     = useRef(false)

  // Canvas drag
  const dragging  = useRef(null)
  const dragOff   = useRef({ x:0, y:0 })
  const didDrag   = useRef(false)
  const isPanning = useRef(false)
  const panStart  = useRef({ x:0, y:0 })
  const svgRef    = useRef()

  // Drop target
  const [dropTargetId, setDropTargetId] = useState(null)
  const [dragGhost, setDragGhost]       = useState(null)

  // List drag
  const dragListItem  = useRef(null)
  const [listDragOver, setListDragOver] = useState(null)

  // Resizable comment panel
  const [commentHeight, setCommentHeight] = useState(200)
  const isResizingComment = useRef(false)
  const resizeCommentY    = useRef(0)
  const resizeCommentH    = useRef(0)

  // Resizable splitter (map | list)
  const containerRef      = useRef(null)
  const [listWidth, setListWidth] = useState(320)
  const isResizingList    = useRef(false)
  const resizeListX       = useRef(0)
  const resizeListW       = useRef(0)

  const commentTimer   = useRef(null)
  const selectedIdRef  = useRef(null)

  // ── Strg+Z ───────────────────────────────────────────────
  useEffect(() => {
    const handler = (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key === 'z' && !renameId) {
        e.preventDefault()
        if (undoStack.current.length === 0) return
        const prev = undoStack.current[undoStack.current.length - 1]
        undoStack.current = undoStack.current.slice(0, -1)
        setNodes(prev)
        supabase.from('nodes').delete().eq('map_id', mapId).then(() =>
          supabase.from('nodes').insert(prev)
        )
        setSaved(true)
      }
    }
    window.addEventListener('keydown', handler)
    return () => window.removeEventListener('keydown', handler)
  }, [renameId, mapId])

  useEffect(() => {
    loadMap()
    const handleResize = () => setIsMobile(window.innerWidth < 768)
    window.addEventListener('resize', handleResize)
    return () => window.removeEventListener('resize', handleResize)
  }, [mapId])

  useEffect(() => {
    if (!selectedId) { setComment(''); return }
    const node = nodes.find(n => n.id === selectedId)
    setComment(node?.comment || '')
  }, [selectedId])

  // Focus rename input
  useEffect(() => {
    if (renameId && renameInputRef.current) {
      setTimeout(() => { renameInputRef.current?.focus(); renameInputRef.current?.select() }, 30)
    }
  }, [renameId])

  // Global mouse handlers
  useEffect(() => {
    const onMouseMove = (e) => {
      // Comment resize
      if (isResizingComment.current) {
        const delta = resizeCommentY.current - e.clientY
        setCommentHeight(Math.max(80, Math.min(500, resizeCommentH.current + delta)))
      }
      // List width resize
      if (isResizingList.current) {
        const delta = resizeListX.current - e.clientX
        setListWidth(Math.max(200, Math.min(600, resizeListW.current + delta)))
      }
    }
    const onMouseUp = () => {
      isResizingComment.current = false
      isResizingList.current    = false
      document.body.style.cursor = ''
    }
    window.addEventListener('mousemove', onMouseMove)
    window.addEventListener('mouseup', onMouseUp)
    return () => {
      window.removeEventListener('mousemove', onMouseMove)
      window.removeEventListener('mouseup', onMouseUp)
    }
  }, [])

  // ── panRef synchron halten ───────────────────────────────

  // ── Viewport: Browser-Pinch-Zoom verhindern ──────────────
  useEffect(() => {
    let meta = document.querySelector('meta[name=viewport]')
    if (!meta) { meta = document.createElement('meta'); meta.name = 'viewport'; document.head.appendChild(meta) }
    meta.content = 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no'
    return () => { meta.content = 'width=device-width, initial-scale=1' }
  }, [])

  // ── Touch-Handler: Pinch-Zoom + Pan ──────────────────────
  const touch1     = useRef(null)   // erster Touch
  const touch2     = useRef(null)   // zweiter Touch (Pinch)
  const pinchDist0 = useRef(0)      // Ausgangs-Pinch-Abstand
  const pinchZoom0 = useRef(1)      // Zoom bei Pinch-Start
  const pinchMid0  = useRef({ x:0, y:0 }) // Mittelpunkt bei Pinch-Start
  const pinchPan0  = useRef({ x:0, y:0 }) // Pan bei Pinch-Start

  const svgContainerRef = useRef()

  const getDist = (a, b) => Math.hypot(b.clientX - a.clientX, b.clientY - a.clientY)
  const getMid  = (a, b) => ({ x: (a.clientX + b.clientX) / 2, y: (a.clientY + b.clientY) / 2 })

  const onTouchStart = (e) => {
    if (e.touches.length === 1) {
      touch1.current = e.touches[0]
      touch2.current = null
      // Pan starten (wie onSvgDown bei Maus)
      const t = e.touches[0]
      isPanning.current = true
      panStart.current  = { x: t.clientX - panRef.current.x, y: t.clientY - panRef.current.y }
    } else if (e.touches.length === 2) {
      touch1.current = e.touches[0]
      touch2.current = e.touches[1]
      isPanning.current = false
      pinchDist0.current = getDist(e.touches[0], e.touches[1])
      pinchZoom0.current = zoomRef.current
      pinchMid0.current  = getMid(e.touches[0], e.touches[1])
      pinchPan0.current  = { ...panRef.current }
    }
  }

  const onTouchMove = (e) => {
    e.preventDefault()
    if (e.touches.length === 1 && isPanning.current) {
      const t = e.touches[0]
      const np = { x: t.clientX - panStart.current.x, y: t.clientY - panStart.current.y }
      setPanSync(np)
    } else if (e.touches.length === 2) {
      const newDist = getDist(e.touches[0], e.touches[1])
      const newMid  = getMid(e.touches[0], e.touches[1])
      const scale   = Math.min(3, Math.max(0.2, pinchZoom0.current * newDist / pinchDist0.current))
      // Zoom um den Mittelpunkt: pan anpassen damit Mittelpunkt fixiert bleibt
      const dx = (pinchMid0.current.x - pinchPan0.current.x) / pinchZoom0.current
      const dy = (pinchMid0.current.y - pinchPan0.current.y) / pinchZoom0.current
      const np = {
        x: pinchMid0.current.x - dx * scale + (newMid.x - pinchMid0.current.x),
        y: pinchMid0.current.y - dy * scale + (newMid.y - pinchMid0.current.y),
      }
      zoomRef.current = scale
      setZoom(scale)
      setPanSync(np)
    }
  }

  const onTouchEnd = (e) => {
    if (e.touches.length < 2) { touch2.current = null }
    if (e.touches.length === 0) {
      isPanning.current = false
      touch1.current = null
    }
  }


  const loadMap = async () => {
    const { data: map } = await supabase.from('maps').select('title').eq('id', mapId).single()
    if (map) setMapTitle(map.title)
    const { data } = await supabase.from('nodes').select('*').eq('map_id', mapId)
    if (data && data.length > 0) {
      setNodes(data)
    } else {
      const root = { id: crypto.randomUUID(), map_id: mapId, label: map?.title || 'Start', x: 120, y: 300, parent_id: null, comment: '' }
      setNodes([root])
      await supabase.from('nodes').insert(root)
    }
  }

  // ── Auto-Layout anwenden + speichern ─────────────────────
  const applyAutoLayout = async (currentNodes) => {
    const laid = autoLayout(currentNodes)
    setNodes(laid)
    for (const n of laid) {
      await supabase.from('nodes').update({ x: n.x, y: n.y }).eq('id', n.id)
    }
    return laid
  }

  // ── Rename Popup ─────────────────────────────────────────
  const openRename = (node) => {
    setSelectedId(node.id)
    setRenameId(node.id)
    setRenameText(node.label)
  }
  const confirmRename = async () => {
    if (!renameId) return
    const t = renameText.trim() || 'Punkt'
    pushUndo(nodes)
    setNodes(prev => prev.map(n => n.id === renameId ? { ...n, label: t } : n))
    setRenameId(null)
    setSaved(false)
    await supabase.from('nodes').update({ label: t }).eq('id', renameId)
    setSaved(true)
  }

  // ── Kind hinzufügen ───────────────────────────────────────
  const addChild = async (parentId, e) => {
    if (e) e.stopPropagation()
    pushUndo(nodes)
    const newNode = {
      id: crypto.randomUUID(), map_id: mapId, label: 'Neuer Punkt',
      x: 0, y: 0, parent_id: parentId, comment: ''
    }
    const next = [...nodes, newNode]
    await supabase.from('nodes').insert(newNode)
    const laid = await applyAutoLayout(next)
    setSelectedId(newNode.id)
    const added = laid.find(n => n.id === newNode.id)
    if (added) setTimeout(() => openRename(added), 80)
  }

  // ── Löschen ───────────────────────────────────────────────
  const deleteNode = async (nodeId, e) => {
    if (e) e.stopPropagation()
    pushUndo(nodes)
    const del  = getDesc(nodeId, nodes)
    const next = nodes.filter(n => !del.includes(n.id))
    if (del.includes(selectedId)) setSelectedId(null)
    await supabase.from('nodes').delete().in('id', del)
    await applyAutoLayout(next)
  }

  // ── Kommentar ─────────────────────────────────────────────
  const handleCommentChange = (val) => {
    setComment(val)
    const curId = selectedIdRef.current
    if (!curId) return
    clearTimeout(commentTimer.current)
    commentTimer.current = setTimeout(async () => {
      setNodes(prev => prev.map(n => n.id === curId ? { ...n, comment: val } : n))
      setSaved(false)
      await supabase.from('nodes').update({ comment: val }).eq('id', curId)
      setSaved(true)
    }, 800)
  }

  // ── Canvas drag ───────────────────────────────────────────
  const onNodeMouseDown = (e, id) => {
    e.stopPropagation()
    if (e.ctrlKey || e.metaKey) {
      // Strg+Klick: toggle Auswahl
      setSelectedIds(prev => {
        const next = new Set(prev)
        if (next.has(id)) next.delete(id)
        else next.add(id)
        return next
      })
      setSelectedId(id)
      return
    }
    // Normaler Klick: wenn id schon in Multi-Auswahl, Multi-Drag starten
    const inMulti = selectedIds.has(id) && selectedIds.size > 1
    if (!inMulti) {
      setSelectedIds(new Set([id]))
      setSelectedId(id)
    }
    const node = nodes.find(n => n.id === id)
    dragging.current = id
    didDrag.current  = false
    dragOff.current  = {
      x: (e.clientX - panRef.current.x) / zoomRef.current - node.x,
      y: (e.clientY - panRef.current.y) / zoomRef.current - node.y
    }
  }
  const onSvgDown = (e) => {
    if (e.target === svgRef.current || e.target.tagName === 'svg') {
      if (!e.ctrlKey && !e.metaKey) {
        // Auswahlrahmen starten
        isSelecting.current = true
        selRectStart.current = { x: e.clientX, y: e.clientY }
        setSelRect({ x1: e.clientX, y1: e.clientY, x2: e.clientX, y2: e.clientY })
        setSelectedId(null)
        setSelectedIds(new Set())
      }
      isPanning.current = true
      panStart.current  = { x: e.clientX - panRef.current.x, y: e.clientY - panRef.current.y }
    }
  }
  const onMove = useCallback((e) => {
    if (dragging.current) {
      didDrag.current = true
      const dx = (e.clientX - panRef.current.x) / zoomRef.current - dragOff.current.x
      const dy = (e.clientY - panRef.current.y) / zoomRef.current - dragOff.current.y
      const draggedNode = nodes.find(n => n.id === dragging.current)
      if (!draggedNode) return
      const ddx = dx - draggedNode.x, ddy = dy - draggedNode.y
      // Multi-Drag: alle ausgewählten verschieben
      const moveIds = selectedIds.size > 1 ? selectedIds : new Set([dragging.current])
      setNodes(prev => prev.map(n => moveIds.has(n.id) ? { ...n, x: n.x + ddx, y: n.y + ddy } : n))
      setDragGhost({ x: e.clientX, y: e.clientY })
      // Drop-Target nur beim Einzel-Drag
      if (selectedIds.size <= 1) {
        const nx = dx, ny = dy
        const W_BOX = 150, H_BOX = 50
        const others = nodes.filter(n => n.id !== dragging.current && !getDesc(dragging.current, nodes).slice(1).includes(n.id))
        let hit = null
        others.forEach(n => {
          if (nx > n.x - W_BOX/2 && nx < n.x + W_BOX/2 && ny > n.y - H_BOX/2 && ny < n.y + H_BOX/2) hit = n.id
        })
        setDropTargetId(hit)
      }
    } else if (isSelecting.current) {
      setSelRect({ x1: selRectStart.current.x, y1: selRectStart.current.y, x2: e.clientX, y2: e.clientY })
      // Knoten im Rahmen ermitteln
      const rect = svgRef.current?.getBoundingClientRect()
      if (!rect) return
      const toSvg = (cx, cy) => ({
        x: (cx - rect.left - panRef.current.x) / zoomRef.current,
        y: (cy - rect.top  - panRef.current.y) / zoomRef.current
      })
      const a = toSvg(selRectStart.current.x, selRectStart.current.y)
      const b = toSvg(e.clientX, e.clientY)
      const minX = Math.min(a.x, b.x), maxX = Math.max(a.x, b.x)
      const minY = Math.min(a.y, b.y), maxY = Math.max(a.y, b.y)
      const inside = new Set(nodes.filter(n => n.x >= minX && n.x <= maxX && n.y >= minY && n.y <= maxY).map(n => n.id))
      setSelectedIds(inside)
      if (inside.size === 1) setSelectedId([...inside][0])
    } else if (isPanning.current) {
      const np = { x: e.clientX - panStart.current.x, y: e.clientY - panStart.current.y }
      setPanSync(np)
    }
  }, [nodes, selectedIds])

  const onUp = useCallback(async () => {
    if (dragging.current) {
      const draggedId = dragging.current
      if (didDrag.current) {
        pushUndo(nodes)
        const moveIds = selectedIds.size > 1 ? selectedIds : new Set([draggedId])
        if (dropTargetId && moveIds.size === 1) {
          // Reparent + neu layouten (nur Einzel-Drag)
          const next = nodes.map(n => n.id === draggedId ? { ...n, parent_id: dropTargetId } : n)
          await supabase.from('nodes').update({ parent_id: dropTargetId }).eq('id', draggedId)
          await applyAutoLayout(next)
        } else {
          // Positionen aller verschobenen Knoten speichern
          setSaved(false)
          const moved = nodes.filter(n => moveIds.has(n.id))
          await Promise.all(moved.map(n => supabase.from('nodes').update({ x: n.x, y: n.y }).eq('id', n.id)))
          setSaved(true)
        }
      }
      dragging.current = null
      didDrag.current  = false
      setDropTargetId(null)
      setDragGhost(null)
    }
    if (isSelecting.current) {
      isSelecting.current = false
      setSelRect(null)
    }
    isPanning.current = false
  }, [nodes, dropTargetId, pushUndo, selectedIds])

  // ── List drag ─────────────────────────────────────────────
  const onListDragStart = (e, id) => { dragListItem.current = id; e.dataTransfer.effectAllowed = 'move' }
  const onListDragOver  = (e, id) => { e.preventDefault(); setListDragOver(id) }
  const onListDrop      = async (e, targetId) => {
    e.preventDefault(); setListDragOver(null)
    const srcId = dragListItem.current
    if (!srcId || srcId === targetId) return
    const target = nodes.find(n => n.id === targetId)
    if (!target) return
    pushUndo(nodes)
    const next = nodes.map(n => n.id === srcId ? { ...n, parent_id: target.parent_id } : n)
    await supabase.from('nodes').update({ parent_id: target.parent_id }).eq('id', srcId)
    await applyAutoLayout(next)
  }

  const selectedNode  = nodes.find(n => n.id === selectedId)
  selectedIdRef.current = selectedId
  const orderedNodes  = getOrderedNodes(nodes)
  const visibleNodes  = getVisibleNodes(nodes)
  const canUndo       = undoStack.current.length > 0
  const hasChildren   = (id) => nodes.some(n => n.parent_id === id)

  // ── HTML → docx TextRuns (bold, italic, underline, background) ──
  const htmlToDocxRuns = (html) => {
    if (!html) return []
    const runs = []
    const div = document.createElement('div')
    div.innerHTML = html

    const hexColor = (str) => {
      if (!str) return undefined
      // rgb(r,g,b) → hex
      const m = str.match(/rgb\((\d+),\s*(\d+),\s*(\d+)\)/)
      if (m) return [m[1],m[2],m[3]].map(n => parseInt(n).toString(16).padStart(2,'0')).join('')
      // #rrggbb oder #rgb
      return str.replace('#','').length === 3
        ? str.replace('#','').split('').map(c=>c+c).join('')
        : str.replace('#','')
    }

    const walk = (el, fmt) => {
      if (el.nodeType === 3) { // Textknoten
        const text = el.textContent.replace(/\u00A0/g,' ')
        if (text) runs.push(new TextRun({
          text,
          bold:      fmt.bold,
          italics:   fmt.italic,
          underline: fmt.underline ? {} : undefined,
          strike:    fmt.strike,
          highlight: fmt.highlight,
          shading:   fmt.shading,
          size: 18,
          color: fmt.color || '334155',
        }))
        return
      }
      if (el.nodeType !== 1) return
      const tag  = el.tagName.toLowerCase()
      const style = el.style || {}
      const newFmt = { ...fmt }
      if (tag==='b'||tag==='strong') newFmt.bold = true
      if (tag==='i'||tag==='em')     newFmt.italic = true
      if (tag==='u')                 newFmt.underline = true
      if (tag==='s'||tag==='strike') newFmt.strike = true
      // Schriftfarbe aus style
      const fc = style.color
      if (fc) { const hex = hexColor(fc); if (hex) newFmt.color = hex }
      // Hintergrundfarbe aus style
      const bg = style.backgroundColor || style.background
      if (bg) {
        const hex = hexColor(bg)
        // Docx unterstützt named highlights + custom shading
        const namedMap = {
          'fef08a':'yellow','bbf7d0':'green','bfdbfe':'cyan',
          'fecaca':'red','e9d5ff':'magenta','fed7aa':'darkYellow'
        }
        const named = hex ? namedMap[hex.toLowerCase()] : null
        if (named) {
          newFmt.highlight = named
          newFmt.shading   = undefined
        } else if (hex) {
          newFmt.highlight = undefined
          newFmt.shading   = { type: 'clear', color: 'auto', fill: hex }
        }
      }
      // Zeilenumbruch bei Block-Elementen
      if (tag==='p'||tag==='div'||tag==='li') {
        el.childNodes.forEach(c => walk(c, newFmt))
        runs.push(new TextRun({ text: '', break: 1 }))
      } else {
        el.childNodes.forEach(c => walk(c, newFmt))
      }
    }

    div.childNodes.forEach(c => walk(c, { bold:false, italic:false, underline:false, strike:false }))
    // Letzten leeren Break entfernen
    while (runs.length > 0 && runs[runs.length-1].text === '' ) runs.pop()
    return runs
  }

  // ── Word Export ──────────────────────────────────────────
  const exportToWord = async () => {
    const ordered = getOrderedNodes(nodes)
    const docChildren = []

    // Titel
    docChildren.push(new Paragraph({
      text: mapTitle,
      heading: HeadingLevel.TITLE,
      spacing: { after: 400 },
    }))

    // Jeden Knoten als Gliederungspunkt
    ordered.forEach(node => {
      const depth = getDepth(node, nodes)
      const num   = getOutlineNumber(node, nodes)

      // Überschrift je Tiefe
      const headingMap = [HeadingLevel.HEADING_1, HeadingLevel.HEADING_2, HeadingLevel.HEADING_3]
      docChildren.push(new Paragraph({
        children: [
          new TextRun({ text: num + '  ', bold: true, color: '2563eb', size: Math.max(24 - depth * 2, 18) }),
          new TextRun({ text: node.label, bold: depth === 0, size: Math.max(24 - depth * 2, 18) }),
        ],
        heading: headingMap[Math.min(depth, 2)],
        indent: { left: depth * 360 },
        spacing: { before: depth === 0 ? 300 : 100, after: 80 },
      }))

      // Kommentar: HTML zu docx TextRuns konvertieren (bold, italic, background)
      if (node.comment && node.comment.trim()) {
        const runs = htmlToDocxRuns(node.comment)
        if (runs.length > 0) {
          docChildren.push(new Paragraph({
            children: runs,
            indent: { left: depth * 360 + 360 },
            spacing: { after: 60 },
          }))
        }
      }
    })

    const doc = new Document({
      sections: [{ properties: {}, children: docChildren }],
      styles: {
        default: {
          document: { run: { font: 'Calibri', size: 22 } }
        }
      }
    })

    const blob = await Packer.toBlob(doc)
    saveAs(blob, `${mapTitle || 'MindMap'}.docx`)
  }

  return (
    <div ref={containerRef}
      style={{ width:'100vw', height:'100vh', display:'flex', flexDirection:'column',
        background:'white', fontFamily:'"DM Sans", system-ui, sans-serif', userSelect:'none' }}
      onMouseMove={onMove} onMouseUp={onUp}>

      {/* ── Rename Popup ── */}
      {renameId && (
        <div style={{ position:'fixed', inset:0, background:'rgba(15,23,42,0.4)',
          display:'flex', alignItems:'center', justifyContent:'center', zIndex:1000 }}>
          <div style={{ background:'white', borderRadius:14, padding:28, width:340,
            boxShadow:'0 20px 60px rgba(0,0,0,0.2)' }}
            onMouseDown={e => e.stopPropagation()}>
            <div style={{ fontSize:15, fontWeight:700, color:'#1e293b', marginBottom:4 }}>
              Knoten umbenennen
            </div>
            <div style={{ fontSize:12, color:'#94a3b8', marginBottom:14 }}>
              Enter zum Speichern · Escape zum Abbrechen
            </div>
            <input ref={renameInputRef} value={renameText}
              onChange={e => setRenameText(e.target.value)}
              onKeyDown={e => { if (e.key==='Enter') confirmRename(); if (e.key==='Escape') setRenameId(null) }}
              style={{ width:'100%', padding:'10px 14px', fontSize:14,
                border:'2px solid #2563eb', borderRadius:8, outline:'none',
                color:'#1e293b', fontFamily:'inherit', boxSizing:'border-box' }}/>
            <div style={{ display:'flex', gap:8, marginTop:16, justifyContent:'flex-end' }}>
              <button onClick={() => setRenameId(null)}
                style={{ padding:'8px 18px', border:'1px solid #e2e8f0', borderRadius:7,
                  background:'white', color:'#64748b', cursor:'pointer', fontSize:13 }}>
                Abbrechen
              </button>
              <button onClick={confirmRename}
                style={{ padding:'8px 20px', border:'none', borderRadius:7,
                  background:'#2563eb', color:'white', cursor:'pointer', fontSize:13, fontWeight:600 }}>
                Speichern
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ── Header (2 Zeilen) ── */}
      <div style={{ borderBottom:'1px solid #e2e8f0', background:'white', flexShrink:0, zIndex:10 }}>

        {/* Zeile 1: Titel + Status */}
        <div style={{ display:'flex', alignItems:'center', gap:10, padding:'0 16px', height:44 }}>
          <button onClick={() => navigate('/dashboard')}
            style={{ background:'transparent', border:'1px solid #e2e8f0', color:'#64748b',
              padding:'4px 10px', borderRadius:6, cursor:'pointer', fontSize:11, whiteSpace:'nowrap' }}>
            ← Dashboard
          </button>
          <span style={{ fontWeight:700, fontSize:15, color:'#1e293b', flex:1,
            overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{mapTitle}</span>
          <span style={{ fontSize:11, color: saved ? '#22c55e' : '#f59e0b', fontWeight:600, flexShrink:0 }}>
            {saved ? '✓ Gespeichert' : '⟳ Speichern…'}
          </span>
          {isMobile && (
            <div style={{ display:'flex', gap:3, background:'#f1f5f9', borderRadius:7, padding:3 }}>
              {[['map','🗺'],['list','☰']].map(([tab, icon]) => (
                <button key={tab} onClick={() => setActiveTab(tab)}
                  style={{ padding:'4px 10px', borderRadius:5, border:'none', cursor:'pointer', fontSize:12,
                    background: activeTab===tab ? 'white' : 'transparent',
                    color: activeTab===tab ? '#2563eb' : '#64748b',
                    fontWeight: activeTab===tab ? 600 : 400,
                    boxShadow: activeTab===tab ? '0 1px 3px rgba(0,0,0,0.1)' : 'none' }}>
                  {icon}
                </button>
              ))}
            </div>
          )}
        </div>

        {/* Zeile 2: Toolbar */}
        <div style={{ display:'flex', alignItems:'center', gap:4, padding:'0 12px 8px', flexWrap:'wrap', rowGap:4 }}>

          {/* ── Knoten-Aktionen (nur wenn ausgewählt) ── */}
          {selectedNode ? (
            <>
              <span style={{ fontSize:10, color:'#94a3b8', fontWeight:600, marginRight:4, textTransform:'uppercase', letterSpacing:0.5 }}>
                {getOutlineNumber(selectedNode, nodes)}
              </span>
              <div style={{ width:1, height:20, background:'#e2e8f0', margin:'0 4px' }}/>
              <ToolBtn icon="✏️" label={isMobile ? '' : 'Umbenennen'} onClick={() => openRename(selectedNode)} color="#475569"/>
              <ToolBtn icon="＋" label={isMobile ? '' : 'Unterpunkt'} onClick={() => addChild(selectedNode.id)} color="#16a34a"/>
              {hasChildren(selectedNode.id) && (
                <ToolBtn
                  icon={collapsed[selectedNode.id] ? '▶' : '▼'}
                  label={isMobile ? '' : (collapsed[selectedNode.id] ? 'Aufklappen' : 'Zuklappen')}
                  onClick={() => toggleCollapse(selectedNode.id)}
                  color="#7c3aed"/>
              )}
              {selectedNode.parent_id && (
                <ToolBtn icon="🗑" label={isMobile ? '' : 'Löschen'} onClick={() => deleteNode(selectedNode.id)} color="#dc2626"/>
              )}
              <div style={{ width:1, height:20, background:'#e2e8f0', margin:'0 4px' }}/>
            </>
          ) : (
            <span style={{ fontSize:11, color:'#94a3b8', marginRight:8 }}>{isMobile ? 'Knoten antippen' : 'Knoten auswählen für Aktionen'}</span>
          )}

          {/* ── Allgemeine Aktionen ── */}
          <ToolBtn icon="⊞" label={isMobile ? '' : 'Layout'} onClick={() => applyAutoLayout(nodes)} color="#475569"/>
          <ToolBtn icon="↩" label={isMobile ? '' : `Rückgängig${canUndo ? ` (${undoStack.current.length})` : ''}`}
            onClick={() => {
              if (!canUndo) return
              const prev = undoStack.current[undoStack.current.length - 1]
              undoStack.current = undoStack.current.slice(0, -1)
              setNodes(prev)
              supabase.from('nodes').delete().eq('map_id', mapId).then(() => supabase.from('nodes').insert(prev))
              setSaved(true)
            }}
            color={canUndo ? '#475569' : '#cbd5e1'} disabled={!canUndo}/>
          <ToolBtn icon="📄" label={isMobile ? '' : '→ Word'} onClick={exportToWord} color="#2563eb" primary/>
        </div>
      </div>

      {/* ── Body ── */}
      <div style={{ flex:1, display:'flex', overflow:'hidden', minHeight:0 }}>

        {/* ── Mindmap ── */}
        {(!isMobile || activeTab==='map') && (
          <div ref={svgContainerRef} style={{ flex:1, position:'relative', overflow:'hidden', minWidth:0, touchAction:'none' }}
            onTouchStart={onTouchStart}
            onTouchMove={onTouchMove}
            onTouchEnd={onTouchEnd}>
            <svg ref={svgRef}
              style={{ width:'100%', height:'100%', background:'#f8fafc', cursor:'default', display:'block' }}
              onMouseDown={onSvgDown}
              onWheel={e => {
                e.preventDefault()
                const factor = e.deltaY < 0 ? 1.1 : 0.9
                const newZoom = Math.min(3, Math.max(0.2, zoomRef.current * factor))
                // Zoom um Mausposition
                const rect = svgRef.current.getBoundingClientRect()
                const mx = e.clientX - rect.left
                const my = e.clientY - rect.top
                const np = {
                  x: mx - (mx - panRef.current.x) * (newZoom / zoomRef.current),
                  y: my - (my - panRef.current.y) * (newZoom / zoomRef.current),
                }
                zoomRef.current = newZoom
                setZoom(newZoom)
                setPanSync(np)
              }}>
              <defs>
                <filter id="sh"><feDropShadow dx="0" dy="2" stdDeviation="4" floodOpacity="0.1"/></filter>
                <filter id="shd"><feDropShadow dx="0" dy="6" stdDeviation="10" floodOpacity="0.18"/></filter>
                {/* Glow filter für selected node - stdDeviation wird via CSS animiert */}
                <filter id="glow-blue"  x="-40%" y="-40%" width="180%" height="180%"><feGaussianBlur stdDeviation="8" result="blur"/><feMerge><feMergeNode in="blur"/><feMergeNode in="blur"/><feMergeNode in="SourceGraphic"/></feMerge></filter>
                <filter id="glow-green" x="-40%" y="-40%" width="180%" height="180%"><feGaussianBlur stdDeviation="8" result="blur"/><feMerge><feMergeNode in="blur"/><feMergeNode in="blur"/><feMergeNode in="SourceGraphic"/></feMerge></filter>
                <filter id="glow-red"   x="-40%" y="-40%" width="180%" height="180%"><feGaussianBlur stdDeviation="8" result="blur"/><feMerge><feMergeNode in="blur"/><feMergeNode in="blur"/><feMergeNode in="SourceGraphic"/></feMerge></filter>
                <filter id="glow-amber" x="-40%" y="-40%" width="180%" height="180%"><feGaussianBlur stdDeviation="8" result="blur"/><feMerge><feMergeNode in="blur"/><feMergeNode in="blur"/><feMergeNode in="SourceGraphic"/></feMerge></filter>
                <filter id="glow-violet" x="-40%" y="-40%" width="180%" height="180%"><feGaussianBlur stdDeviation="8" result="blur"/><feMerge><feMergeNode in="blur"/><feMergeNode in="blur"/><feMergeNode in="SourceGraphic"/></feMerge></filter>
                <filter id="glow-cyan"  x="-40%" y="-40%" width="180%" height="180%"><feGaussianBlur stdDeviation="8" result="blur"/><feMerge><feMergeNode in="blur"/><feMergeNode in="blur"/><feMergeNode in="SourceGraphic"/></feMerge></filter>
                <filter id="glow-pink"  x="-40%" y="-40%" width="180%" height="180%"><feGaussianBlur stdDeviation="8" result="blur"/><feMerge><feMergeNode in="blur"/><feMergeNode in="blur"/><feMergeNode in="SourceGraphic"/></feMerge></filter>
                <style>{`
                  @keyframes nodeGlow {
                    0%,100% { opacity: 0.25; transform: scale(1); }
                    50%     { opacity: 0.85; transform: scale(1.04); }
                  }
                  @keyframes nodeGlowInner {
                    0%,100% { opacity: 0.5; }
                    50%     { opacity: 1; }
                  }
                  .node-halo-outer { animation: nodeGlow 1.8s ease-in-out infinite; transform-box: fill-box; transform-origin: center; }
                  .node-halo-inner { animation: nodeGlowInner 1.8s ease-in-out infinite; }
                `}</style>
              </defs>

              {/* Auswahlrahmen (screen-Koordinaten, außerhalb der g-Transformation) */}
              {selRect && (() => {
                const x = Math.min(selRect.x1, selRect.x2)
                const y = Math.min(selRect.y1, selRect.y2)
                const svgRect = svgRef.current?.getBoundingClientRect() || { left:0, top:0 }
                return (
                  <rect
                    x={x - svgRect.left} y={y - svgRect.top}
                    width={Math.abs(selRect.x2 - selRect.x1)} height={Math.abs(selRect.y2 - selRect.y1)}
                    fill="rgba(37,99,235,0.06)" stroke="#2563eb" strokeWidth="1.5"
                    strokeDasharray="5 3" style={{ pointerEvents:'none' }}/>
                )
              })()}

              <g transform={`translate(${pan.x},${pan.y}) scale(${zoom})`}>
                {/* Linien - nur für sichtbare nodes */}
                {visibleNodes.map(node => {
                  if (!node.parent_id) return null
                  const p = nodes.find(n => n.id === node.parent_id)
                  if (!p) return null
                  const dx    = node.x - p.x
                  const bend  = Math.max(Math.abs(dx) * 0.7, 90)
                  const c1x   = p.x + bend
                  const c2x   = node.x - bend
                  const color = getColor(node, nodes)
                  const isDrg = node.id === dragging.current
                  return (
                    <path key={`e-${node.id}`}
                      d={`M ${p.x} ${p.y} C ${c1x} ${p.y}, ${c2x} ${node.y}, ${node.x} ${node.y}`}
                      fill="none" stroke={color} strokeWidth={isDrg ? 2.5 : 2}
                      strokeOpacity={isDrg ? 0.6 : 0.4}
                      strokeDasharray={isDrg ? '6 3' : 'none'}/>
                  )
                })}
                {/* Nodes - nur sichtbare */}
                {visibleNodes.map(node => {
                  const color    = getColor(node, nodes)
                  const isRoot   = !node.parent_id
                  const depth    = getDepth(node, nodes)
                  const isSel    = node.id === selectedId
                  const isMultiSel = selectedIds.has(node.id) && selectedIds.size > 1
                  const isDrg    = node.id === dragging.current
                  const isDropT  = node.id === dropTargetId
                  const isCollapsed = !!collapsed[node.id]
                  const childCount  = nodes.filter(n => n.parent_id === node.id).length

                  // Ab Tiefe 3: nur Text, kein Box
                  if (depth >= 3) {
                    return (
                      <g key={node.id} style={{ opacity: isDrg ? 0.6 : 1 }}
                        onMouseDown={e => onNodeMouseDown(e, node.id)}
                        onDoubleClick={e => { e.stopPropagation(); openRename(node) }}
                        onTouchEnd={e => { e.stopPropagation(); setSelectedId(node.id) }}
                        onClick={() => setSelectedId(node.id)}>
                        {(isSel || isMultiSel) && (
                          <rect x={node.x-60} y={node.y-12} width={120} height={24} rx="4"
                            fill={isMultiSel ? '#2563eb' : color} style={{ filter:'blur(8px)', opacity:0.35 }}/>
                        )}
                        {/* Weiße Fläche damit Linie unsichtbar hinter Text */}
                        <rect x={node.x-65} y={node.y-11} width={130} height={22} rx="3" fill="white" stroke="none" style={{ pointerEvents:'none' }}/>
                        <text x={node.x} y={node.y+5} textAnchor="middle"
                          fill={isSel || isMultiSel ? color : '#475569'} fontSize="11" fontWeight={isSel || isMultiSel ? '700' : '400'}
                          style={{ cursor:'pointer', userSelect:'none' }}>
                          {node.label.length > 20 ? node.label.slice(0,19)+'…' : node.label}
                        </text>
                      </g>
                    )
                  }

                  const W = isRoot ? 168 : 150, H = isRoot ? 32 : 28

                  return (
                    <g key={node.id} style={{ opacity: isDrg ? 0.65 : 1 }}>
                      {/* Glow */}
                      {(isSel || isMultiSel) && !isDropT && (
                        <>
                          <rect className="node-halo-outer"
                            x={node.x-W/2-16} y={node.y-H/2-16} width={W+32} height={H+32} rx="22"
                            fill={isMultiSel ? '#2563eb' : color} stroke="none" style={{ filter:'blur(14px)' }}/>
                          <rect className="node-halo-inner"
                            x={node.x-W/2-8} y={node.y-H/2-8} width={W+16} height={H+16} rx="16"
                            fill={isMultiSel ? '#2563eb' : color} stroke="none" style={{ filter:'blur(6px)' }}/>
                        </>
                      )}
                      {/* Drop-Ziel */}
                      {isDropT && (
                        <rect x={node.x-W/2-10} y={node.y-H/2-10} width={W+20} height={H+20} rx="18"
                          fill={color+'18'} stroke={color} strokeWidth="2.5">
                          <animate attributeName="strokeOpacity" values="0.4;1;0.4" dur="0.7s" repeatCount="indefinite"/>
                        </rect>
                      )}
                      {/* Weiße Fläche hinter Linie (damit Linie unsichtbar hinter Knoten) */}
                      <rect x={node.x-W/2} y={node.y-H/2} width={W} height={H} rx="8" fill="white" stroke="none"/>
                      {/* Box */}
                      <rect x={node.x-W/2} y={node.y-H/2} width={W} height={H} rx="8"
                        fill="white"
                        stroke={isDropT ? color : (isSel || isMultiSel) ? color : '#e2e8f0'}
                        strokeWidth={isDropT ? 2.5 : (isSel || isMultiSel) ? 2 : 1}
                        filter={isDrg ? 'url(#shd)' : 'url(#sh)'}
                        style={{ cursor:'grab' }}
                        onMouseDown={e => onNodeMouseDown(e, node.id)}
                        onDoubleClick={e => { e.stopPropagation(); openRename(node) }}
                        onClick={() => setSelectedId(node.id)}
                      />
                      {/* Farbstreifen */}
                      <rect x={node.x-W/2} y={node.y-H/2} width={5} height={H} rx="3" fill={color} style={{ pointerEvents:'none' }}/>
                      {/* Label */}
                      <text x={node.x-W/2+14} y={node.y+5}
                        fill="#1e293b" fontSize={isRoot ? 13 : 12} fontWeight={isRoot ? '700' : '500'}
                        style={{ pointerEvents:'none', userSelect:'none' }}>
                        {node.label.length > 18 ? node.label.slice(0,17)+'…' : node.label}
                      </text>
                      {/* Collapse-Button mit Anzahl */}
                      {childCount > 0 && (
                        <g style={{ cursor:'pointer' }} onClick={e => { e.stopPropagation(); toggleCollapse(node.id) }}>
                          <circle cx={node.x+W/2-10} cy={node.y} r={9} fill={color} fillOpacity={isCollapsed ? 0.9 : 0.15}/>
                          <text x={node.x+W/2-10} y={node.y+4}
                            textAnchor="middle" fill={isCollapsed ? 'white' : color} fontSize="9" fontWeight="700"
                            style={{ pointerEvents:'none' }}>
                            {isCollapsed ? childCount : '▼'}
                          </text>
                        </g>
                      )}
                    </g>
                  )
                })}
              </g>
            </svg>

            {/* Drag Ghost */}
            {dragGhost && dragging.current && (() => {
              const node = nodes.find(n => n.id === dragging.current)
              return node ? (
                <div style={{ position:'fixed', left: dragGhost.x+14, top: dragGhost.y-12,
                  background:'#1e293b', color:'white', fontSize:11, padding:'4px 10px',
                  borderRadius:6, pointerEvents:'none', zIndex:999, whiteSpace:'nowrap',
                  boxShadow:'0 4px 12px rgba(0,0,0,0.25)' }}>
                  {dropTargetId
                    ? `↳ Zu "${nodes.find(n=>n.id===dropTargetId)?.label||'?'}" verschieben`
                    : node.label}
                </div>
              ) : null
            })()}

            <div style={{ position:'absolute', bottom:10, left:10, fontSize:10, color:'#94a3b8',
              background:'rgba(255,255,255,0.88)', padding:'3px 10px', borderRadius:20, pointerEvents:'none' }}>
              Doppelklick = umbenennen · + = Kind · Drag auf Knoten = einordnen
            </div>
          </div>
        )}

        {/* ── Ziehbare Trennlinie Map | Liste ── */}
        {!isMobile && (
          <div
            onMouseDown={e => {
              isResizingList.current = true
              resizeListX.current    = e.clientX
              resizeListW.current    = listWidth
              document.body.style.cursor = 'col-resize'
            }}
            style={{ width:6, background:'#f1f5f9', cursor:'col-resize', flexShrink:0,
              display:'flex', alignItems:'center', justifyContent:'center',
              borderLeft:'1px solid #e2e8f0', borderRight:'1px solid #e2e8f0',
              transition:'background 0.15s' }}
            onMouseEnter={e => e.currentTarget.style.background='#dbeafe'}
            onMouseLeave={e => e.currentTarget.style.background='#f1f5f9'}>
            <div style={{ width:2, height:32, borderRadius:2, background:'#cbd5e1' }}/>
          </div>
        )}

        {/* ── Liste + Kommentar ── */}
        {(!isMobile || activeTab==='list') && (
          <div style={{ width: isMobile ? '100%' : listWidth, flexShrink:0,
            display:'flex', flexDirection:'column', overflow:'hidden' }}>

            {/* Gliederungsliste */}
            <div style={{ flex:1, overflowY:'auto', padding:'12px 12px 0' }}>
              <div style={{ fontSize:10, fontWeight:700, color:'#94a3b8', letterSpacing:1,
                marginBottom:10, textTransform:'uppercase', paddingLeft:4 }}>Gliederung</div>
              {orderedNodes.map(node => {
                const depth  = getDepth(node, nodes)
                const color  = getColor(node, nodes)
                const num    = getOutlineNumber(node, nodes)
                const isSel  = node.id === selectedId
                const isOver = node.id === listDragOver
                const isColl = !!collapsed[node.id]
                const hasKids = nodes.some(n => n.parent_id === node.id)

                // Versteckt wenn ein Vorfahre in der Liste eingeklappt ist
                const isHidden = (() => {
                  let cur = nodes.find(n => n.id === node.parent_id)
                  while (cur) {
                    if (collapsed[cur.id]) return true
                    cur = nodes.find(n => n.id === cur.parent_id)
                  }
                  return false
                })()
                if (isHidden) return null

                return (
                  <div key={node.id}
                    draggable
                    onDragStart={e => onListDragStart(e, node.id)}
                    onDragOver={e => onListDragOver(e, node.id)}
                    onDrop={e => onListDrop(e, node.id)}
                    onDragLeave={() => setListDragOver(null)}
                    onClick={() => setSelectedId(node.id)}
                    onDoubleClick={() => openRename(node)}
                    style={{ display:'flex', alignItems:'center', gap:6, padding:'7px 8px',
                      marginLeft: depth * 16, marginBottom:2, borderRadius:7,
                      background: isOver ? `${color}15` : isSel ? `${color}10` : 'transparent',
                      border: isOver ? `1.5px solid ${color}60`
                            : isSel ? `1.5px solid ${color}30` : '1.5px solid transparent',
                      cursor:'grab', transition:'all 0.12s',
                      boxShadow: isOver ? `0 0 0 3px ${color}20` : 'none' }}>
                    <span style={{ color:'#cbd5e1', fontSize:13, userSelect:'none', flexShrink:0 }}>⠿</span>
                    <span style={{ width:7, height:7, borderRadius:'50%', background:color, flexShrink:0 }}/>
                    <span style={{ fontSize:10, fontWeight:700, color, fontFamily:'monospace', minWidth:28, flexShrink:0 }}>{num}</span>
                    <span style={{ flex:1, fontSize:13, color:'#1e293b',
                      fontWeight: depth===0 ? 600 : 400,
                      whiteSpace:'nowrap', overflow:'hidden', textOverflow:'ellipsis' }}>
                      {node.label}
                    </span>
                    {node.comment && <span style={{ fontSize:10, color:'#94a3b8', flexShrink:0 }}>💬</span>}
                    {/* Collapse-Button in der Liste */}
                    {hasKids && (
                      <button onClick={e => { e.stopPropagation(); toggleCollapse(node.id) }}
                        title={isColl ? 'Aufklappen' : 'Zuklappen'}
                        style={{ background: isColl ? color : 'transparent', border:`1px solid ${color}`,
                          color: isColl ? 'white' : color, cursor:'pointer',
                          fontSize:9, padding:'1px 5px', borderRadius:4, lineHeight:1.4, flexShrink:0, fontWeight:700 }}>
                        {isColl ? `+${nodes.filter(n=>n.parent_id===node.id).length}` : '▾'}
                      </button>
                    )}
                    <button onClick={e => { e.stopPropagation(); addChild(node.id) }}
                      style={{ background:'transparent', border:'none', color, cursor:'pointer',
                        fontSize:16, padding:'0 2px', lineHeight:1, flexShrink:0 }}>+</button>
                    {node.parent_id && (
                      <button onClick={e => { e.stopPropagation(); deleteNode(node.id) }}
                        style={{ background:'transparent', border:'none', color:'#fca5a5',
                          cursor:'pointer', fontSize:14, padding:'0 2px', lineHeight:1, flexShrink:0 }}>×</button>
                    )}
                  </div>
                )
              })}
            </div>

            {/* ── Ziehbare Trennlinie Liste | Kommentar ── */}
            {selectedNode && (
              <>
                <div
                  onMouseDown={e => {
                    isResizingComment.current = true
                    resizeCommentY.current    = e.clientY
                    resizeCommentH.current    = commentHeight
                    document.body.style.cursor = 'row-resize'
                  }}
                  style={{ height:10, cursor:'row-resize', display:'flex',
                    alignItems:'center', justifyContent:'center',
                    borderTop:'1px solid #e2e8f0', flexShrink:0, background:'#fafafa',
                    transition:'background 0.15s' }}
                  onMouseEnter={e => e.currentTarget.style.background='#dbeafe'}
                  onMouseLeave={e => e.currentTarget.style.background='#fafafa'}>
                  <div style={{ width:36, height:3, borderRadius:2, background:'#cbd5e1' }}/>
                </div>

                <div style={{ height: commentHeight, flexShrink:0, display:'flex',
                  flexDirection:'column', padding:'10px 12px', background:'#f8fafc', overflow:'hidden' }}>
                  <div style={{ display:'flex', alignItems:'center', gap:6, marginBottom:8 }}>
                    <span style={{ width:8, height:8, borderRadius:'50%', background:getColor(selectedNode, nodes), flexShrink:0 }}/>
                    <span style={{ fontSize:11, fontWeight:600, color:'#475569', flex:1,
                      overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>
                      {getOutlineNumber(selectedNode, nodes)} · {selectedNode.label}
                    </span>
                    <span style={{ fontSize:10, color: saved ? '#22c55e' : '#f59e0b' }}>
                      {saved ? '✓' : '⟳'}
                    </span>
                  </div>
                  <RichTextEditor
                    value={comment}
                    onChange={handleCommentChange}
                    accentColor={getColor(selectedNode, nodes)}
                  />
                </div>
              </>
            )}
          </div>
        )}
      </div>
    </div>
  )
}
