// src/pages/Editor.jsx
import { useState, useEffect, useRef, useCallback } from 'react'
import { useParams, useNavigate } from 'react-router-dom'
import { supabase } from '../lib/supabase'
import { Document, Packer, Paragraph, TextRun, HeadingLevel } from 'docx'
import { saveAs } from 'file-saver'

// ── Farben je Hauptast ────────────────────────────────────────
const BRANCH_COLORS = [
  '#e05c5c','#e07c3c','#d4a017','#5baa5b',
  '#3b9ec4','#6c6cdc','#b05cbf','#3cb0a0'
]

// ── Hilfsfunktionen ───────────────────────────────────────────
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
// Farbe: jeder Hauptast (depth=1) bekommt eine feste Farbe, alle Kinder erben sie
function getBranchColor(node, all) {
  if (!node.parent_id) return '#1e293b'
  let cur = node
  while (cur.parent_id) {
    const parent = all.find(n => n.id === cur.parent_id)
    if (!parent || !parent.parent_id) break
    cur = parent
  }
  const roots = all.filter(n => !n.parent_id)
  if (roots.length === 0) return BRANCH_COLORS[0]
  const root = roots[0]
  const mainBranches = all.filter(n => n.parent_id === root.id)
  const idx = mainBranches.findIndex(n => n.id === cur.id)
  return BRANCH_COLORS[Math.abs(idx) % BRANCH_COLORS.length]
}
function getOutlineNumber(node, all) {
  if (!node.parent_id) return ''
  const parent = all.find(n => n.id === node.parent_id)
  if (!parent) return '?'
  const parentNum = getOutlineNumber(parent, all)
  const siblings = all.filter(n => n.parent_id === node.parent_id)
  const idx = siblings.findIndex(n => n.id === node.id)
  const num = `${idx + 1}`
  return parentNum ? `${parentNum}.${num}` : num
}
function getOrderedNodes(all) {
  const result = []
  function walk(parentId) {
    all.filter(n => n.parent_id === parentId).forEach(child => { result.push(child); walk(child.id) })
  }
  all.filter(n => !n.parent_id).forEach(root => { result.push(root); walk(root.id) })
  return result
}

// ── Radiales Auto-Layout ──────────────────────────────────────
// Root in der Mitte, Hauptäste links/rechts aufgeteilt, Unterknoten als Zeilen
function radialLayout(nodes, cx, cy) {
  if (nodes.length === 0) return nodes
  const positioned = {}
  const root = nodes.find(n => !n.parent_id)
  if (!root) return nodes

  positioned[root.id] = { x: cx, y: cy }

  const mainBranches = nodes.filter(n => n.parent_id === root.id)
  const half = Math.ceil(mainBranches.length / 2)
  const leftBranches  = mainBranches.slice(0, half)
  const rightBranches = mainBranches.slice(half)

  const MAIN_X_OFFSET = 220   // Abstand Root -> Hauptast
  const MAIN_Y_GAP    = 120   // Abstand zwischen Hauptästen
  const SUB_X_STEP    = 160   // Horizontaler Schritt pro Tiefe
  const SUB_Y_GAP     = 32    // Vertikaler Abstand zwischen Unterknoten

  // Berechne Gesamthöhe eines Subtrees (Anzahl Blätter * GAP)
  function subtreeH(id) {
    const ch = nodes.filter(n => n.parent_id === id)
    if (ch.length === 0) return SUB_Y_GAP
    return ch.reduce((s, c) => s + subtreeH(c.id), 0)
  }

  function placeBranch(branchNode, baseX, side) {
    // Hauptast-Position
    const mainBranchIdx = mainBranches.indexOf(branchNode)
    // wird weiter unten gesetzt

    // Rekursiv Unterknoten platzieren
    function placeChildren(nodeId, depth, yStart) {
      const children = nodes.filter(n => n.parent_id === nodeId)
      const totalH = children.reduce((s, c) => s + subtreeH(c.id), 0)
      const parentY = positioned[nodeId]?.y ?? cy
      let curY = parentY - totalH / 2

      children.forEach(child => {
        const h = subtreeH(child.id)
        const childY = curY + h / 2
        const childX = baseX + (side === 'right' ? depth * SUB_X_STEP : -depth * SUB_X_STEP)
        positioned[child.id] = { x: childX, y: childY }
        placeChildren(child.id, depth + 1, curY)
        curY += h
      })
    }

    placeChildren(branchNode.id, 1, positioned[branchNode.id]?.y ?? cy)
  }

  // Platziere linke Äste
  const leftTotalH = leftBranches.reduce((s, b) => s + subtreeH(b.id), 0)
  let leftY = cy - leftTotalH / 2
  leftBranches.forEach(b => {
    const h = subtreeH(b.id)
    positioned[b.id] = { x: cx - MAIN_X_OFFSET, y: leftY + h / 2 }
    leftY += h
    placeBranch(b, cx - MAIN_X_OFFSET, 'left')
  })

  // Platziere rechte Äste
  const rightTotalH = rightBranches.reduce((s, b) => s + subtreeH(b.id), 0)
  let rightY = cy - rightTotalH / 2
  rightBranches.forEach(b => {
    const h = subtreeH(b.id)
    positioned[b.id] = { x: cx + MAIN_X_OFFSET, y: rightY + h / 2 }
    rightY += h
    placeBranch(b, cx + MAIN_X_OFFSET, 'right')
  })

  return nodes.map(n => positioned[n.id]
    ? { ...n, x: positioned[n.id].x, y: positioned[n.id].y }
    : n
  )
}

const MAX_UNDO = 5

// ── Rich Text Editor (Quill) ──────────────────────────────────
function RichTextEditor({ value, onChange, accentColor }) {
  const containerRef = useRef(null)
  const quillRef     = useRef(null)
  const isInternal   = useRef(false)

  useEffect(() => {
    if (!document.getElementById('quill-css')) {
      const link = document.createElement('link')
      link.id = 'quill-css'; link.rel = 'stylesheet'
      link.href = 'https://cdn.quilljs.com/1.3.7/quill.snow.css'
      document.head.appendChild(link)
    }
    const init = () => {
      if (!containerRef.current || quillRef.current) return
      const Q = window.Quill; if (!Q) return
      quillRef.current = new Q(containerRef.current, {
        theme: 'snow', placeholder: 'Kommentar schreiben...',
        modules: { toolbar: [['bold','italic','underline','strike'],[{list:'ordered'},{list:'bullet'}],['clean']] }
      })
      if (value) quillRef.current.clipboard.dangerouslyPasteHTML(value)
      quillRef.current.on('text-change', () => {
        isInternal.current = true
        const h = quillRef.current.root.innerHTML
        onChange(h === '<p><br></p>' ? '' : h)
        setTimeout(() => { isInternal.current = false }, 0)
      })
    }
    if (window.Quill) { init() } else {
      const s = document.createElement('script')
      s.src = 'https://cdn.quilljs.com/1.3.7/quill.min.js'
      s.onload = init; document.head.appendChild(s)
    }
    return () => { quillRef.current = null }
  }, [])

  useEffect(() => {
    if (!quillRef.current || isInternal.current) return
    const cur = quillRef.current.root.innerHTML
    const norm = cur === '<p><br></p>' ? '' : cur
    if (norm !== value) quillRef.current.clipboard.dangerouslyPasteHTML(value || '')
  }, [value])

  return (
    <div style={{ flex:1, display:'flex', flexDirection:'column', overflow:'hidden', minHeight:0 }}>
      <style>{`
        .ql-toolbar{border-radius:8px 8px 0 0!important;border-color:#e2e8f0!important;background:white;padding:4px 8px!important;flex-shrink:0}
        .ql-container{border-radius:0 0 8px 8px!important;border-color:#e2e8f0!important;flex:1;overflow-y:auto;font-family:inherit!important;font-size:13px!important}
        .ql-editor{min-height:60px;line-height:1.7;color:#334155;padding:10px 12px}
        .ql-editor.ql-blank::before{color:#94a3b8;font-style:normal}
        .ql-toolbar button:hover .ql-stroke,.ql-toolbar button.ql-active .ql-stroke{stroke:${accentColor}!important}
        .ql-toolbar button:hover .ql-fill,.ql-toolbar button.ql-active .ql-fill{fill:${accentColor}!important}
      `}</style>
      <div style={{ flex:1, display:'flex', flexDirection:'column', overflow:'hidden', border:'1.5px solid #e2e8f0', borderRadius:8 }}>
        <div ref={containerRef} style={{ flex:1, display:'flex', flexDirection:'column', overflow:'hidden' }}/>
      </div>
    </div>
  )
}

// ── Icon Button ───────────────────────────────────────────────
function IconBtn({ icon, label, onClick, color='#475569', disabled, active, primary }) {
  const [hover, setHover] = useState(false)
  return (
    <button onClick={onClick} disabled={disabled} title={label}
      onMouseEnter={() => setHover(true)} onMouseLeave={() => setHover(false)}
      style={{
        display:'flex', flexDirection:'column', alignItems:'center', justifyContent:'center',
        gap:2, width:48, height:48, borderRadius:10,
        border: active ? `2px solid ${color}` : '1px solid transparent',
        background: primary ? color : active ? color+'18' : hover && !disabled ? color+'12' : 'transparent',
        color: primary ? 'white' : disabled ? '#cbd5e1' : color,
        cursor: disabled ? 'default' : 'pointer',
        transition:'all 0.12s', flexShrink:0, opacity: disabled ? 0.4 : 1,
        padding:0
      }}>
      <span style={{ fontSize:18, lineHeight:1 }}>{icon}</span>
      <span style={{ fontSize:8, fontWeight:600, letterSpacing:0.3, textTransform:'uppercase', lineHeight:1 }}>{label}</span>
    </button>
  )
}

// ── Divider ───────────────────────────────────────────────────
function Divider() {
  return <div style={{ width:1, height:36, background:'#e2e8f0', flexShrink:0, margin:'0 2px' }}/>
}

export default function Editor() {
  const { mapId } = useParams()
  const navigate  = useNavigate()

  const [nodes, setNodes]       = useState([])
  const [mapTitle, setMapTitle] = useState('')
  const [selectedId, setSelectedId] = useState(null)
  const [comment, setComment]   = useState('')
  const [saved, setSaved]       = useState(true)
  const [pan, setPan]           = useState({ x: 0, y: 0 })
  const panRef    = useRef({ x: 0, y: 0 })
  const nodesRef  = useRef([])
  const [collapsed, setCollapsed] = useState({})
  const [isMobile, setIsMobile] = useState(window.innerWidth < 768)
  const [activeTab, setActiveTab] = useState('map')

  // Rename popup
  const [renameId, setRenameId]     = useState(null)
  const [renameText, setRenameText] = useState('')
  const renameInputRef = useRef(null)

  // Undo
  const undoStack = useRef([])
  const pushUndo  = useCallback((s) => {
    undoStack.current = [...undoStack.current.slice(-MAX_UNDO+1), JSON.parse(JSON.stringify(s))]
  }, [])

  // Canvas drag
  const dragging  = useRef(null)
  const dragOff   = useRef({ x:0, y:0 })
  const didDrag   = useRef(false)
  const isPanning = useRef(false)
  const panStart  = useRef({ x:0, y:0 })
  const svgRef    = useRef()
  const svgContainerRef = useRef()

  // Drop target
  const [dropTargetId, setDropTargetId] = useState(null)
  const [dragGhost, setDragGhost]       = useState(null)

  // List drag
  const dragListItem = useRef(null)
  const [listDragOver, setListDragOver] = useState(null)

  // Resizable panels
  const [listWidth, setListWidth]       = useState(300)
  const [commentHeight, setCommentHeight] = useState(200)
  const isResizingList    = useRef(false)
  const isResizingComment = useRef(false)
  const resizeListX       = useRef(0); const resizeListW = useRef(0)
  const resizeCommentY    = useRef(0); const resizeCommentH = useRef(0)
  const commentTimer      = useRef(null)

  useEffect(() => {
    loadMap()
    const onResize = () => setIsMobile(window.innerWidth < 768)
    window.addEventListener('resize', onResize)
    return () => window.removeEventListener('resize', onResize)
  }, [mapId])

  useEffect(() => {
    if (!selectedId) { setComment(''); return }
    setComment(nodes.find(n => n.id === selectedId)?.comment || '')
  }, [selectedId])

  useEffect(() => {
    if (renameId && renameInputRef.current)
      setTimeout(() => { renameInputRef.current?.focus(); renameInputRef.current?.select() }, 30)
  }, [renameId])

  // Global resize + undo
  useEffect(() => {
    const onMove = (e) => {
      if (isResizingList.current)
        setListWidth(Math.max(200, Math.min(600, resizeListW.current + (resizeListX.current - e.clientX))))
      if (isResizingComment.current)
        setCommentHeight(Math.max(80, Math.min(500, resizeCommentH.current + (resizeCommentY.current - e.clientY))))
    }
    const onUp = () => {
      isResizingList.current = false
      isResizingComment.current = false
      document.body.style.cursor = ''
    }
    const onKey = (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key === 'z' && !renameId) {
        e.preventDefault()
        if (!undoStack.current.length) return
        const prev = undoStack.current.pop()
        setNodes(prev)
        supabase.from('nodes').delete().eq('map_id', mapId)
          .then(() => supabase.from('nodes').insert(prev))
        setSaved(true)
      }
    }
    window.addEventListener('mousemove', onMove)
    window.addEventListener('mouseup', onUp)
    window.addEventListener('keydown', onKey)
    return () => {
      window.removeEventListener('mousemove', onMove)
      window.removeEventListener('mouseup', onUp)
      window.removeEventListener('keydown', onKey)
    }
  }, [renameId, mapId])

  const getCenterXY = () => {
    const el = svgContainerRef.current
    if (!el) return { cx: 500, cy: 350 }
    return { cx: el.clientWidth / 2, cy: el.clientHeight / 2 }
  }

  const loadMap = async () => {
    const { data: map } = await supabase.from('maps').select('title').eq('id', mapId).single()
    if (map) setMapTitle(map.title)
    const { data } = await supabase.from('nodes').select('*').eq('map_id', mapId)
    if (data && data.length > 0) {
      setNodes(data)
    } else {
      const { cx, cy } = getCenterXY()
      const root = { id: crypto.randomUUID(), map_id: mapId, label: map?.title || 'Start', x: cx, y: cy, parent_id: null, comment: '' }
      setNodes([root])
      await supabase.from('nodes').insert(root)
    }
  }

  const applyLayout = async (ns) => {
    const { cx, cy } = getCenterXY()
    const laid = radialLayout(ns, cx, cy)
    setNodes(laid)
    for (const n of laid) await supabase.from('nodes').update({ x: n.x, y: n.y }).eq('id', n.id)
    return laid
  }

  // ── Rename ────────────────────────────────────────────────
  const openRename = (node) => { setSelectedId(node.id); setRenameId(node.id); setRenameText(node.label) }
  const confirmRename = async () => {
    if (!renameId) return
    const t = renameText.trim() || 'Punkt'
    pushUndo(nodes)
    setNodes(prev => prev.map(n => n.id === renameId ? { ...n, label: t } : n))
    setRenameId(null); setSaved(false)
    await supabase.from('nodes').update({ label: t }).eq('id', renameId)
    setSaved(true)
  }

  // ── Kind hinzufügen ───────────────────────────────────────
  const addChild = async (parentId) => {
    pushUndo(nodes)
    const newNode = { id: crypto.randomUUID(), map_id: mapId, label: 'Neuer Punkt', x: 0, y: 0, parent_id: parentId, comment: '' }
    const next = [...nodes, newNode]
    await supabase.from('nodes').insert(newNode)
    const laid = await applyLayout(next)
    setSelectedId(newNode.id)
    setTimeout(() => openRename(laid.find(n => n.id === newNode.id) || newNode), 80)
  }

  // ── Root-Knoten hinzufügen ────────────────────────────────
  const addRoot = async () => {
    pushUndo(nodes)
    const { cx, cy } = getCenterXY()
    const newNode = { id: crypto.randomUUID(), map_id: mapId, label: 'Neues Thema', x: cx+300, y: cy+200, parent_id: null, comment: '' }
    setNodes(prev => [...prev, newNode])
    await supabase.from('nodes').insert(newNode)
    setSelectedId(newNode.id)
    setTimeout(() => openRename(newNode), 80)
  }

  // ── Löschen ───────────────────────────────────────────────
  const deleteNode = async (nodeId) => {
    pushUndo(nodes)
    const del = getDesc(nodeId, nodes)
    const next = nodes.filter(n => !del.includes(n.id))
    if (del.includes(selectedId)) setSelectedId(null)
    await supabase.from('nodes').delete().in('id', del)
    await applyLayout(next)
  }

  // ── Collapse ──────────────────────────────────────────────
  const toggleCollapse = (id) => setCollapsed(prev => ({ ...prev, [id]: !prev[id] }))

  const getVisibleNodes = useCallback((all) => {
    const hidden = new Set()
    all.forEach(n => {
      if (!n.parent_id) return
      let cur = n
      while (cur.parent_id) {
        if (collapsed[cur.parent_id]) { hidden.add(n.id); break }
        cur = all.find(x => x.id === cur.parent_id) || {}
      }
    })
    return all.filter(n => !hidden.has(n.id))
  }, [collapsed])

  // ── Kommentar ─────────────────────────────────────────────
  const handleCommentChange = (val) => {
    setComment(val)
    if (!selectedId) return
    clearTimeout(commentTimer.current)
    commentTimer.current = setTimeout(async () => {
      setNodes(prev => prev.map(n => n.id === selectedId ? { ...n, comment: val } : n))
      setSaved(false)
      await supabase.from('nodes').update({ comment: val }).eq('id', selectedId)
      setSaved(true)
    }, 800)
  }

  // nodesRef + panRef: stale-closure-freier Zugriff in Event-Handlern
  useEffect(() => { nodesRef.current = nodes }, [nodes])
  useEffect(() => { panRef.current = pan }, [pan])

  // ── Canvas drag ───────────────────────────────────────────
  const dropTargetIdRef = useRef(null)
  useEffect(() => { dropTargetIdRef.current = dropTargetId }, [dropTargetId])

  const onNodeDown = (e, id) => {
    e.stopPropagation()
    setSelectedId(id)
    const node = nodesRef.current.find(n => n.id === id)
    dragging.current = id; didDrag.current = false
    dragOff.current = { x: e.clientX - (node.x + panRef.current.x), y: e.clientY - (node.y + panRef.current.y) }
  }
  const onSvgDown = (e) => {
    if (e.target === svgRef.current || e.target.tagName === 'svg') {
      isPanning.current = true
      panStart.current  = { x: e.clientX - pan.x, y: e.clientY - pan.y }
      setSelectedId(null)
    }
  }
  const onMove = useCallback((e) => {
    if (dragging.current) {
      didDrag.current = true
      const nx = e.clientX - dragOff.current.x - panRef.current.x
      const ny = e.clientY - dragOff.current.y - panRef.current.y
      setNodes(prev => prev.map(n => n.id === dragging.current ? { ...n, x: nx, y: ny } : n))
      setDragGhost({ x: e.clientX, y: e.clientY })
      const W_BOX = 150, H_BOX = 50
      const others = nodesRef.current.filter(n => n.id !== dragging.current)
      let hit = null
      others.forEach(n => {
        if (nx > n.x-W_BOX/2 && nx < n.x+W_BOX/2 && ny > n.y-H_BOX/2 && ny < n.y+H_BOX/2) hit = n.id
      })
      setDropTargetId(hit)
    } else if (isPanning.current) {
      setPan({ x: e.clientX - panStart.current.x, y: e.clientY - panStart.current.y })
    }
  }, [])

  const onUp = useCallback(async () => {
    if (dragging.current) {
      const id = dragging.current
      const node = nodesRef.current.find(n => n.id === id)
      if (didDrag.current) {
        pushUndo(nodesRef.current)
        if (dropTargetIdRef.current) {
          const next = nodesRef.current.map(n => n.id === id ? { ...n, parent_id: dropTargetIdRef.current } : n)
          await supabase.from('nodes').update({ parent_id: dropTargetIdRef.current }).eq('id', id)
          await applyLayout(next)
        } else if (node) {
          setSaved(false)
          await supabase.from('nodes').update({ x: node.x, y: node.y }).eq('id', id)
          setSaved(true)
        }
      }
      dragging.current = null; didDrag.current = false
      setDropTargetId(null); setDragGhost(null)
    }
    isPanning.current = false
  }, [pushUndo])

  // ── List drag ─────────────────────────────────────────────
  const onListDragStart = (e, id) => { dragListItem.current = id; e.dataTransfer.effectAllowed = 'move' }
  const onListDragOver  = (e, id) => { e.preventDefault(); setListDragOver(id) }
  const onListDrop      = async (e, targetId) => {
    e.preventDefault(); setListDragOver(null)
    const srcId = dragListItem.current
    if (!srcId || srcId === targetId) return
    const target = nodes.find(n => n.id === targetId); if (!target) return
    pushUndo(nodes)
    const next = nodes.map(n => n.id === srcId ? { ...n, parent_id: target.parent_id } : n)
    await supabase.from('nodes').update({ parent_id: target.parent_id }).eq('id', srcId)
    await applyLayout(next)
  }

  // ── Word Export ───────────────────────────────────────────
  const exportToWord = async () => {
    const ordered = getOrderedNodes(nodes)
    const children = [
      new Paragraph({ text: mapTitle, heading: HeadingLevel.TITLE, spacing: { after: 400 } })
    ]
    ordered.forEach(node => {
      const depth = getDepth(node, nodes)
      const num   = getOutlineNumber(node, nodes)
      const hMap  = [HeadingLevel.HEADING_1, HeadingLevel.HEADING_2, HeadingLevel.HEADING_3]
      children.push(new Paragraph({
        children: [
          new TextRun({ text: num ? num+'  ' : '', bold: true, color: '2563eb', size: Math.max(24-depth*2,18) }),
          new TextRun({ text: node.label, bold: depth===0, size: Math.max(24-depth*2,18) })
        ],
        heading: hMap[Math.min(depth,2)],
        indent: { left: depth*360 }, spacing: { before: depth===0?300:100, after:80 }
      }))
      if (node.comment) {
        const plain = node.comment.replace(/<[^>]*>/g,'').replace(/&nbsp;/g,' ').trim()
        if (plain) children.push(new Paragraph({
          children: [new TextRun({ text: plain, italics: true, color: '64748b', size: 18 })],
          indent: { left: depth*360+360 }, spacing: { after: 60 }
        }))
      }
    })
    const doc  = new Document({ sections: [{ properties:{}, children }] })
    const blob = await Packer.toBlob(doc)
    saveAs(blob, `${mapTitle||'MindMap'}.docx`)
  }

  // ── Computed ──────────────────────────────────────────────
  const selectedNode  = nodes.find(n => n.id === selectedId)
  const orderedNodes  = getOrderedNodes(nodes)
  const visibleNodes  = getVisibleNodes(nodes)
  const canUndo       = undoStack.current.length > 0
  const hasChildren   = (id) => nodes.some(n => n.parent_id === id)
  const isCollapsed   = (id) => !!collapsed[id]

  // ── SVG Node rendering ────────────────────────────────────
  const renderNode = (node) => {
    const color    = getBranchColor(node, nodes)
    const depth    = getDepth(node, nodes)
    const isRoot   = depth === 0
    const isSel    = node.id === selectedId
    const isDrg    = node.id === dragging.current
    const isDropT  = node.id === dropTargetId
    const isColl   = isCollapsed(node.id)
    const hasKids  = hasChildren(node.id)

    // Root: großer gefüllter Kreis mit Text
    if (isRoot) {
      const R = 54
      return (
        <g key={node.id} style={{ opacity: isDrg ? 0.7 : 1 }}>
          {isSel && (
            <>
              <circle className="node-halo-outer" cx={node.x} cy={node.y} r={R+18}
                fill={color} style={{ filter:'blur(16px)' }}/>
              <circle className="node-halo-inner" cx={node.x} cy={node.y} r={R+9}
                fill={color} style={{ filter:'blur(7px)' }}/>
            </>
          )}
          <circle cx={node.x} cy={node.y} r={R}
            fill="white" stroke={isSel ? color : '#e2e8f0'} strokeWidth={isSel ? 2.5 : 1.5}
            filter="url(#sh)" style={{ cursor:'grab' }}
            onMouseDown={e => onNodeDown(e, node.id)}
            onDoubleClick={e => { e.stopPropagation(); openRename(node) }}
            onClick={() => setSelectedId(node.id)}/>
          <text x={node.x} y={node.y+5} textAnchor="middle"
            fill="#1e293b" fontSize="15" fontWeight="700"
            style={{ pointerEvents:'none', userSelect:'none' }}>
            {node.label.length > 14 ? node.label.slice(0,13)+'…' : node.label}
          </text>
        </g>
      )
    }

    // Hauptast (depth=1): abgerundetes Rechteck mit Farbe
    if (depth === 1) {
      const W = Math.max(node.label.length * 8 + 32, 80), H = 36
      return (
        <g key={node.id} style={{ opacity: isDrg ? 0.7 : 1 }}>
          {isSel && (
            <>
              <rect className="node-halo-outer"
                x={node.x-W/2-14} y={node.y-H/2-14} width={W+28} height={H+28} rx="20"
                fill={color} style={{ filter:'blur(14px)' }}/>
              <rect className="node-halo-inner"
                x={node.x-W/2-7} y={node.y-H/2-7} width={W+14} height={H+14} rx="14"
                fill={color} style={{ filter:'blur(6px)' }}/>
            </>
          )}
          {isDropT && (
            <rect x={node.x-W/2-8} y={node.y-H/2-8} width={W+16} height={H+16} rx="16"
              fill={color+'22'} stroke={color} strokeWidth="2.5">
              <animate attributeName="strokeOpacity" values="0.3;1;0.3" dur="0.8s" repeatCount="indefinite"/>
            </rect>
          )}
          <rect x={node.x-W/2} y={node.y-H/2} width={W} height={H} rx="10"
            fill={color+'22'} stroke={color} strokeWidth="1.5"
            style={{ cursor:'grab' }}
            onMouseDown={e => onNodeDown(e, node.id)}
            onDoubleClick={e => { e.stopPropagation(); openRename(node) }}
            onClick={() => setSelectedId(node.id)}/>
          <text x={node.x} y={node.y+5} textAnchor="middle"
            fill={color} fontSize="13" fontWeight="700"
            style={{ pointerEvents:'none', userSelect:'none' }}>
            {node.label.length > 18 ? node.label.slice(0,17)+'…' : node.label}
          </text>
          {/* Collapse Button */}
          {hasKids && (
            <g style={{ cursor:'pointer' }} onClick={e => { e.stopPropagation(); toggleCollapse(node.id) }}>
              <circle cx={node.x+W/2+12} cy={node.y} r={9} fill={color} fillOpacity="0.2" stroke={color} strokeWidth="1"/>
              <text x={node.x+W/2+12} y={node.y+4} textAnchor="middle"
                fill={color} fontSize="9" fontWeight="700" style={{ pointerEvents:'none' }}>
                {isColl ? '▶' : '▼'}
              </text>
            </g>
          )}
        </g>
      )
    }

    // Tiefe 2+: nur Text mit Unterstrich-Linie
    const textLen = node.label.length * 7 + 10
    const root    = nodes.find(n => !n.parent_id)
    const isRight = root ? node.x > root.x : true

    return (
      <g key={node.id} style={{ opacity: isDrg ? 0.6 : 1 }}
        onMouseDown={e => onNodeDown(e, node.id)}
        onDoubleClick={e => { e.stopPropagation(); openRename(node) }}
        onClick={() => setSelectedId(node.id)}>
        {isSel && (
          <rect x={isRight ? node.x-4 : node.x-textLen+4} y={node.y-13}
            width={textLen} height={22} rx="4"
            fill={color} style={{ filter:'blur(7px)', opacity:0.4 }}/>
        )}
        {/* Unterstrich */}
        <line
          x1={isRight ? node.x-4 : node.x-textLen+4}
          y1={node.y+10}
          x2={isRight ? node.x+textLen-4 : node.x+4}
          y2={node.y+10}
          stroke={color} strokeWidth="1" strokeOpacity="0.4"/>
        <text
          x={isRight ? node.x : node.x}
          y={node.y+4}
          textAnchor={isRight ? 'start' : 'end'}
          fill={isSel ? color : '#334155'}
          fontSize={depth===2 ? 12 : 11}
          fontWeight={isSel ? '600' : depth===2 ? '500' : '400'}
          style={{ cursor:'pointer', userSelect:'none' }}>
          {node.label.length > 22 ? node.label.slice(0,21)+'…' : node.label}
        </text>
        {/* Kleine Collapse-Markierung */}
        {hasKids && (
          <g style={{ cursor:'pointer' }} onClick={e => { e.stopPropagation(); toggleCollapse(node.id) }}>
            <text x={isRight ? node.x-8 : node.x+8} y={node.y+4}
              textAnchor={isRight ? 'end' : 'start'}
              fill={color} fontSize="9" style={{ pointerEvents:'none' }}>
              {isColl ? '▶' : '▾'}
            </text>
          </g>
        )}
      </g>
    )
  }

  // ── Verbindungslinien ─────────────────────────────────────
  const renderEdge = (node) => {
    if (!node.parent_id) return null
    const p = nodes.find(n => n.id === node.parent_id)
    if (!p) return null
    const color  = getBranchColor(node, nodes)
    const depth  = getDepth(node, nodes)
    const isDrg  = node.id === dragging.current
    const root   = nodes.find(n => !n.parent_id)
    const isRight = root ? p.x >= root.x : true

    // Root -> Hauptast: S-Kurve
    if (depth === 1) {
      const bend = Math.abs(node.x - p.x) * 0.6
      const c1x = p.x + (isRight ? bend : -bend)
      const c2x = node.x + (isRight ? -bend : bend)
      return (
        <path key={`e-${node.id}`}
          d={`M ${p.x} ${p.y} C ${c1x} ${p.y}, ${c2x} ${node.y}, ${node.x} ${node.y}`}
          fill="none" stroke={color} strokeWidth="2.5" strokeOpacity={isDrg ? 0.6 : 0.5}/>
      )
    }

    // Tiefe 2+: horizontale Linie zum Text
    const lineEndX = isRight ? node.x - 4 : node.x + 4
    const parentEndX = isRight ? p.x + Math.max(p.label.length * 8/2, 40) : p.x - Math.max(p.label.length * 8/2, 40)

    return (
      <path key={`e-${node.id}`}
        d={`M ${parentEndX} ${p.y} C ${parentEndX + (isRight?40:-40)} ${p.y}, ${lineEndX + (isRight?-40:40)} ${node.y}, ${lineEndX} ${node.y}`}
        fill="none" stroke={color} strokeWidth={depth===2 ? 1.5 : 1}
        strokeOpacity={isDrg ? 0.6 : 0.35}/>
    )
  }

  return (
    <div style={{ width:'100vw', height:'100vh', display:'flex', flexDirection:'column',
      background:'white', fontFamily:'"DM Sans", system-ui, sans-serif', userSelect:'none' }}
      onMouseMove={onMove} onMouseUp={onUp}>

      {/* ── Rename Popup ── */}
      {renameId && (
        <div style={{ position:'fixed', inset:0, background:'rgba(15,23,42,0.4)',
          display:'flex', alignItems:'center', justifyContent:'center', zIndex:1000 }}>
          <div style={{ background:'white', borderRadius:14, padding:28, width:340,
            boxShadow:'0 20px 60px rgba(0,0,0,0.2)' }} onMouseDown={e => e.stopPropagation()}>
            <div style={{ fontSize:15, fontWeight:700, color:'#1e293b', marginBottom:4 }}>Umbenennen</div>
            <div style={{ fontSize:12, color:'#94a3b8', marginBottom:14 }}>Enter = Speichern · Escape = Abbrechen</div>
            <input ref={renameInputRef} value={renameText}
              onChange={e => setRenameText(e.target.value)}
              onKeyDown={e => { if (e.key==='Enter') confirmRename(); if (e.key==='Escape') setRenameId(null) }}
              style={{ width:'100%', padding:'10px 14px', fontSize:14, border:'2px solid #2563eb',
                borderRadius:8, outline:'none', color:'#1e293b', fontFamily:'inherit', boxSizing:'border-box' }}/>
            <div style={{ display:'flex', gap:8, marginTop:16, justifyContent:'flex-end' }}>
              <button onClick={() => setRenameId(null)}
                style={{ padding:'8px 18px', border:'1px solid #e2e8f0', borderRadius:7,
                  background:'white', color:'#64748b', cursor:'pointer', fontSize:13 }}>Abbrechen</button>
              <button onClick={confirmRename}
                style={{ padding:'8px 20px', border:'none', borderRadius:7,
                  background:'#2563eb', color:'white', cursor:'pointer', fontSize:13, fontWeight:600 }}>Speichern</button>
            </div>
          </div>
        </div>
      )}

      {/* ── Header ── */}
      <div style={{ display:'flex', alignItems:'center', gap:8, padding:'0 12px',
        height:50, borderBottom:'1px solid #e2e8f0', background:'white', flexShrink:0, zIndex:10 }}>
        <button onClick={() => navigate('/dashboard')}
          style={{ background:'transparent', border:'1px solid #e2e8f0', color:'#64748b',
            padding:'5px 10px', borderRadius:6, cursor:'pointer', fontSize:12, whiteSpace:'nowrap' }}>
          ←
        </button>
        <span style={{ fontWeight:700, fontSize:14, color:'#1e293b', flex:1,
          overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{mapTitle}</span>
        <span style={{ fontSize:11, color: saved ? '#22c55e' : '#f59e0b', fontWeight:600, flexShrink:0 }}>
          {saved ? '✓' : '⟳'}
        </span>
        {isMobile && (
          <div style={{ display:'flex', gap:3, background:'#f1f5f9', borderRadius:7, padding:3 }}>
            {[['map','🗺'],['list','☰']].map(([tab,icon]) => (
              <button key={tab} onClick={() => setActiveTab(tab)}
                style={{ padding:'4px 10px', borderRadius:5, border:'none', cursor:'pointer', fontSize:12,
                  background: activeTab===tab ? 'white' : 'transparent',
                  color: activeTab===tab ? '#2563eb' : '#64748b', fontWeight: activeTab===tab ? 600 : 400,
                  boxShadow: activeTab===tab ? '0 1px 3px rgba(0,0,0,0.1)' : 'none' }}>
                {icon}
              </button>
            ))}
          </div>
        )}
      </div>

      {/* ── Body ── */}
      <div style={{ flex:1, display:'flex', overflow:'hidden', minHeight:0 }}>

        {/* ── Vertikale Icon-Toolbar ── */}
        <div style={{ width:64, flexShrink:0, display:'flex', flexDirection:'column',
          alignItems:'center', gap:4, padding:'10px 0',
          background:'white', borderRight:'1px solid #e2e8f0', overflowY:'auto' }}>

          {/* Allgemein */}
          <IconBtn icon="⊞" label="Layout" onClick={() => applyLayout(nodes)} color="#475569"/>
          <IconBtn icon="↩" label="Undo" onClick={() => {
            if (!canUndo) return
            const prev = undoStack.current.pop()
            setNodes(prev)
            supabase.from('nodes').delete().eq('map_id', mapId).then(() => supabase.from('nodes').insert(prev))
            setSaved(true)
          }} color="#475569" disabled={!canUndo}/>
          <IconBtn icon="📄" label="Word" onClick={exportToWord} color="#2563eb" primary/>

          <Divider/>

          {/* Knoten hinzufügen */}
          <IconBtn icon="🌱" label="Neu" onClick={addRoot} color="#16a34a"
            active={false}/>
          {selectedNode && (
            <IconBtn icon="➕" label="Kind" onClick={() => addChild(selectedNode.id)} color="#16a34a"/>
          )}

          <Divider/>

          {/* Ausgewählter Knoten */}
          {selectedNode ? (
            <>
              <IconBtn icon="✏️" label="Name" onClick={() => openRename(selectedNode)} color="#475569"/>
              {hasChildren(selectedNode.id) && (
                <IconBtn icon={isCollapsed(selectedNode.id) ? '▶' : '▼'} label="Klapp"
                  onClick={() => toggleCollapse(selectedNode.id)} color="#7c3aed"
                  active={isCollapsed(selectedNode.id)}/>
              )}
              {selectedNode.parent_id && (
                <IconBtn icon="🗑" label="Löschen" onClick={() => deleteNode(selectedNode.id)} color="#dc2626"/>
              )}
            </>
          ) : (
            <span style={{ fontSize:9, color:'#cbd5e1', textAlign:'center', padding:'0 4px', lineHeight:1.4 }}>
              Knoten<br/>wählen
            </span>
          )}
        </div>

        {/* ── Mindmap ── */}
        {(!isMobile || activeTab==='map') && (
          <div ref={svgContainerRef} style={{ flex:1, position:'relative', overflow:'hidden', minWidth:0 }}>
            <svg ref={svgRef}
              style={{ width:'100%', height:'100%', background:'#f9fafb', cursor:'default', display:'block' }}
              onMouseDown={onSvgDown}>
              <defs>
                <filter id="sh"><feDropShadow dx="0" dy="2" stdDeviation="5" floodOpacity="0.08"/></filter>
                <filter id="shd"><feDropShadow dx="0" dy="6" stdDeviation="12" floodOpacity="0.15"/></filter>
                <style>{`
                  @keyframes nodeGlow { 0%,100%{opacity:0.2;} 50%{opacity:0.75;} }
                  @keyframes nodeGlowInner { 0%,100%{opacity:0.4;} 50%{opacity:0.9;} }
                  .node-halo-outer{animation:nodeGlow 1.8s ease-in-out infinite;transform-box:fill-box;transform-origin:center;}
                  .node-halo-inner{animation:nodeGlowInner 1.8s ease-in-out infinite;}
                `}</style>
              </defs>
              <g transform={`translate(${pan.x},${pan.y})`}>
                {visibleNodes.map(n => renderEdge(n))}
                {visibleNodes.map(n => renderNode(n))}
              </g>
            </svg>

            {/* Drag ghost */}
            {dragGhost && dragging.current && (() => {
              const n = nodes.find(x => x.id === dragging.current)
              return n ? (
                <div style={{ position:'fixed', left:dragGhost.x+14, top:dragGhost.y-12,
                  background:'#1e293b', color:'white', fontSize:11, padding:'4px 10px',
                  borderRadius:6, pointerEvents:'none', zIndex:999, whiteSpace:'nowrap',
                  boxShadow:'0 4px 12px rgba(0,0,0,0.25)' }}>
                  {dropTargetId ? `↳ Zu "${nodes.find(x=>x.id===dropTargetId)?.label||'?'}"` : n.label}
                </div>
              ) : null
            })()}

            <div style={{ position:'absolute', bottom:10, left:10, fontSize:10, color:'#94a3b8',
              background:'rgba(255,255,255,0.88)', padding:'3px 10px', borderRadius:20, pointerEvents:'none' }}>
              Doppelklick = umbenennen · Drag auf Knoten = einordnen
            </div>
          </div>
        )}

        {/* ── Ziehbare Trennlinie Map | Liste ── */}
        {!isMobile && (
          <div onMouseDown={e => { isResizingList.current=true; resizeListX.current=e.clientX; resizeListW.current=listWidth; document.body.style.cursor='col-resize' }}
            style={{ width:6, background:'#f1f5f9', cursor:'col-resize', flexShrink:0,
              display:'flex', alignItems:'center', justifyContent:'center',
              borderLeft:'1px solid #e2e8f0', borderRight:'1px solid #e2e8f0' }}
            onMouseEnter={e => e.currentTarget.style.background='#dbeafe'}
            onMouseLeave={e => e.currentTarget.style.background='#f1f5f9'}>
            <div style={{ width:2, height:32, borderRadius:2, background:'#cbd5e1' }}/>
          </div>
        )}

        {/* ── Liste + Kommentar ── */}
        {(!isMobile || activeTab==='list') && (
          <div style={{ width:isMobile?'100%':listWidth, flexShrink:0, display:'flex', flexDirection:'column', overflow:'hidden' }}>

            {/* Gliederungsliste */}
            <div style={{ flex:1, overflowY:'auto', padding:'12px 12px 0' }}>
              <div style={{ fontSize:10, fontWeight:700, color:'#94a3b8', letterSpacing:1,
                marginBottom:10, textTransform:'uppercase', paddingLeft:4 }}>Gliederung</div>
              {orderedNodes.map(node => {
                const depth  = getDepth(node, nodes)
                const color  = getBranchColor(node, nodes)
                const num    = getOutlineNumber(node, nodes)
                const isSel  = node.id === selectedId
                const isOver = node.id === listDragOver
                return (
                  <div key={node.id} draggable
                    onDragStart={e => onListDragStart(e, node.id)}
                    onDragOver={e => onListDragOver(e, node.id)}
                    onDrop={e => onListDrop(e, node.id)}
                    onDragLeave={() => setListDragOver(null)}
                    onClick={() => setSelectedId(node.id)}
                    onDoubleClick={() => openRename(node)}
                    style={{ display:'flex', alignItems:'center', gap:6, padding:'6px 8px',
                      marginLeft:depth*14, marginBottom:2, borderRadius:7,
                      background: isOver ? color+'15' : isSel ? color+'10' : 'transparent',
                      border: isOver ? `1.5px solid ${color}50` : isSel ? `1.5px solid ${color}28` : '1.5px solid transparent',
                      cursor:'grab', transition:'all 0.12s' }}>
                    <span style={{ color:'#cbd5e1', fontSize:12, userSelect:'none', flexShrink:0 }}>⠿</span>
                    <span style={{ width:6, height:6, borderRadius:'50%', background:color, flexShrink:0 }}/>
                    <span style={{ fontSize:10, fontWeight:700, color, fontFamily:'monospace', minWidth:26, flexShrink:0 }}>{num}</span>
                    <span style={{ flex:1, fontSize:12, color:'#1e293b', fontWeight:depth<=1?600:400,
                      whiteSpace:'nowrap', overflow:'hidden', textOverflow:'ellipsis' }}>{node.label}</span>
                    {node.comment && <span style={{ fontSize:10, color:'#94a3b8', flexShrink:0 }}>💬</span>}
                    <button onClick={e => { e.stopPropagation(); addChild(node.id) }}
                      style={{ background:'transparent', border:'none', color, cursor:'pointer', fontSize:15, padding:'0 2px', lineHeight:1, flexShrink:0 }}>+</button>
                    {node.parent_id && (
                      <button onClick={e => { e.stopPropagation(); deleteNode(node.id) }}
                        style={{ background:'transparent', border:'none', color:'#fca5a5', cursor:'pointer', fontSize:13, padding:'0 2px', lineHeight:1, flexShrink:0 }}>×</button>
                    )}
                  </div>
                )
              })}
            </div>

            {/* Ziehbare Trennlinie + Kommentar */}
            {selectedNode && (
              <>
                <div onMouseDown={e => { isResizingComment.current=true; resizeCommentY.current=e.clientY; resizeCommentH.current=commentHeight; document.body.style.cursor='row-resize' }}
                  style={{ height:10, cursor:'row-resize', display:'flex', alignItems:'center', justifyContent:'center',
                    borderTop:'1px solid #e2e8f0', flexShrink:0, background:'#fafafa' }}
                  onMouseEnter={e => e.currentTarget.style.background='#dbeafe'}
                  onMouseLeave={e => e.currentTarget.style.background='#fafafa'}>
                  <div style={{ width:36, height:3, borderRadius:2, background:'#cbd5e1' }}/>
                </div>
                <div style={{ height:commentHeight, flexShrink:0, display:'flex', flexDirection:'column',
                  padding:'10px 12px', background:'#f8fafc', overflow:'hidden' }}>
                  <div style={{ display:'flex', alignItems:'center', gap:6, marginBottom:8 }}>
                    <span style={{ width:8, height:8, borderRadius:'50%', background:getBranchColor(selectedNode,nodes), flexShrink:0 }}/>
                    <span style={{ fontSize:11, fontWeight:600, color:'#475569', flex:1,
                      overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>
                      {getOutlineNumber(selectedNode,nodes)} · {selectedNode.label}
                    </span>
                  </div>
                  <RichTextEditor value={comment} onChange={handleCommentChange}
                    accentColor={getBranchColor(selectedNode,nodes)}/>
                </div>
              </>
            )}
          </div>
        )}
      </div>
    </div>
  )
}
