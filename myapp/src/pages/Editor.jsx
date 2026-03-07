// src/pages/Editor.jsx
import { useState, useEffect, useRef, useCallback } from 'react'
import { useParams, useNavigate } from 'react-router-dom'
import { supabase } from '../lib/supabase'

const DEPTH_COLORS = ['#2563eb','#16a34a','#dc2626','#d97706','#7c3aed','#0891b2','#be185d']

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
  if (!node.parent_id) {
    const idx = all.filter(n => !n.parent_id).findIndex(n => n.id === node.id)
    return `${idx + 1}`
  }
  const parent = all.find(n => n.id === node.parent_id)
  if (!parent) return '?'
  const parentNum = getOutlineNumber(parent, all)
  const idx = all.filter(n => n.parent_id === node.parent_id).findIndex(n => n.id === node.id)
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

  const NODE_W = 230   // horizontaler Abstand
  const NODE_H = 110   // vertikaler Abstand zwischen Geschwistern
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

export default function Editor() {
  const { mapId } = useParams()
  const navigate  = useNavigate()

  const [nodes, setNodes]           = useState([])
  const [mapTitle, setMapTitle]     = useState('')
  const [selectedId, setSelectedId] = useState(null)
  const [comment, setComment]       = useState('')
  const [saved, setSaved]           = useState(true)
  const [pan, setPan]               = useState({ x: 0, y: 0 })
  const [isMobile, setIsMobile]     = useState(window.innerWidth < 768)
  const [activeTab, setActiveTab]   = useState('map')

  // Rename Popup
  const [renameId, setRenameId]     = useState(null)
  const [renameText, setRenameText] = useState('')
  const renameInputRef              = useRef(null)

  // Undo
  const undoStack = useRef([])
  const pushUndo  = useCallback((snap) => {
    undoStack.current = [...undoStack.current.slice(-MAX_UNDO + 1), JSON.parse(JSON.stringify(snap))]
  }, [])

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

  const commentTimer = useRef(null)

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
    if (!selectedId) return
    clearTimeout(commentTimer.current)
    commentTimer.current = setTimeout(async () => {
      setNodes(prev => prev.map(n => n.id === selectedId ? { ...n, comment: val } : n))
      setSaved(false)
      await supabase.from('nodes').update({ comment: val }).eq('id', selectedId)
      setSaved(true)
    }, 800)
  }

  // ── Canvas drag ───────────────────────────────────────────
  const onNodeMouseDown = (e, id) => {
    e.stopPropagation()
    setSelectedId(id)
    const node = nodes.find(n => n.id === id)
    dragging.current = id
    didDrag.current  = false
    dragOff.current  = { x: e.clientX - node.x, y: e.clientY - node.y }
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
      const nx = e.clientX - dragOff.current.x
      const ny = e.clientY - dragOff.current.y
      setNodes(prev => prev.map(n => n.id === dragging.current ? { ...n, x: nx, y: ny } : n))
      setDragGhost({ x: e.clientX, y: e.clientY })
      // Drop-Ziel: nur wenn Knoten wirklich ÜBER einem anderen liegt (Box-Überlappung)
      const W_BOX = 150, H_BOX = 50
      const others = nodes.filter(n => n.id !== dragging.current && !getDesc(dragging.current, nodes).slice(1).includes(n.id))
      let hit = null
      others.forEach(n => {
        if (
          nx > n.x - W_BOX/2 && nx < n.x + W_BOX/2 &&
          ny > n.y - H_BOX/2 && ny < n.y + H_BOX/2
        ) { hit = n.id }
      })
      setDropTargetId(hit)
    } else if (isPanning.current) {
      setPan({ x: e.clientX - panStart.current.x, y: e.clientY - panStart.current.y })
    }
  }, [nodes])

  const onUp = useCallback(async () => {
    if (dragging.current) {
      const draggedId = dragging.current
      const node = nodes.find(n => n.id === draggedId)
      if (didDrag.current) {
        pushUndo(nodes)
        if (dropTargetId) {
          // Reparent + neu layouten
          const next = nodes.map(n => n.id === draggedId ? { ...n, parent_id: dropTargetId } : n)
          await supabase.from('nodes').update({ parent_id: dropTargetId }).eq('id', draggedId)
          await applyAutoLayout(next)
        } else {
          // Nur Position speichern
          setSaved(false)
          await supabase.from('nodes').update({ x: node.x, y: node.y }).eq('id', draggedId)
          setSaved(true)
        }
      }
      dragging.current = null
      didDrag.current  = false
      setDropTargetId(null)
      setDragGhost(null)
    }
    isPanning.current = false
  }, [nodes, dropTargetId, pushUndo])

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

  const selectedNode = nodes.find(n => n.id === selectedId)
  const orderedNodes = getOrderedNodes(nodes)
  const canUndo      = undoStack.current.length > 0

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

      {/* ── Header ── */}
      <div style={{ display:'flex', alignItems:'center', gap:10, padding:'0 16px',
        height:50, borderBottom:'1px solid #e2e8f0', background:'white', flexShrink:0, zIndex:10 }}>
        <button onClick={() => navigate('/dashboard')}
          style={{ background:'transparent', border:'1px solid #e2e8f0', color:'#64748b',
            padding:'5px 12px', borderRadius:6, cursor:'pointer', fontSize:12, whiteSpace:'nowrap' }}>
          ← Dashboard
        </button>
        <span style={{ fontWeight:700, fontSize:15, color:'#1e293b', flex:1,
          overflow:'hidden', textOverflow:'ellipsis', whiteSpace:'nowrap' }}>{mapTitle}</span>

        {/* Auto-Layout Button */}
        <button onClick={() => applyAutoLayout(nodes)}
          title="Knoten neu anordnen"
          style={{ background:'#f1f5f9', border:'1px solid #e2e8f0', color:'#334155',
            padding:'5px 10px', borderRadius:6, cursor:'pointer', fontSize:12, whiteSpace:'nowrap' }}>
          ⊞ Layout
        </button>

        {/* Undo */}
        <button onClick={() => {
          if (!canUndo) return
          const prev = undoStack.current[undoStack.current.length - 1]
          undoStack.current = undoStack.current.slice(0, -1)
          setNodes(prev)
          supabase.from('nodes').delete().eq('map_id', mapId).then(() => supabase.from('nodes').insert(prev))
          setSaved(true)
        }}
          disabled={!canUndo}
          title="Rückgängig (Strg+Z)"
          style={{ background: canUndo ? '#f1f5f9' : 'transparent',
            border:'1px solid #e2e8f0', color: canUndo ? '#334155' : '#cbd5e1',
            padding:'5px 10px', borderRadius:6, cursor: canUndo ? 'pointer' : 'default',
            fontSize:13, whiteSpace:'nowrap' }}>
          ↩ {canUndo ? `(${undoStack.current.length})` : ''}
        </button>

        <span style={{ fontSize:11, color: saved ? '#22c55e' : '#f59e0b', fontWeight:600, flexShrink:0 }}>
          {saved ? '✓' : '⟳'}
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

      {/* ── Body ── */}
      <div style={{ flex:1, display:'flex', overflow:'hidden', minHeight:0 }}>

        {/* ── Mindmap ── */}
        {(!isMobile || activeTab==='map') && (
          <div style={{ flex:1, position:'relative', overflow:'hidden', minWidth:0 }}>
            <svg ref={svgRef}
              style={{ width:'100%', height:'100%', background:'#f8fafc', cursor:'default', display:'block' }}
              onMouseDown={onSvgDown}>
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
              <g transform={`translate(${pan.x},${pan.y})`}>
                {/* Linien */}
                {nodes.map(node => {
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
                {/* Nodes */}
                {nodes.map(node => {
                  const color   = getColor(node, nodes)
                  const isRoot  = !node.parent_id
                  const isSel   = node.id === selectedId
                  const isDrg   = node.id === dragging.current
                  const isDropT = node.id === dropTargetId
                  const num     = getOutlineNumber(node, nodes)
                  const W = isRoot ? 168 : 150, H = 50

                  return (
                    <g key={node.id} style={{ opacity: isDrg ? 0.65 : 1 }}>
                      {/* Glow bei Auswahl - CSS animiert */}
                      {isSel && !isDropT && (
                        <>
                          {/* Äußerer weicher Glow */}
                          <rect className="node-halo-outer"
                            x={node.x-W/2-16} y={node.y-H/2-16} width={W+32} height={H+32} rx="22"
                            fill={color} stroke="none" style={{ filter:`blur(14px)` }}/>
                          {/* Innerer schärferer Glow */}
                          <rect className="node-halo-inner"
                            x={node.x-W/2-8} y={node.y-H/2-8} width={W+16} height={H+16} rx="16"
                            fill={color} stroke="none" style={{ filter:`blur(6px)` }}/>
                        </>
                      )}
                      {/* Drop-Ziel Glühen */}
                      {isDropT && (
                        <rect x={node.x-W/2-10} y={node.y-H/2-10} width={W+20} height={H+20} rx="18"
                          fill={`${color}18`} stroke={color} strokeWidth="2.5">
                          <animate attributeName="strokeOpacity" values="0.4;1;0.4" dur="0.7s" repeatCount="indefinite"/>
                        </rect>
                      )}
                      {/* Box */}
                      <rect x={node.x-W/2} y={node.y-H/2} width={W} height={H} rx="11"
                        fill="white"
                        stroke={isDropT ? color : isSel ? color : '#e2e8f0'}
                        strokeWidth={isDropT ? 2.5 : isSel ? 2 : 1}
                        filter={isDrg ? 'url(#shd)' : 'url(#sh)'}
                        style={{ cursor:'grab' }}
                        onMouseDown={e => onNodeMouseDown(e, node.id)}
                        onDoubleClick={e => { e.stopPropagation(); openRename(node) }}
                      />
                      {/* Farbstreifen */}
                      <rect x={node.x-W/2} y={node.y-H/2} width={5} height={H} rx="3" fill={color}/>
                      {/* Nummer */}
                      <text x={node.x-W/2+12} y={node.y-H/2+14}
                        fill={color} fontSize="9" fontWeight="700"
                        style={{ pointerEvents:'none', fontFamily:'monospace' }}>{num}</text>
                      {/* Label */}
                      <text x={node.x-W/2+12} y={node.y+8}
                        fill="#1e293b" fontSize={isRoot ? 13 : 12} fontWeight={isRoot ? '700' : '500'}
                        style={{ pointerEvents:'none', userSelect:'none' }}>
                        {node.label.length > 17 ? node.label.slice(0,16)+'…' : node.label}
                      </text>
                      {/* + */}
                      <g style={{ cursor:'pointer' }} onClick={e => addChild(node.id, e)}>
                        <circle cx={node.x+W/2+16} cy={node.y} r={12} fill={color}/>
                        <text x={node.x+W/2+16} y={node.y+1} textAnchor="middle" dominantBaseline="middle"
                          fill="white" fontSize="18" fontWeight="700" style={{ pointerEvents:'none' }}>+</text>
                      </g>
                      {/* × */}
                      {node.parent_id && (
                        <g style={{ cursor:'pointer' }} onClick={e => deleteNode(node.id, e)}>
                          <circle cx={node.x-W/2-16} cy={node.y} r={12} fill="#fee2e2"/>
                          <text x={node.x-W/2-16} y={node.y+1} textAnchor="middle" dominantBaseline="middle"
                            fill="#dc2626" fontSize="16" style={{ pointerEvents:'none' }}>×</text>
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
                  <textarea placeholder="Kommentar schreiben..."
                    value={comment}
                    onChange={e => handleCommentChange(e.target.value)}
                    style={{ flex:1, width:'100%', padding:'10px 12px',
                      border:'1.5px solid #e2e8f0', borderRadius:8, fontSize:13,
                      color:'#334155', background:'white', resize:'none', outline:'none',
                      fontFamily:'inherit', lineHeight:1.7, boxSizing:'border-box',
                      transition:'border-color 0.15s' }}
                    onFocus={e => e.target.style.borderColor = getColor(selectedNode, nodes)}
                    onBlur={e => e.target.style.borderColor = '#e2e8f0'}
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
