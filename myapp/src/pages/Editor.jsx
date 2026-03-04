import { useState, useEffect, useRef, useCallback } from 'react'
import { useParams, useNavigate } from 'react-router-dom'
import { supabase } from '../lib/supabase'

const COLORS = ['#4ade80','#60a5fa','#f472b6','#fb923c','#a78bfa','#34d399','#fbbf24']

function getDepth(node, all) {
  if (!node.parent_id) return 0
  const p = all.find(n => n.id === node.parent_id)
  return p ? getDepth(p, all) + 1 : 1
}
function getColor(node, all) { return COLORS[getDepth(node, all) % COLORS.length] }
function getDesc(id, all) {
  const r = [id]
  all.forEach(n => { if (n.parent_id === id) r.push(...getDesc(n.id, all)) })
  return r
}

export default function Editor() {
  const { mapId }  = useParams()
  const navigate   = useNavigate()
  const [nodes, setNodes]         = useState([])
  const [mapTitle, setMapTitle]   = useState('')
  const [editingId, setEditingId] = useState(null)
  const [editText, setEditText]   = useState('')
  const [saved, setSaved]         = useState(true)
  const [pan, setPan]             = useState({ x: 0, y: 0 })
  const dragging  = useRef(null)
  const dragOff   = useRef({ x:0, y:0 })
  const isPanning = useRef(false)
  const panStart  = useRef({ x:0, y:0 })
  const svgRef    = useRef()

  useEffect(() => { loadMap() }, [mapId])

  const loadMap = async () => {
    const { data: map } = await supabase.from('maps').select('title').eq('id', mapId).single()
    if (map) setMapTitle(map.title)
    const { data } = await supabase.from('nodes').select('*').eq('map_id', mapId)
    if (data && data.length > 0) {
      setNodes(data)
    } else {
      const root = { id: crypto.randomUUID(), map_id: mapId, label: map?.title || 'Start', x: 500, y: 350, parent_id: null }
      setNodes([root])
      await supabase.from('nodes').insert(root)
    }
  }

  const addChild = async (parentId) => {
    const parent   = nodes.find(n => n.id === parentId)
    const siblings = nodes.filter(n => n.parent_id === parentId)
    const newNode  = {
      id: crypto.randomUUID(), map_id: mapId, label: 'Neue Idee',
      x: parent.x + 200, y: parent.y + (siblings.length - 1) * 70, parent_id: parentId
    }
    setNodes(prev => [...prev, newNode])
    await supabase.from('nodes').insert(newNode)
    setEditingId(newNode.id)
    setEditText('Neue Idee')
  }

  const saveLabel = async (nodeId) => {
    const t = editText.trim() || 'Node'
    setNodes(prev => prev.map(n => n.id === nodeId ? { ...n, label: t } : n))
    setEditingId(null); setSaved(false)
    await supabase.from('nodes').update({ label: t }).eq('id', nodeId)
    setSaved(true)
  }

  const deleteNode = async (nodeId) => {
    const del = getDesc(nodeId, nodes)
    setNodes(prev => prev.filter(n => !del.includes(n.id)))
    await supabase.from('nodes').delete().in('id', del)
  }

  const onNodeDown = (e, id) => {
    e.stopPropagation(); if (editingId) return
    const node = nodes.find(n => n.id === id)
    dragging.current = id
    dragOff.current  = { x: e.clientX - node.x, y: e.clientY - node.y }
  }

  const onSvgDown = (e) => {
    if (e.target === svgRef.current || e.target.tagName === 'svg') {
      isPanning.current = true
      panStart.current  = { x: e.clientX - pan.x, y: e.clientY - pan.y }
    }
  }

  const onMove = useCallback((e) => {
    if (dragging.current) {
      setNodes(prev => prev.map(n => n.id === dragging.current
        ? { ...n, x: e.clientX - dragOff.current.x, y: e.clientY - dragOff.current.y } : n))
    } else if (isPanning.current) {
      setPan({ x: e.clientX - panStart.current.x, y: e.clientY - panStart.current.y })
    }
  }, [])

  const onUp = useCallback(async () => {
    if (dragging.current) {
      const node = nodes.find(n => n.id === dragging.current)
      if (node) {
        setSaved(false)
        await supabase.from('nodes').update({ x: node.x, y: node.y }).eq('id', node.id)
        setSaved(true)
      }
      dragging.current = null
    }
    isPanning.current = false
  }, [nodes])

  return (
    <div style={{ width:'100vw', height:'100vh', background:'#0f0f0f', overflow:'hidden', position:'relative', userSelect:'none' }}
      onMouseMove={onMove} onMouseUp={onUp}>

      <div style={{ position:'absolute', top:0, left:0, right:0, display:'flex', alignItems:'center', gap:20, padding:'14px 24px', background:'rgba(15,15,15,0.95)', borderBottom:'1px solid #1e1e1e', zIndex:10 }}>
        <button onClick={() => navigate('/dashboard')}
          style={{ background:'transparent', border:'1px solid #2a2a2a', color:'#888', padding:'6px 14px', borderRadius:6, cursor:'pointer', fontSize:13 }}>
          ← Dashboard
        </button>
        <span style={{ color:'#fff', fontFamily:'Georgia,serif', fontSize:18 }}>{mapTitle}</span>
        <span style={{ color: saved ? '#4ade80' : '#fbbf24', fontSize:12 }}>
          {saved ? '✓ Gespeichert' : '⟳ Speichern…'}
        </span>
        <span style={{ color:'#444', fontSize:12, marginLeft:'auto' }}>
          Doppelklick = bearbeiten | + = Kind | × = löschen | Drag = verschieben
        </span>
      </div>

      <svg ref={svgRef} style={{ width:'100%', height:'100%', paddingTop:56 }} onMouseDown={onSvgDown}>
        <g transform={`translate(${pan.x},${pan.y})`}>

          {nodes.map(node => {
            if (!node.parent_id) return null
            const p = nodes.find(n => n.id === node.parent_id)
            if (!p) return null
            const mx = (p.x + node.x) / 2
            return (
              <path key={`e-${node.id}`}
                d={`M ${p.x} ${p.y} C ${mx} ${p.y}, ${mx} ${node.y}, ${node.x} ${node.y}`}
                fill="none" stroke={getColor(node, nodes)} strokeWidth="1.5" strokeOpacity="0.35" />
            )
          })}

          {nodes.map(node => {
            const color  = getColor(node, nodes)
            const isRoot = !node.parent_id
            const W = isRoot ? 150 : 130, H = 44

            return (
              <g key={node.id}>
                <rect x={node.x-W/2+3} y={node.y-H/2+3} width={W} height={H} rx="10" fill="rgba(0,0,0,0.35)" />
                <rect x={node.x-W/2} y={node.y-H/2} width={W} height={H} rx="10"
                  fill="#1a1a1a" stroke={color} strokeWidth={isRoot ? 2.5 : 1.5}
                  style={{ cursor:'grab' }}
                  onMouseDown={e => onNodeDown(e, node.id)}
                  onDoubleClick={() => { setEditingId(node.id); setEditText(node.label) }} />

                {editingId === node.id ? (
                  <foreignObject x={node.x-W/2+8} y={node.y-H/2+8} width={W-16} height={H-16}>
                    <input xmlns="http://www.w3.org/1999/xhtml" autoFocus value={editText}
                      onChange={e => setEditText(e.target.value)}
                      onBlur={() => saveLabel(node.id)}
                      onKeyDown={e => { if (e.key==='Enter') saveLabel(node.id); if (e.key==='Escape') setEditingId(null) }}
                      style={{ width:'100%', height:'100%', background:'transparent', border:'none', outline:'none', color:'#fff', fontSize:13, textAlign:'center', fontFamily:'inherit' }} />
                  </foreignObject>
                ) : (
                  <text x={node.x} y={node.y} textAnchor="middle" dominantBaseline="middle"
                    fill={isRoot ? color : '#e5e7eb'} fontSize={isRoot ? 14 : 13} fontWeight={isRoot ? '700' : '400'}
                    style={{ pointerEvents:'none', userSelect:'none' }}>
                    {node.label.length > 16 ? node.label.slice(0,15)+'…' : node.label}
                  </text>
                )}

                <g style={{ cursor:'pointer' }} onClick={() => addChild(node.id)}>
                  <circle cx={node.x+W/2+14} cy={node.y} r={11} fill={color} />
                  <text x={node.x+W/2+14} y={node.y+1} textAnchor="middle" dominantBaseline="middle"
                    fill="#000" fontSize="18" fontWeight="700" style={{ pointerEvents:'none' }}>+</text>
                </g>

                {node.parent_id && (
                  <g style={{ cursor:'pointer' }} onClick={() => deleteNode(node.id)}>
                    <circle cx={node.x-W/2-14} cy={node.y} r={11} fill="#1f2937" />
                    <text x={node.x-W/2-14} y={node.y+1} textAnchor="middle" dominantBaseline="middle"
                      fill="#f87171" fontSize="16" style={{ pointerEvents:'none' }}>×</text>
                  </g>
                )}
              </g>
            )
          })}
        </g>
      </svg>
    </div>
  )
}
