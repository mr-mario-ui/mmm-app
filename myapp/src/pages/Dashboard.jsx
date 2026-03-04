import { useState } from 'react'
import { useNavigate } from 'react-router-dom'
import { useAuth } from '../hooks/useAuth'
import { useMaps } from '../hooks/useMaps'

export default function Dashboard() {
  const { user, logout }                        = useAuth()
  const { maps, loading, createMap, deleteMap } = useMaps(user?.id)
  const [newTitle, setNewTitle]                 = useState('')
  const navigate                                = useNavigate()

  const handleCreate = async () => {
    if (!newTitle.trim()) return
    const map = await createMap(newTitle)
    if (map) navigate(`/editor/${map.id}`)
    setNewTitle('')
  }

  return (
    <div style={{ minHeight:'100vh', background:'#0f0f0f', color:'#fff' }}>
      <header style={{ display:'flex', justifyContent:'space-between', alignItems:'center', padding:'20px 32px', borderBottom:'1px solid #1e1e1e' }}>
        <h1 style={{ margin:0, fontFamily:'Georgia,serif', fontSize:22, color:'#4ade80' }}>MindMap</h1>
        <div style={{ display:'flex', alignItems:'center', gap:16 }}>
          <span style={{ color:'#888', fontSize:13 }}>{user?.email}</span>
          <button onClick={logout}
            style={{ background:'transparent', border:'1px solid #333', color:'#888', padding:'6px 14px', borderRadius:6, cursor:'pointer', fontSize:13 }}>
            Abmelden
          </button>
        </div>
      </header>

      <main style={{ maxWidth:900, margin:'0 auto', padding:'40px 32px' }}>
        <div style={{ display:'flex', gap:12, marginBottom:40 }}>
          <input
            style={{ flex:1, padding:'12px 16px', background:'#1a1a1a', border:'1px solid #2a2a2a', borderRadius:8, color:'#fff', fontSize:14 }}
            placeholder="Neue Map benennen..."
            value={newTitle}
            onChange={e => setNewTitle(e.target.value)}
            onKeyDown={e => e.key === 'Enter' && handleCreate()}
          />
          <button onClick={handleCreate}
            style={{ padding:'12px 24px', background:'#4ade80', border:'none', borderRadius:8, color:'#000', fontWeight:700, cursor:'pointer' }}>
            + Erstellen
          </button>
        </div>

        {loading ? (
          <p style={{ color:'#555', textAlign:'center', marginTop:60 }}>Laden...</p>
        ) : maps.length === 0 ? (
          <p style={{ color:'#555', textAlign:'center', marginTop:60 }}>Noch keine Maps. Erstelle deine erste!</p>
        ) : (
          <div style={{ display:'grid', gridTemplateColumns:'repeat(auto-fill, minmax(240px, 1fr))', gap:16 }}>
            {maps.map(map => (
              <div key={map.id} style={{ background:'#1a1a1a', border:'1px solid #2a2a2a', borderRadius:12, overflow:'hidden' }}>
                <div onClick={() => navigate(`/editor/${map.id}`)}
                  style={{ padding:20, fontSize:16, fontWeight:600, cursor:'pointer' }}>
                  {map.title}
                </div>
                <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', padding:'12px 20px', borderTop:'1px solid #222' }}>
                  <span style={{ color:'#555', fontSize:12 }}>{new Date(map.updated_at).toLocaleDateString('de-DE')}</span>
                  <button onClick={() => deleteMap(map.id)}
                    style={{ background:'transparent', border:'none', color:'#f87171', cursor:'pointer', fontSize:12 }}>
                    Löschen
                  </button>
                </div>
              </div>
            ))}
          </div>
        )}
      </main>
    </div>
  )
}
