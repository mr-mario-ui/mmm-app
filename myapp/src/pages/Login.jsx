import { useState } from 'react'
import { useNavigate } from 'react-router-dom'
import { useAuth } from '../hooks/useAuth'

export default function Login() {
  const [email, setEmail]       = useState('')
  const [password, setPassword] = useState('')
  const [isRegister, setIsReg]  = useState(false)
  const [error, setError]       = useState('')
  const { login, register }     = useAuth()
  const navigate                = useNavigate()

  const handleSubmit = async () => {
    setError('')
    try {
      isRegister ? await register(email, password) : await login(email, password)
      navigate('/dashboard')
    } catch (err) {
      setError(err.message)
    }
  }

  const inp = {
    display:'block', width:'100%', padding:'12px 16px', marginBottom:12,
    background:'#111', border:'1px solid #333', borderRadius:8,
    color:'#fff', fontSize:14, boxSizing:'border-box'
  }

  return (
    <div style={{ display:'flex', justifyContent:'center', alignItems:'center', height:'100vh', background:'#0f0f0f' }}>
      <div style={{ background:'#1a1a1a', padding:40, borderRadius:16, width:360, border:'1px solid #2a2a2a' }}>
        <h1 style={{ color:'#4ade80', fontFamily:'Georgia,serif', margin:'0 0 4px' }}>MindMap</h1>
        <p style={{ color:'#888', margin:'0 0 28px', fontSize:14 }}>
          {isRegister ? 'Konto erstellen' : 'Anmelden'}
        </p>

        <input style={inp} type="email" placeholder="E-Mail"
          value={email} onChange={e => setEmail(e.target.value)} />
        <input style={inp} type="password" placeholder="Passwort"
          value={password} onChange={e => setPassword(e.target.value)}
          onKeyDown={e => e.key === 'Enter' && handleSubmit()} />

        {error && <p style={{ color:'#f87171', fontSize:13, margin:'0 0 12px' }}>{error}</p>}

        <button onClick={handleSubmit}
          style={{ width:'100%', padding:13, background:'#4ade80', border:'none', borderRadius:8, color:'#000', fontWeight:700, fontSize:15, cursor:'pointer' }}>
          {isRegister ? 'Registrieren' : 'Anmelden'}
        </button>
        <p onClick={() => setIsReg(!isRegister)}
          style={{ color:'#888', fontSize:13, textAlign:'center', marginTop:16, cursor:'pointer', textDecoration:'underline' }}>
          {isRegister ? 'Bereits ein Konto? Anmelden' : 'Noch kein Konto? Registrieren'}
        </p>
      </div>
    </div>
  )
}
