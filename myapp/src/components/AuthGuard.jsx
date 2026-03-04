import { Navigate } from 'react-router-dom'
import { useAuth } from '../hooks/useAuth'

export default function AuthGuard({ children }) {
  const { user, loading } = useAuth()

  if (loading) {
    return (
      <div style={{ display:'flex', justifyContent:'center', alignItems:'center', height:'100vh', background:'#0f0f0f' }}>
        <div style={{ width:32, height:32, border:'3px solid #1e1e1e', borderTop:'3px solid #4ade80', borderRadius:'50%' }} />
      </div>
    )
  }

  if (!user) return <Navigate to="/login" replace />
  return children
}
