import { useState, useEffect } from 'react'
import { supabase } from '../lib/supabase'

export function useMaps(userId) {
  const [maps, setMaps] = useState([])
  const [loading, setLoading] = useState(true)

  useEffect(() => {
    if (!userId) return
    loadMaps()
  }, [userId])

  const loadMaps = async () => {
    const { data, error } = await supabase
      .from('maps')
      .select('*')
      .eq('user_id', userId)
      .order('updated_at', { ascending: false })

    if (!error) setMaps(data)
    setLoading(false)
  }

  const createMap = async (title) => {
    const { data, error } = await supabase
      .from('maps')
      .insert({ title, user_id: userId })
      .select()
      .single()

    if (!error) {
      setMaps(prev => [data, ...prev])
      return data
    }
  }

  const deleteMap = async (mapId) => {
    await supabase.from('maps').delete().eq('id', mapId)
    setMaps(prev => prev.filter(m => m.id !== mapId))
  }

  return { maps, loading, createMap, deleteMap, reload: loadMaps }
}
