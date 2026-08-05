import { createClient } from '@supabase/supabase-js'

const url = import.meta.env.VITE_SUPABASE_URL
const key = import.meta.env.VITE_SUPABASE_ANON_KEY

if (!url || !key) {
  console.error('Nedostaju VITE_SUPABASE_URL i VITE_SUPABASE_ANON_KEY u .env datoteci!')
}

export const supabase = createClient(url, key)
