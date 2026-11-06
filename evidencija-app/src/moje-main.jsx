import React from 'react'
import { createRoot } from 'react-dom/client'
import Moje from './Moje.jsx'

// Zasebna ulazna točka za radnički portal (/moje) — ne učitava aplikaciju.
createRoot(document.getElementById('root')).render(<Moje />)
