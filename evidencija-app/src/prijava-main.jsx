import React from 'react'
import { createRoot } from 'react-dom/client'
import Prijava from './Prijava.jsx'

// Zasebna ulazna točka za javnu prijavnicu (/prijava) — ne učitava aplikaciju.
createRoot(document.getElementById('root')).render(<Prijava />)
