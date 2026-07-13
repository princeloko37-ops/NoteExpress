import React, { useEffect, useMemo, useRef, useState } from 'react'
import * as XLSX from 'xlsx'
import {
  Users, Edit3, BookOpen, Trophy, MoreHorizontal, Lock, Moon, Sun, Plus,
  Trash2, ArrowUp, ArrowDown, Upload, Download, ChevronRight, X, Check,
  RotateCcw, Home, KeyRound, ShieldCheck, Copy, AlertCircle, BarChart3,
  LogOut, Pencil, ChevronLeft, CircleCheck, CircleX, Layers, ClipboardList,
} from 'lucide-react'

/* ============================================================================
   CONSTANTES MÉTIER
   ========================================================================== */

const EVALUATION_TYPES = [
  'Évaluation Formative 1',
  'Évaluation Sommative 1',
  'Évaluation Sommative 2',
  'Évaluation Sommative 3',
]

const PASSING_AVERAGE = 9 // seuil admis / non admis sur la moyenne générale
const SUBJECT_PASS = 10 // seuil de réussite par matière

const APP_NAME = 'NoteExpress'
const APP_VERSION = '2.0.0'
const CACHE_HINT = 'v2' // à faire correspondre avec public/sw.js -> CACHE_NAME

const LICENSE_SALT = 'NX-EDUCMASTER-BENIN'
const ADMIN_PASSCODE = '379246' // code réservé à Prince pour générer des licences — à changer si besoin

/* ============================================================================
   CLÉS LOCALSTORAGE
   ========================================================================== */

const LS_CLASSES = 'noteexpress:classes'
const LS_ACTIVE = 'noteexpress:activeId'
const LS_LEGACY = 'carnet-notes:current' // migration depuis l'ancienne app
const LS_PIN = 'noteexpress:pin'
const LS_DARK = 'noteexpress:dark'
const LS_LICENSE = 'noteexpress:license'
const LS_DEFAULT_MODE = 'noteexpress:defaultEntryMode'

/* ============================================================================
   HELPERS GÉNÉRAUX
   ========================================================================== */

function uid(prefix = 'id') {
  return `${prefix}_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`
}

function formatNum(n) {
  if (n === null || n === undefined || n === '' || Number.isNaN(n)) return ''
  const rounded = Math.round(n * 100) / 100
  return rounded.toString().replace('.', ',')
}

function toNum(v) {
  if (v === null || v === undefined || v === '') return 0
  if (typeof v === 'number') return Number.isNaN(v) ? 0 : v
  const cleaned = String(v).trim().replace(',', '.')
  const n = parseFloat(cleaned)
  return Number.isNaN(n) ? 0 : n
}

function cleanMatricule(v) {
  // Certains exports EducMaster stockent le matricule avec une apostrophe
  // collée au début (ex: "'2141025034359") — on l'enlève pour l'affichage
  // et les clés internes, mais rawMatricule garde la valeur d'origine pour
  // que l'export réimporte exactement comme EducMaster l'attend.
  return String(v ?? '').trim().replace(/^['\u2018\u2019]/, '')
}

function normalizeLabel(s) {
  return String(s ?? '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .trim()
}

/* ============================================================================
   PARSING DU CODE DE SAISIE (obtenue + perfectionnement en un seul champ)
   Formats acceptés :
     "1202"   -> 12 + 02 = 14 (entier, sans séparateur)
     "120"    -> 12 + 0  = 12
     "12,52"  -> 12,5 + 2 = 14,5 (décimal, sans séparateur)
     "14,5 2" / "12:0" / "18,5;1" (séparateur explicite espace / : / ; / ')
   ========================================================================== */

function sanitizeCode(raw) {
  return String(raw ?? '').replace(/[^0-9,.\s:;']/g, '')
}

function parseCode(raw) {
  const s = sanitizeCode(raw).trim()
  if (!s) return null

  const clampPerf = (n) => (n === null || n === undefined ? n : Math.max(0, Math.min(2, Math.round(n))))

  // Séparateur explicite entre obtenue et perfectionnement
  const sep = s.match(/^([0-9]+(?:[.,][0-9]+)?)[\s:;']+([0-9]+(?:[.,][0-9]+)?)$/)
  if (sep) {
    const obtenue = toNum(sep[1])
    const perfectionnement = clampPerf(toNum(sep[2]))
    return { obtenue, perfectionnement, incomplete: false }
  }

  // Décimal sans séparateur : "12,52" -> obtenue "12,5" + perfectionnement "2"
  if (/[.,]/.test(s)) {
    const m = s.match(/^([0-9]+)[.,]([0-9]*)$/)
    if (!m) return null
    const intPart = m[1]
    const decPart = m[2]
    if (decPart.length >= 2) {
      const obtenueDec = decPart.slice(0, -1)
      const perf = decPart.slice(-1)
      return { obtenue: toNum(`${intPart}.${obtenueDec}`), perfectionnement: clampPerf(toNum(perf)), incomplete: false }
    }
    return { obtenue: toNum(`${intPart}.${decPart || '0'}`), perfectionnement: null, incomplete: true }
  }

  // Entier sans séparateur
  if (/^[0-9]+$/.test(s)) {
    if (s.length <= 2) {
      return { obtenue: parseInt(s, 10), perfectionnement: null, incomplete: true }
    }
    if (s.length === 3) {
      return { obtenue: parseInt(s.slice(0, 2), 10), perfectionnement: clampPerf(parseInt(s.slice(2), 10)), incomplete: false }
    }
    // 4 chiffres et plus : les 2 derniers sont le perfectionnement
    return {
      obtenue: parseInt(s.slice(0, s.length - 2), 10),
      perfectionnement: clampPerf(parseInt(s.slice(-2), 10)),
      incomplete: false,
    }
  }

  return null
}

function isCodeComplete(raw) {
  const p = parseCode(raw)
  if (!p) return false
  return !p.incomplete && p.perfectionnement !== null && !Number.isNaN(p.obtenue) && !Number.isNaN(p.perfectionnement)
}

function codeFromValues(obtenue, perfectionnement) {
  if (obtenue === null || obtenue === undefined || obtenue === '') return ''
  if (perfectionnement === null || perfectionnement === undefined || perfectionnement === '') {
    return formatNum(obtenue)
  }
  return `${formatNum(obtenue)} ${formatNum(perfectionnement)}`
}

/* ============================================================================
   CLASSEMENT / MOYENNES
   ========================================================================== */

function computeRanking(roster, subjects, grades, attendance) {
  const rows = roster.map((student) => {
    const isAbsent = attendance?.[student.matricule] === false
    const studentGrades = grades?.[student.matricule] || {}
    let sum = 0
    let count = 0
    subjects.forEach((subj) => {
      const g = studentGrades[subj.key]
      if (g && (g.obtenue !== null && g.obtenue !== undefined)) {
        sum += toNum(g.obtenue) + toNum(g.perfectionnement)
        count += 1
      }
    })
    const average = !isAbsent && count > 0 ? sum / count : null
    return { ...student, average, isAbsent, gradedCount: count }
  })

  const ranked = [...rows].sort((a, b) => {
    if (a.average === null && b.average === null) return 0
    if (a.average === null) return 1
    if (b.average === null) return -1
    return b.average - a.average
  })

  let lastAvg = null
  let lastRank = 0
  ranked.forEach((row, idx) => {
    if (row.average === null) {
      row.rank = null
      return
    }
    if (lastAvg !== null && Math.abs(row.average - lastAvg) < 1e-9) {
      row.rank = lastRank
    } else {
      row.rank = idx + 1
      lastRank = row.rank
      lastAvg = row.average
    }
  })

  ranked.forEach((row) => {
    row.decision = row.average === null ? null : row.average >= PASSING_AVERAGE ? 'Admis' : 'Non admis'
  })

  return ranked
}

function computeSubjectStats(roster, subjects, grades, attendance) {
  return subjects.map((subj) => {
    let sum = 0
    let count = 0
    let success = 0
    roster.forEach((student) => {
      if (attendance?.[student.matricule] === false) return
      const g = grades?.[student.matricule]?.[subj.key]
      if (g && g.obtenue !== null && g.obtenue !== undefined) {
        const total = toNum(g.obtenue) + toNum(g.perfectionnement)
        sum += total
        count += 1
        if (total >= SUBJECT_PASS) success += 1
      }
    })
    return {
      key: subj.key,
      label: subj.label,
      average: count > 0 ? sum / count : null,
      successRate: count > 0 ? (success / count) * 100 : null,
      graded: count,
      total: roster.length,
    }
  })
}

/* ============================================================================
   SYSTÈME DE LICENCE (sans serveur) — btoa/atob + checksum
   Le code encode : nom de l'école + date d'expiration.
   Ce n'est pas un système de sécurité incassable : c'est un verrou "raisonnable"
   côté client, dans l'esprit décrit pour ce projet (pas de backend).
   ========================================================================== */

function simpleHash(str) {
  let hash = 0
  for (let i = 0; i < str.length; i++) {
    hash = (hash * 31 + str.charCodeAt(i)) >>> 0
  }
  return hash
}

function toBase64Url(str) {
  const b64 = btoa(unescape(encodeURIComponent(str)))
  return b64.replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')
}

function fromBase64Url(b64url) {
  const b64 = b64url.replace(/-/g, '+').replace(/_/g, '/')
  const padded = b64 + '='.repeat((4 - (b64.length % 4)) % 4)
  return decodeURIComponent(escape(atob(padded)))
}

function generateLicenseCode({ school, days }) {
  const expiresAt = Date.now() + days * 24 * 60 * 60 * 1000
  const payload = JSON.stringify({ s: school.trim(), e: expiresAt, v: 1 })
  const b64url = toBase64Url(payload)
  const checksum = (simpleHash(b64url + LICENSE_SALT) % 46656).toString(36).toUpperCase().padStart(3, '0')
  const raw = b64url + checksum
  const blocks = raw.match(/.{1,5}/g) || [raw]
  return `NX-${blocks.join('-')}`
}

function verifyLicenseCode(code) {
  try {
    const raw = String(code ?? '').trim().replace(/^NX-/i, '').replace(/[\s-]/g, '')
    if (raw.length < 8) return { valid: false, reason: 'format' }
    const checksum = raw.slice(-3)
    const b64url = raw.slice(0, -3)
    const expected = (simpleHash(b64url + LICENSE_SALT) % 46656).toString(36).toUpperCase().padStart(3, '0')
    if (checksum.toUpperCase() !== expected) return { valid: false, reason: 'invalid' }
    const json = fromBase64Url(b64url)
    const payload = JSON.parse(json)
    if (!payload.s || !payload.e) return { valid: false, reason: 'format' }
    if (Date.now() > payload.e) return { valid: false, reason: 'expired', school: payload.s, expiresAt: payload.e }
    return { valid: true, school: payload.s, expiresAt: payload.e, code }
  } catch (e) {
    return { valid: false, reason: 'format' }
  }
}

function loadStoredLicense() {
  try {
    const raw = localStorage.getItem(LS_LICENSE)
    if (!raw) return null
    const { code } = JSON.parse(raw)
    const result = verifyLicenseCode(code)
    return result.valid ? result : { ...result, code }
  } catch {
    return null
  }
}

function saveLicense(code) {
  localStorage.setItem(LS_LICENSE, JSON.stringify({ code }))
}

/* ============================================================================
   PERSISTANCE DES CLASSES
   ========================================================================== */

function loadClasses() {
  try {
    const raw = localStorage.getItem(LS_CLASSES)
    if (raw) return JSON.parse(raw)
  } catch {}
  // Migration depuis l'ancien format mono-classe
  try {
    const legacy = localStorage.getItem(LS_LEGACY)
    if (legacy) {
      const old = JSON.parse(legacy)
      const migrated = { ...old, id: old.id || uid('cls') }
      return [migrated]
    }
  } catch {}
  return []
}

function saveClasses(classes) {
  localStorage.setItem(LS_CLASSES, JSON.stringify(classes))
}

/* ============================================================================
   IMPORT EXCEL EDUCMASTER
   Hypothèse de structure : ligne 0 = en-têtes de matière (label répété / fusionné
   visuellement), ligne 1 = sous-en-têtes (Matricule, Nom, Prénoms, puis paires
   "Note obtenue" / "Note perfectionnement"), lignes suivantes = données élèves.
   ========================================================================== */

function slugify(s) {
  return normalizeLabel(s).replace(/[^a-z0-9]+/g, '_').replace(/^_+|_+$/g, '') || 'matiere'
}

function parseEducMasterWorkbook(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: 'array' })
  const ws = wb.Sheets[wb.SheetNames[0]]
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true })
  if (!rows.length) throw new Error('Fichier vide')

  const headerRow0 = rows[0] || []
  const headerRow1 = rows[1] || []

  // Repère les 3 colonnes fixes par leur libellé (insensible à la casse/accents).
  // Certains exports mettent ces libellés sur la 1re ligne (fusionnée verticalement
  // avec la 2e), d'autres sur la 2e : on scanne les deux.
  let matCol = -1, nomCol = -1, prenomCol = -1
  ;[headerRow0, headerRow1].forEach((headerRow) => {
    headerRow.forEach((cell, c) => {
      const label = normalizeLabel(cell)
      if (matCol === -1 && label.includes('matricule')) matCol = c
      else if (nomCol === -1 && label.includes('nom') && !label.includes('prenom')) nomCol = c
      else if (prenomCol === -1 && label.includes('prenom')) prenomCol = c
    })
  })
  if (matCol === -1) matCol = 0
  if (nomCol === -1) nomCol = 1
  if (prenomCol === -1) prenomCol = 2

  const firstSubjectCol = Math.max(matCol, nomCol, prenomCol) + 1

  // Propage les libellés de matière fusionnés (cellule vide = même matière que la précédente)
  const subjectLabelByCol = {}
  let lastLabel = ''
  for (let c = firstSubjectCol; c < headerRow0.length; c++) {
    const raw = String(headerRow0[c] ?? '').trim()
    if (raw) lastLabel = raw
    subjectLabelByCol[c] = lastLabel
  }

  // Regroupe les colonnes deux par deux (obtenue / perfectionnement) par matière
  const subjects = []
  let c = firstSubjectCol
  let subjIndex = 0
  while (c < headerRow0.length) {
    const label = subjectLabelByCol[c]
    if (!label) { c += 1; continue }
    // cherche la fin de la plage de cette matière (colonnes consécutives avec le même libellé)
    let end = c
    while (end + 1 < headerRow0.length && subjectLabelByCol[end + 1] === label) end += 1
    const key = `${slugify(label)}_${subjIndex}`
    subjects.push({ key, label, col: c })
    subjIndex += 1
    c = end + 1
  }

  const roster = []
  for (let r = 2; r < rows.length; r++) {
    const row = rows[r]
    if (!row) continue
    const rawMat = row[matCol]
    if (rawMat === undefined || rawMat === null || String(rawMat).trim() === '') continue
    roster.push({
      matricule: cleanMatricule(rawMat),
      rawMatricule: rawMat,
      nom: String(row[nomCol] ?? '').trim(),
      prenoms: String(row[prenomCol] ?? '').trim(),
    })
  }

  if (!roster.length) throw new Error("Aucun élève détecté dans le fichier.")
  if (!subjects.length) throw new Error('Aucune matière détectée dans le fichier.')

  return { roster, subjects }
}

function buildExportWorkbook(klass) {
  const { roster, subjects, grades, attendance, className } = klass
  const headerRow0 = ['', '', '']
  const headerRow1 = ['Matricule', 'Nom', 'Prénoms']
  subjects.forEach((s) => {
    headerRow0.push(s.label, '')
    headerRow1.push('Note obtenue', 'Note perfectionnement')
  })

  const dataRows = roster.map((student) => {
    const row = [student.rawMatricule ?? student.matricule, student.nom, student.prenoms]
    subjects.forEach((s) => {
      const g = grades?.[student.matricule]?.[s.key]
      row.push(g?.obtenue ?? '', g?.perfectionnement ?? '')
    })
    return row
  })

  const sheetData = [headerRow0, headerRow1, ...dataRows]
  const ws = XLSX.utils.aoa_to_sheet(sheetData)

  const merges = []
  subjects.forEach((s, i) => {
    const col = 3 + i * 2
    merges.push({ s: { r: 0, c: col }, e: { r: 0, c: col + 1 } })
  })
  ws['!merges'] = merges
  ws['!cols'] = headerRow1.map(() => ({ wch: 14 }))

  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, ws, 'Notes')

  const ranking = computeRanking(roster, subjects, grades, attendance)
  const recapData = [
    ['Rang', 'Matricule', 'Nom', 'Prénoms', 'Moyenne / 20', 'Décision'],
    ...ranking.map((r) => [
      r.isAbsent ? 'ABS' : r.rank ?? '',
      r.rawMatricule ?? r.matricule,
      r.nom,
      r.prenoms,
      r.isAbsent ? 'Absent' : r.average !== null ? formatNum(r.average) : '',
      r.isAbsent ? '' : r.decision ?? '',
    ]),
  ]
  const wsRecap = XLSX.utils.aoa_to_sheet(recapData)
  wsRecap['!cols'] = [{ wch: 8 }, { wch: 12 }, { wch: 18 }, { wch: 18 }, { wch: 12 }, { wch: 14 }]
  XLSX.utils.book_append_sheet(wb, wsRecap, 'Récapitulatif')

  const fileName = `${(className || 'Classe').replace(/[^a-z0-9]+/gi, '_')}_NoteExpress.xlsx`
  XLSX.writeFile(wb, fileName)
}

/* ============================================================================
   PETITS COMPOSANTS D'INTERFACE PARTAGÉS
   ========================================================================== */

function Screen({ children, className = '' }) {
  return (
    <div className={`min-h-screen bg-[#F3F4F6] dark:bg-[#0F172A] text-[#1E3A8A] dark:text-[#E2E8F0] transition-colors ${className}`}>
      {children}
    </div>
  )
}

function TopBar({ title, subtitle, onBack, right }) {
  return (
    <div className="sticky top-0 z-20 bg-[#1E3A8A] dark:bg-[#0F172A] text-[#F3F4F6] px-4 pt-[calc(env(safe-area-inset-top)+0.9rem)] pb-3 flex items-center gap-3 border-b border-[#F59E0B]/30 shadow-sm">
      {onBack && (
        <button onClick={onBack} className="p-1.5 -ml-1.5 rounded-lg active:bg-white/10">
          <ChevronLeft size={22} />
        </button>
      )}
      <div className="flex-1 min-w-0">
        <div className="font-semibold text-[1.05rem] truncate">{title}</div>
        {subtitle && <div className="text-xs text-[#F59E0B] truncate">{subtitle}</div>}
      </div>
      {right}
    </div>
  )
}

function PrimaryButton({ children, onClick, disabled, className = '', type = 'button' }) {
  return (
    <button
      type={type}
      onClick={onClick}
      disabled={disabled}
      className={`w-full py-3.5 rounded-2xl bg-[#1E3A8A] dark:bg-[#F59E0B] text-white dark:text-[#0F172A] font-semibold text-[0.98rem] active:scale-[0.98] transition disabled:opacity-40 disabled:active:scale-100 shadow-sm ${className}`}
    >
      {children}
    </button>
  )
}

function GhostButton({ children, onClick, className = '' }) {
  return (
    <button
      onClick={onClick}
      className={`w-full py-3 rounded-2xl border border-[#1E3A8A]/20 dark:border-[#F59E0B]/30 font-medium text-[0.95rem] active:bg-black/5 dark:active:bg-white/5 transition ${className}`}
    >
      {children}
    </button>
  )
}

function EmptyState({ icon: Icon, title, hint }) {
  return (
    <div className="flex flex-col items-center justify-center text-center py-16 px-8 text-[#1E3A8A]/50 dark:text-[#E2E8F0]/40">
      {Icon && <Icon size={36} className="mb-3 opacity-60" />}
      <div className="font-medium">{title}</div>
      {hint && <div className="text-sm mt-1">{hint}</div>}
    </div>
  )
}

/* ============================================================================
   LICENCE — écran de verrouillage + panneau admin de génération de codes
   ========================================================================== */

function LicenseGate({ onActivated }) {
  const [code, setCode] = useState('')
  const [error, setError] = useState('')
  const [adminOpen, setAdminOpen] = useState(false)
  const [adminUnlocked, setAdminUnlocked] = useState(false)
  const [adminPass, setAdminPass] = useState('')
  const [school, setSchool] = useState('')
  const [days, setDays] = useState(365)
  const [generated, setGenerated] = useState('')
  const [copied, setCopied] = useState(false)
  const [logoTaps, setLogoTaps] = useState(0)

  function handleActivate() {
    const result = verifyLicenseCode(code)
    if (!result.valid) {
      setError(
        result.reason === 'expired'
          ? `Cette licence a expiré. Contactez le fournisseur pour la renouveler.`
          : 'Code invalide. Vérifiez la saisie (copier-coller conseillé).'
      )
      return
    }
    saveLicense(code)
    setError('')
    onActivated(result)
  }

  function handleLogoTap() {
    const next = logoTaps + 1
    setLogoTaps(next)
    if (next >= 5) {
      setAdminOpen(true)
      setLogoTaps(0)
    }
  }

  function handleGenerate() {
    if (!school.trim()) return
    setGenerated(generateLicenseCode({ school, days }))
    setCopied(false)
  }

  function handleCopy() {
    navigator.clipboard?.writeText(generated).then(() => {
      setCopied(true)
      setTimeout(() => setCopied(false), 1500)
    })
  }

  if (adminOpen) {
    return (
      <Screen>
        <TopBar title="Espace administrateur" subtitle="Génération de licences" onBack={() => { setAdminOpen(false); setAdminUnlocked(false) }} />
        <div className="p-5 max-w-md mx-auto space-y-5">
          {!adminUnlocked ? (
            <>
              <p className="text-sm opacity-70">Réservé au responsable de {APP_NAME}. Saisissez le code d'accès administrateur.</p>
              <input
                type="password"
                inputMode="numeric"
                value={adminPass}
                onChange={(e) => setAdminPass(e.target.value)}
                placeholder="Code administrateur"
                className="w-full px-4 py-3 rounded-xl border border-[#1E3A8A]/20 dark:border-[#F59E0B]/30 bg-white dark:bg-[#1E293B] outline-none focus:ring-2 focus:ring-[#F59E0B]"
              />
              <PrimaryButton onClick={() => adminPass === ADMIN_PASSCODE ? setAdminUnlocked(true) : setError('Code administrateur incorrect')}>
                Déverrouiller
              </PrimaryButton>
              {error && <p className="text-sm text-red-500 flex items-center gap-1.5"><AlertCircle size={15} />{error}</p>}
            </>
          ) : (
            <>
              <div className="flex items-center gap-2 text-[#F59E0B] mb-1">
                <ShieldCheck size={18} />
                <span className="font-medium text-sm">Générateur de licence</span>
              </div>
              <div>
                <label className="text-sm font-medium block mb-1.5">Nom de l'école</label>
                <input
                  value={school}
                  onChange={(e) => setSchool(e.target.value)}
                  placeholder="Ex : EPP Ganmi A"
                  className="w-full px-4 py-3 rounded-xl border border-[#1E3A8A]/20 dark:border-[#F59E0B]/30 bg-white dark:bg-[#1E293B] outline-none focus:ring-2 focus:ring-[#F59E0B]"
                />
              </div>
              <div>
                <label className="text-sm font-medium block mb-1.5">Durée de la licence</label>
                <div className="grid grid-cols-4 gap-2">
                  {[30, 90, 180, 365].map((d) => (
                    <button
                      key={d}
                      onClick={() => setDays(d)}
                      className={`py-2 rounded-xl text-sm font-medium border ${days === d ? 'bg-[#1E3A8A] dark:bg-[#F59E0B] text-white dark:text-[#0F172A] border-transparent' : 'border-[#1E3A8A]/20 dark:border-[#F59E0B]/30'}`}
                    >
                      {d === 365 ? '1 an' : `${d} j`}
                    </button>
                  ))}
                </div>
              </div>
              <PrimaryButton onClick={handleGenerate} disabled={!school.trim()}>Générer le code</PrimaryButton>
              {generated && (
                <div className="p-4 rounded-2xl bg-[#1E3A8A]/5 dark:bg-white/5 space-y-2">
                  <div className="text-xs opacity-60">Code à transmettre à l'école (WhatsApp, SMS...)</div>
                  <div className="font-mono text-sm break-all">{generated}</div>
                  <button onClick={handleCopy} className="flex items-center gap-1.5 text-sm text-[#F59E0B] font-medium">
                    <Copy size={15} /> {copied ? 'Copié !' : 'Copier le code'}
                  </button>
                </div>
              )}
            </>
          )}
        </div>
      </Screen>
    )
  }

  return (
    <Screen>
      <div className="min-h-screen flex flex-col justify-center px-6 py-10 max-w-md mx-auto">
        <div className="flex flex-col items-center mb-8">
          <button onClick={handleLogoTap} className="w-16 h-16 rounded-2xl bg-[#1E3A8A] dark:bg-[#F59E0B] flex items-center justify-center mb-4 active:scale-95 transition">
            <span className="text-white dark:text-[#0F172A] font-bold text-xl">NE</span>
          </button>
          <h1 className="text-xl font-bold">{APP_NAME}</h1>
          <p className="text-sm opacity-60 mt-1 text-center">Saisie rapide des notes EducMaster</p>
        </div>

        <div className="p-5 rounded-2xl bg-white dark:bg-[#1E293B] shadow-sm space-y-4">
          <div className="flex items-center gap-2 text-sm font-medium">
            <KeyRound size={16} className="text-[#F59E0B]" />
            Activation requise
          </div>
          <p className="text-sm opacity-70">
            Cette application nécessite un code de licence fourni par votre établissement ou le distributeur de {APP_NAME}.
          </p>
          <input
            value={code}
            onChange={(e) => { setCode(e.target.value); setError('') }}
            placeholder="NX-XXXXX-XXXXX-XXXXX"
            className="w-full px-4 py-3 rounded-xl border border-[#1E3A8A]/20 dark:border-[#F59E0B]/30 bg-transparent outline-none focus:ring-2 focus:ring-[#F59E0B] font-mono text-sm"
          />
          {error && <p className="text-sm text-red-500 flex items-center gap-1.5"><AlertCircle size={15} />{error}</p>}
          <PrimaryButton onClick={handleActivate} disabled={!code.trim()}>Activer</PrimaryButton>
        </div>

        <p className="text-xs opacity-40 text-center mt-6">Pas de code ? Contactez la personne qui vous a fourni {APP_NAME}.</p>
      </div>
    </Screen>
  )
}

function LicenseExpiryBanner({ license, onDismiss }) {
  if (!license?.expiresAt) return null
  const daysLeft = Math.ceil((license.expiresAt - Date.now()) / (24 * 60 * 60 * 1000))
  if (daysLeft > 7) return null
  return (
    <div className="bg-amber-100 dark:bg-amber-900/40 text-amber-900 dark:text-amber-200 text-xs px-4 py-2 flex items-center justify-between gap-2">
      <span className="flex items-center gap-1.5"><AlertCircle size={14} /> Licence expire dans {Math.max(daysLeft, 0)} jour{daysLeft > 1 ? 's' : ''}.</span>
      <button onClick={onDismiss}><X size={14} /></button>
    </div>
  )
}

/* ============================================================================
   ÉCRAN PIN
   ========================================================================== */

function PinLockScreen({ pin, onUnlock }) {
  const [entry, setEntry] = useState('')
  const [error, setError] = useState(false)

  function press(d) {
    if (entry.length >= 4) return
    const next = entry + d
    setEntry(next)
    setError(false)
    if (next.length === 4) {
      if (next === pin) {
        setTimeout(() => onUnlock(), 120)
      } else {
        setError(true)
        setTimeout(() => setEntry(''), 400)
      }
    }
  }

  return (
    <Screen>
      <div className="min-h-screen flex flex-col items-center justify-center px-6">
        <Lock size={28} className="mb-4 text-[#F59E0B]" />
        <p className="text-sm opacity-60 mb-6">Entrez le code PIN</p>
        <div className={`flex gap-3 mb-8 ${error ? 'animate-pulse' : ''}`}>
          {[0, 1, 2, 3].map((i) => (
            <div key={i} className={`w-3.5 h-3.5 rounded-full ${i < entry.length ? (error ? 'bg-red-500' : 'bg-[#1E3A8A] dark:bg-[#F59E0B]') : 'bg-[#1E3A8A]/15 dark:bg-white/15'}`} />
          ))}
        </div>
        <div className="grid grid-cols-3 gap-4 max-w-[280px]">
          {['1', '2', '3', '4', '5', '6', '7', '8', '9', '', '0', '⌫'].map((d, i) =>
            d === '' ? <div key={i} /> : (
              <button
                key={i}
                onClick={() => (d === '⌫' ? setEntry(entry.slice(0, -1)) : press(d))}
                className="w-16 h-16 rounded-full flex items-center justify-center text-lg font-medium bg-white dark:bg-[#1E293B] active:bg-[#1E3A8A]/10 dark:active:bg-white/10 shadow-sm"
              >
                {d}
              </button>
            )
          )}
        </div>
      </div>
    </Screen>
  )
}

/* ============================================================================
   SÉLECTEUR / GESTIONNAIRE DE CLASSES
   ========================================================================== */

function ClassSwitcher({ classes, activeId, onSelect, onCreate, onDelete, onClose }) {
  const [creating, setCreating] = useState(false)
  const [name, setName] = useState('')

  return (
    <div className="fixed inset-0 z-30 bg-black/40 flex items-end" onClick={onClose}>
      <div className="w-full bg-[#F3F4F6] dark:bg-[#0F172A] rounded-t-3xl max-h-[80vh] flex flex-col" onClick={(e) => e.stopPropagation()}>
        <div className="p-5 pb-3 flex items-center justify-between border-b border-[#1E3A8A]/10 dark:border-white/10">
          <h2 className="font-semibold">Mes classes</h2>
          <button onClick={onClose}><X size={20} /></button>
        </div>
        <div className="flex-1 overflow-y-auto p-4 space-y-2">
          {classes.map((c) => (
            <div
              key={c.id}
              onClick={() => onSelect(c.id)}
              className={`p-4 rounded-2xl flex items-center justify-between cursor-pointer ${c.id === activeId ? 'bg-[#1E3A8A] text-white dark:bg-[#F59E0B] dark:text-[#0F172A]' : 'bg-white dark:bg-[#1E293B]'}`}
            >
              <div className="min-w-0">
                <div className="font-medium truncate">{c.className}</div>
                <div className="text-xs opacity-70">{c.roster?.length ?? 0} élèves · {c.evaluationType}</div>
              </div>
              {classes.length > 1 && (
                <button
                  onClick={(e) => { e.stopPropagation(); onDelete(c.id) }}
                  className="p-2 opacity-60 active:opacity-100"
                >
                  <Trash2 size={17} />
                </button>
              )}
            </div>
          ))}
          {classes.length === 0 && <EmptyState icon={Layers} title="Aucune classe" hint="Créez votre première classe" />}
        </div>
        <div className="p-4 border-t border-[#1E3A8A]/10 dark:border-white/10">
          {creating ? (
            <div className="space-y-2">
              <input
                autoFocus
                value={name}
                onChange={(e) => setName(e.target.value)}
                placeholder="Nom de la classe (ex : CM2 Ganmi A)"
                className="w-full px-4 py-3 rounded-xl border border-[#1E3A8A]/20 dark:border-[#F59E0B]/30 bg-white dark:bg-[#1E293B] outline-none"
              />
              <div className="flex gap-2">
                <GhostButton onClick={() => setCreating(false)}>Annuler</GhostButton>
                <PrimaryButton onClick={() => { if (name.trim()) { onCreate(name.trim()); setName(''); setCreating(false) } }}>Créer</PrimaryButton>
              </div>
            </div>
          ) : (
            <PrimaryButton onClick={() => setCreating(true)}>
              <span className="flex items-center justify-center gap-2"><Plus size={18} /> Nouvelle classe</span>
            </PrimaryButton>
          )}
        </div>
      </div>
    </div>
  )
}

/* ============================================================================
   RÉORGANISATION DES MATIÈRES
   ========================================================================== */

function ReorderPanel({ subjects, onReorder, onClose }) {
  const [list, setList] = useState(subjects)

  function move(idx, dir) {
    const next = [...list]
    const target = idx + dir
    if (target < 0 || target >= next.length) return
    ;[next[idx], next[target]] = [next[target], next[idx]]
    setList(next)
  }

  return (
    <div className="fixed inset-0 z-30 bg-black/40 flex items-end" onClick={onClose}>
      <div className="w-full bg-[#F3F4F6] dark:bg-[#0F172A] rounded-t-3xl max-h-[80vh] flex flex-col" onClick={(e) => e.stopPropagation()}>
        <div className="p-5 pb-3 flex items-center justify-between border-b border-[#1E3A8A]/10 dark:border-white/10">
          <h2 className="font-semibold">Réorganiser les matières</h2>
          <button onClick={onClose}><X size={20} /></button>
        </div>
        <div className="flex-1 overflow-y-auto p-4 space-y-2">
          {list.map((s, idx) => (
            <div key={s.key} className="p-3.5 rounded-xl bg-white dark:bg-[#1E293B] flex items-center justify-between">
              <span className="font-medium text-sm">{s.label}</span>
              <div className="flex gap-1">
                <button onClick={() => move(idx, -1)} disabled={idx === 0} className="p-2 disabled:opacity-30"><ArrowUp size={16} /></button>
                <button onClick={() => move(idx, 1)} disabled={idx === list.length - 1} className="p-2 disabled:opacity-30"><ArrowDown size={16} /></button>
              </div>
            </div>
          ))}
        </div>
        <div className="p-4 border-t border-[#1E3A8A]/10 dark:border-white/10">
          <PrimaryButton onClick={() => { onReorder(list); onClose() }}>Enregistrer l'ordre</PrimaryButton>
        </div>
      </div>
    </div>
  )
}

/* ============================================================================
   NAVIGATION BASSE / PANNEAU "PLUS"
   ========================================================================== */

const TABS = [
  { key: 'eleves', label: 'Élèves', icon: Users },
  { key: 'saisie', label: 'Saisie', icon: Edit3 },
  { key: 'matieres', label: 'Matières', icon: BookOpen },
  { key: 'classement', label: 'Classement', icon: Trophy },
  { key: 'plus', label: 'Plus', icon: MoreHorizontal },
]

function BottomNav({ active, onChange, onMore }) {
  return (
    <div className="fixed bottom-0 left-0 right-0 z-20 bg-white dark:bg-[#1E293B] border-t border-[#1E3A8A]/10 dark:border-white/10 pb-[env(safe-area-inset-bottom)]">
      <div className="flex">
        {TABS.map((t) => {
          const Icon = t.icon
          const isActive = active === t.key
          return (
            <button
              key={t.key}
              onClick={() => (t.key === 'plus' ? onMore() : onChange(t.key))}
              className="flex-1 flex flex-col items-center gap-1 py-2.5"
            >
              <Icon size={20} className={isActive ? 'text-[#F59E0B]' : 'opacity-50'} />
              <span className={`text-[0.68rem] ${isActive ? 'text-[#F59E0B] font-medium' : 'opacity-50'}`}>{t.label}</span>
            </button>
          )
        })}
      </div>
    </div>
  )
}

function MorePanel({
  onClose, onSwitchClass, onReorder, dark, onToggleDark, pin, onSetPin, onRemovePin,
  onUndo, canUndo, onDeleteClass, defaultMode, onSetDefaultMode, onGoWelcome, onStats,
}) {
  const [pinSetup, setPinSetup] = useState(false)
  const [newPin, setNewPin] = useState('')

  const Row = ({ icon: Icon, label, onClick, danger }) => (
    <button onClick={onClick} className={`w-full flex items-center gap-3 px-4 py-3.5 rounded-xl bg-white dark:bg-[#1E293B] ${danger ? 'text-red-500' : ''}`}>
      <Icon size={18} />
      <span className="flex-1 text-left text-sm font-medium">{label}</span>
      <ChevronRight size={16} className="opacity-40" />
    </button>
  )

  return (
    <div className="fixed inset-0 z-30 bg-black/40 flex items-end" onClick={onClose}>
      <div className="w-full bg-[#F3F4F6] dark:bg-[#0F172A] rounded-t-3xl max-h-[85vh] flex flex-col" onClick={(e) => e.stopPropagation()}>
        <div className="p-5 pb-3 flex items-center justify-between border-b border-[#1E3A8A]/10 dark:border-white/10">
          <h2 className="font-semibold">Plus</h2>
          <button onClick={onClose}><X size={20} /></button>
        </div>
        <div className="flex-1 overflow-y-auto p-4 space-y-2">
          <Row icon={Layers} label="Changer de classe" onClick={onSwitchClass} />
          <Row icon={BarChart3} label="Statistiques par matière" onClick={onStats} />
          <Row icon={ArrowUp} label="Réorganiser les matières" onClick={onReorder} />
          <Row icon={RotateCcw} label={canUndo ? 'Annuler la dernière saisie' : 'Rien à annuler'} onClick={canUndo ? onUndo : undefined} />

          <button onClick={onToggleDark} className="w-full flex items-center gap-3 px-4 py-3.5 rounded-xl bg-white dark:bg-[#1E293B]">
            {dark ? <Sun size={18} /> : <Moon size={18} />}
            <span className="flex-1 text-left text-sm font-medium">Mode {dark ? 'clair' : 'sombre'}</span>
          </button>

          <button onClick={() => onSetDefaultMode(defaultMode === 'eleve' ? 'matiere' : 'eleve')} className="w-full flex items-center gap-3 px-4 py-3.5 rounded-xl bg-white dark:bg-[#1E293B]">
            <ClipboardList size={18} />
            <span className="flex-1 text-left text-sm font-medium">Mode de saisie par défaut : {defaultMode === 'eleve' ? 'par élève' : 'par matière'}</span>
          </button>

          {!pinSetup ? (
            <Row icon={Lock} label={pin ? 'Changer / retirer le PIN' : 'Verrouiller par code PIN'} onClick={() => setPinSetup(true)} />
          ) : (
            <div className="p-4 rounded-xl bg-white dark:bg-[#1E293B] space-y-2">
              <input
                value={newPin}
                onChange={(e) => setNewPin(e.target.value.replace(/\D/g, '').slice(0, 4))}
                placeholder="Nouveau code à 4 chiffres"
                inputMode="numeric"
                className="w-full px-3 py-2.5 rounded-lg border border-[#1E3A8A]/20 dark:border-[#F59E0B]/30 bg-transparent outline-none"
              />
              <div className="flex gap-2">
                {pin && <GhostButton onClick={() => { onRemovePin(); setPinSetup(false) }}>Retirer le PIN</GhostButton>}
                <PrimaryButton onClick={() => { if (newPin.length === 4) { onSetPin(newPin); setPinSetup(false); setNewPin('') } }} disabled={newPin.length !== 4}>
                  Valider
                </PrimaryButton>
              </div>
            </div>
          )}

          <Row icon={Home} label="Retour à l'écran d'accueil" onClick={onGoWelcome} />
          <Row icon={Trash2} label="Supprimer cette classe" onClick={onDeleteClass} danger />
        </div>
        <div className="p-4 text-center text-xs opacity-40">{APP_NAME} v{APP_VERSION}</div>
      </div>
    </div>
  )
}

/* ============================================================================
   PARCOURS INITIAL D'UNE CLASSE NEUVE
   ========================================================================== */

function WelcomeScreen({ onStart, onReset }) {
  return (
    <Screen>
      <div className="min-h-screen flex flex-col items-center justify-center px-6 text-center">
        <div className="mb-2 px-3.5 py-1.5 rounded-full bg-[#F59E0B]/10 border border-[#F59E0B]/30 flex items-center gap-2">
          <span className="w-1.5 h-1.5 rounded-full bg-[#F59E0B]" />
          <span className="text-[0.68rem] font-semibold tracking-wide text-[#F59E0B] uppercase">Saisie ultra-rapide</span>
        </div>

        <h1 className="mt-4 mb-1 font-black tracking-tight leading-none text-[2.6rem] sm:text-5xl">
          <span className="text-[#1E3A8A] dark:text-white">Note</span>
          <span className="text-[#F59E0B]">Express</span>
        </h1>
        <div className="w-16 h-1 rounded-full bg-gradient-to-r from-[#1E3A8A] to-[#F59E0B] mb-6" />

        <p className="opacity-70 mb-6 max-w-xs text-[0.95rem] leading-relaxed">
          La saisie intelligente qui facilite <span className="font-semibold">EducMaster</span>.
        </p>

        <div className="w-full max-w-xs p-4 rounded-2xl border border-[#1E3A8A]/15 dark:border-[#F59E0B]/20 bg-[#1E3A8A]/[0.03] dark:bg-white/5 mb-8">
          <p className="text-sm leading-relaxed">
            Une nouvelle façon de saisir les notes d'évaluation, sans le stress des cellules Excel.
          </p>
        </div>

        <div className="w-full max-w-xs space-y-3">
          <PrimaryButton onClick={onStart}>
            <span className="flex items-center justify-center gap-2">Commencer <ChevronRight size={18} /></span>
          </PrimaryButton>
          {onReset && (
            <button onClick={onReset} className="w-full py-2 text-xs font-medium text-red-500/80 flex items-center justify-center gap-1.5">
              <Trash2 size={13} /> Réinitialiser l'application
            </button>
          )}
        </div>
      </div>
    </Screen>
  )
}

function ImportScreen({ onImported, onBack, onReset }) {
  const [error, setError] = useState('')
  const [loading, setLoading] = useState(false)
  const fileRef = useRef(null)

  async function handleFile(file) {
    setError('')
    setLoading(true)
    try {
      const buf = await file.arrayBuffer()
      const { roster, subjects } = parseEducMasterWorkbook(buf)
      onImported({ roster, subjects })
    } catch (e) {
      setError(e.message || "Impossible de lire ce fichier. Vérifiez qu'il s'agit bien d'un export EducMaster.")
    } finally {
      setLoading(false)
    }
  }

  return (
    <Screen>
      <TopBar title="Importation" onBack={onBack} />
      <div className="p-6 flex flex-col items-center text-center pt-10">
        <div className="w-16 h-16 rounded-2xl bg-[#F59E0B]/10 flex items-center justify-center mb-5">
          <Upload size={28} className="text-[#F59E0B]" />
        </div>
        <h2 className="font-bold text-xl mb-3">Importation</h2>
        <div className="w-full max-w-xs p-4 rounded-2xl bg-[#1E3A8A]/[0.04] dark:bg-white/5 mb-6">
          <p className="text-sm leading-relaxed">
            Importez le fichier Excel EducMaster de la liste de saisie des notes de votre classe.
          </p>
        </div>

        <input ref={fileRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={(e) => e.target.files[0] && handleFile(e.target.files[0])} />
        <button
          onClick={() => fileRef.current?.click()}
          disabled={loading}
          className="w-full max-w-xs p-8 rounded-3xl border-2 border-dashed border-[#1E3A8A]/25 dark:border-[#F59E0B]/25 flex flex-col items-center gap-3 active:scale-[0.99] transition disabled:opacity-50"
        >
          <Upload size={26} className="text-[#1E3A8A] dark:text-[#F59E0B]" />
          <div className="font-semibold text-sm">{loading ? 'Lecture en cours…' : 'Fichier Excel'}</div>
          <div className="text-xs opacity-50">Touchez pour importer</div>
        </button>

        {error && (
          <div className="mt-5 p-3.5 rounded-xl bg-red-50 dark:bg-red-900/20 text-red-600 dark:text-red-300 text-sm flex items-start gap-2 max-w-xs text-left">
            <AlertCircle size={16} className="shrink-0 mt-0.5" />
            {error}
          </div>
        )}

        <div className="mt-8 space-y-3 w-full max-w-xs">
          {onBack && <button onClick={onBack} className="text-sm font-medium opacity-60">Retour à l'accueil</button>}
          {onReset && (
            <button onClick={onReset} className="w-full py-1 text-xs font-medium text-red-500/80 flex items-center justify-center gap-1.5">
              <Trash2 size={13} /> Réinitialiser l'application
            </button>
          )}
        </div>
      </div>
    </Screen>
  )
}

function EvaluationPicker({ onPick }) {
  return (
    <Screen>
      <TopBar title="Type d'évaluation" />
      <div className="p-5 space-y-2.5 pt-8">
        {EVALUATION_TYPES.map((ev) => (
          <button
            key={ev}
            onClick={() => onPick(ev)}
            className="w-full p-4 rounded-2xl bg-white dark:bg-[#1E293B] text-left flex items-center justify-between shadow-sm active:scale-[0.99] transition"
          >
            <span className="font-medium text-sm">{ev}</span>
            <ChevronRight size={18} className="opacity-40" />
          </button>
        ))}
      </div>
    </Screen>
  )
}

function EntryModePicker({ onPick }) {
  return (
    <Screen>
      <TopBar title="Mode de saisie" />
      <div className="p-5 pt-8 space-y-3">
        <p className="text-sm opacity-60 px-1 mb-2">Comment souhaitez-vous saisir les notes ? Vous pourrez changer à tout moment.</p>
        <button onClick={() => onPick('eleve')} className="w-full p-5 rounded-2xl bg-white dark:bg-[#1E293B] text-left shadow-sm active:scale-[0.99] transition flex items-center gap-4">
          <Users size={26} className="text-[#F59E0B]" />
          <div>
            <div className="font-semibold">Par élève</div>
            <div className="text-xs opacity-60">Un élève à la fois, toutes ses matières</div>
          </div>
        </button>
        <button onClick={() => onPick('matiere')} className="w-full p-5 rounded-2xl bg-white dark:bg-[#1E293B] text-left shadow-sm active:scale-[0.99] transition flex items-center gap-4">
          <BookOpen size={26} className="text-[#F59E0B]" />
          <div>
            <div className="font-semibold">Par matière</div>
            <div className="text-xs opacity-60">Une matière à la fois, tous les élèves</div>
          </div>
        </button>
      </div>
    </Screen>
  )
}

function GradeBreakdown({ grade }) {
  if (!grade || grade.obtenue === null || grade.obtenue === undefined) return null
  const perf = grade.perfectionnement ?? 0
  const total = toNum(grade.obtenue) + toNum(perf)
  return (
    <div className="mt-2 pt-2 border-t border-[#1E3A8A]/10 dark:border-[#F59E0B]/15 flex items-center justify-between">
      <div className="text-xs opacity-60">
        Note : <span className="font-semibold opacity-100">{formatNum(grade.obtenue)}</span>
        {'  '}| Perf : <span className="font-semibold opacity-100">{formatNum(perf)}</span>
      </div>
      <div className="px-2.5 py-1 rounded-lg bg-[#1E3A8A] dark:bg-[#F59E0B] text-white dark:text-[#0F172A] text-xs font-bold">
        {formatNum(total)}/20
      </div>
    </div>
  )
}

/* ============================================================================
   SAISIE PAR ÉLÈVE
   ========================================================================== */

function StudentEntryTab({ klass, onSetGrade, onToggleAttendance }) {
  const { roster, subjects, grades, attendance } = klass
  const [index, setIndex] = useState(0)
  const student = roster[index]
  const inputRefs = useRef({})

  if (!student) return <EmptyState icon={Users} title="Aucun élève" />

  const isAbsent = attendance?.[student.matricule] === false

  function handleChange(subjKey, raw) {
    onSetGrade(student.matricule, subjKey, raw)
    if (isCodeComplete(raw)) {
      const currentIdx = subjects.findIndex((s) => s.key === subjKey)
      const next = subjects[currentIdx + 1]
      if (next) {
        setTimeout(() => inputRefs.current[next.key]?.focus(), 30)
      }
      if (navigator.vibrate) navigator.vibrate(15)
    }
  }

  return (
    <div className="pb-28">
      <div className="px-4 pt-4 pb-2 flex items-center justify-between">
        <button onClick={() => setIndex(Math.max(0, index - 1))} disabled={index === 0} className="p-2 disabled:opacity-30"><ChevronLeft size={20} /></button>
        <div className="text-center min-w-0 flex-1">
          <div className="font-semibold truncate">{student.nom} {student.prenoms}</div>
          <div className="text-xs opacity-50">{index + 1} / {roster.length} · Matricule {student.matricule}</div>
        </div>
        <button onClick={() => setIndex(Math.min(roster.length - 1, index + 1))} disabled={index === roster.length - 1} className="p-2 disabled:opacity-30"><ChevronRight size={20} /></button>
      </div>

      <div className="px-4 pb-3">
        <button
          onClick={() => onToggleAttendance(student.matricule)}
          className={`w-full py-2.5 rounded-xl text-sm font-medium flex items-center justify-center gap-2 ${isAbsent ? 'bg-red-100 text-red-600 dark:bg-red-900/30 dark:text-red-300' : 'bg-green-100 text-green-700 dark:bg-green-900/30 dark:text-green-300'}`}
        >
          {isAbsent ? <CircleX size={16} /> : <CircleCheck size={16} />}
          {isAbsent ? 'Élève absent' : 'Élève présent'}
        </button>
      </div>

      {isAbsent ? (
        <EmptyState icon={CircleX} title="Élève marqué absent" hint="Aucune note à saisir pour cette évaluation." />
      ) : (
        <div className="px-4 space-y-2.5">
          {subjects.map((s) => {
            const g = grades?.[student.matricule]?.[s.key]
            return (
              <div key={s.key} className="p-3.5 rounded-2xl bg-white dark:bg-[#1E293B]">
                <div className="flex items-center justify-between gap-3">
                  <span className="text-sm font-medium flex-1 min-w-0 truncate">{s.label}</span>
                  <input
                    ref={(el) => (inputRefs.current[s.key] = el)}
                    value={g?.rawCode ?? ''}
                    onChange={(e) => handleChange(s.key, e.target.value)}
                    placeholder="1402"
                    inputMode="decimal"
                    className="w-24 text-center px-2 py-2 rounded-lg border border-[#1E3A8A]/15 dark:border-[#F59E0B]/25 bg-transparent outline-none focus:ring-2 focus:ring-[#F59E0B] font-mono"
                  />
                </div>
                <GradeBreakdown grade={g} />
              </div>
            )
          })}
        </div>
      )}
    </div>
  )
}

/* ============================================================================
   LISTE DES ÉLÈVES
   ========================================================================== */

function ElevesTab({ klass, onToggleAttendance, onRemoveStudent }) {
  const { roster, subjects, grades, attendance } = klass
  const ranking = useMemo(() => computeRanking(roster, subjects, grades, attendance), [roster, subjects, grades, attendance])
  const byMatricule = useMemo(() => Object.fromEntries(ranking.map((r) => [r.matricule, r])), [ranking])

  if (!roster.length) return <EmptyState icon={Users} title="Aucun élève importé" />

  return (
    <div className="px-4 pt-4 pb-28 space-y-2">
      {roster.map((student) => {
        const r = byMatricule[student.matricule]
        const isAbsent = attendance?.[student.matricule] === false
        return (
          <div key={student.matricule} className="p-3.5 rounded-2xl bg-white dark:bg-[#1E293B] flex items-center gap-3">
            <div className="min-w-0 flex-1">
              <div className="font-medium text-sm truncate">{student.nom} {student.prenoms}</div>
              <div className="text-xs opacity-50 truncate">Matricule {student.matricule}</div>
            </div>
            <div className="text-right shrink-0">
              {isAbsent ? (
                <span className="text-xs font-medium text-red-500">Absent</span>
              ) : (
                <span className="text-sm font-semibold">{r?.average !== null && r?.average !== undefined ? `${formatNum(r.average)}/20` : '—'}</span>
              )}
            </div>
            <button onClick={() => onToggleAttendance(student.matricule)} className="p-2 shrink-0">
              {isAbsent ? <CircleX size={17} className="text-red-500" /> : <CircleCheck size={17} className="text-green-600" />}
            </button>
            <button onClick={() => onRemoveStudent(student.matricule)} className="p-2 shrink-0 opacity-50">
              <Trash2 size={16} />
            </button>
          </div>
        )
      })}
    </div>
  )
}

/* ============================================================================
   SAISIE PAR MATIÈRE
   ========================================================================== */

function SubjectTab({ klass, onSetGrade }) {
  const { roster, subjects, grades, attendance } = klass
  const [subjIndex, setSubjIndex] = useState(0)
  const subject = subjects[subjIndex]
  const inputRefs = useRef({})

  if (!subject) return <EmptyState icon={BookOpen} title="Aucune matière" />

  function handleChange(matricule, raw, idx) {
    onSetGrade(matricule, subject.key, raw)
    if (isCodeComplete(raw)) {
      const nextStudent = roster[idx + 1]
      if (nextStudent) setTimeout(() => inputRefs.current[nextStudent.matricule]?.focus(), 30)
      if (navigator.vibrate) navigator.vibrate(15)
    }
  }

  return (
    <div className="pb-28">
      <div className="px-4 pt-4 pb-3 flex items-center justify-between">
        <button onClick={() => setSubjIndex(Math.max(0, subjIndex - 1))} disabled={subjIndex === 0} className="p-2 disabled:opacity-30"><ChevronLeft size={20} /></button>
        <div className="text-center min-w-0 flex-1">
          <div className="font-semibold truncate">{subject.label}</div>
          <div className="text-xs opacity-50">{subjIndex + 1} / {subjects.length}</div>
        </div>
        <button onClick={() => setSubjIndex(Math.min(subjects.length - 1, subjIndex + 1))} disabled={subjIndex === subjects.length - 1} className="p-2 disabled:opacity-30"><ChevronRight size={20} /></button>
      </div>

      <div className="px-4 space-y-2.5">
        {roster.map((student, idx) => {
          const isAbsent = attendance?.[student.matricule] === false
          const g = grades?.[student.matricule]?.[subject.key]
          return (
            <div key={student.matricule} className={`p-3.5 rounded-2xl bg-white dark:bg-[#1E293B] ${isAbsent ? 'opacity-40' : ''}`}>
              <div className="flex items-center justify-between gap-3">
                <span className="text-sm font-medium flex-1 min-w-0 truncate">{student.nom} {student.prenoms}</span>
                <input
                  ref={(el) => (inputRefs.current[student.matricule] = el)}
                  value={g?.rawCode ?? ''}
                  onChange={(e) => handleChange(student.matricule, e.target.value, idx)}
                  placeholder="1402"
                  inputMode="decimal"
                  disabled={isAbsent}
                  className="w-24 text-center px-2 py-2 rounded-lg border border-[#1E3A8A]/15 dark:border-[#F59E0B]/25 bg-transparent outline-none focus:ring-2 focus:ring-[#F59E0B] font-mono disabled:bg-black/5"
                />
              </div>
              {!isAbsent && <GradeBreakdown grade={g} />}
            </div>
          )
        })}
      </div>
    </div>
  )
}

/* ============================================================================
   CLASSEMENT
   ========================================================================== */

function RecapTab({ klass, onExport }) {
  const { roster, subjects, grades, attendance } = klass
  const ranking = useMemo(() => computeRanking(roster, subjects, grades, attendance), [roster, subjects, grades, attendance])
  const admisCount = ranking.filter((r) => r.decision === 'Admis').length
  const notesCount = ranking.filter((r) => !r.isAbsent).length

  if (!roster.length) return <EmptyState icon={Trophy} title="Aucun élève à classer" />

  return (
    <div className="px-4 pt-4 pb-28 space-y-3">
      <div className="p-4 rounded-2xl bg-[#1E3A8A] dark:bg-[#F59E0B] text-white dark:text-[#0F172A] flex items-center justify-between">
        <div>
          <div className="text-xs opacity-80">Résumé de la classe</div>
          <div className="font-semibold">{admisCount} admis sur {notesCount}</div>
        </div>
        <button onClick={onExport} className="p-2.5 rounded-xl bg-white/15 dark:bg-black/10 active:scale-95 transition">
          <Download size={18} />
        </button>
      </div>

      {ranking.map((r) => (
        <div key={r.matricule} className="p-3.5 rounded-2xl bg-white dark:bg-[#1E293B] flex items-center gap-3">
          <div className="w-9 h-9 rounded-full bg-[#1E3A8A]/5 dark:bg-white/5 flex items-center justify-center text-xs font-semibold shrink-0">
            {r.isAbsent ? '—' : r.rank}
          </div>
          <div className="min-w-0 flex-1">
            <div className="font-medium text-sm truncate">{r.nom} {r.prenoms}</div>
            {!r.isAbsent && (
              <div className={`text-xs font-medium ${r.decision === 'Admis' ? 'text-green-600' : 'text-red-500'}`}>{r.decision}</div>
            )}
          </div>
          <div className="text-right shrink-0">
            {r.isAbsent ? (
              <span className="text-xs font-medium text-red-500">Absent</span>
            ) : (
              <span className="font-semibold text-sm">{formatNum(r.average)}/20</span>
            )}
          </div>
        </div>
      ))}
    </div>
  )
}

/* ============================================================================
   STATISTIQUES PAR MATIÈRE
   ========================================================================== */

function StatsTab({ klass, onClose }) {
  const { roster, subjects, grades, attendance } = klass
  const stats = useMemo(() => computeSubjectStats(roster, subjects, grades, attendance), [roster, subjects, grades, attendance])

  return (
    <Screen>
      <TopBar title="Statistiques par matière" onBack={onClose} />
      <div className="px-4 pt-4 pb-10 space-y-2.5">
        {stats.map((s) => (
          <div key={s.key} className="p-4 rounded-2xl bg-white dark:bg-[#1E293B]">
            <div className="font-medium text-sm mb-2">{s.label}</div>
            <div className="flex items-center justify-between text-xs opacity-60 mb-1">
              <span>Moyenne classe</span>
              <span className="font-semibold text-[#1E3A8A] dark:text-[#F59E0B] text-sm">{s.average !== null ? `${formatNum(s.average)}/20` : '—'}</span>
            </div>
            <div className="flex items-center justify-between text-xs opacity-60">
              <span>Taux de réussite (≥ {SUBJECT_PASS}/20)</span>
              <span className="font-semibold text-sm">{s.successRate !== null ? `${Math.round(s.successRate)}%` : '—'}</span>
            </div>
            <div className="mt-2 h-1.5 rounded-full bg-black/5 dark:bg-white/10 overflow-hidden">
              <div className="h-full bg-[#F59E0B]" style={{ width: `${s.successRate ?? 0}%` }} />
            </div>
          </div>
        ))}
        {stats.length === 0 && <EmptyState icon={BarChart3} title="Aucune matière" />}
      </div>
    </Screen>
  )
}

/* ============================================================================
   COMPOSANT RACINE
   ========================================================================== */

export default function App() {
  const [licenseStatus, setLicenseStatus] = useState('checking') // checking | locked | unlocked
  const [license, setLicense] = useState(null)
  const [bannerDismissed, setBannerDismissed] = useState(false)

  const [classes, setClasses] = useState([])
  const [activeId, setActiveId] = useState(null)
  const [dark, setDark] = useState(false)
  const [pin, setPin] = useState(null)
  const [pinLocked, setPinLocked] = useState(false)
  const [defaultMode, setDefaultMode] = useState('eleve')

  const [activeTab, setActiveTab] = useState('eleves')
  const [classSwitcherOpen, setClassSwitcherOpen] = useState(false)
  const [morePanelOpen, setMorePanelOpen] = useState(false)
  const [reorderOpen, setReorderOpen] = useState(false)
  const [statsOpen, setStatsOpen] = useState(false)
  const [undoStack, setUndoStack] = useState([])

  // --- Initialisation ---
  useEffect(() => {
    const stored = loadStoredLicense()
    if (stored?.valid) {
      setLicense(stored)
      setLicenseStatus('unlocked')
    } else {
      setLicenseStatus('locked')
    }

    setClasses(loadClasses())
    setActiveId(localStorage.getItem(LS_ACTIVE) || null)
    setDark(localStorage.getItem(LS_DARK) === '1')
    setPin(localStorage.getItem(LS_PIN) || null)
    setPinLocked(!!localStorage.getItem(LS_PIN))
    setDefaultMode(localStorage.getItem(LS_DEFAULT_MODE) || 'eleve')
  }, [])

  useEffect(() => {
    document.documentElement.classList.toggle('dark', dark)
  }, [dark])

  useEffect(() => {
    if (classes.length) saveClasses(classes)
  }, [classes])

  useEffect(() => {
    if (activeId) localStorage.setItem(LS_ACTIVE, activeId)
  }, [activeId])

  const activeClass = useMemo(() => classes.find((c) => c.id === activeId) || null, [classes, activeId])

  function updateClass(id, updater) {
    setClasses((prev) => prev.map((c) => (c.id === id ? updater(c) : c)))
  }

  function pushUndo(matricule, subjKey, prevGrade) {
    setUndoStack((prev) => [...prev.slice(-19), { classId: activeId, matricule, subjKey, prevGrade }])
  }

  function handleSetGrade(matricule, subjKey, rawCode) {
    if (!activeClass) return
    const prevGrade = activeClass.grades?.[matricule]?.[subjKey] || null
    pushUndo(matricule, subjKey, prevGrade)
    const parsed = parseCode(rawCode)
    updateClass(activeClass.id, (c) => ({
      ...c,
      grades: {
        ...c.grades,
        [matricule]: {
          ...c.grades?.[matricule],
          [subjKey]: {
            rawCode,
            obtenue: parsed && !Number.isNaN(parsed.obtenue) ? parsed.obtenue : null,
            perfectionnement: parsed && parsed.perfectionnement !== null && !Number.isNaN(parsed.perfectionnement) ? parsed.perfectionnement : null,
          },
        },
      },
      updatedAt: Date.now(),
    }))
  }

  function handleUndo() {
    const last = undoStack[undoStack.length - 1]
    if (!last) return
    setUndoStack((prev) => prev.slice(0, -1))
    updateClass(last.classId, (c) => ({
      ...c,
      grades: {
        ...c.grades,
        [last.matricule]: {
          ...c.grades?.[last.matricule],
          [last.subjKey]: last.prevGrade,
        },
      },
    }))
  }

  function handleToggleAttendance(matricule) {
    if (!activeClass) return
    updateClass(activeClass.id, (c) => {
      const wasAbsent = c.attendance?.[matricule] === false
      const nextAttendance = { ...c.attendance }
      if (wasAbsent) delete nextAttendance[matricule]
      else nextAttendance[matricule] = false
      return { ...c, attendance: nextAttendance }
    })
  }

  function handleRemoveStudent(matricule) {
    if (!activeClass) return
    updateClass(activeClass.id, (c) => ({ ...c, roster: c.roster.filter((s) => s.matricule !== matricule) }))
  }

  function handleCreateClass(name) {
    const newClass = {
      id: uid('cls'),
      className: name,
      evaluationType: null,
      entryModeChosen: false,
      roster: [],
      subjects: [],
      grades: {},
      attendance: {},
      updatedAt: Date.now(),
    }
    setClasses((prev) => [...prev, newClass])
    setActiveId(newClass.id)
    setClassSwitcherOpen(false)
  }

  function handleDeleteClass(id) {
    setClasses((prev) => {
      const next = prev.filter((c) => c.id !== id)
      if (id === activeId) setActiveId(next[0]?.id || null)
      return next
    })
  }

  function handleSetPin(newPin) {
    localStorage.setItem(LS_PIN, newPin)
    setPin(newPin)
  }

  function handleRemovePin() {
    localStorage.removeItem(LS_PIN)
    setPin(null)
  }

  function handleSetDefaultMode(mode) {
    localStorage.setItem(LS_DEFAULT_MODE, mode)
    setDefaultMode(mode)
  }

  function handleReorderSubjects(newSubjects) {
    if (!activeClass) return
    updateClass(activeClass.id, (c) => ({ ...c, subjects: newSubjects }))
  }

  function handleGoWelcome() {
    if (!activeClass) return
    updateClass(activeClass.id, (c) => ({ ...c, entryModeChosen: false, evaluationType: null, roster: [], subjects: [], grades: {}, attendance: {} }))
    setMorePanelOpen(false)
  }

  function handleResetApplication() {
    if (!window.confirm('Réinitialiser NoteExpress ? Toutes les classes et notes enregistrées sur cet appareil seront définitivement supprimées.')) return
    localStorage.removeItem(LS_CLASSES)
    localStorage.removeItem(LS_ACTIVE)
    localStorage.removeItem(LS_PIN)
    localStorage.removeItem(LS_DARK)
    localStorage.removeItem(LS_DEFAULT_MODE)
    setClasses([])
    setActiveId(null)
    setPin(null)
    setPinLocked(false)
    setDark(false)
  }

  /* ------------------------- Rendus conditionnels ------------------------- */

  if (licenseStatus === 'checking') return null

  if (licenseStatus === 'locked') {
    return (
      <LicenseGate
        onActivated={(result) => {
          setLicense(result)
          setLicenseStatus('unlocked')
        }}
      />
    )
  }

  if (pinLocked) {
    return <PinLockScreen pin={pin} onUnlock={() => setPinLocked(false)} />
  }

  // Aucune classe encore créée
  if (!activeClass) {
    return (
      <Screen>
        <WelcomeScreen onStart={() => handleCreateClass('Ma classe')} onReset={classes.length ? handleResetApplication : undefined} />
      </Screen>
    )
  }

  // Parcours de configuration initiale d'une classe neuve
  if (!activeClass.roster.length) {
    return (
      <ImportScreen
        onImported={({ roster, subjects }) => updateClass(activeClass.id, (c) => ({ ...c, roster, subjects }))}
        onBack={classes.length > 1 ? () => handleDeleteClass(activeClass.id) : undefined}
        onReset={handleResetApplication}
      />
    )
  }
  if (!activeClass.evaluationType) {
    return (
      <EvaluationPicker
        onPick={(ev) => updateClass(activeClass.id, (c) => ({ ...c, evaluationType: ev }))}
      />
    )
  }
  if (!activeClass.entryModeChosen) {
    return (
      <EntryModePicker
        onPick={(mode) => {
          setActiveTab(mode === 'eleve' ? 'saisie' : 'matieres')
          updateClass(activeClass.id, (c) => ({ ...c, entryModeChosen: true }))
        }}
      />
    )
  }

  if (statsOpen) {
    return <StatsTab klass={activeClass} onClose={() => setStatsOpen(false)} />
  }

  return (
    <Screen>
      <TopBar
        title={activeClass.className}
        subtitle={activeClass.evaluationType}
        right={
          <div className="flex items-center gap-1 -mr-1.5">
            <button
              onClick={() => buildExportWorkbook(activeClass)}
              title="Exporter"
              className="p-2 rounded-lg active:bg-white/10"
            >
              <Download size={18} />
            </button>
            <button
              onClick={() => (pin ? setPinLocked(true) : setMorePanelOpen(true))}
              title="Verrouiller"
              className="p-2 rounded-lg active:bg-white/10"
            >
              <Lock size={18} className={pin ? '' : 'opacity-40'} />
            </button>
            <button onClick={() => setClassSwitcherOpen(true)} title="Changer de classe" className="p-2 rounded-lg active:bg-white/10">
              <Layers size={18} />
            </button>
          </div>
        }
      />
      <LicenseExpiryBanner license={license} onDismiss={() => setBannerDismissed(true)} />

      {activeTab === 'eleves' && (
        <ElevesTab klass={activeClass} onToggleAttendance={handleToggleAttendance} onRemoveStudent={handleRemoveStudent} />
      )}
      {activeTab === 'saisie' && (
        <StudentEntryTab klass={activeClass} onSetGrade={handleSetGrade} onToggleAttendance={handleToggleAttendance} />
      )}
      {activeTab === 'matieres' && (
        <SubjectTab klass={activeClass} onSetGrade={handleSetGrade} />
      )}
      {activeTab === 'classement' && (
        <RecapTab klass={activeClass} onExport={() => buildExportWorkbook(activeClass)} />
      )}

      <BottomNav active={activeTab} onChange={setActiveTab} onMore={() => setMorePanelOpen(true)} />

      {classSwitcherOpen && (
        <ClassSwitcher
          classes={classes}
          activeId={activeId}
          onSelect={(id) => { setActiveId(id); setClassSwitcherOpen(false) }}
          onCreate={handleCreateClass}
          onDelete={handleDeleteClass}
          onClose={() => setClassSwitcherOpen(false)}
        />
      )}

      {reorderOpen && (
        <ReorderPanel subjects={activeClass.subjects} onReorder={handleReorderSubjects} onClose={() => setReorderOpen(false)} />
      )}

      {morePanelOpen && (
        <MorePanel
          onClose={() => setMorePanelOpen(false)}
          onSwitchClass={() => { setMorePanelOpen(false); setClassSwitcherOpen(true) }}
          onReorder={() => { setMorePanelOpen(false); setReorderOpen(true) }}
          onStats={() => { setMorePanelOpen(false); setStatsOpen(true) }}
          dark={dark}
          onToggleDark={() => { setDark(!dark); localStorage.setItem(LS_DARK, !dark ? '1' : '0') }}
          pin={pin}
          onSetPin={handleSetPin}
          onRemovePin={handleRemovePin}
          onUndo={handleUndo}
          canUndo={undoStack.length > 0}
          onDeleteClass={() => { handleDeleteClass(activeClass.id); setMorePanelOpen(false) }}
          defaultMode={defaultMode}
          onSetDefaultMode={handleSetDefaultMode}
          onGoWelcome={handleGoWelcome}
        />
      )}
    </Screen>
  )
}
