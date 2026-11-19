import { useEffect, useState, type ChangeEvent } from 'react'
import { Header } from '../GradebookConverter/Header'
import { Footer } from '../GradebookConverter/Footer'
import { Alert, Button } from '../ui'
import './EarlyYearsApp.css'

type Profile = { value: string; label: string; subjects: string[] }
type EarlyYearsResult = {
  grade: string; term: string; student_count: number; fields_transferred: number; subjects: string[]; warnings: string[]
  output: { available: boolean; filename: string; download: string }
}

export const EarlyYearsApp = ({ onSwitch }: { onSwitch: () => void }) => {
  const [profiles, setProfiles] = useState<Profile[]>([])
  const [grade, setGrade] = useState('K1')
  const [term, setTerm] = useState('T1')
  const [files, setFiles] = useState<File[]>([])
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState('')
  const [result, setResult] = useState<EarlyYearsResult | null>(null)
  const apiBase = import.meta.env.VITE_API_BASE_URL || ''

  useEffect(() => {
    fetch(`${apiBase}/api/early-years/profiles/`).then(async response => {
      if (!response.ok) throw new Error()
      const data = await response.json()
      setProfiles(data.profiles || [])
    }).catch(() => setError('Early Years service is not available. Start the Django backend and reload.'))
  }, [apiBase])

  const chooseFile = (event: ChangeEvent<HTMLInputElement>) => {
    const selected = Array.from(event.target.files || [])
    const invalid = selected.find(file => !file.name.toLowerCase().endsWith('.xls'))
    if (invalid) {
      setError(`'${invalid.name}' is not a legacy .xls E-Progress workbook.`); return
    }
    const unique = new Map(files.map(file => [`${file.name}-${file.size}`, file]))
    selected.forEach(file => unique.set(`${file.name}-${file.size}`, file))
    setFiles(Array.from(unique.values())); setError(''); setResult(null)
    event.target.value = ''
  }

  const removeFile = (file: File) => setFiles(current => current.filter(item => item !== file))

  const submit = async () => {
    if (!files.length) { setError('Choose one or more .xls E-Progress workbooks.'); return }
    setBusy(true); setError(''); setResult(null)
    const data = new FormData()
    files.forEach(file => data.append('e_progress', file))
    data.append('grade', grade); data.append('term', term)
    try {
      const response = await fetch(`${apiBase}/api/early-years/process/`, { method: 'POST', body: data })
      const payload = await response.json()
      if (!response.ok) throw new Error(payload.error || 'Early Years import failed.')
      setResult(payload as EarlyYearsResult)
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : 'Early Years import failed.')
    } finally { setBusy(false) }
  }

  const selected = profiles.find(item => item.value === grade)
  return <div className="app early-years-app">
    <Header><Button variant="secondary" size="sm" onClick={onSwitch}>G1–G9 Gradebook</Button></Header>
    <main className="app-main"><section className="early-years-panel">
      <div className="ey-heading"><span>Early Years • separate workflow</span><h2>K1–K3 E‑Progress Import</h2><p>This does not use or change E‑Gradebook files.</p></div>
      <div className="ey-controls">
        <label>Grade<select value={grade} disabled={busy} onChange={event => setGrade(event.target.value)}>{profiles.length ? profiles.map(item => <option key={item.value} value={item.value}>{item.label}</option>) : <option>K1</option>}</select></label>
        <label>Term<select value={term} disabled={busy} onChange={event => setTerm(event.target.value)}>{['T1', 'T2', 'T3', 'T4'].map(value => <option key={value} value={value}>Term {value.slice(1)}</option>)}</select></label>
      </div>
      <div className="ey-subjects"><strong>Expected subject sheets</strong><span>{selected?.subjects.join(' • ') || 'Loading profiles…'}</span></div>
      <label className="ey-upload"><input type="file" multiple accept=".xls,application/vnd.ms-excel" onChange={chooseFile} disabled={busy}/><strong>Choose E‑Progress .xls files</strong><small>Add separate English, Khmer, Math, Science, or CP files for the same class and term</small></label>
      {files.length > 0 && <ul className="ey-file-list">{files.map(file => <li key={`${file.name}-${file.size}`}><span>{file.name}</span><button type="button" onClick={() => removeFile(file)} disabled={busy} aria-label={`Remove ${file.name}`}>×</button></li>)}</ul>}
      {error && <Alert type="error" message={error}/>} 
      <Button onClick={submit} disabled={!files.length || busy} className="ey-action">{busy ? 'Converting and importing…' : `Create ${term.replace('T', 'Term ')} Score System`}</Button>
      {result && <div className="ey-result"><h3>Import complete</h3><p>{result.student_count} students · {result.fields_transferred} verified outcome values transferred</p><a className="ey-download" href={`${apiBase}${result.output.download}`}>Download {result.output.filename}</a><h4>Notes</h4><ul>{result.warnings.map((note, index) => <li key={index}>{note}</li>)}</ul></div>}
    </section></main>
    <Footer />
  </div>
}
