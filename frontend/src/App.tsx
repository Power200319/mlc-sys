import { useState, useCallback } from 'react'
import type { ChangeEvent, DragEvent } from 'react'
import mlcLogo from './assets/mlc_logo-removebg-preview.png'
import './App.css'

type Output = { available: boolean; download?: string; reason?: string; filename?: string; count?: number }
type Subject = { name: string; source_subject: string; source_period: string | null }
type Result = { 
  period: string; 
  print_label?: string | null; 
  student_count: number; 
  populated_students: number; 
  subjects: Subject[]; 
  warnings: string[]; 
  outputs: { scoring: Output; report_cards: Output } 
}

const printPackages = [
  { 
    value: 'P1', 
    period: 'Q1', 
    label: 'Print 1', 
    subtitle: 'Q1 & Teacher Comments',
    description: 'Upload Q1 subject E-gradebooks. Creates scoring workbook with Q1 data and student PDF reports.',
    requiresScoring: false,
    icon: '📝'
  },
  { 
    value: 'P2', 
    period: 'Q2', 
    label: 'Print 2', 
    subtitle: 'Q2, Exam 1 & Final SEM1',
    description: 'Upload previous scoring workbook + Q2 subjects. Enter Exam 1 manually in Excel after download.',
    requiresScoring: true,
    icon: '📊'
  },
  { 
    value: 'P3', 
    period: 'Q3', 
    label: 'Print 3', 
    subtitle: 'Q3 & Teacher Comments',
    description: 'Upload continued scoring workbook + Q3 subjects. Preserves Q1, Q2, Exam 1. Creates PDF reports.',
    requiresScoring: true,
    icon: '📝'
  },
  { 
    value: 'P4', 
    period: 'Q4', 
    label: 'Print 4', 
    subtitle: 'Q4, Exam 2 & Final SEM2',
    description: 'Upload continued scoring workbook + Q4 subjects. Enter Exam 2 manually. Creates Final Semester 2 report.',
    requiresScoring: true,
    icon: '📊'
  },
]

const uniqueFiles = (incoming: FileList | File[]): File[] => {
  const files = Array.from(incoming)
  return files.filter((file, index) => 
    files.findIndex((candidate) => candidate.name === file.name && candidate.size === file.size) === index
  )
}

const formatFileSize = (bytes: number): string => {
  if (bytes < 1024) return `${bytes} B`
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`
}

const App = () => {
  const [files, setFiles] = useState<File[]>([])
  const [printStage, setPrintStage] = useState('P1')
  const [scoringWorkbook, setScoringWorkbook] = useState<File | null>(null)
  const [dragging, setDragging] = useState(false)
  const [processing, setProcessing] = useState(false)
  const [error, setError] = useState('')
  const [result, setResult] = useState<Result | null>(null)
  const [progress, setProgress] = useState<{ stage: string; percent: number } | null>(null)

  const selectedPackage = printPackages.find((p) => p.value === printStage) || printPackages[0]

  const addFiles = useCallback((incoming?: FileList | File[]) => {
    setError('')
    setResult(null)
    if (!incoming) return
    const candidates = uniqueFiles(incoming)
    const invalid = candidates.find((file) => !file.name.toLowerCase().endsWith('.xlsx'))
    if (invalid) {
      setError(`"${invalid.name}" is not a valid .xlsx file. Please select E-gradebook files only.`)
      return
    }
    setFiles((current) => uniqueFiles([...current, ...candidates]))
  }, [])

  const removeFile = useCallback((file: File) => {
    setFiles((current) => current.filter((candidate) => candidate !== file))
  }, [])

  const clearAllFiles = useCallback(() => {
    setFiles([])
  }, [])

  const handleScoringWorkbookChange = useCallback((event: ChangeEvent<HTMLInputElement>) => {
    const selected = event.target.files?.[0] || null
    if (selected && !selected.name.toLowerCase().endsWith('.xlsm')) {
      setError('Please select a valid .xlsm scoring workbook.')
      return
    }
    setError('')
    setResult(null)
    setScoringWorkbook(selected)
  }, [])

  const processGradebooks = async () => {
    if (!files.length) {
      setError('Please add at least one E-gradebook file before importing.')
      return
    }
    if (selectedPackage.requiresScoring && !scoringWorkbook) {
      setError(`Print ${printStage.slice(1)} requires a continued scoring workbook. Please upload the .xlsm file from the previous print.`)
      return
    }

    setProcessing(true)
    setProgress({ stage: 'Preparing files…', percent: 10 })
    setError('')
    setResult(null)

    const data = new FormData()
    files.forEach((file) => data.append('gradebooks', file))
    data.append('period', selectedPackage.period)
    data.append('print_stage', selectedPackage.value)
    if (scoringWorkbook) data.append('scoring_workbook', scoringWorkbook)

    try {
      setProgress({ stage: 'Uploading & validating…', percent: 30 })
      const response = await fetch('/api/process/', { method: 'POST', body: data })
      
      setProgress({ stage: 'Processing gradebooks…', percent: 60 })
      const payload = await response.json()
      
      setProgress({ stage: 'Finalizing…', percent: 90 })
      if (!response.ok) throw new Error(payload.error || 'Import failed. Please try again.')
      
      setResult(payload as Result)
      setProgress({ stage: 'Complete', percent: 100 })
    } catch (requestError) {
      setError(requestError instanceof Error ? requestError.message : 'An unexpected error occurred.')
      setProgress(null)
    } finally {
      setProcessing(false)
      setTimeout(() => setProgress(null), 1000)
    }
  }

  return (
    <main className="app">
      <header className="app-header">
        <div className="header-content">
          <div className="brand">
            <img src={mlcLogo} alt="MLC International School" className="brand-logo" />
            <div className="brand-text">
              <h1>MLCIS Gradebook Converter</h1>
              <p>MLC International School</p>
            </div>
          </div>
          <div className="header-actions">
            <span className="version-badge">v2024-2025</span>
          </div>
        </div>
      </header>

      <div className="app-main">
        <div className="layout">
          <aside className="sidebar">
            <nav className="package-selector" aria-label="Print package selection">
              {printPackages.map((pkg) => (
                <button
                  key={pkg.value}
                  className={`package-card ${printStage === pkg.value ? 'active' : ''}`}
                  onClick={() => { setPrintStage(pkg.value); setError(''); setResult(null); }}
                  disabled={processing}
                >
                  <span className="package-icon" aria-hidden="true">{pkg.icon}</span>
                  <div className="package-info">
                    <span className="package-label">{pkg.label}</span>
                    <span className="package-subtitle">{pkg.subtitle}</span>
                  </div>
                  <span className="package-indicator" aria-hidden="true"></span>
                </button>
              ))}
            </nav>
            
            <div className="sidebar-footer">
              <p className="sidebar-hint">Select a print package to begin</p>
            </div>
          </aside>

          <section className="content" aria-labelledby="content-title">
            <header className="content-header">
              <h2 id="content-title">{selectedPackage.label}: {selectedPackage.subtitle}</h2>
              <p className="content-description">{selectedPackage.description}</p>
            </header>

            {selectedPackage.requiresScoring && (
              <div className="card scoring-workbook-card">
                <div className="card-header">
                  <span className="card-icon" aria-hidden="true">📁</span>
                  <h3>Continued Scoring Workbook</h3>
                </div>
                <p className="card-description">
                  Required for Print 2, 3 & 4. Upload the <strong>.xlsm scoring workbook</strong> downloaded from the previous print package.
                </p>
                <label className="file-input-wrapper">
                  <input
                    type="file"
                    accept=".xlsm,application/vnd.ms-excel.sheet.macroEnabled.12"
                    onChange={handleScoringWorkbookChange}
                    disabled={processing}
                  />
                  <span className="file-input-label">
                    <span className="file-input-icon" aria-hidden="true">⬆</span>
                    <span>Choose Scoring Workbook (.xlsm)</span>
                  </span>
                </label>
                {scoringWorkbook && (
                  <div className="selected-file">
                    <span className="selected-file-icon" aria-hidden="true">✓</span>
                    <span className="selected-file-name">{scoringWorkbook.name}</span>
                    <span className="selected-file-size">{formatFileSize(scoringWorkbook.size)}</span>
                    <button 
                      type="button" 
                      className="remove-file-btn" 
                      onClick={() => setScoringWorkbook(null)}
                      disabled={processing}
                      aria-label="Remove scoring workbook"
                    >✕</button>
                  </div>
                )}
              </div>
            )}

            <div className="card upload-card">
              <div className="card-header">
                <span className="card-icon" aria-hidden="true">📤</span>
                <h3>E-Gradebook Files</h3>
              </div>
              <p className="card-description">
                Drop <strong>.xlsx subject files</strong> here. You can add multiple files before importing.
                {files.length > 0 && <span> <strong>{files.length}</strong> file{files.length !== 1 ? 's' : ''} ready.</span>}
              </p>

              <label 
                className={`drop-zone ${dragging ? 'dragging' : ''} ${files.length > 0 ? 'has-files' : ''}`}
                onDragEnter={(e) => { e.preventDefault(); setDragging(true); }}
                onDragLeave={(e) => { e.preventDefault(); setDragging(false); }}
                onDragOver={(e) => { e.preventDefault(); }}
                onDrop={(e: DragEvent<HTMLLabelElement>) => { e.preventDefault(); setDragging(false); addFiles(e.dataTransfer.files); }}
              >
                <input
                  type="file"
                  multiple
                  accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                  onChange={(e: ChangeEvent<HTMLInputElement>) => addFiles(e.target.files || undefined)}
                  disabled={processing}
                />
                <div className="drop-zone-content">
                  <div className="drop-zone-icon" aria-hidden="true">📁</div>
                  <strong>Drop E-gradebook files here</strong>
                  <span>or click to browse</span>
                  <span className="drop-zone-hint">Supports: .xlsx files only</span>
                </div>
              </label>

              {files.length > 0 && (
                <div className="file-list">
                  <div className="file-list-header">
                    <h4>Files Ready to Import ({files.length})</h4>
                    <button type="button" className="clear-all-btn" onClick={clearAllFiles} disabled={processing}>
                      Clear All
                    </button>
                  </div>
                  <ul className="file-list-items">
                    {files.map((file) => (
                      <li key={`${file.name}-${file.size}`} className="file-item">
                        <div className="file-info">
                          <span className="file-icon" aria-hidden="true">📄</span>
                          <div className="file-details">
                            <span className="file-name">{file.name}</span>
                            <span className="file-size">{formatFileSize(file.size)}</span>
                          </div>
                        </div>
                        <button 
                          type="button" 
                          className="remove-file-btn" 
                          onClick={() => removeFile(file)}
                          disabled={processing}
                          aria-label={`Remove ${file.name}`}
                        >✕</button>
                      </li>
                    ))}
                  </ul>
                </div>
              )}
            </div>

            {error && (
              <div className="alert alert-error" role="alert">
                <span className="alert-icon" aria-hidden="true">⚠</span>
                <span>{error}</span>
              </div>
            )}

            <div className="action-bar">
              <button 
                className="btn btn-primary btn-lg" 
                onClick={processGradebooks} 
                disabled={!files.length || processing}
              >
                {processing ? (
                  <>
                    <span className="btn-spinner" aria-hidden="true"></span>
                    <span>Processing…</span>
                  </>
                ) : (
                  `Create ${selectedPackage.label} Output`
                )}
              </button>
              {processing && progress && (
                <div className="progress-indicator">
                  <div className="progress-bar">
                    <div className="progress-fill" style={{ width: `${progress.percent}%` }}></div>
                  </div>
                  <span className="progress-text">{progress.stage}</span>
                </div>
              )}
            </div>

            {result && (
              <section className="result-section" aria-live="polite">
                <div className="card result-card success">
                  <div className="card-header">
                    <span className="card-icon success" aria-hidden="true">✓</span>
                    <div>
                      <h3>{result.print_label || result.period} Complete</h3>
                      <p className="result-summary">
                        {result.subjects.length} subject{result.subjects.length !== 1 ? 's' : ''} imported · 
                        {result.student_count} student{result.student_count !== 1 ? 's' : ''} processed
                      </p>
                    </div>
                  </div>

                  <div className="result-subjects">
                    <h4>Imported Subjects</h4>
                    <ul className="subject-chips">
                      {result.subjects.map((subject) => (
                        <li key={subject.name} className="subject-chip">
                          <span>{subject.name}</span>
                          {subject.source_period && (
                            <span className="subject-period">{subject.source_period}</span>
                          )}
                        </li>
                      ))}
                    </ul>
                  </div>

                  <div className="result-downloads">
                    {result.outputs.scoring.available && result.outputs.scoring.download && (
                      <a 
                        href={result.outputs.scoring.download} 
                        className="btn btn-secondary btn-lg download-btn"
                        download
                      >
                        <span className="btn-icon" aria-hidden="true">📊</span>
                        <span>Download Scoring Workbook</span>
                        <span className="btn-meta">{result.outputs.scoring.filename}</span>
                      </a>
                    )}
                    {result.outputs.report_cards.available && result.outputs.report_cards.download && (
                      <a 
                        href={result.outputs.report_cards.download} 
                        className="btn btn-primary btn-lg download-btn"
                        download
                      >
                        <span className="btn-icon" aria-hidden="true">📄</span>
                        <span>Download Student PDFs (ZIP)</span>
                        <span className="btn-meta">{result.outputs.report_cards.count} reports</span>
                      </a>
                    )}
                    {!result.outputs.report_cards.available && (
                      <div className="alert alert-warning">
                        <span className="alert-icon" aria-hidden="true">⚠</span>
                        <span>Report Cards unavailable: {result.outputs.report_cards.reason}</span>
                      </div>
                    )}
                  </div>

                  {result.warnings.length > 0 && (
                    <details className="warnings-section">
                      <summary>Warnings & Notes ({result.warnings.length})</summary>
                      <ul className="warnings-list">
                        {result.warnings.map((warning, i) => (
                          <li key={i}>{warning}</li>
                        ))}
                      </ul>
                    </details>
                  )}
                </div>
              </section>
            )}
          </section>
        </div>
      </div>

      <footer className="app-footer">
        <p>MLC International School · Gradebook Converter System</p>
        <p className="footer-note">For privacy, downloads are available once. Uploaded files are deleted after processing.</p>
      </footer>
    </main>
  )
}

export default App