import { useState, useCallback, useEffect } from 'react'
import type { Result, GradeOption, Output } from '../../../types'
import { printPackages } from '../../../utils/helpers'
import { useFileSelection, useProcessing } from '../../../hooks'
import { Header } from '../Header'
import { PackageSelector } from '../PackageSelector'
import { GradeSelector } from '../GradeSelector'
import { ScoringWorkbookUpload } from '../ScoringWorkbookUpload'
import { FileDropZone } from '../FileDropZone'
import { ActionBar } from '../ActionBar'
import { ResultSection } from '../ResultSection'
import { Footer } from '../Footer'
import { Alert } from '../../ui/Alert'
import { EarlyYearsApp } from '../../EarlyYears'
import './App.css'

const App = () => {
  const [workflow, setWorkflow] = useState<'gradebook' | 'early-years'>('gradebook')
  const [printStage, setPrintStage] = useState('P1')
  const [grade, setGrade] = useState(6)
  const [grades, setGrades] = useState<GradeOption[]>([])
  const [scoringWorkbook, setScoringWorkbook] = useState<File | null>(null)
  const [generatingReports, setGeneratingReports] = useState(false)

  const { files, addFiles, removeFile, clearAllFiles } = useFileSelection()
  const { processing, progress, error, result, startProcessing, updateProgress, finishProcessing, setError } =
    useProcessing()

  const selectedPackage = printPackages.find((p) => p.value === printStage) || printPackages[0]

  useEffect(() => {
    const fetchGrades = async () => {
      try {
        const apiBase = import.meta.env.VITE_API_BASE_URL || ''
        const response = await fetch(`${apiBase}/api/grades/`)
        if (response.ok) {
          const data = await response.json()
          setGrades(data.grades || [])
        }
      } catch {
        // Silently fail; grades will be empty and selector shows loading
      }
    }
    fetchGrades()
  }, [])

  const handleAddFiles = useCallback(
    (incoming?: FileList | File[]) => {
      const result = addFiles(incoming)
      if (result?.error) {
        setError(result.error)
      }
    },
    [addFiles, setError]
  )

  const handleScoringChange = useCallback((file: File | null) => {
    setScoringWorkbook(file)
    setError('')
  }, [setError])

  const processGradebooks = async () => {
    if (!files.length) {
      setError('Please add at least one E-gradebook file before importing.')
      return
    }
    if (selectedPackage.requiresScoring && !scoringWorkbook) {
      setError(`Print ${printStage.slice(1)} requires a continued scoring workbook. Please upload the .xlsm file from the previous print.`)
      return
    }

    startProcessing()

    const data = new FormData()
    files.forEach((file) => data.append('gradebooks', file))
    data.append('period', selectedPackage.period)
    data.append('print_stage', selectedPackage.value)
    data.append('grade', String(grade))
    if (scoringWorkbook) data.append('scoring_workbook', scoringWorkbook)

    let processingNotice: number | undefined
    try {
      updateProgress('Uploading & validating…', 30)
      const apiBase = import.meta.env.VITE_API_BASE_URL || ''
      processingNotice = window.setTimeout(() => {
        updateProgress('Writing scoring workbook…', 70)
      }, 1200)
      const response = await fetch(`${apiBase}/api/process/`, { method: 'POST', body: data })
      window.clearTimeout(processingNotice)

      updateProgress('Processing gradebooks…', 60)
      const payload = await response.json()

      updateProgress('Finalizing…', 90)
      if (!response.ok) throw new Error(payload.error || 'Import failed. Please try again.')

      finishProcessing(true, payload as Result)
    } catch (requestError) {
      finishProcessing(false, requestError instanceof Error ? requestError : new Error('An unexpected error occurred.'))
    } finally {
      if (processingNotice !== undefined) window.clearTimeout(processingNotice)
    }
  }

  const generateReports = async () => {
    if (!result) return
    setGeneratingReports(true)
    setError('')
    const data = new FormData()
    data.append('job_id', result.job_id)
    data.append('grade', String(result.grade))
    try {
      const apiBase = import.meta.env.VITE_API_BASE_URL || ''
      const response = await fetch(`${apiBase}/api/generate-reports/`, { method: 'POST', body: data })
      const payload = await response.json()
      if (!response.ok) throw new Error(payload.error || 'Student card generation failed.')
      finishProcessing(true, {
        ...result,
        warnings: [...result.warnings, ...(payload.warnings || [])],
        outputs: { ...result.outputs, report_cards: payload.report_cards as Output },
      })
    } catch (requestError) {
      setError(requestError instanceof Error ? requestError.message : 'Student card generation failed.')
    } finally {
      setGeneratingReports(false)
    }
  }

  if (workflow === 'early-years') return <EarlyYearsApp onSwitch={() => setWorkflow('gradebook')} />

  return (
    <div className="app">
      <Header><button className="workflow-switch" type="button" onClick={() => setWorkflow('early-years')}>Early Years K1–K3</button></Header>

      <main className="app-main">
        <div className="layout">
          <aside className="sidebar">
            <GradeSelector
              selectedGrade={grade}
              onChange={setGrade}
              grades={grades}
              disabled={processing}
            />
            <PackageSelector selectedValue={printStage} onChange={setPrintStage} disabled={processing} />

            <div className="sidebar-footer">
              <p className="sidebar-hint">Select grade and print package to begin</p>
            </div>
          </aside>

          <section className="content" aria-labelledby="content-title">
            <header className="content-header">
              <h2 id="content-title">
                {selectedPackage.label}: {selectedPackage.subtitle}
              </h2>
              <p className="content-description">{selectedPackage.description}</p>
            </header>

            {selectedPackage.requiresScoring && (
              <ScoringWorkbookUpload
                file={scoringWorkbook}
                onChange={handleScoringChange}
                disabled={processing}
                required
              />
            )}

            <FileDropZone
              files={files}
              onAddFiles={handleAddFiles}
              onRemoveFile={removeFile}
              onClearAll={clearAllFiles}
              disabled={processing}
            />

            {error && <Alert type="error" message={error} />}

            <ActionBar
              onProcess={processGradebooks}
              processing={processing}
              progress={progress}
              disabled={!files.length}
              label={selectedPackage.label}
            />

            {result && <ResultSection result={result} onGenerateReports={selectedPackage.value === 'P1' || selectedPackage.value === 'P3' ? generateReports : undefined} generatingReports={generatingReports} />}
          </section>
        </div>
      </main>

      <Footer />
    </div>
  )
}

export default App
