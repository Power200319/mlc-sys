import { useState, useCallback } from 'react'
import type { Result } from '../../../types'
import { printPackages } from '../../../utils/helpers'
import { useFileSelection, useProcessing } from '../../../hooks'
import { Header } from '../Header'
import { PackageSelector } from '../PackageSelector'
import { ScoringWorkbookUpload } from '../ScoringWorkbookUpload'
import { FileDropZone } from '../FileDropZone'
import { ActionBar } from '../ActionBar'
import { ResultSection } from '../ResultSection'
import { Footer } from '../Footer'
import { Alert } from '../../ui/Alert'
import './App.css'

const App = () => {
  const [printStage, setPrintStage] = useState('P1')
  const [scoringWorkbook, setScoringWorkbook] = useState<File | null>(null)

  const { files, addFiles, removeFile, clearAllFiles } = useFileSelection()
  const { processing, progress, error, result, startProcessing, updateProgress, finishProcessing, setError } =
    useProcessing()

  const selectedPackage = printPackages.find((p) => p.value === printStage) || printPackages[0]

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
    if (scoringWorkbook) data.append('scoring_workbook', scoringWorkbook)

    try {
      updateProgress('Uploading & validating…', 30)
      const response = await fetch('/api/process/', { method: 'POST', body: data })

      updateProgress('Processing gradebooks…', 60)
      const payload = await response.json()

      updateProgress('Finalizing…', 90)
      if (!response.ok) throw new Error(payload.error || 'Import failed. Please try again.')

      finishProcessing(true, payload as Result)
    } catch (requestError) {
      finishProcessing(false, requestError instanceof Error ? requestError : new Error('An unexpected error occurred.'))
    }
  }

  return (
    <div className="app">
      <Header />

      <main className="app-main">
        <div className="layout">
          <aside className="sidebar">
            <PackageSelector selectedValue={printStage} onChange={setPrintStage} disabled={processing} />

            <div className="sidebar-footer">
              <p className="sidebar-hint">Select a print package to begin</p>
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

            {result && <ResultSection result={result} />}
          </section>
        </div>
      </main>

      <Footer />
    </div>
  )
}

export default App