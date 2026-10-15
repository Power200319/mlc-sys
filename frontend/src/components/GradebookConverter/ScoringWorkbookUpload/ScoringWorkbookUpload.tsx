import { useState, useCallback, type ChangeEvent } from 'react'
import { formatFileSize } from '../../../utils/helpers'
import { Card, CardHeader, CardBody, Button, Input } from '../../ui'
import './ScoringWorkbookUpload.css'

export interface ScoringWorkbookUploadProps {
  file: File | null
  onChange: (file: File | null) => void
  disabled?: boolean
  required?: boolean
}

export const ScoringWorkbookUpload = ({
  file,
  onChange,
  disabled = false,
  required = false,
}: ScoringWorkbookUploadProps) => {
  const [error, setError] = useState('')

  const handleChange = useCallback((event: ChangeEvent<HTMLInputElement>) => {
    const selected = event.target.files?.[0] || null
    if (selected && !selected.name.toLowerCase().endsWith('.xlsm')) {
      setError('Please select a valid .xlsm scoring workbook.')
      return
    }
    setError('')
    onChange(selected)
  }, [onChange])

  const handleRemove = useCallback(() => {
    setError('')
    onChange(null)
  }, [onChange])

  return (
    <Card variant="outlined">
      <CardHeader
        icon={<span aria-hidden="true">📁</span>}
        title="Continued Scoring Workbook"
        subtitle={required ? 'Required for Print 2, 3 & 4' : 'Optional for Print 1'}
      >
        {required && <span className="required-badge">Required</span>}
      </CardHeader>

      <CardBody>
        <p className="card-description">
          Upload the <strong>.xlsm scoring workbook</strong> downloaded from the previous print package.
        </p>

        <div className="file-input-wrapper">
          <Input
            type="file"
            id="scoring-workbook"
            accept=".xlsm,application/vnd.ms-excel.sheet.macroEnabled.12"
            onChange={handleChange}
            disabled={disabled}
            error={error}
            helperText={!error ? 'Choose Scoring Workbook (.xlsm)' : undefined}
            fullWidth
            className="file-input-hidden"
          />
          <label htmlFor="scoring-workbook" className="file-input-label">
            <span className="file-input-icon" aria-hidden="true">⬆</span>
            <span>Choose Scoring Workbook (.xlsm)</span>
          </label>
        </div>

        {file && (
          <div className="selected-file">
            <span className="selected-file-icon" aria-hidden="true">✓</span>
            <span className="selected-file-name">{file.name}</span>
            <span className="selected-file-size">{formatFileSize(file.size)}</span>
            <Button
              variant="ghost"
              size="sm"
              onClick={handleRemove}
              disabled={disabled}
              aria-label="Remove scoring workbook"
            >
              ✕
            </Button>
          </div>
        )}
      </CardBody>
    </Card>
  )
}