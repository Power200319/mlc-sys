import { useState, useCallback, type ChangeEvent, type DragEvent } from 'react'
import { formatFileSize } from '../../../utils/helpers'
import { Card, CardHeader, CardBody, Button } from '../../ui'
import './FileDropZone.css'

export interface FileDropZoneProps {
  files: File[]
  onAddFiles: (files: File[]) => void
  onRemoveFile: (file: File) => void
  onClearAll: () => void
  disabled?: boolean
}

export const FileDropZone = ({
  files,
  onAddFiles,
  onRemoveFile,
  onClearAll,
  disabled = false,
}: FileDropZoneProps) => {
  const [dragging, setDragging] = useState(false)

  const handleDragEnter = useCallback((e: DragEvent<HTMLLabelElement>) => {
    e.preventDefault()
    setDragging(true)
  }, [])

  const handleDragLeave = useCallback((e: DragEvent<HTMLLabelElement>) => {
    e.preventDefault()
    setDragging(false)
  }, [])

  const handleDragOver = useCallback((e: DragEvent<HTMLLabelElement>) => {
    e.preventDefault()
  }, [])

  const handleDrop = useCallback((e: DragEvent<HTMLLabelElement>) => {
    e.preventDefault()
    setDragging(false)
    onAddFiles(Array.from(e.dataTransfer.files))
  }, [onAddFiles])

  const handleFileSelect = useCallback((e: ChangeEvent<HTMLInputElement>) => {
    onAddFiles(Array.from(e.target.files || []))
  }, [onAddFiles])

  return (
    <Card variant="outlined">
      <CardHeader
        icon={<span aria-hidden="true">📤</span>}
        title="E-Gradebook Files"
        subtitle={`${files.length} file${files.length !== 1 ? 's' : ''} ready`}
      />

      <CardBody>
        <p className="card-description">
          Drop <strong>.xlsx subject files</strong> here. You can add multiple files before importing.
        </p>

        <label
          className={`drop-zone ${dragging ? 'dragging' : ''} ${files.length > 0 ? 'has-files' : ''}`}
          onDragEnter={handleDragEnter}
          onDragLeave={handleDragLeave}
          onDragOver={handleDragOver}
          onDrop={handleDrop}
        >
          <input
            type="file"
            multiple
            accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            onChange={handleFileSelect}
            disabled={disabled}
            className="drop-zone-input"
          />
          <div className="drop-zone-content">
            <div className="drop-zone-icon" aria-hidden="true">📁</div>
            <strong>Drop E-gradebook files here</strong>
            <span>or click to browse</span>
            <span className="drop-zone-hint">Supports: .xlsx files only</span>
          </div>
        </label>

        {files.length > 0 && <FileList files={files} onRemoveFile={onRemoveFile} onClearAll={onClearAll} disabled={disabled} />}
      </CardBody>
    </Card>
  )
}

const FileList = ({
  files,
  onRemoveFile,
  onClearAll,
  disabled,
}: {
  files: File[]
  onRemoveFile: (file: File) => void
  onClearAll: () => void
  disabled: boolean
}) => {
  return (
    <div className="file-list">
      <div className="file-list-header">
        <h4>Files Ready to Import ({files.length})</h4>
        <Button variant="ghost" size="sm" onClick={onClearAll} disabled={disabled}>
          Clear All
        </Button>
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
            <Button
              variant="ghost"
              size="sm"
              onClick={() => onRemoveFile(file)}
              disabled={disabled}
              aria-label={`Remove ${file.name}`}
            >
              ✕
            </Button>
          </li>
        ))}
      </ul>
    </div>
  )
}