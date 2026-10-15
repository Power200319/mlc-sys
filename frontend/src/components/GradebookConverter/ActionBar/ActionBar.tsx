import { Button } from '../../ui'
import './ActionBar.css'

export interface ActionBarProps {
  onProcess: () => void
  processing: boolean
  progress: { stage: string; percent: number } | null
  disabled: boolean
  label: string
}

export const ActionBar = ({
  onProcess,
  processing,
  progress,
  disabled,
  label,
}: ActionBarProps) => {
  return (
    <div className="action-bar">
      <Button
        variant="primary"
        size="lg"
        fullWidth
        onClick={onProcess}
        disabled={disabled || processing}
        isLoading={processing}
      >
        {processing ? 'Processing…' : `Create ${label} Output`}
      </Button>

      {processing && progress && (
        <div className="progress-indicator">
          <div className="progress-bar">
            <div className="progress-fill" style={{ width: `${progress.percent}%` }} />
          </div>
          <span className="progress-text">{progress.stage}</span>
        </div>
      )}
    </div>
  )
}