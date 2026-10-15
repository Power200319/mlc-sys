import './Alert.css'

export interface AlertProps {
  type: 'error' | 'warning' | 'success' | 'info'
  message: string
  onDismiss?: () => void
}

export const Alert = ({ type, message, onDismiss }: AlertProps) => {
  const icons = {
    error: (
      <svg width="20" height="20" viewBox="0 0 20 20" fill="none" stroke="currentColor" strokeWidth="2" aria-hidden="true">
        <circle cx="10" cy="10" r="9.5" />
        <path d="M10 6V11M10 13H10.01" />
      </svg>
    ),
    warning: (
      <svg width="20" height="20" viewBox="0 0 20 20" fill="none" stroke="currentColor" strokeWidth="2" aria-hidden="true">
        <path d="M10 2L18 17H2L10 2Z" />
        <path d="M10 7V11M10 13H10.01" />
      </svg>
    ),
    success: (
      <svg width="20" height="20" viewBox="0 0 20 20" fill="none" stroke="currentColor" strokeWidth="2.5" aria-hidden="true">
        <path d="M16 5L8 13L4 9" />
      </svg>
    ),
    info: (
      <svg width="20" height="20" viewBox="0 0 20 20" fill="none" stroke="currentColor" strokeWidth="2" aria-hidden="true">
        <circle cx="10" cy="10" r="9.5" />
        <path d="M10 6V11M10 13H10.01" />
      </svg>
    ),
  }

  return (
    <div className={`alert alert-${type}`} role="alert">
      <span className="alert-icon" aria-hidden="true">{icons[type]}</span>
      <span>{message}</span>
      {onDismiss && (
        <button className="alert-dismiss" onClick={onDismiss} aria-label="Dismiss">
          <svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="2">
            <path d="M5 5L11 11M11 5L5 11" />
          </svg>
        </button>
      )}
    </div>
  )
}