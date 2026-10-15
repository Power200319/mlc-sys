import type { InputHTMLAttributes } from 'react'
import './Input.css'

export interface InputProps extends InputHTMLAttributes<HTMLInputElement> {
  label?: string
  error?: string
  helperText?: string
  leftIcon?: React.ReactNode
  rightIcon?: React.ReactNode
  fullWidth?: boolean
}

export const Input = ({
  label,
  error,
  helperText,
  leftIcon,
  rightIcon,
  fullWidth = false,
  className = '',
  id,
  ...props
}: InputProps) => {
  const inputId = id || `input-${Math.random().toString(36).slice(2, 9)}`
  const errorId = error ? `${inputId}-error` : undefined
  const helperId = helperText ? `${inputId}-helper` : undefined

  const classNames = ['input-wrapper', fullWidth ? 'input-full' : '', error ? 'input-has-error' : '', className]
    .filter(Boolean)
    .join(' ')

  const inputClassNames = ['input', error ? 'input-error' : ''].filter(Boolean).join(' ')

  const describedBy = [errorId, helperId].filter(Boolean).join(' ') || undefined

  return (
    <div className={classNames}>
      {label && <label htmlFor={inputId} className="input-label">{label}</label>}
      <div className="input-field-wrapper">
        {leftIcon && <span className="input-icon input-icon-left" aria-hidden="true">{leftIcon}</span>}
        <input
          id={inputId}
          className={inputClassNames}
          aria-invalid={!!error}
          aria-describedby={describedBy}
          {...props}
        />
        {rightIcon && <span className="input-icon input-icon-right" aria-hidden="true">{rightIcon}</span>}
      </div>
      {error && <p id={errorId} className="input-error" role="alert">{error}</p>}
      {helperText && !error && <p id={helperId} className="input-helper">{helperText}</p>}
    </div>
  )
}