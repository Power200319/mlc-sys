import type { ReactNode } from 'react'
import mlcLogo from '../../../assets/mlc_logo-removebg-preview.png'
import './Header.css'

export interface HeaderProps {
  children?: ReactNode
}

export const Header = ({ children }: HeaderProps) => {
  return (
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
          {children}
        </div>
      </div>
    </header>
  )
}