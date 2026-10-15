import type { ReactNode } from 'react'
import { printPackages } from '../../../utils/helpers'
import './PackageSelector.css'

export interface PackageSelectorProps {
  selectedValue: string
  onChange: (value: string) => void
  disabled?: boolean
  children?: ReactNode
}

export const PackageSelector = ({
  selectedValue,
  onChange,
  disabled = false,
}: PackageSelectorProps) => {
  return (
    <nav className="package-selector" aria-label="Print package selection">
      {printPackages.map((pkg) => (
        <button
          key={pkg.value}
          className={`package-card ${selectedValue === pkg.value ? 'active' : ''}`}
          onClick={() => onChange(pkg.value)}
          disabled={disabled}
          aria-pressed={selectedValue === pkg.value}
        >
          <span className="package-icon" aria-hidden="true">{pkg.icon}</span>
          <div className="package-info">
            <span className="package-label">{pkg.label}</span>
            <span className="package-subtitle">{pkg.subtitle}</span>
          </div>
          <span className="package-indicator" aria-hidden="true" />
        </button>
      ))}
    </nav>
  )
}