import { useCallback, type ChangeEvent } from 'react'
import type { GradeOption } from '../../../types'
import './GradeSelector.css'

export interface GradeSelectorProps {
  selectedGrade: number
  onChange: (grade: number) => void
  grades: GradeOption[]
  disabled?: boolean
}

export const GradeSelector = ({
  selectedGrade,
  onChange,
  grades,
  disabled = false,
}: GradeSelectorProps) => {
  const handleChange = useCallback((event: ChangeEvent<HTMLSelectElement>) => {
    const value = parseInt(event.target.value, 10)
    if (!isNaN(value)) {
      onChange(value)
    }
  }, [onChange])

  if (grades.length === 0) {
    return (
      <div className="grade-selector grade-selector-loading">
        Loading grades…
      </div>
    )
  }

  return (
    <div className="grade-selector">
      <label className="grade-label" htmlFor="grade-select">Grade</label>
      <select
        id="grade-select"
        value={String(selectedGrade)}
        onChange={handleChange}
        disabled={disabled}
        className="grade-select"
      >
        {grades.map((g) => (
          <option key={g.value} value={g.value}>
            Grade {g.value}
          </option>
        ))}
      </select>
    </div>
  )
}
