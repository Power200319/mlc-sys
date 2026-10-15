import type { Result } from '../../../types'
import { Card, CardHeader, CardBody } from '../../ui'
import './ResultSection.css'

export interface ResultSectionProps {
  result: Result
}

export const ResultSection = ({ result }: ResultSectionProps) => {
  const scoringOutput = result.outputs.scoring
  const reportOutput = result.outputs.report_cards

  return (
    <section className="result-section" aria-live="polite">
      <Card variant="elevated" className="result-card success">
        <CardHeader
          icon={
            <span className="card-header-icon success" aria-hidden="true">
              <svg width="20" height="20" viewBox="0 0 20 20" fill="none" stroke="currentColor" strokeWidth="2.5">
                <path d="M16 4L8 12L4 8" />
              </svg>
            </span>
          }
          title={`${result.print_label || result.period} Complete`}
          subtitle={
            <p className="result-summary">
              {result.subjects.length} subject{result.subjects.length !== 1 ? 's' : ''} imported ·{' '}
              {result.student_count} student{result.student_count !== 1 ? 's' : ''} processed
            </p>
          }
        />

        <CardBody className="result-subjects">
          <h4>Imported Subjects</h4>
          <ul className="subject-chips">
            {result.subjects.map((subject) => (
              <li key={subject.name} className="subject-chip">
                <span>{subject.name}</span>
                {subject.source_period && (
                  <span className="subject-period">{subject.source_period}</span>
                )}
              </li>
            ))}
          </ul>
        </CardBody>

        <CardBody className="result-downloads">
          {scoringOutput.available && scoringOutput.download && (
            <a
              href={scoringOutput.download}
              className="btn btn-secondary btn-lg download-btn"
              download
            >
              <span className="btn-icon" aria-hidden="true">
                <svg width="20" height="20" viewBox="0 0 20 20" fill="none" stroke="currentColor" strokeWidth="1.5">
                  <path d="M14 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h8a2 2 0 0 0 2-2V8L14 3z" />
                  <path d="M14 3V8H19" />
                </svg>
              </span>
              <div className="download-btn-content">
                <span>Download Scoring Workbook</span>
                <span className="download-btn-meta">{scoringOutput.filename}</span>
              </div>
            </a>
          )}

          {reportOutput.available && reportOutput.download && (
            <a
              href={reportOutput.download}
              className="btn btn-primary btn-lg download-btn"
              download
            >
              <span className="btn-icon" aria-hidden="true">
                <svg width="20" height="20" viewBox="0 0 20 20" fill="none" stroke="currentColor" strokeWidth="1.5">
                  <path d="M14 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h8a2 2 0 0 0 2-2V8L14 3z" />
                  <path d="M14 3V8H19" />
                </svg>
              </span>
              <div className="download-btn-content">
                <span>Download Student PDFs (ZIP)</span>
                <span className="download-btn-meta">{reportOutput.count} reports</span>
              </div>
            </a>
          )}

          {!reportOutput.available && (
            <div className="alert alert-warning">
              <span className="alert-icon" aria-hidden="true">
                <svg width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="2">
                  <circle cx="8" cy="8" r="7.5" />
                  <path d="M8 5V9M8 11H8.01" />
                </svg>
              </span>
              <span>Report Cards unavailable: {reportOutput.reason}</span>
            </div>
          )}
        </CardBody>

        {result.warnings.length > 0 && (
          <details className="warnings-section">
            <summary>Warnings & Notes ({result.warnings.length})</summary>
            <ul className="warnings-list">
              {result.warnings.map((warning, i) => (
                <li key={i}>{warning}</li>
              ))}
            </ul>
          </details>
        )}
      </Card>
    </section>
  )
}