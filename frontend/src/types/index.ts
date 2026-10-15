export interface Output {
  available: boolean
  download?: string
  reason?: string
  filename?: string
  count?: number
}

export interface Subject {
  name: string
  source_subject: string
  source_period: string | null
}

export interface Result {
  period: string
  print_label?: string | null
  student_count: number
  populated_students: number
  subjects: Subject[]
  warnings: string[]
  outputs: {
    scoring: Output
    report_cards: Output
  }
}

export interface PrintPackage {
  value: string
  period: string
  label: string
  subtitle: string
  description: string
  requiresScoring: boolean
  icon: string
}