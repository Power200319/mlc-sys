export const printPackages = [
  {
    value: 'P1',
    period: 'Q1',
    label: 'Print 1',
    subtitle: 'Q1 & Teacher Comments',
    description: 'Upload Q1 subject E-gradebooks. Creates scoring workbook with Q1 data and student PDF reports.',
    requiresScoring: false,
    icon: '📝',
  },
  {
    value: 'P2',
    period: 'Q2',
    label: 'Print 2',
    subtitle: 'Q2, Exam 1 & Final SEM1',
    description: 'Upload previous scoring workbook + Q2 subjects. Enter Exam 1 manually in Excel after download.',
    requiresScoring: true,
    icon: '📊',
  },
  {
    value: 'P3',
    period: 'Q3',
    label: 'Print 3',
    subtitle: 'Q3 & Teacher Comments',
    description: 'Upload continued scoring workbook + Q3 subjects. Preserves Q1, Q2, Exam 1. Creates PDF reports.',
    requiresScoring: true,
    icon: '📝',
  },
  {
    value: 'P4',
    period: 'Q4',
    label: 'Print 4',
    subtitle: 'Q4, Exam 2 & Final SEM2',
    description: 'Upload continued scoring workbook + Q4 subjects. Enter Exam 2 manually. Creates Final Semester 2 report.',
    requiresScoring: true,
    icon: '📊',
  },
] as const

export type PrintPackage = typeof printPackages[number]

export const uniqueFiles = (incoming: FileList | File[]): File[] => {
  const files = Array.from(incoming)
  return files.filter((file, index) =>
    files.findIndex((candidate) => candidate.name === file.name && candidate.size === file.size) === index
  )
}

export const formatFileSize = (bytes: number): string => {
  if (bytes < 1024) return `${bytes} B`
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`
}