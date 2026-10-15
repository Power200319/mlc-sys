import { useState, useCallback } from 'react'
import type { Result } from '../types'

interface ProgressStage {
  stage: string
  percent: number
}

export const useProcessing = () => {
  const [processing, setProcessing] = useState(false)
  const [progress, setProgress] = useState<ProgressStage | null>(null)
  const [error, setError] = useState('')
  const [result, setResult] = useState<Result | null>(null)

  const startProcessing = useCallback(() => {
    setProcessing(true)
    setProgress({ stage: 'Preparing files…', percent: 10 })
    setError('')
    setResult(null)
  }, [])

  const updateProgress = useCallback((stage: string, percent: number) => {
    setProgress({ stage, percent })
  }, [])

  const finishProcessing = useCallback((success: boolean, data?: Result | Error) => {
    setProcessing(false)
    if (success && data) {
      setResult(data as Result)
      setProgress({ stage: 'Complete', percent: 100 })
      setTimeout(() => setProgress(null), 1000)
    } else if (!success && data) {
      setError(data instanceof Error ? data.message : 'An unexpected error occurred.')
      setProgress(null)
    }
  }, [])

  const reset = useCallback(() => {
    setProcessing(false)
    setProgress(null)
    setError('')
    setResult(null)
  }, [])

  return {
    processing,
    progress,
    error,
    result,
    setError,
    setResult,
    startProcessing,
    updateProgress,
    finishProcessing,
    reset,
  }
}