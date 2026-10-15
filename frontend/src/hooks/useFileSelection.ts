import { useState, useCallback } from 'react'

export const useFileSelection = () => {
  const [files, setFiles] = useState<File[]>([])

  const addFiles = useCallback((incoming?: FileList | File[]) => {
    if (!incoming) return
    const candidates = Array.from(incoming)
    const invalid = candidates.find((file) => !file.name.toLowerCase().endsWith('.xlsx'))
    if (invalid) {
      return { error: `"${invalid.name}" is not a valid .xlsx file. Please select E-gradebook files only.` }
    }
    setFiles((current: File[]) => {
      const combined: File[] = [...current, ...candidates]
      return combined.filter((file: File, index: number) =>
        combined.findIndex((c: File) => c.name === file.name && c.size === file.size) === index
      )
    })
    return { error: null }
  }, [])

  const removeFile = useCallback((file: File) => {
    setFiles((current: File[]) => current.filter((candidate: File) => candidate !== file))
  }, [])

  const clearAllFiles = useCallback(() => {
    setFiles([])
  }, [])

  return {
    files,
    addFiles,
    removeFile,
    clearAllFiles,
    setFiles,
  }
}