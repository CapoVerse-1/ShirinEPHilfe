import { useState, useCallback } from 'react'
import { useDropzone } from 'react-dropzone'
import * as XLSX from 'xlsx'
import { saveAs } from 'file-saver'

interface PromotionData {
  dayOfWeek: string
  date: string
  startTime: string
  endTime: string
  totalHours: number
  pointOfSales: string
  district: string
  marketName: string
  coffeeAdvisor: string
}

export default function Home() {
  const [file, setFile] = useState<File | null>(null)
  const [processedData, setProcessedData] = useState<PromotionData[]>([])
  const [status, setStatus] = useState<{ type: 'success' | 'error' | '', message: string }>({ type: '', message: '' })

  const getDayOfWeek = (date: Date): string => {
    const days = ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa']
    return days[date.getDay()]
  }

  const getWorkingHours = (dayOfWeek: string): { start: string, end: string } => {
    if (dayOfWeek === 'Sa' || dayOfWeek === 'So') {
      return { start: '09:00', end: '18:00' }
    }
    return { start: '09:30', end: '18:30' }
  }

  const processExcelData = (data: any[][]) => {
    const result: PromotionData[] = []
    
    // Skip header row, start from row 1
    for (let rowIndex = 1; rowIndex < data.length; rowIndex++) {
      const row = data[rowIndex]
      if (!row || row.length < 5) continue
      
      const marketName = row[0] || ''
      const district = row[1] || ''
      
      // Process each date column (starting from column E = index 4)
      for (let colIndex = 4; colIndex < row.length && colIndex < 35; colIndex++) { // 31 days max
        const value = row[colIndex]
        if (!value || value === 0) continue
        
        // Calculate the date (July 2025)
        const day = colIndex - 3 // Column E = day 1, F = day 2, etc.
        const date = new Date(2025, 6, day) // July = month 6 (0-indexed)
        const dayOfWeek = getDayOfWeek(date)
        const workingHours = getWorkingHours(dayOfWeek)
        
        // Format date as DD.MM.YYYY
        const formattedDate = `${day.toString().padStart(2, '0')}.07.2025`
        
        // Clean market name (remove MM prefix)
        const cleanMarketName = marketName.replace(/^MM\s*/, '')
        
        // Determine hours based on value
        let totalHours = 8
        if (value === 0.75) {
          totalHours = 6
        }
        
        // If value is 2, create two separate entries
        const numPromotions = value === 2 ? 2 : 1
        
        for (let i = 0; i < numPromotions; i++) {
          result.push({
            dayOfWeek,
            date: formattedDate,
            startTime: workingHours.start,
            endTime: workingHours.end,
            totalHours,
            pointOfSales: marketName,
            district,
            marketName: cleanMarketName,
            coffeeAdvisor: ''
          })
        }
      }
    }
    
    return result
  }

  const onDrop = useCallback((acceptedFiles: File[]) => {
    const file = acceptedFiles[0]
    if (!file) return
    
    setFile(file)
    setStatus({ type: '', message: '' })
    
    const reader = new FileReader()
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer)
        const workbook = XLSX.read(data, { type: 'array' })
        const sheetName = workbook.SheetNames[0]
        const worksheet = workbook.Sheets[sheetName]
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 })
        
        const processed = processExcelData(jsonData as any[][])
        setProcessedData(processed)
        setStatus({ 
          type: 'success', 
          message: `Successfully processed ${processed.length} promotion entries from ${file.name}` 
        })
      } catch (error) {
        setStatus({ 
          type: 'error', 
          message: `Error processing file: ${error instanceof Error ? error.message : 'Unknown error'}` 
        })
      }
    }
    reader.readAsArrayBuffer(file)
  }, [])

  const { getRootProps, getInputProps, isDragActive } = useDropzone({
    onDrop,
    accept: {
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
      'application/vnd.ms-excel': ['.xls']
    },
    multiple: false
  })

  const exportToExcel = () => {
    if (processedData.length === 0) return
    
    const headers = [
      'Tag der Woche',
      'Datum',
      'Startzeit',
      'Endzeit',
      'Gesamtstunden',
      'Point of Sales',
      'Bezirk',
      'Marktname',
      'Coffee Advisor'
    ]
    
    const exportData = [
      headers,
      ...processedData.map(item => [
        item.dayOfWeek,
        item.date,
        item.startTime,
        item.endTime,
        item.totalHours,
        item.pointOfSales,
        item.district,
        item.marketName,
        item.coffeeAdvisor
      ])
    ]
    
    const worksheet = XLSX.utils.aoa_to_sheet(exportData)
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Promotions')
    
    const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' })
    const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
    
    const fileName = file ? `processed_${file.name}` : 'promotion_schedule.xlsx'
    saveAs(blob, fileName)
  }

  return (
    <div className="container">
      <h1 className="title">Promotion Planner</h1>
      <p className="subtitle">
        Upload your Excel file to convert promotion schedules into a workable format
      </p>
      
      <div className="card">
        <div
          {...getRootProps()}
          className={`dropzone ${isDragActive ? 'active' : ''}`}
        >
          <input {...getInputProps()} />
          {isDragActive ? (
            <p>Drop the Excel file here...</p>
          ) : (
            <div>
              <p>Drag and drop an Excel file here, or click to select</p>
              <p style={{ marginTop: '0.5rem', fontSize: '0.9rem', color: '#666' }}>
                Supports .xlsx and .xls files
              </p>
            </div>
          )}
        </div>
        
        {status.message && (
          <div className={`status ${status.type}`}>
            {status.message}
          </div>
        )}
        
        {file && (
          <div style={{ marginTop: '1rem' }}>
            <p><strong>Selected file:</strong> {file.name}</p>
            <p><strong>Size:</strong> {(file.size / 1024).toFixed(2)} KB</p>
          </div>
        )}
      </div>
      
      {processedData.length > 0 && (
        <div className="card">
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '1rem' }}>
            <h2>Processed Data Preview ({processedData.length} entries)</h2>
            <button className="button" onClick={exportToExcel}>
              Export Excel
            </button>
          </div>
          
          <div className="preview">
            <table>
              <thead>
                <tr>
                  <th>Tag</th>
                  <th>Datum</th>
                  <th>Start</th>
                  <th>Ende</th>
                  <th>Stunden</th>
                  <th>Point of Sales</th>
                  <th>Bezirk</th>
                  <th>Marktname</th>
                  <th>Coffee Advisor</th>
                </tr>
              </thead>
              <tbody>
                {processedData.slice(0, 50).map((item, index) => (
                  <tr key={index}>
                    <td>{item.dayOfWeek}</td>
                    <td>{item.date}</td>
                    <td>{item.startTime}</td>
                    <td>{item.endTime}</td>
                    <td>{item.totalHours}</td>
                    <td>{item.pointOfSales}</td>
                    <td>{item.district}</td>
                    <td>{item.marketName}</td>
                    <td>{item.coffeeAdvisor}</td>
                  </tr>
                ))}
              </tbody>
            </table>
            {processedData.length > 50 && (
              <p style={{ padding: '1rem', textAlign: 'center', color: '#666' }}>
                Showing first 50 entries. Export to see all {processedData.length} entries.
              </p>
            )}
          </div>
        </div>
      )}
    </div>
  )
} 