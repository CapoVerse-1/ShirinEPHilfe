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
    const days = ['Sonntag', 'Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag']
    return days[date.getDay()]
  }

  const getWorkingHours = (dayOfWeek: string): { start: string, end: string } => {
    if (dayOfWeek === 'Samstag' || dayOfWeek === 'Sonntag') {
      return { start: '09:00', end: '18:00' }
    }
    return { start: '09:30', end: '18:30' }
  }

  const getMonthInfo = (monthAbbr: string, actualYear?: number): { monthNumber: number, year: number, monthName: string } => {
    // Normalize the month abbreviation to handle case variations
    const normalizedMonth = monthAbbr.charAt(0).toUpperCase() + monthAbbr.slice(1).toLowerCase()
    
    const germanMonths: { [key: string]: { number: number, name: string } } = {
      'Jan': { number: 0, name: 'Januar' },
      'Feb': { number: 1, name: 'Februar' },
      'Mär': { number: 2, name: 'März' },
      'Mar': { number: 2, name: 'März' }, // Alternative for März
      'Apr': { number: 3, name: 'April' },
      'Mai': { number: 4, name: 'Mai' },
      'Jun': { number: 5, name: 'Juni' },
      'Jul': { number: 6, name: 'Juli' },
      'Aug': { number: 7, name: 'August' },
      'Sep': { number: 8, name: 'September' },
      'Okt': { number: 9, name: 'Oktober' },
      'Oct': { number: 9, name: 'Oktober' }, // Alternative for Oktober
      'Nov': { number: 10, name: 'November' },
      'Dez': { number: 11, name: 'Dezember' },
      'Dec': { number: 11, name: 'Dezember' } // Alternative for Dezember
    }
    
    console.log('Looking for month:', normalizedMonth) // Debug log
    const monthInfo = germanMonths[normalizedMonth]
    if (!monthInfo) {
      console.log('Available months:', Object.keys(germanMonths)) // Debug log
      throw new Error(`Unknown month abbreviation: ${monthAbbr} (normalized: ${normalizedMonth})`)
    }
    
    // Use the actual year from Excel data if provided, otherwise use current year
    const year = actualYear || new Date().getFullYear()
    
    return {
      monthNumber: monthInfo.number,
      year: year,
      monthName: monthInfo.name
    }
  }

  const getDaysInMonth = (month: number, year: number): number => {
    return new Date(year, month + 1, 0).getDate()
  }

  const excelDateToJSDate = (excelDate: number): Date => {
    // Excel date serial number to JavaScript Date
    // Excel counts from January 1, 1900, but has a leap year bug for 1900
    // JavaScript Date counts from January 1, 1970
    const excelEpoch = new Date(1899, 11, 30) // December 30, 1899
    const jsDate = new Date(excelEpoch.getTime() + excelDate * 24 * 60 * 60 * 1000)
    return jsDate
  }

  const extractMonthFromHeader = (data: any[][]): { monthAbbr: string, year: number } => {
    // Look for month from the first date in the header row, starting from column E (index 4)
    const headerRow = data[0]
    if (!headerRow) throw new Error('No header row found')
    
    console.log('Header row:', headerRow) // Debug log
    
    for (let i = 4; i < headerRow.length; i++) {
      const cellValue = headerRow[i]
      console.log(`Column ${i} value:`, cellValue, typeof cellValue) // Debug log
      
      if (cellValue) {
        // Check if it's a number (Excel date serial number)
        if (typeof cellValue === 'number') {
          try {
            const jsDate = excelDateToJSDate(cellValue)
            const month = jsDate.getMonth() // 0-11
            const year = jsDate.getFullYear() // Get the actual year from Excel
            
            const monthAbbreviations = [
              'Jan', 'Feb', 'Mär', 'Apr', 'Mai', 'Jun',
              'Jul', 'Aug', 'Sep', 'Okt', 'Nov', 'Dez'
            ]
            
            const monthAbbr = monthAbbreviations[month]
            console.log('Converted Excel date', cellValue, 'to JS date:', jsDate, 'month:', monthAbbr, 'year:', year) // Debug log
            return { monthAbbr, year }
          } catch (error) {
            console.log('Error converting Excel date:', error) // Debug log
            continue
          }
        }
        
        // Fallback: try to parse as string (in case some Excel files have text dates)
        const cellStr = String(cellValue).trim()
        
        // Try multiple patterns to extract month abbreviation
        // Pattern 1: "01.Aug", "02.Aug", etc.
        let match = cellStr.match(/\d{1,2}\.([A-Za-z]{3})/i)
        if (match) {
          console.log('Found month with pattern 1:', match[1]) // Debug log
          return { monthAbbr: match[1], year: new Date().getFullYear() }
        }
        
        // Pattern 2: "01Aug", "02Aug", etc. (without dot)
        match = cellStr.match(/\d{1,2}([A-Za-z]{3})/i)
        if (match) {
          console.log('Found month with pattern 2:', match[1]) // Debug log
          return { monthAbbr: match[1], year: new Date().getFullYear() }
        }
        
        // Pattern 3: Just the month abbreviation "Aug", "Jul", etc.
        match = cellStr.match(/^([A-Za-z]{3})$/i)
        if (match) {
          console.log('Found month with pattern 3:', match[1]) // Debug log
          return { monthAbbr: match[1], year: new Date().getFullYear() }
        }
      }
    }
    
    console.log('Could not find month in any column from E onwards') // Debug log
    throw new Error('Could not detect month from Excel headers')
  }

  const processExcelData = (data: any[][]) => {
    const result: PromotionData[] = []
    
    // Extract month and year from header row
    const { monthAbbr, year: excelYear } = extractMonthFromHeader(data)
    const currentYear = new Date().getFullYear() // Use current year (2025)
    const { monthNumber, year, monthName } = getMonthInfo(monthAbbr, currentYear)
    const daysInMonth = getDaysInMonth(monthNumber, currentYear)
    
    console.log(`Using Excel month: ${monthAbbr}, Excel year: ${excelYear}, Current year: ${currentYear}`) // Debug log
    
    // Skip header row, start from row 1
    for (let rowIndex = 1; rowIndex < data.length; rowIndex++) {
      const row = data[rowIndex]
      if (!row || row.length < 5) continue
      
      const marketName = row[0] || ''
      const district = row[1] || ''
      
      // Process each date column (starting from column E = index 4)
      const maxColumns = 4 + daysInMonth // E + number of days in month
      for (let colIndex = 4; colIndex < row.length && colIndex < maxColumns; colIndex++) {
        const value = row[colIndex]
        if (!value || value === 0) continue
        
        // Calculate the date using CURRENT YEAR but same day/month from Excel
        const day = colIndex - 3 // Column E = day 1, F = day 2, etc.
        const dateInCurrentYear = new Date(currentYear, monthNumber, day)
        const dayOfWeek = getDayOfWeek(dateInCurrentYear)
        const workingHours = getWorkingHours(dayOfWeek)
        
        // Format date as DD.MM.YYYY with CURRENT YEAR
        const formattedDate = `${day.toString().padStart(2, '0')}.${(monthNumber + 1).toString().padStart(2, '0')}.${currentYear}`
        
        console.log(`Day ${day} of ${monthAbbr} ${currentYear} is a ${dayOfWeek}`) // Debug log
        
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
    
    // Add autofilter to the data range
    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1')
    worksheet['!autofilter'] = { ref: worksheet['!ref'] || 'A1' }
    
    // Set column widths for better readability
    worksheet['!cols'] = [
      { wch: 12 }, // Tag der Woche
      { wch: 12 }, // Datum
      { wch: 10 }, // Startzeit
      { wch: 10 }, // Endzeit
      { wch: 12 }, // Gesamtstunden
      { wch: 20 }, // Point of Sales
      { wch: 15 }, // Bezirk
      { wch: 20 }, // Marktname
      { wch: 15 }  // Coffee Advisor
    ]
    
    // Style the Coffee Advisor column (column I) with yellow background
    for (let row = 0; row <= processedData.length; row++) {
      const cellAddress = XLSX.utils.encode_cell({ r: row, c: 8 }) // Column I (index 8)
      if (!worksheet[cellAddress]) worksheet[cellAddress] = { t: 's', v: '' }
      
      worksheet[cellAddress].s = {
        fill: {
          fgColor: { rgb: 'FFFF00' } // Yellow background
        },
        border: {
          top: { style: 'thin', color: { rgb: '000000' } },
          bottom: { style: 'thin', color: { rgb: '000000' } },
          left: { style: 'thin', color: { rgb: '000000' } },
          right: { style: 'thin', color: { rgb: '000000' } }
        }
      }
    }
    
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