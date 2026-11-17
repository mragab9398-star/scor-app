
import { useState } from 'react'
import * as XLSX from 'xlsx'

function NpiReport({ onBack }) {
  const [columns, setColumns] = useState([])
  const [row, setRow] = useState(null)
  const [fileName, setFileName] = useState('')
  const [error, setError] = useState('')

  const testSku = '01-10-10'

  function handleFileUpload(event) {
    const file = event.target.files[0]
    if (!file) return

    setError('')
    setRow(null)
    setColumns([])
    setFileName(file.name)

    const reader = new FileReader()

    reader.onload = (e) => {
      try {
        const workbook = XLSX.read(e.target.result, {
          type: 'array',
          cellDates: true,
        })

        const sheet = workbook.Sheets[workbook.SheetNames[0]]
        const data = XLSX.utils.sheet_to_json(sheet, {
          header: 1,
          defval: '',
        })

        const headers = data[0]
        const skuIndex = headers.findIndex(
          (header) => String(header).trim().toUpperCase() === 'SKU'
        )

        if (skuIndex === -1) {
          setError('SKU column was not found in the Excel file.')
          return
        }

        const item = data.slice(1).find(
          (values) => String(values[skuIndex]).trim() === testSku
        )

        if (!item) {
          setError(`Test SKU ${testSku} was not found.`)
          return
        }

        const validColumns = headers
          .map((header, index) => ({
            name: String(header).trim(),
            index,
          }))
          .filter((column) => column.name !== '')

        setColumns(validColumns)
        setRow(item)
      } catch {
        setError('Unable to read this Excel file. Please check the file.')
      }
    }

    reader.readAsArrayBuffer(file)
  }

  function moveColumn(index, direction) {
    const newColumns = [...columns]
    const targetIndex = index + direction

    if (targetIndex < 0 || targetIndex >= newColumns.length) return

    ;[newColumns[index], newColumns[targetIndex]] = [
      newColumns[targetIndex],
      newColumns[index],
    ]

    setColumns(newColumns)
  }

  return (
    <div className="npi-report-page">
      <div className="npi-report-header">
        <button type="button" onClick={onBack}>
          ← Back to Reports
        </button>

        <div>
          <h1>NPI Report</h1>
          <p>New Product Introduction Analysis</p>
        </div>
      </div>

      <div className="npi-upload-panel">
        <h3>Upload NPI Excel File</h3>
        <p>For the first test, the report will display SKU: {testSku}</p>

        <input
          type="file"
          accept=".xlsx,.xls"
          onChange={handleFileUpload}
        />

        {fileName && <p>Selected file: {fileName}</p>}
        {error && <p className="npi-error">{error}</p>}
      </div>

      {row && (
        <div className="npi-results">
          <h2>Test Item: {testSku}</h2>
          <p>Use the arrows to change the column order.</p>

          <div className="npi-column-order">
            {columns.map((column, index) => (
              <div className="npi-column-item" key={column.index}>
                <span>{index + 1}. {column.name}</span>

                <button
                  type="button"
                  disabled={index === 0}
                  onClick={() => moveColumn(index, -1)}
                  title="Move left"
                >
                  ←
                </button>

                <button
                  type="button"
                  disabled={index === columns.length - 1}
                  onClick={() => moveColumn(index, 1)}
                  title="Move right"
                >
                  →
                </button>
              </div>
            ))}
          </div>

          <div className="npi-table-wrapper">
            <table className="npi-table">
              <thead>
                <tr>
                  {columns.map((column) => (
                    <th key={column.index}>{column.name}</th>
                  ))}
                </tr>
              </thead>

              <tbody>
                <tr>
                  {columns.map((column) => (
                    <td key={column.index}>
                      {row[column.index] instanceof Date
                        ? row[column.index].toLocaleDateString()
                        : String(row[column.index] ?? '')}
                    </td>
                  ))}
                </tr>
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  )
}

export default NpiReport
