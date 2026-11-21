import { useState } from 'react'
import * as XLSX from 'xlsx'

// These headers are taken from the supplied ON HAND 16-11-2026 workbook.
// Keep the original spelling, language, spacing and warehouse codes unchanged.
const DEFAULT_COLUMNS = [
  'رقم الصنف',
  'اسم المنتج',
  'الفئة ',
  'حالة دورة حياة المنتج',
  'Vend group',
  'Brand',
  'Level 1',
  'Level 2',
  'Level 3',
  'Level 4',
  'تعبئة الكرتون',
  'اجمالي المخزون ',
  '010',
  '0110',
  '0115',
  '013',
  '0182',
  '0221',
  '0225',
  '0231',
  '0235',
  '028',
  '0311',
  '0315',
  '0321',
  '0325',
  '038',
  '0415',
  '0421',
  '0425',
  '0441',
  '0445',
  '048',
  '0482',
  '0511',
  '0515',
  '0521',
  '0525',
  '0545',
  '0611',
  '0615',
  '058',
  '0581',
  '0582',
]

const ON_HAND_STORAGE_KEY = 'scor-on-hand-report-v1'

function loadSavedOnHand() {
  try {
    const saved = JSON.parse(localStorage.getItem(ON_HAND_STORAGE_KEY) || 'null')
    if (saved && Array.isArray(saved.columns) && Array.isArray(saved.rows)) return saved
  } catch (error) {
    console.warn('Unable to load saved On Hand data:', error)
  }
  return null
}

function OnHandReport({ onBack }) {
  const [initialSaved] = useState(() => loadSavedOnHand())
  const [columns, setColumns] = useState(() => initialSaved?.columns || DEFAULT_COLUMNS)
  const [rows, setRows] = useState(() => initialSaved?.rows || [])
  const [fileName, setFileName] = useState(() => initialSaved?.fileName || '')
  const [hasUnsavedChanges, setHasUnsavedChanges] = useState(false)
  const [saveMessage, setSaveMessage] = useState(initialSaved ? 'Saved data restored from this browser.' : '')
  const [search, setSearch] = useState('')
  const [message, setMessage] = useState('')
  const [pageSize, setPageSize] = useState(100)
  const [currentPage, setCurrentPage] = useState(1)
  const [duplicateCount, setDuplicateCount] = useState(0)
  const [blankSkuCount, setBlankSkuCount] = useState(0)

  function handleImport(event) {
    const file = event.target.files?.[0]
    if (!file) return

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

        if (!data.length || !data[0].some((cell) => String(cell).trim())) {
          setMessage('The Excel file is empty or has no column headers.')
          return
        }

        const headers = data[0].map((header, index) => {
          const name = String(header).trim()
          return name || `Column ${index + 1}`
        })

        const importedRows = data.slice(1)
          .filter((row) => row.some((cell) => cell !== ''))
          .map((row) => headers.map((_, index) => row[index] ?? ''))

        // First column in the supplied On Hand sheet is the item/SKU identifier.
        const skuIndex = Math.max(0, headers.findIndex((header) =>
          /sku|item.?code|رقم الصنف|كود الصنف/i.test(String(header))
        ))
        const skuValues = importedRows
          .map((row) => String(row[skuIndex] ?? '').trim())
          .filter(Boolean)
        const duplicateValues = skuValues.filter((sku, index) => skuValues.indexOf(sku) !== index)
        const uniqueDuplicates = new Set(duplicateValues)
        const blanks = importedRows.filter((row) => !String(row[skuIndex] ?? '').trim()).length

        setColumns(headers)
        setRows(importedRows)
        setHasUnsavedChanges(true)
        setSaveMessage('Imported data is not saved yet.')
        setFileName(file.name)
        setDuplicateCount(uniqueDuplicates.size)
        setBlankSkuCount(blanks)
        setCurrentPage(1)
        setMessage(
          `Imported ${importedRows.length.toLocaleString()} rows. ` +
          `${uniqueDuplicates.size} duplicated item codes; ${blanks} blank item codes.`
        )
      } catch {
        setMessage('Unable to read the Excel file. Please check the workbook format.')
      }
    }

    reader.readAsArrayBuffer(file)
    event.target.value = ''
  }

  function updateCell(rowIndex, columnIndex, value) {
    setHasUnsavedChanges(true)
    setSaveMessage('You have unsaved changes.')
    setRows((currentRows) =>
      currentRows.map((row, index) =>
        index === rowIndex
          ? row.map((cell, cellIndex) =>
              cellIndex === columnIndex ? value : cell
            )
          : row
      )
    )
  }

  function addItem() {
    setHasUnsavedChanges(true)
    setSaveMessage('New item added. Save to keep your changes.')
    setRows((currentRows) => [
      ...currentRows,
      columns.map(() => ''),
    ])
    setCurrentPage(Math.ceil((rows.length + 1) / pageSize))
    setMessage('New item row added.')
  }

  function deleteItem(rowIndex) {
    setHasUnsavedChanges(true)
    setSaveMessage('Item deleted. Save to keep your changes.')
    setRows((currentRows) =>
      currentRows.filter((_, index) => index !== rowIndex)
    )
    setCurrentPage(1)
    setMessage('Item row deleted.')
  }

  function moveColumn(index, direction) {
    setHasUnsavedChanges(true)
    setSaveMessage('Column order changed. Save to keep your changes.')
    const targetIndex = index + direction
    if (targetIndex < 0 || targetIndex >= columns.length) return

    setColumns((current) => {
      const next = [...current]
      ;[next[index], next[targetIndex]] = [next[targetIndex], next[index]]
      return next
    })

    setRows((current) =>
      current.map((row) => {
        const next = [...row]
        ;[next[index], next[targetIndex]] = [next[targetIndex], next[index]]
        return next
      })
    )
  }

  function downloadTemplate() {
    const blankRows = Array.from({ length: 100 }, () =>
      DEFAULT_COLUMNS.map(() => '')
    )
    const worksheet = XLSX.utils.aoa_to_sheet([DEFAULT_COLUMNS, ...blankRows])
    worksheet['!cols'] = DEFAULT_COLUMNS.map((header) => ({
      wch: Math.min(Math.max(String(header).length + 3, 12), 24),
    }))
    worksheet['!autofilter'] = {
      ref: `A1:${XLSX.utils.encode_col(DEFAULT_COLUMNS.length - 1)}101`,
    }
    worksheet['!freeze'] = { xSplit: 0, ySplit: 1 }
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, worksheet, 'On Hand')
    XLSX.writeFile(workbook, 'SCOR_On_Hand_Import_Template_100_Items.xlsx')
    setMessage('100-row On Hand import template downloaded.')
  }

  function saveAllChanges() {
    try {
      localStorage.setItem(ON_HAND_STORAGE_KEY, JSON.stringify({
        columns,
        rows,
        fileName,
        savedAt: new Date().toISOString(),
      }))
      setHasUnsavedChanges(false)
      setSaveMessage(`Saved ${rows.length.toLocaleString()} item rows in this browser.`)
    } catch (error) {
      console.error('Unable to save On Hand data:', error)
      setSaveMessage('Save failed. Your browser storage may be full or unavailable.')
    }
  }

  function exportExcel() {
    if (!columns.length) {
      setMessage('There are no columns to export.')
      return
    }
    const worksheet = XLSX.utils.aoa_to_sheet([
      columns,
      ...rows,
    ])
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, worksheet, 'On Hand')
    XLSX.writeFile(workbook, 'SCOR_On_Hand_Report.xlsx')
    setMessage(`Exported ${rows.length} item rows successfully.`)
  }

  const filteredRows = rows
    .map((row, index) => ({ row, originalIndex: index }))
    .filter(({ row }) =>
      row.some((cell) =>
        String(cell).toLowerCase().includes(search.toLowerCase())
      )
    )

  const totalPages = Math.max(1, Math.ceil(filteredRows.length / pageSize))
  const safePage = Math.min(currentPage, totalPages)
  const pageRows = filteredRows.slice((safePage - 1) * pageSize, safePage * pageSize)

  return (
    <div className="onhand-page">
      <div className="onhand-header">
        <div>
          <button type="button" onClick={onBack}>
            ← Back to Reports
          </button>
          <h1>On Hand Report</h1>
          <p>Inventory balances by item and warehouse</p>
        </div>
      </div>

      <div className="onhand-toolbar">
        <label className="onhand-upload">
          Import Excel
          <input
            type="file"
            accept=".xlsx,.xls"
            onChange={handleImport}
            hidden
          />
        </label>

        <button type="button" onClick={downloadTemplate}>
          Download 100-Row Template
        </button>

        <button type="button" onClick={addItem}>
          + Add Item
        </button>

        <button type="button" onClick={saveAllChanges} disabled={!hasUnsavedChanges}>
          Save Changes
        </button>

        <button type="button" onClick={exportExcel} disabled={!rows.length}>
          Export Excel
        </button>

        <input
          type="text"
          placeholder="Search SKU, product or warehouse..."
          value={search}
          onChange={(e) => {
            setSearch(e.target.value)
            setCurrentPage(1)
          }}
        />
      </div>

      {fileName && (
        <p className="onhand-file-name">
          Source file: {fileName}
        </p>
      )}

      {saveMessage && (
        <p className="onhand-message" role="status">{saveMessage}</p>
      )}
      {message && (
        <p className="onhand-message" role="status">{message}</p>
      )}

      <div className="onhand-summary">
        <strong>{rows.length}</strong>
        <span>Total Items</span>
        <span>{columns.length} Columns</span>
        <span>{filteredRows.length.toLocaleString()} Matching Rows</span>
        <span>{duplicateCount} Duplicate Codes</span>
        <span>{blankSkuCount} Blank Item Codes</span>
      </div>

      <div className="onhand-pagination" style={{
        display: 'flex',
        alignItems: 'center',
        flexWrap: 'wrap',
        gap: '10px',
        padding: '10px 0'
      }}>
        <label>
          Rows per page:{' '}
          <select
            value={pageSize}
            onChange={(e) => {
              setPageSize(Number(e.target.value))
              setCurrentPage(1)
            }}
          >
            <option value={100}>100</option>
            <option value={250}>250</option>
            <option value={500}>500</option>
          </select>
        </label>
        <button type="button" disabled={safePage <= 1} onClick={() => setCurrentPage(1)}>
          First
        </button>
        <button type="button" disabled={safePage <= 1} onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}>
          Previous
        </button>
        <span>Page {safePage.toLocaleString()} of {totalPages.toLocaleString()}</span>
        <button type="button" disabled={safePage >= totalPages} onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}>
          Next
        </button>
        <button type="button" disabled={safePage >= totalPages} onClick={() => setCurrentPage(totalPages)}>
          Last
        </button>
      </div>

      <details className="onhand-column-settings" style={{ margin: '12px 0' }}>
        <summary style={{ cursor: 'pointer', fontWeight: 600, padding: '8px 0' }}>
          Column Order & Settings ({columns.length} columns)
        </summary>
        <div style={{ display: 'flex', flexWrap: 'wrap', gap: '8px', padding: '10px 0' }}>
          {columns.map((column, index) => (
            <div className="onhand-column-item" key={`${column}-${index}`}
              style={{ display: 'inline-flex', alignItems: 'center', gap: '4px', border: '1px solid #d5dce5', borderRadius: '6px', padding: '4px 6px' }}>
              <span>{index + 1}. {column}</span>
              <button
                type="button"
                disabled={index === 0}
                onClick={() => moveColumn(index, -1)}
                aria-label={`Move ${column} left`}
              >
                ←
              </button>
              <button
                type="button"
                disabled={index === columns.length - 1}
                onClick={() => moveColumn(index, 1)}
                aria-label={`Move ${column} right`}
              >
                →
              </button>
            </div>
          ))}
        </div>
      </details>

      <div className="onhand-table-wrapper" style={{
        width: '100%',
        maxWidth: '100%',
        overflowX: 'auto',
        overflowY: 'auto',
        border: '1px solid #d7dee8',
        borderRadius: '8px',
        background: '#fff'
      }}>
        <table className="onhand-table" style={{
          borderCollapse: 'separate',
          borderSpacing: 0,
          width: 'max-content',
          minWidth: '100%',
          whiteSpace: 'nowrap'
        }}>
          <thead>
            <tr>
              {columns.map((column, index) => (
                <th key={`${column}-${index}`} style={{
                  position: 'sticky',
                  top: 0,
                  zIndex: 2,
                  background: '#17365d',
                  color: '#fff',
                  padding: '12px 14px',
                  textAlign: 'center',
                  whiteSpace: 'nowrap',
                  borderRight: '1px solid #355575'
                }}>{column}</th>
              ))}
            </tr>
          </thead>

          <tbody>
            {pageRows.length > 0 ? (
              pageRows.map(({ row, originalIndex }) => (
                <tr key={originalIndex}>
                  {columns.map((column, columnIndex) => (
                    <td key={`${column}-${columnIndex}`} style={{
                      padding: '6px 8px',
                      borderBottom: '1px solid #e5eaf0',
                      borderRight: '1px solid #edf0f4',
                      minWidth: '130px'
                    }}>
                      <input
                        style={{
                          width: '115px',
                          minWidth: '115px',
                          boxSizing: 'border-box',
                          padding: '7px 8px',
                          border: '1px solid #d7dee8',
                          borderRadius: '4px'
                        }}
                        value={String(row[columnIndex] ?? '')}
                        onChange={(e) =>
                          updateCell(originalIndex, columnIndex, e.target.value)
                        }
                        aria-label={`${column}, row ${originalIndex + 1}`}
                      />
                    </td>
                  ))}
                </tr>
              ))
            ) : (
              <tr>
                <td colSpan={columns.length} className="onhand-no-rows" style={{
                  padding: '24px',
                  textAlign: 'center',
                  color: '#64748b'
                }}>
                  {search
                    ? 'No matching items found.'
                    : 'No inventory rows yet. Import Excel or click Add Item.'}
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  )
}

export default OnHandReport
