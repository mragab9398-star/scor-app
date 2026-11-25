import { useMemo, useState } from 'react'
import * as XLSX from 'xlsx'

const DEFAULT_COLUMNS = [
  'رقم الصنف', 'اسم المنتج', 'الفئة ', 'حالة دورة حياة المنتج',
  'Vend group', 'Brand', 'Level 1', 'Level 2', 'Level 3', 'Level 4',
  'تعبئة الكرتون', 'اجمالي المخزون ', '010', '0110', '0115', '013',
  '0182', '0221', '0225', '0231', '0235', '028', '0311', '0315',
  '0321', '0325', '038', '0415', '0421', '0425', '0441', '0445',
  '048', '0482', '0511', '0515', '0521', '0525', '0545', '0611',
  '0615', '058', '0581', '0582',
]

const STORAGE_KEY = 'scor-on-hand-report-v1'

function loadSavedData() {
  try {
    const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || 'null')
    if (saved && Array.isArray(saved.columns) && Array.isArray(saved.rows)) return saved
  } catch (error) {
    console.warn('Unable to load On Hand data:', error)
  }
  return null
}

const colors = {
  navy: '#12304d',
  navy2: '#1c456b',
  teal: '#078b91',
  bg: '#eef3f8',
  surface: '#ffffff',
  border: '#d9e3ed',
  text: '#20364b',
  muted: '#718398',
  green: '#16845b',
  amber: '#a96812',
}

function OnHandReport({ onBack }) {
  const [initialSaved] = useState(() => loadSavedData())
  const [columns, setColumns] = useState(() => initialSaved?.columns || DEFAULT_COLUMNS)
  const [rows, setRows] = useState(() => initialSaved?.rows || [])
  const [fileName, setFileName] = useState(() => initialSaved?.fileName || '')
  const [hasUnsavedChanges, setHasUnsavedChanges] = useState(false)
  const [saveMessage, setSaveMessage] = useState(
    initialSaved ? 'Saved data restored from this browser.' : 'No saved changes yet.'
  )
  const [search, setSearch] = useState('')
  const [message, setMessage] = useState('')
  const [pageSize, setPageSize] = useState(100)
  const [currentPage, setCurrentPage] = useState(1)
  const [duplicateCount, setDuplicateCount] = useState(0)
  const [blankSkuCount, setBlankSkuCount] = useState(0)

  const skuIndex = Math.max(0, columns.findIndex((column) =>
    /sku|item.?code|رقم الصنف|كود الصنف/i.test(String(column))
  ))

  const filteredRows = useMemo(() => rows
    .map((row, index) => ({ row, originalIndex: index }))
    .filter(({ row }) => row.some((cell) =>
      String(cell ?? '').toLowerCase().includes(search.toLowerCase())
    )), [rows, search])

  const totalPages = Math.max(1, Math.ceil(filteredRows.length / pageSize))
  const safePage = Math.min(currentPage, totalPages)
  const pageRows = filteredRows.slice((safePage - 1) * pageSize, safePage * pageSize)

  function markDirty(text = 'You have unsaved changes.') {
    setHasUnsavedChanges(true)
    setSaveMessage(text)
  }

  function handleImport(event) {
    const file = event.target.files?.[0]
    if (!file) return
    const reader = new FileReader()

    reader.onload = (e) => {
      try {
        const workbook = XLSX.read(e.target.result, { type: 'array', cellDates: true })
        const sheet = workbook.Sheets[workbook.SheetNames[0]]
        const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' })

        if (!data.length || !data[0].some((cell) => String(cell).trim())) {
          setMessage('The Excel file is empty or has no column headers.')
          return
        }

        const headers = data[0].map((header, index) =>
          String(header).trim() || `Column ${index + 1}`
        )
        const importedRows = data.slice(1)
          .filter((row) => row.some((cell) => cell !== ''))
          .map((row) => headers.map((_, index) => row[index] ?? ''))

        const importedSkuIndex = Math.max(0, headers.findIndex((header) =>
          /sku|item.?code|رقم الصنف|كود الصنف/i.test(String(header))
        ))
        const skuValues = importedRows
          .map((row) => String(row[importedSkuIndex] ?? '').trim())
          .filter(Boolean)
        const duplicateValues = skuValues.filter((sku, index) => skuValues.indexOf(sku) !== index)

        setColumns(headers)
        setRows(importedRows)
        setFileName(file.name)
        setDuplicateCount(new Set(duplicateValues).size)
        setBlankSkuCount(importedRows.filter((row) =>
          !String(row[importedSkuIndex] ?? '').trim()
        ).length)
        setCurrentPage(1)
        markDirty('Imported data is ready. Click Save Changes to keep it in this browser.')
        setMessage(`Imported ${importedRows.length.toLocaleString()} rows from ${file.name}.`)
      } catch {
        setMessage('Unable to read the Excel file. Please check the workbook format.')
      }
    }

    reader.readAsArrayBuffer(file)
    event.target.value = ''
  }

  function updateCell(rowIndex, columnIndex, value) {
    markDirty()
    setRows((current) => current.map((row, index) =>
      index === rowIndex
        ? row.map((cell, cellIndex) => cellIndex === columnIndex ? value : cell)
        : row
    ))
  }

  function addItem() {
    markDirty('New row added. Complete the item details, then save.')
    setRows((current) => [...current, columns.map(() => '')])
    setCurrentPage(Math.max(1, Math.ceil((rows.length + 1) / pageSize)))
    setMessage('A new editable item row was added.')
  }

  function deleteItem(rowIndex) {
    markDirty('A row was removed. Click Save Changes to keep the update.')
    setRows((current) => current.filter((_, index) => index !== rowIndex))
    setCurrentPage(1)
    setMessage('Item row removed from the current report.')
  }

  function moveColumn(index, direction) {
    const target = index + direction
    if (target < 0 || target >= columns.length) return
    markDirty('Column order changed. Click Save Changes to keep the new order.')

    setColumns((current) => {
      const next = [...current]
      ;[next[index], next[target]] = [next[target], next[index]]
      return next
    })
    setRows((current) => current.map((row) => {
      const next = [...row]
      ;[next[index], next[target]] = [next[target], next[index]]
      return next
    }))
  }

  function downloadTemplate() {
    const blankRows = Array.from({ length: 100 }, () => DEFAULT_COLUMNS.map(() => ''))
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
      localStorage.setItem(STORAGE_KEY, JSON.stringify({
        columns, rows, fileName, savedAt: new Date().toISOString(),
      }))
      setHasUnsavedChanges(false)
      setSaveMessage(`Saved ${rows.length.toLocaleString()} item rows in this browser.`)
      setMessage('Your On Hand report changes have been saved.')
    } catch (error) {
      console.error('Unable to save On Hand data:', error)
      setSaveMessage('Save failed. Browser storage may be full or unavailable.')
    }
  }

  function exportExcel() {
    if (!columns.length) return setMessage('There are no columns to export.')
    const worksheet = XLSX.utils.aoa_to_sheet([columns, ...rows])
    worksheet['!cols'] = columns.map((column) => ({
      wch: Math.min(Math.max(String(column).length + 3, 12), 24),
    }))
    worksheet['!autofilter'] = {
      ref: `A1:${XLSX.utils.encode_col(columns.length - 1)}${Math.max(rows.length + 1, 1)}`,
    }
    worksheet['!freeze'] = { xSplit: 0, ySplit: 1 }
    const workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, worksheet, 'On Hand')
    XLSX.writeFile(workbook, 'SCOR_On_Hand_Report.xlsx')
    setMessage(`Exported all ${rows.length.toLocaleString()} item rows to Excel.`)
  }

  const buttonBase = {
    minHeight: 42, borderRadius: 9, padding: '0 16px', border: `1px solid ${colors.border}`,
    background: '#fff', color: colors.navy, fontWeight: 700, fontSize: 13,
    cursor: 'pointer', display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
    gap: 8, whiteSpace: 'nowrap',
  }
  const cardStyle = {
    background: colors.surface, border: `1px solid ${colors.border}`, borderRadius: 13,
    padding: '15px 18px', boxShadow: '0 3px 12px rgba(19,48,77,.045)',
  }
  const smallLabel = {
    display: 'block', color: colors.muted, fontSize: 11, fontWeight: 800,
    letterSpacing: '.07em', textTransform: 'uppercase', marginBottom: 8,
  }
  const statValue = { fontSize: 25, lineHeight: 1.1, fontWeight: 800, color: colors.navy }

  return (
    <div style={{
      width: '100%', maxWidth: 'none', minWidth: 0, boxSizing: 'border-box',
      minHeight: 'calc(100vh - 105px)', padding: '20px 22px 26px',
      background: colors.bg, color: colors.text,
      fontFamily: 'Segoe UI, Tahoma, Arial, sans-serif',
    }}>
      {/* Report heading */}
      <div style={{
        display: 'flex', justifyContent: 'space-between', alignItems: 'center',
        gap: 16, flexWrap: 'wrap', marginBottom: 18,
      }}>
        <div>
          <button type="button" onClick={onBack} style={{
            border: 0, background: 'transparent', color: colors.navy2,
            fontWeight: 700, padding: '0 0 9px', cursor: 'pointer', fontSize: 13,
          }}>← Back to Reports</button>
          <div style={{ display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap' }}>
            <h1 style={{ margin: 0, color: colors.navy, fontSize: 30, lineHeight: 1.15, fontWeight: 800 }}>
              On Hand Report
            </h1>
            <span style={{
              padding: '6px 11px', borderRadius: 20, background: '#e1f2f2',
              color: '#08777d', fontSize: 11, fontWeight: 800, letterSpacing: '.04em',
            }}>INVENTORY</span>
          </div>
          <p style={{ margin: '7px 0 0', color: colors.muted, fontSize: 14 }}>
            Inventory balances by item and warehouse
          </p>
        </div>
        <div style={{
          color: colors.muted, fontSize: 12, fontWeight: 600,
          background: '#fff', border: `1px solid ${colors.border}`,
          borderRadius: 9, padding: '10px 13px',
        }}>
          {columns.length} fields · {rows.length.toLocaleString()} records
        </div>
      </div>

      {/* Toolbar */}
      <div style={{
        ...cardStyle, display: 'flex', alignItems: 'center', gap: 10,
        flexWrap: 'wrap', marginBottom: 16, padding: 14,
      }}>
        <label style={{
          ...buttonBase, background: colors.navy, color: '#fff',
          borderColor: colors.navy, cursor: 'pointer',
        }}>
          <span>↑</span> Import Excel
          <input type="file" accept=".xlsx,.xls" onChange={handleImport} hidden />
        </label>
        <button type="button" onClick={downloadTemplate} style={buttonBase}>
          ↓ Download Template
        </button>
        <button type="button" onClick={addItem} style={{
          ...buttonBase, background: '#e7f4f4', color: '#08777d', borderColor: '#c7e6e6',
        }}>
          + Add Item
        </button>
        <button type="button" onClick={saveAllChanges} disabled={!hasUnsavedChanges} style={{
          ...buttonBase,
          background: hasUnsavedChanges ? colors.teal : '#edf1f5',
          color: hasUnsavedChanges ? '#fff' : '#94a3b2',
          borderColor: hasUnsavedChanges ? colors.teal : '#e0e6ed',
          cursor: hasUnsavedChanges ? 'pointer' : 'not-allowed',
        }}>
          ✓ Save Changes
        </button>
        <button type="button" onClick={exportExcel} disabled={!rows.length} style={{
          ...buttonBase, background: rows.length ? '#fff' : '#edf1f5',
          color: rows.length ? colors.navy : '#94a3b2',
          cursor: rows.length ? 'pointer' : 'not-allowed',
        }}>
          ↓ Export Excel
        </button>
        <div style={{ flex: '1 1 260px', minWidth: 230, marginLeft: 'auto' }}>
          <input
            type="search" value={search}
            onChange={(e) => { setSearch(e.target.value); setCurrentPage(1) }}
            placeholder="Search SKU, product, brand or warehouse..."
            style={{
              width: '100%', height: 42, boxSizing: 'border-box', padding: '0 14px',
              border: `1px solid ${colors.border}`, borderRadius: 9,
              background: '#fbfdff', color: colors.text, fontSize: 13, outlineColor: colors.teal,
            }}
          />
        </div>
      </div>

      {/* Save/import feedback */}
      {(saveMessage || message || fileName) && (
        <div style={{
          display: 'flex', alignItems: 'center', justifyContent: 'space-between',
          flexWrap: 'wrap', gap: 8, padding: '10px 14px', marginBottom: 15,
          borderRadius: 9, border: `1px solid ${hasUnsavedChanges ? '#f0d8b6' : '#cfe7dc'}`,
          background: hasUnsavedChanges ? '#fff8ed' : '#f0f8f4',
          color: hasUnsavedChanges ? '#8a581c' : '#276448', fontSize: 12, fontWeight: 600,
        }}>
          <span>{saveMessage}</span>
          {fileName && <span style={{ color: colors.muted }}>Source: {fileName}</span>}
          {message && <span>{message}</span>}
        </div>
      )}

      {/* KPI strip */}
      <div style={{
        display: 'grid', gridTemplateColumns: 'repeat(4, minmax(145px, 1fr))',
        gap: 12, marginBottom: 15,
      }}>
        <div style={cardStyle}>
          <span style={smallLabel}>Total items</span>
          <strong style={statValue}>{rows.length.toLocaleString()}</strong>
        </div>
        <div style={cardStyle}>
          <span style={smallLabel}>Report columns</span>
          <strong style={statValue}>{columns.length}</strong>
        </div>
        <div style={cardStyle}>
          <span style={smallLabel}>Duplicate item codes</span>
          <strong style={{ ...statValue, color: duplicateCount ? colors.amber : colors.green }}>
            {duplicateCount.toLocaleString()}
          </strong>
        </div>
        <div style={cardStyle}>
          <span style={smallLabel}>Missing item codes</span>
          <strong style={{ ...statValue, color: blankSkuCount ? colors.amber : colors.green }}>
            {blankSkuCount.toLocaleString()}
          </strong>
        </div>
      </div>

      {/* Table controls */}
      <div style={{
        ...cardStyle, display: 'flex', alignItems: 'center', gap: 10,
        flexWrap: 'wrap', marginBottom: 12, padding: '10px 14px',
      }}>
        <label style={{ color: colors.muted, fontSize: 12, fontWeight: 700 }}>
          Rows per page&nbsp;
          <select value={pageSize} onChange={(e) => {
            setPageSize(Number(e.target.value)); setCurrentPage(1)
          }} style={{
            height: 34, padding: '0 9px', border: `1px solid ${colors.border}`,
            borderRadius: 7, background: '#fff', color: colors.text, fontWeight: 700,
          }}>
            <option value={100}>100</option>
            <option value={250}>250</option>
            <option value={500}>500</option>
          </select>
        </label>
        <span style={{ marginLeft: 'auto', color: colors.muted, fontSize: 12 }}>
          Showing {filteredRows.length ? ((safePage - 1) * pageSize + 1).toLocaleString() : 0}
          –{Math.min(safePage * pageSize, filteredRows.length).toLocaleString()}
          {' '}of {filteredRows.length.toLocaleString()}
        </span>
        {[
          ['First', () => setCurrentPage(1), safePage <= 1],
          ['Previous', () => setCurrentPage((p) => Math.max(1, p - 1)), safePage <= 1],
          [`Page ${safePage} / ${totalPages}`, null, true],
          ['Next', () => setCurrentPage((p) => Math.min(totalPages, p + 1)), safePage >= totalPages],
          ['Last', () => setCurrentPage(totalPages), safePage >= totalPages],
        ].map(([label, action, disabled], index) => (
          <button key={index} type="button" disabled={disabled || !action}
            onClick={action || undefined}
            style={{
              minHeight: 34, padding: '0 11px', borderRadius: 7,
              border: `1px solid ${colors.border}`,
              background: index === 2 ? '#eaf1f7' : '#fff',
              color: disabled && index !== 2 ? '#a4b0bd' : colors.navy2,
              fontSize: 12, fontWeight: 700,
              cursor: disabled || !action ? 'default' : 'pointer',
            }}>
            {label}
          </button>
        ))}
      </div>

      {/* Compact, collapsible column ordering */}
      <details style={{
        marginBottom: 12, background: '#f8fafc',
        border: `1px solid ${colors.border}`, borderRadius: 10, padding: '0 14px',
      }}>
        <summary style={{
          cursor: 'pointer', padding: '12px 0', color: colors.navy2,
          fontSize: 13, fontWeight: 800,
        }}>
          Column Order & Settings ({columns.length} columns)
        </summary>
        <div style={{
          display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(220px, 1fr))',
          gap: 7, padding: '0 0 14px',
        }}>
          {columns.map((column, index) => (
            <div key={`${column}-${index}`} style={{
              display: 'flex', alignItems: 'center', gap: 7,
              minWidth: 0, padding: '6px 8px', borderRadius: 7,
              background: '#fff', border: `1px solid ${colors.border}`,
            }}>
              <span style={{
                flex: 1, minWidth: 0, overflow: 'hidden', textOverflow: 'ellipsis',
                whiteSpace: 'nowrap', color: colors.text, fontSize: 11, fontWeight: 600,
              }}>{index + 1}. {column}</span>
              <button type="button" disabled={index === 0} onClick={() => moveColumn(index, -1)}
                style={{ ...buttonBase, minHeight: 27, width: 30, padding: 0 }}>←</button>
              <button type="button" disabled={index === columns.length - 1} onClick={() => moveColumn(index, 1)}
                style={{ ...buttonBase, minHeight: 27, width: 30, padding: 0 }}>→</button>
            </div>
          ))}
        </div>
      </details>

      {/* Expanded report grid */}
    {/* Expanded report grid */}
<div style={{
  width: '100%',
  height: '72vh',
  minHeight: 620,
  overflow: 'auto',
  border: '1px solid #d8e2ed',
  borderRadius: 14,
  background: '#ffffff',
  boxShadow: '0 6px 24px rgba(19,48,77,.09)',
}}>
     
<table style={{
  borderCollapse: 'separate',
  borderSpacing: 0,
  width: 'max-content',
  minWidth: '100%',
  whiteSpace: 'nowrap',
  fontSize: 14,
}}>
          <thead>
            <tr>
              {columns.map((column, index) => (
                <th key={`${column}-${index}`} style={{
                  position: 'sticky', top: 0, zIndex: 3,
                  background: colors.navy, color: '#fff',
padding: '10px 12px',
minWidth: index < 4 ? 180 : 150,
borderRight: '1px solid #e7edf3',
borderBottom: '1px solid #e7edf3',
                }}>{column}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {pageRows.length ? pageRows.map(({ row, originalIndex }, visibleIndex) => (
              <tr key={originalIndex} style={{
                background: visibleIndex % 2 ? '#f7fafd' : '#fff',
              }}>
                {columns.map((column, columnIndex) => (
                  <td key={`${column}-${columnIndex}`} style={{
                    padding: '7px 9px', minWidth: columnIndex < 4 ? 180 : 150,
                    borderRight: '1px solid #e7edf3', borderBottom: '1px solid #e7edf3',
                  }}>
                    <input
                      value={String(row[columnIndex] ?? '')}
                      onChange={(e) => updateCell(originalIndex, columnIndex, e.target.value)}
                      aria-label={`${column}, row ${originalIndex + 1}`}
                      style={{
                        width: '100%', minWidth: columnIndex < 4 ? 137 : 108,
                       height: 44,
boxSizing: 'border-box',
padding: '0 12px',
border: '1px solid #dce5ee',
borderRadius: 7,
background: '#fff',
color: colors.text,
fontSize: 13,
outlineColor: colors.teal,
                      }}
                    />
                  </td>
                ))}
              </tr>
            )) : (
              <tr>
                <td colSpan={columns.length} style={{
                  padding: '60px 24px', textAlign: 'center', color: colors.muted,
                  fontSize: 14,
                }}>
                  {search
                    ? 'No matching items found. Try another search.'
                    : 'No inventory rows yet. Import an Excel file or click Add Item.'}
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
      <div style={{
        display: 'flex', justifyContent: 'space-between', gap: 12,
        flexWrap: 'wrap', padding: '10px 2px 0', color: colors.muted, fontSize: 11,
      }}>
        <span>Tip: use the horizontal scrollbar to view all warehouse columns.</span>
        <span>{columns.length} columns · {pageRows.length.toLocaleString()} rows on this page</span>
      </div>
    </div>
  )
}

export default OnHandReport
