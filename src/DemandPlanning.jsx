import { useState } from 'react'
import * as XLSX from 'xlsx'

const sampleData = [
  { SKU: 'LGT-1024', 'Product Name': 'LED Lamp', Category: 'Lighting', Month: 'Jan-2026', 'Sales Qty': 120 },
  { SKU: 'LGT-1024', 'Product Name': 'LED Lamp', Category: 'Lighting', Month: 'Feb-2026', 'Sales Qty': 145 },
  { SKU: 'LGT-1024', 'Product Name': 'LED Lamp', Category: 'Lighting', Month: 'Mar-2026', 'Sales Qty': 130 },
  { SKU: 'LGT-1189', 'Product Name': 'Panel Light', Category: 'Lighting', Month: 'Jan-2026', 'Sales Qty': 90 },
  { SKU: 'LGT-1189', 'Product Name': 'Panel Light', Category: 'Lighting', Month: 'Feb-2026', 'Sales Qty': 110 },
  { SKU: 'SAN-3310', 'Product Name': 'Wash Basin', Category: 'Sanitary', Month: 'Jan-2026', 'Sales Qty': 65 },
  { SKU: 'ELE-2208', 'Product Name': 'Electrical Cable', Category: 'Electrical', Month: 'Jan-2026', 'Sales Qty': 200 },
  { SKU: 'PLB-4417', 'Product Name': 'Water Mixer', Category: 'Plumbing', Month: 'Jan-2026', 'Sales Qty': 75 },
]

const sections = [
  { id: 'master', title: 'Master Data Sheet', short: 'Master Data', icon: '▦', description: 'Product and item planning attributes' },
  { id: 'sales', title: 'Sales History', short: 'Sales History', icon: '▥', description: 'Historical sales by SKU and period' },
  { id: 'forecast', title: 'Forecast', short: 'Forecast', icon: '▤', description: 'Demand forecast methods and results' },
  { id: 'pr', title: 'Purchase Requisition', short: 'PR', icon: '▧', description: 'Create and manage purchase requisitions' },
  { id: 'po', title: 'Purchase Orders', short: 'PO', icon: '▣', description: 'View purchase orders from ERP' },
]

const forecastMethods = [
  '3 Month Average', '6 Month Average', '12 Month Average',
  'Q1 Average', 'Q2 Average', 'Q3 Average', 'Q4 Average',
  'H1 Average', 'H2 Average', 'Min Max', 'Manual Forecast',
]

function DemandPlanning() {
  const [activeSection, setActiveSection] = useState('master')
  const [data, setData] = useState([])
  const [source, setSource] = useState('')
  const [fileName, setFileName] = useState('')
  const [error, setError] = useState('')
  const [search, setSearch] = useState('')
  const [forecastMethod, setForecastMethod] = useState('3 Month Average')
  const [applicationScope, setApplicationScope] = useState('All Items')
  const [selectedCategory, setSelectedCategory] = useState('All Categories')
  const [selectedSkus, setSelectedSkus] = useState([])
  const [prRows, setPrRows] = useState([])
  const [prMessage, setPrMessage] = useState('')

const [masterRows, setMasterRows] = useState([
  {
    SKU: 'LGT-1024',
    'Product Name': 'LED Lamp',
    'Item Group': 'Lighting',
    Category: 'Lighting',
    Subcategory: 'Indoor Lighting',
    Brand: 'Brand A',
    Supplier: 'Supplier A',
    UOM: 'PCS',
    Carton: 12,
    Box: 1,
    CBM: 0.05,
    Color: 'White',
    Size: 'Standard',
    Family: 'LED',
    Display: 'Yes',
    'Back Store': 'Yes',
    'AVG Sales': 15,
    DOH: 30,
    'ABC Pareto': 'A',
    'Turn Over': 1.5,
    'Turn Over Status': 'Slow Moving',
    'ON Hand':100

   
  },
  {
    SKU: 'SAN-3310',
    'Product Name': 'Wash Basin',
    'Item Group': 'Sanitary',
    Category: 'Sanitary',
    Subcategory: 'Basins',
    Brand: 'Brand B',
    Supplier: 'Supplier B',
    UOM: 'PCS',
    Carton: 1,
    Box: 1,
    CBM: 0.2,
    Color: 'White',
    Size: 'Medium',
    Family: 'Wash Basin',
    Display: 'Yes',
    'Back Store': 'No',
    'AVG Sales': 0,
    DOH: 0,
    'ABC Pareto': 'C',
    'Turn Over': 0,
    'Turn Over Status': 'Non Moving'
  }
])


const [editingSku, setEditingSku] = useState(null)
const [selectedMasterSkus, setSelectedMasterSkus] = useState([])
const [masterSearch, setMasterSearch] = useState('')
const [masterScope, setMasterScope] = useState('All Items')
const [masterCategory, setMasterCategory] = useState('All Categories')


  const categories = [...new Set(data.map((row) => row.Category).filter(Boolean))]
  const products = [...new Map(data.map((row) => [row.SKU, row])).values()]

  const scopedProducts = products.filter((product) => {
    if (applicationScope === 'All Items') return true
    if (applicationScope === 'By Category') return selectedCategory === 'All Categories' || product.Category === selectedCategory
    if (applicationScope === 'Multiple SKUs') return selectedSkus.includes(product.SKU)
    return true
  })

  const loadSample = () => {
    setData(sampleData)
    setSource('Sample Data')
    setFileName('SCOR Sample Data')
    setError('')
  }

  const uploadExcel = (event) => {
    const file = event.target.files?.[0]
    if (!file) return
    setError('')
    const reader = new FileReader()
    reader.onload = (e) => {
      try {
        const workbook = XLSX.read(e.target.result, { type: 'array' })
        const worksheet = workbook.Sheets[workbook.SheetNames[0]]
        const rows = XLSX.utils.sheet_to_json(worksheet)
        if (!rows.length) {
          setError('The Excel file is empty.')
          return
        }
        const required = ['SKU', 'Product Name', 'Category', 'Month', 'Sales Qty']
        const missing = required.filter((column) => !Object.keys(rows[0]).includes(column))
        if (missing.length) {
          setError(`Missing columns: ${missing.join(', ')}`)
          return
        }
        setData(rows)
        setSource('Excel Import')
        setFileName(file.name)
      } catch {
        setError('Unable to read this file. Please check the Excel format.')
      }
    }
    reader.readAsArrayBuffer(file)
  }

  const filteredData = data.filter((row) => {
    const q = search.toLowerCase()
    return [row.SKU, row['Product Name'], row.Category].some((v) => String(v ?? '').toLowerCase().includes(q))
  })

  const createPR = () => {
    const newRows = scopedProducts.map((product) => ({
      sku: product.SKU,
      name: product['Product Name'],
      category: product.Category,
      requestedQty: '',
      requiredDate: '',
      status: 'Draft',
    }))
    setPrRows(newRows)
    setPrMessage(`${newRows.length} PR line(s) prepared as draft. Enter quantities and required dates before saving.`)
  }

  const updatePRRow = (index, field, value) => {
    setPrRows((current) => current.map((row, i) => i === index ? { ...row, [field]: value } : row))
  }

  const current = sections.find((section) => section.id === activeSection)

  return (
    <div className="demand-planning-shell">
      <style>{`
        .demand-planning-shell{display:grid;grid-template-columns:205px minmax(0,1fr);gap:22px;color:#17324d;align-items:start}
        .dp-nav{background:#fff;border:1px solid #e1e8f0;border-radius:14px;padding:15px 12px;position:sticky;top:16px}
        .dp-nav-title{font-size:12px;font-weight:700;color:#8190a3;letter-spacing:.08em;text-transform:uppercase;padding:5px 10px 14px}
        .dp-nav-list{display:flex;flex-direction:column;position:relative}
        .dp-nav-list:before{content:"";position:absolute;left:24px;top:25px;bottom:25px;width:2px;background:#d9e7ec}
        .dp-nav-item{position:relative;display:flex;align-items:center;gap:11px;width:100%;padding:11px 9px;border:0;border-radius:9px;background:transparent;text-align:left;color:#52657a;cursor:pointer;z-index:1}
        .dp-nav-item:hover{background:#f2f7f9}
        .dp-nav-item.active{background:#e7f5f6;color:#087f8c;font-weight:700}
        .dp-nav-icon{width:31px;height:31px;flex:0 0 31px;display:grid;place-items:center;border:2px solid #d9e7ec;border-radius:50%;background:#fff;color:#567086;font-size:16px}
        .dp-nav-item.active .dp-nav-icon{border-color:#087f8c;background:#087f8c;color:#fff}
        .dp-nav-label{font-size:13px;line-height:1.25}
        .dp-content{min-width:0}
        .dp-heading{display:flex;justify-content:space-between;align-items:flex-start;gap:16px;margin-bottom:20px}
        .dp-heading h1{margin:0 0 6px;font-size:27px}.dp-heading p{margin:0;color:#64748b}
        .dp-card{background:#fff;border:1px solid #e1e8f0;border-radius:14px;padding:23px;box-shadow:0 4px 15px rgba(20,45,70,.04)}
        .dp-card h2{margin-top:0}
        .dp-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(210px,1fr));gap:15px}
        .dp-tile{border:1px solid #dbe4ee;border-radius:11px;padding:17px;background:#f8fafc}
        .dp-tile h3{margin:0 0 7px}.dp-tile p{font-size:13px;color:#64748b;min-height:36px}
        .dp-button{border:0;border-radius:8px;padding:10px 15px;background:#087f8c;color:#fff;cursor:pointer;font-weight:600}
        .dp-button.secondary{background:#17324d}.dp-button:disabled{opacity:.4;cursor:not-allowed}
        .dp-error{margin-top:14px;padding:12px;border-radius:8px;background:#fff1f2;color:#be123c}
        .dp-note{padding:14px;border-radius:9px;background:#f1f5f9;color:#475569;margin-top:15px;font-size:13px;line-height:1.65}
        .dp-table-wrap{overflow-x:auto;margin-top:16px}.dp-table{width:100%;border-collapse:collapse;font-size:13px}
        .dp-table th,.dp-table td{padding:10px;border-bottom:1px solid #e5eaf0;text-align:left;white-space:nowrap}
        .dp-table th{background:#f1f5f9}
        .dp-search,.dp-select,.dp-input{padding:9px 11px;border:1px solid #dbe4ee;border-radius:8px;background:#fff;color:#17324d;max-width:100%}
        .dp-toolbar{display:flex;gap:12px;align-items:center;flex-wrap:wrap;margin:15px 0}
        .dp-field{display:flex;flex-direction:column;gap:6px;font-size:13px;color:#52657a}
        .dp-badge{display:inline-block;border-radius:20px;padding:5px 9px;background:#e6f5f6;color:#087f8c;font-size:12px}
        @media(max-width:760px){.demand-planning-shell{grid-template-columns:1fr;gap:14px}.dp-nav{position:static}.dp-nav-list{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:4px}.dp-nav-list:before{display:none}.dp-nav-item{padding:8px 5px}.dp-nav-label{font-size:12px}.dp-heading h1{font-size:23px}}
      `}</style>

      <aside className="dp-nav">
        <div className="dp-nav-title">Demand Planning</div>
        <nav className="dp-nav-list" aria-label="Demand Planning sections">
          {sections.map((section) => (
            <button
              key={section.id}
              className={`dp-nav-item ${activeSection === section.id ? 'active' : ''}`}
              onClick={() => setActiveSection(section.id)}
              title={section.title}
            >
              <span className="dp-nav-icon">{section.icon}</span>
              <span className="dp-nav-label">{section.short}</span>
            </button>
          ))}
        </nav>
      </aside>

      <main className="dp-content">
        <div className="dp-heading">
          <div>
            <h1>{current?.title}</h1>
            <p>{current?.description}</p>
          </div>
          <span className="dp-badge">Flexible navigation</span>
        </div>

{activeSection === 'master' && (
  <section className="dp-card master-data-page">
    <div className="master-toolbar">
      <div>
        <h2>Released Products / Master Data</h2>
        <p>Manage product planning attributes</p>
      </div>

      <div className="master-actions">
        <button className="dp-button" onClick={() => {
const newRow = {
  SKU: `NEW-${Date.now()}`,
  'Product Name': '',
  'Item Group': '',
  Category: '',
  Subcategory: '',
  Brand: '',
  Supplier: '',
  UOM: 'PCS',
  Carton: 1,
  Box: 1,
  CBM: 0,
  Color: '',
  Size: '',
  Family: '',
  Display: 'No',
  'Back Store': 'No',
  'AVG Sales': 0,
  DOH: 0,
  'ABC Pareto': 'C',
  'Turn Over': 0,
  'Turn Over Status': 'Non Moving',
  'ON Hand': 0
}
          setMasterRows(current => [...current, newRow])
          setEditingSku(newRow.SKU)
        }}>
          + New
        </button>

        <button className="dp-button secondary" onClick={() => {
          setMasterRows(current =>
            current.filter(row => !selectedMasterSkus.includes(row.SKU))
          )
          setSelectedMasterSkus([])
        }}>
          Delete
        </button>

        <label className="dp-button secondary">
          Import Excel
          <input
            type="file"
            accept=".xlsx,.xls"
            hidden
            onChange={async (e) => {
              const file = e.target.files[0]
              if (!file) return

              try {
                const buffer = await file.arrayBuffer()
                const workbook = XLSX.read(buffer)
                const sheet = workbook.Sheets[workbook.SheetNames[0]]
                const importedRows = XLSX.utils.sheet_to_json(sheet)

                setMasterRows(current => {
                  const combined = [...current]
                  importedRows.forEach(row => {
                    const index = combined.findIndex(item => item.SKU === row.SKU)
                    if (index >= 0) combined[index] = { ...combined[index], ...row }
                    else combined.push(row)
                  })
                  return combined
                })
              } catch {
                alert('Unable to read this Excel file.')
              }

              e.target.value = ''
            }}
          />
        </label>

        <button className="dp-button secondary" onClick={() => {
          const worksheet = XLSX.utils.json_to_sheet(masterRows)
          const workbook = XLSX.utils.book_new()
          XLSX.utils.book_append_sheet(workbook, worksheet, 'Master Data')
          XLSX.writeFile(workbook, 'SCOR_Master_Data.xlsx')
        }}>
          Export Excel
        </button>
      </div>
    </div>

    <div className="dp-toolbar">
      <input
        className="dp-search"
        placeholder="Search SKU, Product, Brand..."
        value={masterSearch}
        onChange={e => setMasterSearch(e.target.value)}
      />

      <select
        className="dp-select"
        value={masterScope}
        onChange={e => setMasterScope(e.target.value)}
      >
        <option>All Items</option>
        <option>By Category</option>
        <option>Multiple SKUs</option>
      </select>

      {masterScope === 'By Category' && (
        <select
          className="dp-select"
          value={masterCategory}
          onChange={e => setMasterCategory(e.target.value)}
        >
          <option>All Categories</option>
          {[...new Set(masterRows.map(row => row.Category).filter(Boolean))].map(category => (
            <option key={category}>{category}</option>
          ))}
        </select>
      )}
    </div>

    <div className="dp-table-wrap">
      <table className="dp-table">
        <thead>
          <tr>
            <th>Select</th>
            {Object.keys(masterRows[0] || {}).map(field => (
              <th key={field}>{field}</th>
            ))}
            <th>Action</th>
          </tr>
        </thead>

        <tbody>
          {masterRows
            .filter(row => {
              const searchMatch = Object.values(row).some(value =>
                String(value).toLowerCase().includes(masterSearch.toLowerCase())
              )
              const categoryMatch =
                masterScope !== 'By Category' ||
                masterCategory === 'All Categories' ||
                row.Category === masterCategory
              const skuMatch =
                masterScope !== 'Multiple SKUs' ||
                selectedMasterSkus.includes(row.SKU)

              return searchMatch && categoryMatch && skuMatch
            })
            .map(row => (
              <tr key={row.SKU}>
                <td>
                  <input
                    type="checkbox"
                    checked={selectedMasterSkus.includes(row.SKU)}
                    onChange={e => setSelectedMasterSkus(current =>
                      e.target.checked
                        ? [...current, row.SKU]
                        : current.filter(sku => sku !== row.SKU)
                    )}
                  />
                </td>

                {Object.keys(row).map(field => (

<td key={field}>
  {field === 'DOH' ? (
    Number(row['AVG Sales']) > 0
      ? ((Number(row['ON Hand'] || 0) / Number(row['AVG Sales'])) * 30).toFixed(1)
      : 0
  ) : editingSku === row.SKU ? (
    <input
      className="dp-input"
      value={row[field] ?? ''}
      onChange={e => setMasterRows(current =>
        current.map(item =>
          item.SKU === row.SKU
            ? { ...item, [field]: e.target.value }
            : item
        )
      )}
    />
  ) : (
    row[field]
  )}
</td>

                ))}

                <td>
                  {editingSku === row.SKU ? (
                    <button className="dp-button" onClick={() => setEditingSku(null)}>
                      Save
                    </button>
                  ) : (
                    <button className="dp-button secondary" onClick={() => setEditingSku(row.SKU)}>
                      Edit
                    </button>
                  )}
                </td>
              </tr>
            ))}
        </tbody>
      </table>
    </div>

    <p className="master-count">
      Total Items: {masterRows.length}
    </p>

    <div className="dp-note">
      Data sources: Manual Entry and Excel Import.
      External ERP integration will be supported in a later phase.
    </div>
  </section>
)}

        {activeSection === 'sales' && (
          <section className="dp-card">
            <h2>Sales History</h2>
            <p>Load, search and review historical sales. You can move to any other section at any time.</p>
            <div className="dp-grid">
              <div className="dp-tile">
                <h3>Excel Import</h3>
                <p>Upload historical sales using the SCOR column format.</p>
                <label className="dp-button" style={{ display: 'inline-block' }}>
                  Choose Excel File
                  <input type="file" accept=".xlsx,.xls" onChange={uploadExcel} hidden />
                </label>
              </div>
              <div className="dp-tile">
                <h3>Sample Data</h3>
                <p>Load sample products and monthly sales to explore the module.</p>
                <button className="dp-button secondary" onClick={loadSample}>Load Sample Data</button>
              </div>
            </div>
            {error && <div className="dp-error">{error}</div>}
            {data.length > 0 && <>
              <div className="dp-toolbar">
                <span><strong>{data.length}</strong> records · {source} · {fileName}</span>
                <input className="dp-search" value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Search SKU, product or category..." />
              </div>
              <div className="dp-table-wrap"><table className="dp-table">
                <thead><tr>{Object.keys(data[0]).map((col) => <th key={col}>{col}</th>)}</tr></thead>
                <tbody>{filteredData.map((row, i) => <tr key={`${row.SKU}-${row.Month}-${i}`}>{Object.keys(data[0]).map((col) => <td key={col}>{row[col]}</td>)}</tr>)}</tbody>
              </table></div>
              <small>Showing {filteredData.length} of {data.length} records.</small>
            </>}
          </section>
        )}

        {activeSection === 'forecast' && (
          <section className="dp-card">
            <h2>Forecast Workspace</h2>
            <p>Configure forecast methods at the scope you choose. No mandatory sequence: open any section whenever you need it.</p>
            {data.length === 0 && <div className="dp-note">Load sales history first from the Sales History icon. You can return here at any time.</div>}
            <div className="dp-toolbar">
              <label className="dp-field">Application Scope
                <select className="dp-select" value={applicationScope} onChange={(e) => setApplicationScope(e.target.value)}>
                  <option>All Items</option><option>By Category</option><option>Multiple SKUs</option>
                </select>
              </label>
              {applicationScope === 'By Category' && <label className="dp-field">Category
                <select className="dp-select" value={selectedCategory} onChange={(e) => setSelectedCategory(e.target.value)}>
                  <option>All Categories</option>{categories.map((cat) => <option key={cat}>{cat}</option>)}
                </select>
              </label>}
            </div>
            {applicationScope === 'Multiple SKUs' && (
              <div className="dp-grid">
                {products.map((product) => <label key={product.SKU} className="dp-tile">
                  <input type="checkbox" checked={selectedSkus.includes(product.SKU)} onChange={(e) => setSelectedSkus((old) => e.target.checked ? [...old, product.SKU] : old.filter((sku) => sku !== product.SKU))} />
                  {' '}{product.SKU} — {product['Product Name']}
                </label>)}
              </div>
            )}
            <div className="dp-toolbar">
              <label className="dp-field">Default Forecast Method
                <select className="dp-select" value={forecastMethod} onChange={(e) => setForecastMethod(e.target.value)}>
                  {forecastMethods.map((method) => <option key={method}>{method}</option>)}
                </select>
              </label>
              <span className="dp-badge">{scopedProducts.length} selected item(s)</span>
            </div>
            <div className="dp-table-wrap"><table className="dp-table">
              <thead><tr><th>SKU</th><th>Product Name</th><th>Category</th><th>Forecast Method</th></tr></thead>
              <tbody>{scopedProducts.map((product) => <tr key={product.SKU}><td>{product.SKU}</td><td>{product['Product Name']}</td><td>{product.Category}</td><td>{forecastMethod}</td></tr>)}</tbody>
            </table></div>
            <div className="dp-note">Quarterly averages will compare the same quarter across previous years to reflect seasonality. Forecast calculations and result publishing will be implemented in the next development step.</div>
          </section>
        )}

        {activeSection === 'pr' && (
          <section className="dp-card">
            <h2>Purchase Requisition</h2>
            <p>SCOR creates PRs. Purchase Orders remain read-only and come from the ERP in a later integration phase.</p>
            <div className="dp-toolbar">
              <label className="dp-field">PR item scope
                <select className="dp-select" value={applicationScope} onChange={(e) => setApplicationScope(e.target.value)}>
                  <option>All Items</option><option>By Category</option><option>Multiple SKUs</option>
                </select>
              </label>
              {applicationScope === 'By Category' && <label className="dp-field">Category
                <select className="dp-select" value={selectedCategory} onChange={(e) => setSelectedCategory(e.target.value)}>
                  <option>All Categories</option>{categories.map((cat) => <option key={cat}>{cat}</option>)}
                </select>
              </label>}
            </div>
            {applicationScope === 'Multiple SKUs' && <div className="dp-grid">
              {products.map((product) => <label className="dp-tile" key={product.SKU}>
                <input type="checkbox" checked={selectedSkus.includes(product.SKU)} onChange={(e) => setSelectedSkus((old) => e.target.checked ? [...old, product.SKU] : old.filter((sku) => sku !== product.SKU))} />
                {' '}{product.SKU} — {product['Product Name']}
              </label>)}
            </div>}
            <div className="dp-toolbar">
              <button className="dp-button" onClick={createPR} disabled={scopedProducts.length === 0}>Create PR Draft</button>
              <span className="dp-badge">{scopedProducts.length} item(s) in scope</span>
            </div>
            {prMessage && <div className="dp-note">{prMessage}</div>}
            {prRows.length > 0 && <div className="dp-table-wrap"><table className="dp-table">
              <thead><tr><th>SKU</th><th>Product</th><th>Category</th><th>Requested Qty</th><th>Required Date</th><th>Status</th></tr></thead>
              <tbody>{prRows.map((row, i) => <tr key={row.sku}>
                <td>{row.sku}</td><td>{row.name}</td><td>{row.category}</td>
                <td><input className="dp-input" type="number" min="0" value={row.requestedQty} onChange={(e) => updatePRRow(i, 'requestedQty', e.target.value)} placeholder="Qty" /></td>
                <td><input className="dp-input" type="date" value={row.requiredDate} onChange={(e) => updatePRRow(i, 'requiredDate', e.target.value)} /></td>
                <td><span className="dp-badge">{row.status}</span></td>
              </tr>)}</tbody>
            </table></div>}
            <div className="dp-note">This first version prepares editable PR draft lines in the browser. Saving, approval routing and ERP submission are not connected yet.</div>
          </section>
        )}

        {activeSection === 'po' && (
          <section className="dp-card">
            <h2>Purchase Orders</h2>
            <p>Read-only workspace for purchase orders received from an external ERP.</p>
            <div className="dp-note">SCOR will not create or edit POs. The future ERP connection will populate PO Number, Supplier, SKU, Ordered Qty, Received Qty, Expected Delivery Date and PO Status here.</div>
            <span className="dp-badge">ERP integration — later phase</span>
          </section>
        )}
      </main>
    </div>
  )
}

export default DemandPlanning
