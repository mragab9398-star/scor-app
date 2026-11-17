
const reports = [
  {
    id: 1,
    name: 'NPI Report',
    description: 'New Product Introduction analysis',
    category: 'Product Analysis',
    icon: '📦',
    status: 'Available',
  },
  {
    id: 2,
    name: 'Sales Report',
    description: 'Sales performance and trends',
    category: 'Sales',
    icon: '📈',
    status: 'Coming Soon',
  },
  {
    id: 3,
    name: 'Inventory Report',
    description: 'Stock levels and inventory health',
    category: 'Inventory',
    icon: '🏬',
    status: 'Available',
  },
  {
    id: 4,
    name: 'Transfer Report',
    description: 'Warehouse and branch transfers',
    category: 'Operations',
    icon: '🚚',
    status: 'Coming Soon',
  },
  {
    id: 5,
    name: 'Aging Report',
    description: 'Slow-moving and non-moving items',
    category: 'Inventory',
    icon: '⏳',
    status: 'Coming Soon',
  },
]

function ReportsIndex({ onOpenReport }) {
  return (
    <div className="reports-page">
      <div className="reports-header">
        <div>
          <h1>Reports Center</h1>
          <p>Choose a report to view its details and analysis.</p>
        </div>
        <div className="reports-count">
          {reports.length} Reports
        </div>
      </div>

      <div className="reports-grid">
        {reports.map((report) => (
          <div className="report-card" key={report.id}>
            <div className="report-card-top">
              <span className="report-icon">{report.icon}</span>
              <span className={`report-status ${report.status === 'Available' ? 'available' : ''}`}>
                {report.status}
              </span>
            </div>

            <span className="report-category">{report.category}</span>
            <h3>{report.name}</h3>
            <p>{report.description}</p>

            <button
              type="button"
              disabled={report.status !== 'Available'}
              onClick={() => onOpenReport(report.id)}
            >
              Open Report →
            </button>
          </div>
        ))}
      </div>
    </div>
  )
}

export default ReportsIndex
    