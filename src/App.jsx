
import { useState } from 'react'
import './App.css'
import DemandPlanning from './DemandPlanning'
import ReportsIndex from './reports/ReportsIndex'
import NpiReport from './reports/NpiReport'
import OnHandReport from './reports/OnHandReport'

function App() {
  const [sidebarOpen, setSidebarOpen] = useState(true)
  const [activePage, setActivePage] = useState('Dashboard')
  const [activeReport, setActiveReport] = useState(null)
  const [language, setLanguage] = useState('en')
  const [period, setPeriod] = useState('Oct 2026')
  const [region, setRegion] = useState('All Regions')
  const [category, setCategory] = useState('All Categories')
function changeLanguage(lang) {
  const select = document.querySelector(".goog-te-combo");

  if (select) {
    select.value = lang;
    select.dispatchEvent(new Event("change"));
  }
}
  const menuItems = [
    'Dashboard',
    'Demand Planning',
    'Inventory Management',
    'Replenishment',
    'Warehouse Transfers',
    'S&OP',
    'Reports',
    'Master Data',
    'Settings'
  ]

  const kpis = [
    { title: 'Total Inventory', value: '12,450', change: '+2.8%', icon: '▦', color: 'teal' },
    { title: 'Forecast Accuracy', value: '85%', change: '+8 pp', icon: '◎', color: 'green' },
    { title: 'Service Level', value: '80%', change: '+5 pp', icon: '⇧', color: 'blue' },
    { title: 'Stock Turnover', value: '5.6x', change: '+0.6x', icon: '⟳', color: 'purple' },
    { title: 'Days on Hand (DOH)', value: '52', change: '-8 days', icon: '▦', color: 'orange' },
    { title: 'OOS %', value: '5.2%', change: '-1.8 pp', icon: '⚠', color: 'red' }
  ]

  const months = ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct']

  const categories = [
    { name: 'Lighting', value: 88 },
    { name: 'Sanitary', value: 85 },
    { name: 'Electrical', value: 83 },
    { name: 'Plumbing', value: 82 },
    { name: 'Others', value: 80 }
  ]

  const slowMoving = [
    { sku: 'LGT-1024', category: 'Lighting', stock: '1,250', doh: 180 },
    { sku: 'SAN-3310', category: 'Sanitary', stock: '840', doh: 150 },
    { sku: 'ELE-2208', category: 'Electrical', stock: '620', doh: 140 },
    { sku: 'PLB-4417', category: 'Plumbing', stock: '510', doh: 135 },
    { sku: 'LGT-1189', category: 'Lighting', stock: '480', doh: 120 }
  ]

  const inventory = [
    { name: 'In Stock', value: '10,680', percent: '85.8%', color: 'green' },
    { name: 'Low Stock', value: '1,125', percent: '9.0%', color: 'orange' },
    { name: 'Out of Stock', value: '645', percent: '5.2%', color: 'red' }
  ]

  const alerts = [
    { type: 'danger', icon: '⚠', text: '125 items below Min level' },
    { type: 'warning', icon: '⚠', text: '48 pending warehouse transfers' },
    { type: 'info', icon: 'ⓘ', text: '20 items with high forecast variance' }
  ]

  return (
<div
  className="app"
dir="ltr"
  lang={language}
>
      {/* Sidebar */}
<aside className={`sidebar ${sidebarOpen ? '' : 'sidebar-collapsed'}`}>   
  <button
  type="button"
  className="sidebar-toggle"
  onClick={() => setSidebarOpen(!sidebarOpen)}
  title={sidebarOpen ? 'Collapse Menu' : 'Expand Menu'}
>
  {sidebarOpen ? '❮' : '❯'}
</button>

       <div className="logo">
<img
  src="/scor-logo.png"
  alt="SCOR Logo"
  className="sidebar-logo"
/>
        </div>

        <nav>
          {menuItems.map((item, index) => (
            <div
              key={item}
              className={activePage === item ? 'active' : ''}
              onClick={() => setActivePage(item)}
            >
              <span className="menu-icon">
                {['⌂', '▥', '▦', '⟳', '⇄', '◎', '▤', '▣', '⚙'][index]}
              </span>
              {item}
            </div>
          ))}
        </nav>
      </aside>

      {/* Main Content */}
      <main className="main-content">

        {/* Topbar */}
        
<header className="topbar">
  <input
    type="text"
    placeholder="Search for SKU, Product, Category..."
  />

  <div className="user-info">

<div className="language-switcher">
  <button
    type="button"
    onClick={() => changeLanguage('en')}
    title="English"
  >
    <img src="https://flagcdn.com/w40/gb.png" alt="English" />
  </button>

  <button
    type="button"
    onClick={() => changeLanguage('ar')}
    title="العربية"
  >
    <img src="https://flagcdn.com/w40/sa.png" alt="Arabic" />
  </button>

  <div
    id="google_translate_element"
    className="hidden-google-translate"
  ></div>
</div>

    <span>🔔</span>
  </div>
</header>
        {activePage === 'Dashboard' ? (
          <>
            {/* Page Header */}
            <section className="page-header">
              <div>
                <h1>Supply Chain Dashboard</h1>
                <p>Planning, Inventory and Replenishment Overview</p>
              </div>

              <div className="filters">
                <select value={period} onChange={(e) => setPeriod(e.target.value)}>
                  <option>Oct 2026</option>
                  <option>Sep 2026</option>
                  <option>Aug 2026</option>
                </select>

                <select value={region} onChange={(e) => setRegion(e.target.value)}>
                  <option>All Regions</option>
                  <option>Riyadh</option>
                  <option>Jeddah</option>
                  <option>Najran</option>
                  <option>Jazan</option>
                  <option>Khamis</option>
                </select>

                <select value={category} onChange={(e) => setCategory(e.target.value)}>
                  <option>All Categories</option>
                  <option>Lighting</option>
                  <option>Sanitary</option>
                  <option>Electrical</option>
                  <option>Plumbing</option>
                </select>
              </div>
            </section>

            {/* KPI Cards */}
            <section className="cards">
              {kpis.map((kpi) => (
                <div className="card" key={kpi.title}>
                  <div className={`kpi-icon ${kpi.color}`}>
                    {kpi.icon}
                  </div>

                  <div className="kpi-content">
                    <h3>{kpi.title}</h3>
                    <h2>{kpi.value}</h2>
                    <p>{kpi.change} vs previous period</p>
                  </div>
                </div>
              ))}
            </section>

            {/* Main Analytics */}
            <section className="analytics">

              {/* Sales vs Forecast */}
              <div className="panel sales-panel">
                <h2>Sales vs Forecast Trend</h2>

                <div className="chart-legend">
                  <span><i className="actual-dot"></i> Actual Sales</span>
                  <span><i className="forecast-dot"></i> Forecast</span>
                </div>

                <div className="line-chart">
                  <svg viewBox="0 0 700 230" preserveAspectRatio="none">
                    {[40, 85, 130, 175, 215].map((y) => (
                      <line
                        key={y}
                        x1="35"
                        y1={y}
                        x2="690"
                        y2={y}
                        stroke="#E2E8F0"
                        strokeWidth="1"
                      />
                    ))}

                    <polyline
                      points="45,160 145,145 245,130 345,115 445,95 545,75 645,55"
                      fill="none"
                      stroke="#087E8B"
                      strokeWidth="4"
                      strokeDasharray="8 6"
                    />

                    <polyline
                      points="45,175 145,165 245,150 345,135 445,120 545,100 645,80"
                      fill="none"
                      stroke="#102A43"
                      strokeWidth="4"
                    />

                    {[45, 145, 245, 345, 445, 545, 645].map((x, i) => (
                      <g key={x}>
                        <circle
                          cx={x}
                          cy={[175, 165, 150, 135, 120, 100, 80][i]}
                          r="5"
                          fill="white"
                          stroke="#102A43"
                          strokeWidth="3"
                        />
                        <circle
                          cx={x}
                          cy={[160, 145, 130, 115, 95, 75, 55][i]}
                          r="5"
                          fill="white"
                          stroke="#087E8B"
                          strokeWidth="3"
                        />
                      </g>
                    ))}
                  </svg>

                  <div className="chart-months">
                    {months.map((month) => (
                      <span key={month}>{month}</span>
                    ))}
                  </div>
                </div>
              </div>

              {/* Inventory Status */}
              <div className="panel inventory-panel">
                <h2>Inventory by Status</h2>

                <div className="inventory-content">
                  <div className="donut-chart">
                    <div className="donut-center">
                      <strong>12,450</strong>
                      <span>Units</span>
                    </div>
                  </div>

                  <div className="inventory-status">
                    {inventory.map((item) => (
                      <div key={item.name}>
                        <span className={`status-dot ${item.color}`}></span>
                        <span>{item.name}</span>
                        <strong>{item.value}</strong>
                        <small>{item.percent}</small>
                      </div>
                    ))}
                  </div>
                </div>
              </div>
            </section>

            {/* Bottom Analytics */}
            <section className="bottom-panels">

              {/* Forecast Accuracy */}
              <div className="panel">
                <h2>Forecast Accuracy by Category</h2>

                {categories.map((item) => (
                  <div className="progress-row" key={item.name}>
                    <span>{item.name}</span>

                    <div className="progress-track">
                      <div
                        className="progress-fill"
                        style={{ width: `${item.value}%` }}
                      ></div>
                    </div>

                    <strong>{item.value}%</strong>
                  </div>
                ))}
              </div>

              {/* Slow Moving Inventory */}
              <div className="panel slow-panel">
                <h2>Slow Moving Inventory (Top 5)</h2>

                <table>
                  <thead>
                    <tr>
                      <th>SKU</th>
                      <th>Category</th>
                      <th>On Hand</th>
                      <th>DOH</th>
                    </tr>
                  </thead>

                  <tbody>
                    {slowMoving.map((item) => (
                      <tr key={item.sku}>
                        <td>{item.sku}</td>
                        <td>{item.category}</td>
                        <td>{item.stock}</td>
                        <td>{item.doh}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              {/* Alerts */}
              <div className="panel alerts-panel">
                <h2>Alerts & Actions</h2>

                {alerts.map((alert) => (
                  <div className={`alert ${alert.type}`} key={alert.text}>
                    <span>{alert.icon}</span>
                    {alert.text}
                    <strong>›</strong>
                  </div>
                ))}
              </div>

            </section>
          </>

) : activePage === 'Demand Planning' ? (
  <DemandPlanning />
) : activePage === 'Reports' && activeReport === 3 ? (
  <OnHandReport
    onBack={() => setActiveReport(null)}
  />
) : activePage === 'Reports' ? (
  <ReportsIndex
    onOpenReport={(reportId) => setActiveReport(reportId)}
  />
) : (
  <section className="module-placeholder">
    <h1>{activePage}</h1>
    <p>This module is under development.</p>
  </section>
)}

      </main>
    </div>
  )
}

export default App
