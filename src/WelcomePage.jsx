import { ResponsiveContainer, PieChart, Pie, Cell, Tooltip } from 'recharts'
import {
  Truck, Upload, CircleCheck, Clock3, AlertTriangle, PackageCheck, ArrowRight,
} from 'lucide-react'
import './Welcome.css'

/*
  Welcome page (self-contained, fits on one screen).
  Props:
    fileName  - name of the uploaded file ('' when nothing uploaded)
    loading   - true while the file is being read
    error     - error message text
    allRows   - all prepared rows (Waiting To Load + Loading), each with _kpi and _group
    stats     - { waiting: { total }, loading: { total } }  (numbers shown on the two buttons)
    onUpload  - file input onChange handler
    onOpen    - called with 'waiting' or 'loading' when a button is clicked
*/

const COLORS = {
  'Within KPI': '#16A34A',
  Exception: '#E9A23B',
  'Major Exception': '#DC4545',
}

const GROUPS = [
  { id: 'drc', label: 'DRC' },
  { id: 'reload', label: 'Reload Warehouses' },
  { id: 'third', label: 'Third-Party Other Pickup Points' },
]

const STATUSES = ['Within KPI', 'Exception', 'Major Exception']

function StatCard({ icon: Icon, title, value, note, tone }) {
  return (
    <article className={`ws-card ws-card-${tone}`}>
      <span className="ws-card-icon"><Icon size={18} /></span>
      <div>
        <p>{title}</p>
        <strong>{value.toLocaleString()}</strong>
        <small>{note}</small>
      </div>
    </article>
  )
}

function GroupBars({ title, rows }) {
  const valid = rows.filter(row => row._kpi !== 'Date Missing')
  const total = valid.length
  return (
    <div className="ws-group">
      <h3 title={title}>{title}</h3>
      <span className="ws-group-total">{total.toLocaleString()} trucks</span>
      {STATUSES.map(status => {
        const count = valid.filter(row => row._kpi === status).length
        const percent = total ? (count / total) * 100 : 0
        return (
          <div className="ws-bar-row" key={status}>
            <div className="ws-bar-label"><span>{status}</span><strong>{count.toLocaleString()} · {percent.toFixed(1)}%</strong></div>
            <div className="ws-bar-track"><div className="ws-bar-fill" style={{ width: `${percent}%`, background: COLORS[status] }} /></div>
          </div>
        )
      })}
    </div>
  )
}

export default function WelcomePage({ fileName, loading, error, allRows, stats, onUpload, onOpen }) {
  const hasFile = Boolean(fileName)
  const counts = {
    total: allRows.length,
    within: allRows.filter(row => row._kpi === 'Within KPI').length,
    exception: allRows.filter(row => row._kpi === 'Exception').length,
    major: allRows.filter(row => row._kpi === 'Major Exception').length,
  }
  const pieData = STATUSES.map(name => ({
    name,
    value: name === 'Within KPI' ? counts.within : name === 'Exception' ? counts.exception : counts.major,
  }))

  return (
    <div className="welcome-page">
      <div className={`welcome-inner ${hasFile ? 'has-file' : ''}`}>
        <div className="welcome-top">
          <header className="welcome-hero">
            <div className="welcome-logo"><Truck size={hasFile ? 24 : 30} /></div>
            <div className="welcome-hero-text">
              <p className="eyebrow">AFRICA OPERATIONS · KPI MONITORING</p>
              <h1>Welcome to Logistics</h1>
              <p className="welcome-sub">Upload your operations file once, review the overall KPI board, then open the Waiting To Load or Loading dashboard.</p>
            </div>
          </header>

          <section className="welcome-upload">
            <div className="welcome-upload-text">
              <h2>Upload your file</h2>
              <p>Excel or CSV (.xlsx, .xls, .csv). Upload once — switching pages keeps your data.</p>
            </div>
            <label className="welcome-upload-button">
              <Upload size={17} />
              {loading ? 'Reading file...' : hasFile ? 'Upload a new file' : 'Choose file'}
              <input type="file" accept=".xlsx,.xls,.csv" hidden disabled={loading} onChange={onUpload} />
            </label>
            <div className="welcome-file">
              <span className={`status-dot ${hasFile ? 'status-dot-ready' : ''}`} />
              <strong>{fileName || 'No file uploaded yet'}</strong>
            </div>
          </section>
        </div>

        {error && <div className="error-banner"><AlertTriangle size={18} /><span>{error}</span></div>}

        {hasFile && (
          <>
            <section className="ws-cards">
              <StatCard icon={Truck} title="All Truck Performance" value={counts.total} note="Waiting To Load + Loading" tone="navy" />
              <StatCard icon={CircleCheck} title="Within KPI" value={counts.within} note="Within target" tone="green" />
              <StatCard icon={Clock3} title="Exception" value={counts.exception} note="Needs attention" tone="orange" />
              <StatCard icon={AlertTriangle} title="Major Exception" value={counts.major} note="Action required" tone="red" />
            </section>

            <section className="ws-middle">
              <article className="ws-panel ws-pie">
                <h2>Overall Performance</h2>
                <div className="ws-pie-wrap">
                  <ResponsiveContainer width="100%" height="100%">
                    <PieChart>
                      <Pie data={pieData} dataKey="value" nameKey="name" cx="50%" cy="50%" innerRadius="62%" outerRadius="92%" paddingAngle={3} stroke="none">
                        {pieData.map(item => <Cell key={item.name} fill={COLORS[item.name]} />)}
                      </Pie>
                      <Tooltip formatter={(value, name) => [value, name]} />
                    </PieChart>
                  </ResponsiveContainer>
                  <div className="ws-pie-center"><strong>{counts.total.toLocaleString()}</strong><span>Total</span></div>
                </div>
                <div className="ws-legend">
                  {pieData.map(item => (
                    <span key={item.name}><i style={{ background: COLORS[item.name] }} />{counts.total ? ((item.value / counts.total) * 100).toFixed(1) : '0.0'}%</span>
                  ))}
                </div>
              </article>

              <article className="ws-panel ws-groups" aria-label="KPI performance by pickup group">
                <h2>KPI by Pickup Group</h2>
                <div className="ws-groups-grid">
                  {GROUPS.map(group => (
                    <GroupBars key={group.id} title={group.label} rows={allRows.filter(row => row._group === group.id)} />
                  ))}
                </div>
              </article>
            </section>

            <section className="welcome-actions">
              <button type="button" className="welcome-open-button" onClick={() => onOpen('waiting')}>
                <Clock3 size={18} /> Waiting To Load <span className="welcome-count">{stats.waiting.total.toLocaleString()}</span> <ArrowRight size={17} />
              </button>
              <button type="button" className="welcome-open-button welcome-open-loading" onClick={() => onOpen('loading')}>
                <PackageCheck size={18} /> Loading <span className="welcome-count">{stats.loading.total.toLocaleString()}</span> <ArrowRight size={17} />
              </button>
            </section>
          </>
        )}
      </div>
    </div>
  )
}
