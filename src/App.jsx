import { useEffect, useMemo, useState } from 'react'
import html2canvas from 'html2canvas'
import * as XLSX from 'xlsx'
import { PieChart, Pie, Cell, Tooltip, ResponsiveContainer } from 'recharts'
import {
  LayoutDashboard, Truck, Upload, Search, FileSpreadsheet,
  AlertTriangle, CircleCheck, Clock3, Eye, Printer, Camera, ChevronDown,
  Menu, X, RefreshCw,
} from 'lucide-react'
import './App.css'
import './App.override.css'

const COLORS = {
  'Within KPI': '#16A34A',
  Exception: '#E9A23B',
  'Major Exception': '#DC4545',
  'Date Missing': '#94A3B8',
}

const GROUPS = [
  { id: 'drc', label: 'DRC' },
  { id: 'reload', label: 'Pickup Points Using Reload' },
  { id: 'third', label: 'Third Party / Other Pickup Points' },
]

const EXCEL_COLUMNS = [
  'Client: Account Name',
  'Depart from Country',
  'Pickup Point: Location Name',
  'Delivery Number',
  'Horse: Asset No [Reg]',
  'Delivery Status',
  'Reporting Date',
  'Loading Date',
  'Document Received Date',
  'Transporter: Account Name',
  'Last Event: Internal Comments',
  'Haulier: Account Name',
]

/*
  KPI RULES (Waiting days = Today - Reporting Date)
  within    -> days from 0 up to and including `within`
  exception -> days above `within` up to and including `exception`
  major     -> days above `exception`

  Waiting to Load
    DRC     : 0-3  | 4-5   | >5
    Reload  : 0-1  | 2     | >2
    Other   : 0-2  | 3     | >3
  Waiting for Documents
    DRC     : 0-7  | 8-12  | >12
    Reload  : 0-2  | 3     | >3
    Other   : 0-3  | 4-5   | >5
*/
const KPI_RULES = {
  'waiting to load': {
    drc: { within: 3, exception: 5 },
    reload: { within: 1, exception: 2 },
    third: { within: 2, exception: 3 },
  },
  'waiting for documents': {
    drc: { within: 7, exception: 12 },
    reload: { within: 2, exception: 3 },
    third: { within: 3, exception: 5 },
  },
}

function clean(value) {
  return String(value ?? '').trim()
}

async function handleScreenshot() {
  const dashboard = document.querySelector('.dashboard-content')

  if (!dashboard) {
    alert('Dashboard nahi mila!')
    return
  }

  try {
    const canvas = await html2canvas(dashboard, {
      scale: 2,
      useCORS: true,
      backgroundColor: '#f5f7fb',
      logging: false,
      windowWidth: Math.max(document.documentElement.clientWidth, dashboard.scrollWidth),
    })

    const link = document.createElement('a')
    link.download = `Africa-Operations-${new Date().toISOString().slice(0, 10)}.png`
    link.href = canvas.toDataURL('image/png')
    link.click()
  } catch (error) {
    console.error('Screenshot error:', error)
    alert('Screenshot nahi ban saka. Browser Console check karo.')
  }
}

function getDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate())
  }
  if (typeof value === 'number' && value > 0) {
    const parsed = XLSX.SSF.parse_date_code(value)
    if (parsed) return new Date(parsed.y, parsed.m - 1, parsed.d)
  }
  if (!value) return null

  const text = clean(value)
  let date
  // Always read text dates as DD/MM/YYYY or DD-MM-YYYY (day first).
  const match = text.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/)
  if (match) {
    const day = Number(match[1])
    const month = Number(match[2])
    const year = Number(match[3])
    date = new Date(year, month - 1, day)
    // Reject impossible dates such as 31/02/2026 (JS would roll them over).
    if (date.getMonth() !== month - 1 || date.getDate() !== day) date = null
  } else {
    const parsed = new Date(text)
    if (!Number.isNaN(parsed.getTime())) {
      date = new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate())
    }
  }
  return date && !Number.isNaN(date.getTime()) ? date : null
}

function formatDate(value) {
  const date = getDate(value)
  return date
    ? date.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' })
    : '—'
}

// Waiting days = Today - Reporting Date (calendar days, never negative).
function getWaitingDays(value) {
  const startDate = getDate(value)
  if (!startDate) return null

  const today = new Date()
  const startUtc = Date.UTC(startDate.getFullYear(), startDate.getMonth(), startDate.getDate())
  const todayUtc = Date.UTC(today.getFullYear(), today.getMonth(), today.getDate())

  return Math.max(0, Math.round((todayUtc - startUtc) / 86400000))
}

function classifyGroup(row) {
  const country = clean(row['Depart from Country']).toLowerCase()
  const pickup = clean(row['Pickup Point: Location Name'])
  const isDRC = country.includes('democratic republic of the congo') || country.includes('democratic republic of congo') || country === 'drc' || country.includes('(drc)') || country.includes('congo, democratic republic')
  if (isDRC) return 'drc'
  if (/\breload\b/i.test(pickup)) return 'reload'
  return 'third'
}

function classifyKPI(days, group, operationalCategory) {
  if (days === null) return 'Date Missing'

  const category = clean(operationalCategory).toLowerCase()
  const rule = KPI_RULES[category]?.[group]
  if (!rule) return 'Date Missing'

  if (days <= rule.within) return 'Within KPI'
  if (days <= rule.exception) return 'Exception'
  return 'Major Exception'
}

function prepareRow(row, index) {
  const group = classifyGroup(row)
  const days = getWaitingDays(row['Reporting Date'])
  const operationalCategory = clean(row['Delivery Status'])

  return {
    ...row,
    _id: `${clean(row['Delivery Number']) || 'record'}-${index}`,
    _group: group,
    _days: days,
    _category: operationalCategory,
    _kpi: classifyKPI(days, group, operationalCategory),
  }
}

const DAYS_FILTER = 'Days Count'
const FILTER_COLUMNS = [DAYS_FILTER, ...EXCEL_COLUMNS]
const BLANK_OPTION = '__BLANK__'

function filterLabel(column) {
  return column === DAYS_FILTER ? 'Days Count (Today − Reporting Date)' : column
}

// The visible text of a cell for a given filter column ('' when empty).
function getColumnValue(row, column) {
  if (column === DAYS_FILTER) return row._days === null ? '' : String(row._days)
  if (column.includes('Date')) return getDate(row[column]) ? formatDate(row[column]) : ''
  return clean(row[column])
}

function StatusBadge({ status }) {
  const className = status.toLowerCase().replaceAll(' ', '-')
  return <span className={`status-badge ${className}`}><span className="status-badge-dot" />{status}</span>
}

function SummaryCard({ icon: Icon, title, value, note, color }) {
  return (
    <article className={`summary-card summary-${color}`}>
      <div className="summary-card-top"><span>{title}</span><span className="summary-icon"><Icon size={19} /></span></div>
      <h2>{value.toLocaleString()}</h2>
      <p>{note}</p>
    </article>
  )
}

function GroupKpiChart({ group, rows }) {
  const chartRows = rows.filter(row => row._kpi !== 'Date Missing')
  const missingDate = rows.filter(row => row._kpi === 'Date Missing').length
  const statuses = ['Within KPI', 'Exception', 'Major Exception']
  const values = statuses.map(status => ({
    name: status,
    count: chartRows.filter(row => row._kpi === status).length,
    color: COLORS[status],
  }))
  const total = chartRows.length
  const title = GROUPS.find(item => item.id === group)?.label || group

  return (
    <article className="panel group-kpi-panel" style={{ minWidth: 0, padding: '20px', overflow: 'hidden' }}>
      <div className="panel-heading" style={{ alignItems: 'flex-start', minHeight: '54px' }}>
        <div><h2 style={{ fontSize: '16px', lineHeight: 1.4 }}>{title}</h2><p>{total.toLocaleString()} records with a valid Reporting Date</p></div>
        <span className="panel-icon"><Truck size={18} /></span>
      </div>
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, minmax(0, 1fr))', gap: '10px', alignItems: 'end', minHeight: '190px', paddingTop: '12px' }}>
        {values.map(item => {
          const percent = total ? (item.count / total) * 100 : 0
          const barHeight = total ? Math.max(item.count > 0 ? 8 : 0, (percent / 100) * 112) : 0
          return (
            <div key={item.name} style={{ minWidth: 0, display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'flex-end', height: '185px', gap: '7px' }}>
              <strong style={{ fontSize: '13px', color: 'var(--text-primary, #172033)', fontVariantNumeric: 'tabular-nums' }}>{item.count.toLocaleString()}</strong>
              <span style={{ fontSize: '11px', color: 'var(--text-secondary, #64748b)' }}>{percent.toFixed(1)}%</span>
              <div title={`${item.name}: ${item.count} trucks (${percent.toFixed(1)}%)`} style={{ height: `${barHeight}px`, width: 'min(100%, 48px)', minHeight: item.count > 0 ? '8px' : '0', background: item.color, borderRadius: '6px 6px 2px 2px', transition: 'height 180ms ease' }} />
              <span style={{ fontSize: '11px', lineHeight: 1.25, minHeight: '29px', textAlign: 'center', overflowWrap: 'anywhere', color: 'var(--text-secondary, #64748b)' }}>{item.name}</span>
            </div>
          )
        })}
      </div>
      {missingDate > 0 && <p style={{ margin: '10px 0 0', fontSize: '11px', color: 'var(--text-secondary, #64748b)' }}>{missingDate} record(s) have a missing/invalid Reporting Date; excluded from these percentages.</p>}
      {total === 0 && <p style={{ margin: '8px 0 0', textAlign: 'center', fontSize: '12px', color: 'var(--text-secondary, #64748b)' }}>No records to chart for this selection.</p>}
    </article>
  )
}

const headerSelectStyle = {
  display: 'block', width: '100%', minWidth: '120px', maxWidth: '220px', marginTop: '6px',
  padding: '4px 6px', fontSize: '12px', fontWeight: 400, textTransform: 'none', letterSpacing: 'normal',
  color: '#172033', background: '#ffffff', border: '1px solid #cbd5e1', borderRadius: '6px', cursor: 'pointer',
}

// Excel-style dropdown that sits under a table column title.
function HeaderFilter({ label, value, onChange, list = [], hasBlank = false, formatValue }) {
  const items = value && value !== BLANK_OPTION && !list.includes(value) ? [...list, value] : list
  return (
    <select
      className="header-filter-select"
      style={headerSelectStyle}
      value={value || ''}
      onChange={event => onChange(event.target.value)}
      aria-label={`Filter by ${label}`}
    >
      <option value="">All</option>
      {hasBlank && <option value={BLANK_OPTION}>(Blank)</option>}
      {items.map(item => <option key={item} value={item}>{formatValue ? formatValue(item) : item}</option>)}
    </select>
  )
}

const KPI_STATUS_LIST = ['Within KPI', 'Exception', 'Major Exception', 'Date Missing']

function RecordTable({ sectionId, title, description, rows, onDetails, options, columnFilters, onColumnFilter, statusFilter, onStatusFilter }) {
  const [expanded, setExpanded] = useState(false)
  const visibleRows = expanded ? rows : rows.slice(0, 8)

  async function downloadFullTableScreenshot() {
    const wasExpanded = expanded
    const section = document.getElementById(sectionId)
    if (!section) {
      alert('Table nahi mila. Page refresh karke dobara try karein.')
      return
    }

    let captureHost
    try {
      // Expand the real React table first so every filtered row is rendered.
      if (!wasExpanded) setExpanded(true)
      await new Promise(resolve => setTimeout(resolve, 180))

      const sourceTable = section.querySelector('.wide-operations-table')
      if (!sourceTable) throw new Error('Table not found')

      // Capture an isolated copy outside the dashboard's scroll/overflow containers.
      const tableClone = sourceTable.cloneNode(true)
      tableClone.querySelectorAll('select').forEach(element => element.remove())
      tableClone.querySelectorAll('th').forEach(th => {
        th.style.position = 'static'
        th.style.top = 'auto'
        th.style.zIndex = 'auto'
      })
      tableClone.style.width = `${Math.max(sourceTable.scrollWidth, sourceTable.offsetWidth)}px`
      tableClone.style.minWidth = `${Math.max(sourceTable.scrollWidth, sourceTable.offsetWidth)}px`
      tableClone.style.maxWidth = 'none'
      tableClone.style.tableLayout = 'auto'
      tableClone.style.overflow = 'visible'

      captureHost = document.createElement('div')
      captureHost.style.position = 'fixed'
      captureHost.style.left = '-100000px'
      captureHost.style.top = '0'
      captureHost.style.width = `${Math.max(sourceTable.scrollWidth, sourceTable.offsetWidth)}px`
      captureHost.style.height = 'auto'
      captureHost.style.maxHeight = 'none'
      captureHost.style.overflow = 'visible'
      captureHost.style.background = '#ffffff'
      captureHost.style.padding = '12px'
      captureHost.style.boxSizing = 'content-box'
      captureHost.appendChild(tableClone)
      document.body.appendChild(captureHost)

      await new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)))
      const captureWidth = Math.ceil(tableClone.scrollWidth || tableClone.offsetWidth)
      const captureHeight = Math.ceil(tableClone.scrollHeight || tableClone.offsetHeight)
      if (!captureWidth || !captureHeight) throw new Error('Table has no drawable size')

      const scale = Math.max(0.35, Math.min(2, 28000 / captureWidth, 28000 / captureHeight))
      const canvas = await html2canvas(captureHost, {
        scale,
        backgroundColor: '#ffffff',
        useCORS: true,
        logging: false,
        width: captureHost.scrollWidth,
        height: captureHost.scrollHeight,
        windowWidth: Math.max(document.documentElement.clientWidth, captureHost.scrollWidth),
        windowHeight: Math.max(document.documentElement.clientHeight, captureHost.scrollHeight),
        scrollX: 0,
        scrollY: 0,
      })

      const link = document.createElement('a')
      const safeTitle = title.replace(/[^a-z0-9]+/gi, '-').replace(/^-|-$/g, '')
      link.download = `Africa-Operations-${safeTitle}-TABLE-ONLY-${new Date().toISOString().slice(0, 10)}.png`
      link.href = canvas.toDataURL('image/png')
      link.click()
    } catch (error) {
      console.error('Table screenshot error:', error)
      alert('Table ka complete screenshot nahi ban saka. Browser Console check karo.')
    } finally {
      if (captureHost?.parentNode) captureHost.parentNode.removeChild(captureHost)
      if (!wasExpanded) setExpanded(false)
    }
  }

  return (
    <section className="records-section" id={sectionId}>
      <div className="records-heading">
        <div>
          <div className="records-title-row"><h2>{title}</h2><span className="record-count">{rows.length.toLocaleString()} records</span></div>
          <p>{description}</p>
        </div>
        <div className="records-heading-actions">
          <button className="table-screenshot-button" onClick={downloadFullTableScreenshot} type="button" title="Download only this table as a PNG, including every row">
            <Camera size={15} /> Table Only PNG
          </button>
          <button className="secondary-button" onClick={() => setExpanded(!expanded)} type="button">
            {expanded ? 'Show Less' : 'Expand Table'}
            <ChevronDown size={16} className={expanded ? 'rotate-icon' : ''} />
          </button>
        </div>
      </div>

      <div className="table-wrap full-columns-table-wrap">
        <table className="wide-operations-table">
          <thead><tr>
            <th className="days-count-th">
              Days Count (Today − Reporting Date)
              <HeaderFilter
                label={filterLabel(DAYS_FILTER)}
                value={columnFilters[DAYS_FILTER]}
                onChange={value => onColumnFilter(DAYS_FILTER, value)}
                list={options[DAYS_FILTER]?.list}
                hasBlank={options[DAYS_FILTER]?.hasBlank}
                formatValue={value => `${value} ${value === '1' ? 'day' : 'days'}`}
              />
            </th>
            {EXCEL_COLUMNS.map(column => (
              <th key={column}>
                {column}
                <HeaderFilter
                  label={column}
                  value={columnFilters[column]}
                  onChange={value => onColumnFilter(column, value)}
                  list={options[column]?.list}
                  hasBlank={options[column]?.hasBlank}
                />
              </th>
            ))}
            <th>
              KPI Status
              <HeaderFilter
                label="KPI Status"
                value={statusFilter === 'all' ? '' : statusFilter}
                onChange={value => onStatusFilter(value || 'all')}
                list={KPI_STATUS_LIST}
              />
            </th>
            <th>Details</th>
          </tr></thead>
          <tbody>
            {visibleRows.map(row => (
              <tr key={row._id}>
                <td className="days-count-cell" style={{ fontWeight: 700, textAlign: 'center', fontVariantNumeric: 'tabular-nums' }}>
                  {row._days === null ? '—' : `${row._days} ${row._days === 1 ? 'day' : 'days'}`}
                </td>
                {EXCEL_COLUMNS.map(column => (
                  <td key={column} className={column === 'Delivery Number' ? 'delivery-cell' : column === 'Pickup Point: Location Name' ? 'pickup-cell' : ''} title={clean(row[column])}>
                    {column.includes('Date') ? formatDate(row[column]) : clean(row[column]) || '—'}
                  </td>
                ))}
                <td><StatusBadge status={row._kpi} /></td>
                <td><button className="icon-button" aria-label="View record details" title="View all record details" onClick={() => onDetails(row)}><Eye size={17} /></button></td>
              </tr>
            ))}
          </tbody>
        </table>
        {rows.length === 0 && <div className="table-empty"><FileSpreadsheet size={25} /><strong>No matching records</strong><p>Try changing your filters or upload another Excel file.</p></div>}
      </div>

      {rows.length > 8 && <button className="show-all-button" onClick={() => setExpanded(!expanded)}>{expanded ? 'Show fewer records' : `Show all ${rows.length} records`}</button>}
    </section>
  )
}

function RecordModal({ row, onClose }) {
  if (!row) return null
  return (
    <div className="modal-backdrop" onClick={onClose}>
      <section className="details-modal" role="dialog" aria-modal="true" aria-label="Delivery record details" onClick={event => event.stopPropagation()}>
        <div className="modal-heading">
          <div><p className="eyebrow">DELIVERY DETAILS</p><h2>{clean(row['Delivery Number']) || 'Record details'}</h2><StatusBadge status={row._kpi} /></div>
          <button className="icon-button" onClick={onClose} aria-label="Close details"><X size={20} /></button>
        </div>
        <div className="details-grid">
          {EXCEL_COLUMNS.map(column => <div className="detail-field" key={column}><span>{column}</span><strong>{column.includes('Date') ? formatDate(row[column]) : clean(row[column]) || '—'}</strong></div>)}
          <div className="detail-field"><span>Pickup Group</span><strong>{GROUPS.find(group => group.id === row._group)?.label || '—'}</strong></div>
          <div className="detail-field"><span>Waiting Days</span><strong>{row._days === null ? 'Reporting date missing' : row._days}</strong></div>
        </div>
        <div className="modal-footer"><button className="primary-button" onClick={onClose}>Done</button></div>
      </section>
    </div>
  )
}

function App() {
  const emptyColumnFilters = () => Object.fromEntries(FILTER_COLUMNS.map(column => [column, '']))

  const [fileName, setFileName] = useState('')
  const [rows, setRows] = useState([])
  const [error, setError] = useState('')
  const [loading, setLoading] = useState(false)
  const [search, setSearch] = useState('')
  const emptyTableFilters = () => Object.fromEntries(GROUPS.map(group => [group.id, emptyColumnFilters()]))
  const [tableFilters, setTableFilters] = useState(emptyTableFilters)
  const [groupFilter, setGroupFilter] = useState('all')
  const [categoryFilter, setCategoryFilter] = useState('all')
  const [statusFilter, setStatusFilter] = useState('all')
  const [selectedRow, setSelectedRow] = useState(null)
  const [sidebarOpen, setSidebarOpen] = useState(false)

  // Check the local date every minute. When midnight passes, recalculate
  // waiting days and KPI statuses automatically without re-uploading Excel.
  const [todayKey, setTodayKey] = useState(() => new Date().toDateString())

  useEffect(() => {
    const checkToday = () => {
      const newTodayKey = new Date().toDateString()
      setTodayKey(currentKey => currentKey === newTodayKey ? currentKey : newTodayKey)
    }

    const timerId = window.setInterval(checkToday, 60 * 1000)
    return () => window.clearInterval(timerId)
  }, [])

  // `rows` holds the raw Excel rows. Days/KPI are recomputed here so they never go stale.
  const currentRows = useMemo(
    () => rows.map((row, index) => prepareRow(row, index)),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [rows, todayKey]
  )

  async function handleFileUpload(event) {
    const file = event.target.files?.[0]
    if (!file) return
    setLoading(true)
    setError('')
    try {
      const buffer = await file.arrayBuffer()
      const workbook = XLSX.read(buffer, { type: 'array' })
      const sheet = workbook.Sheets[workbook.SheetNames[0]]
      if (!sheet) throw new Error('The uploaded workbook has no readable worksheet.')
      const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: '' })
      if (rawRows.length === 0) throw new Error('The selected worksheet contains no records.')

      const headers = Object.keys(rawRows[0])
      const required = ['Delivery Status', 'Depart from Country', 'Pickup Point: Location Name', 'Reporting Date']
      const missing = required.filter(header => !headers.includes(header))
      if (missing.length) throw new Error(`Missing required Excel columns: ${missing.join(', ')}`)

      // Keep raw rows only; prepareRow runs in currentRows.
      const waitingRows = rawRows.filter(row =>
        ['waiting to load', 'waiting for documents'].includes(clean(row['Delivery Status']).toLowerCase())
      )
      setRows(waitingRows)
      setFileName(file.name)
      setSearch('')
      setTableFilters(emptyTableFilters())
      setGroupFilter('all')
      setCategoryFilter('all')
      setStatusFilter('all')
    } catch (err) {
      setError(err.message || 'Could not read this file. Please check the Excel format.')
    } finally {
      setLoading(false)
      event.target.value = ''
    }
  }

  // Toolbar filters (search, pickup group, KPI status) apply to the whole dashboard.
  const baseRows = useMemo(() => {
    const query = search.trim().toLowerCase()
    return currentRows.filter(row => {
      const groupMatches = groupFilter === 'all' || row._group === groupFilter
      const categoryMatches = categoryFilter === 'all' || clean(row._category).toLowerCase() === categoryFilter
      const statusMatches = statusFilter === 'all' || row._kpi === statusFilter
      const searchMatches = !query || [
        ...EXCEL_COLUMNS.map(column => column.includes('Date') ? `${formatDate(row[column])} ${clean(row[column])}` : clean(row[column])),
        String(row._days ?? ''), row._kpi,
      ].some(value => value.toLowerCase().includes(query))
      return groupMatches && categoryMatches && statusMatches && searchMatches
    })
  }, [currentRows, search, groupFilter, categoryFilter, statusFilter])

  // Header dropdowns belong to one table only, so choosing a value in one table never empties the others.
  const rowsByGroup = useMemo(() => Object.fromEntries(GROUPS.map(group => {
    const filters = tableFilters[group.id] || {}
    const groupRows = baseRows
      .filter(row => row._group === group.id)
      .filter(row => FILTER_COLUMNS.every(column => {
        const selected = filters[column]
        if (!selected) return true
        const value = getColumnValue(row, column)
        return selected === BLANK_OPTION ? value === '' : value === selected
      }))
    return [group.id, groupRows]
  })), [baseRows, tableFilters])

  // KPI cards and the overall chart count whatever the three tables currently show.
  const filteredRows = useMemo(() => GROUPS.flatMap(group => rowsByGroup[group.id]), [rowsByGroup])

  // Dropdown options: every distinct value found in the whole uploaded file (all countries, all pickup points...).
  const columnOptions = useMemo(() => {
    const result = {}
    FILTER_COLUMNS.forEach(column => {
      const values = new Set()
      let hasBlank = false
      currentRows.forEach(row => {
        const value = getColumnValue(row, column)
        if (value === '') hasBlank = true
        else values.add(value)
      })
      const list = [...values]
      if (column === DAYS_FILTER) list.sort((a, b) => Number(a) - Number(b))
      else if (column.includes('Date')) list.sort((a, b) => getDate(a)?.getTime() - getDate(b)?.getTime())
      else list.sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' }))
      result[column] = { list, hasBlank }
    })
    return result
  }, [currentRows])

  const counts = useMemo(() => ({
    total: filteredRows.length,
    within: filteredRows.filter(row => row._kpi === 'Within KPI').length,
    exception: filteredRows.filter(row => row._kpi === 'Exception').length,
    major: filteredRows.filter(row => row._kpi === 'Major Exception').length,
    missingDate: filteredRows.filter(row => row._kpi === 'Date Missing').length,
  }), [filteredRows])

  const pieData = [
    { name: 'Within KPI', value: counts.within },
    { name: 'Exception', value: counts.exception },
    { name: 'Major Exception', value: counts.major },
  ]

  function scrollTo(id) {
    document.getElementById(id)?.scrollIntoView({ behavior: 'smooth' })
    setSidebarOpen(false)
  }

  function resetFilters() {
    setSearch('')
    setTableFilters(emptyTableFilters())
    setGroupFilter('all')
    setCategoryFilter('all')
    setStatusFilter('all')
  }

  return (
    <div className="app-shell">
      {sidebarOpen && <button className="sidebar-scrim" aria-label="Close navigation" onClick={() => setSidebarOpen(false)} />}
      <aside className={`sidebar ${sidebarOpen ? 'sidebar-open' : ''}`}>
        <div className="sidebar-brand">
          <div className="brand-mark">A</div>
          <div><strong>Africa Operations</strong><span>KPI MONITORING</span></div>
          <button className="mobile-close" aria-label="Close navigation" onClick={() => setSidebarOpen(false)}><X size={18} /></button>
        </div>
        <p className="sidebar-label">WORKSPACE</p>
        <button className="nav-item active" onClick={() => scrollTo('overview')}><LayoutDashboard size={18} /> Overview</button>
        <button className="nav-item" onClick={() => scrollTo('table-drc')}><Truck size={18} /> DRC</button>
        <button className="nav-item" onClick={() => scrollTo('table-reload')}><RefreshCw size={18} /> Reload Pickup Points</button>
        <button className="nav-item" onClick={() => scrollTo('table-third')}><FileSpreadsheet size={18} /> Third Party</button>
        <button className="nav-item" onClick={() => scrollTo('action-summary')}><AlertTriangle size={18} /> Action Needed</button>
        <div className="sidebar-bottom"><span className="sidebar-status-dot" /><div><strong>Dashboard Ready</strong><small>Waiting To Load + Waiting for Documents</small></div></div>
      </aside>

      <div className="main-area">
        <header className="topbar">
          <button className="mobile-menu" aria-label="Toggle navigation" aria-expanded={sidebarOpen} onClick={() => setSidebarOpen(open => !open)} type="button"><Menu size={21} /></button>
          <div className="topbar-title"><span>Operations /</span> Loading Performance</div>
          <label className="upload-button"><Upload size={17} />{loading ? 'Reading file...' : 'Upload Excel File'}<input type="file" accept=".xlsx,.xls,.csv" hidden disabled={loading} onChange={handleFileUpload} /></label>
        </header>

        <main className="dashboard-content" id="overview">
          <section className="page-heading">
            <div><p className="eyebrow">OPERATIONS OVERVIEW</p><h1>Loading Performance</h1><p className="subtitle">Monitor Waiting To Load and Waiting for Documents aging and KPI exceptions across pickup points.</p></div>
            <div className="file-status"><span className={`status-dot ${fileName ? 'status-dot-ready' : ''}`} /><div><strong>{fileName || 'No file uploaded'}</strong><p>{fileName ? `${rows.length.toLocaleString()} Waiting To Load / Documents records loaded` : 'Upload an Excel file to begin'}</p></div></div>
          </section>

          {error && <div className="error-banner"><AlertTriangle size={18} /><span>{error}</span></div>}
          <div className="scope-note"><Clock3 size={17} /><span>Current view: <strong>Waiting To Load + Waiting for Documents</strong>.</span></div>

          <section className="summary-grid">
            <SummaryCard icon={Truck} title="All Truck Performance" value={counts.total} note="Filtered operational records" color="navy" />
            <SummaryCard icon={CircleCheck} title="Within KPI" value={counts.within} note="Within target" color="green" />
            <SummaryCard icon={Clock3} title="Exception" value={counts.exception} note="Needs attention" color="orange" />
            <SummaryCard icon={AlertTriangle} title="Major Exception" value={counts.major} note="Action required" color="red" />
          </section>

          <section className="charts-grid">
            <article className="panel chart-panel pie-panel" style={{ gridColumn: '1 / -1' }}>
              <div className="panel-heading"><div><h2>Overall Performance</h2><p>KPI status distribution</p></div><span className="panel-icon"><LayoutDashboard size={18} /></span></div>
              <div className="pie-chart-wrap">
                {counts.total === 0 ? <div className="chart-empty"><FileSpreadsheet size={28} /><p>Upload a file to see performance.</p></div> : (
                  <ResponsiveContainer width="100%" height="100%">
                    <PieChart>
                      <Pie data={pieData} dataKey="value" nameKey="name" cx="50%" cy="50%" innerRadius={67} outerRadius={100} paddingAngle={3} stroke="none">
                        {pieData.map(item => <Cell key={item.name} fill={COLORS[item.name]} />)}
                      </Pie>
                      <Tooltip formatter={(value, name) => [value, name]} />
                    </PieChart>
                  </ResponsiveContainer>
                )}
                <div className="pie-center"><strong>{counts.total.toLocaleString()}</strong><span>Total Records</span></div>
              </div>
              <div className="chart-legend">
                {pieData.map(item => <div className="legend-item" key={item.name}><span className="legend-dot" style={{ background: COLORS[item.name] }} /><span>{item.name}</span><strong>{counts.total ? ((item.value / counts.total) * 100).toFixed(1) : '0.0'}%</strong></div>)}
              </div>
            </article>
          </section>

          <section aria-label="KPI performance by pickup group" style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(min(100%, 280px), 1fr))', gap: '18px', marginTop: '20px', marginBottom: '24px' }}>
            {GROUPS.map(group => (
              <GroupKpiChart key={group.id} group={group.id} rows={rowsByGroup[group.id] || []} />
            ))}
          </section>

          <section className="filter-toolbar" style={{ display: 'flex', alignItems: 'center', gap: '14px', flexWrap: 'wrap', margin: '4px 0 22px' }}>
            <label className="filter-search" style={{ flex: '1 1 280px', minWidth: '240px' }}><Search size={17} /><input type="search" placeholder="Search across all columns..." value={search} onChange={event => setSearch(event.target.value)} /></label>
            <label className="filter-select" style={{ display: 'flex', flexDirection: 'row', alignItems: 'center', gap: '8px', margin: 0 }}><span style={{ whiteSpace: 'nowrap' }}>Pickup Group</span><select value={groupFilter} onChange={event => setGroupFilter(event.target.value)}><option value="all">All pickup groups</option>{GROUPS.map(group => <option key={group.id} value={group.id}>{group.label}</option>)}</select></label>
            <label className="filter-select" style={{ display: 'flex', flexDirection: 'row', alignItems: 'center', gap: '8px', margin: 0 }}><span style={{ whiteSpace: 'nowrap' }}>KPI Status</span><select value={statusFilter} onChange={event => setStatusFilter(event.target.value)}><option value="all">All KPI statuses</option><option value="Within KPI">Within KPI</option><option value="Exception">Exception</option><option value="Major Exception">Major Exception</option><option value="Date Missing">Date Missing</option></select></label>
            <button className="reset-button" onClick={resetFilters} type="button">Reset all</button>
          </section>

          <section className="tables-intro">
            <div><p className="eyebrow">RECORD BREAKDOWN</p><h2>Operational Records</h2><p>Three pickup groups, each with its own KPI results and action list.</p></div>
            <div className="dashboard-export-buttons">
              <button className="secondary-button" onClick={handleScreenshot} type="button"><Camera size={16} /> Full Dashboard PNG</button>
              <button className="secondary-button print-button" onClick={() => window.print()} type="button"><Printer size={16} /> Print / Save PDF</button>
            </div>
          </section>

          {GROUPS.map(group => {
            const groupRows = rowsByGroup[group.id] || []
            const description = group.id === 'drc'
              ? 'Depart from Country identifies the Democratic Republic of Congo.'
              : group.id === 'reload'
                ? 'Non-DRC records where Pickup Point: Location Name contains the word Reload.'
                : 'Non-DRC records where the pickup-point name does not contain Reload.'
            return (
              <RecordTable
                key={group.id}
                sectionId={`table-${group.id}`}
                title={group.label}
                description={description}
                rows={groupRows}
                onDetails={setSelectedRow}
                options={columnOptions}
                columnFilters={tableFilters[group.id] || {}}
                onColumnFilter={(column, value) => setTableFilters(current => ({ ...current, [group.id]: { ...current[group.id], [column]: value } }))}
                statusFilter={statusFilter}
                onStatusFilter={setStatusFilter}
              />
            )
          })}

          <section className="overall-action-summary" id="action-summary"><div className="overall-action-icon"><AlertTriangle size={21} /></div><div><h2>Overall Action Needed</h2><p>Major exceptions across all visible pickup groups.</p></div><strong>{counts.major.toLocaleString()}</strong></section>
          {counts.missingDate > 0 && <p className="missing-date-note">{counts.missingDate} record(s) have no valid Reporting Date. They are marked Date Missing and excluded from the three KPI chart categories until the date is available.</p>}
          <footer className="dashboard-footer">Africa Operations <span>·</span> KPI Monitoring Dashboard <span className="footer-dot">·</span> Loading Operations</footer>
        </main>
      </div>
      <RecordModal row={selectedRow} onClose={() => setSelectedRow(null)} />
    </div>
  )
}

export default App
