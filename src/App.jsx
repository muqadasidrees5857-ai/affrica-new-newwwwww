import { useEffect, useMemo, useState } from 'react'
import { createPortal } from 'react-dom'
import html2canvas from 'html2canvas'
import * as XLSX from 'xlsx'
import { PieChart, Pie, Cell, Tooltip, ResponsiveContainer } from 'recharts'
import {
  LayoutDashboard, Truck, Upload, Search, FileSpreadsheet,
  AlertTriangle, CircleCheck, Clock3, Eye, Printer, Camera, ChevronDown,
  Menu, X, RefreshCw, ArrowUp, ArrowDown, ArrowUpDown, Maximize2,
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

// Columns shown in the three tables (Internal Comments and Haulier are only shown in the details popup).
const HIDDEN_TABLE_COLUMNS = ['Last Event: Internal Comments', 'Haulier: Account Name']
const TABLE_COLUMNS = EXCEL_COLUMNS.filter(column => !HIDDEN_TABLE_COLUMNS.includes(column))

const DAYS_FILTER = 'Days Count'
const FILTER_COLUMNS = [DAYS_FILTER, ...TABLE_COLUMNS]
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
  flex: 1, minWidth: 0, width: '100%',
  padding: '3px 4px', fontSize: '11px', fontWeight: 400, textTransform: 'none', letterSpacing: 'normal',
  color: '#172033', background: '#ffffff', border: '1px solid #cbd5e1', borderRadius: '6px', cursor: 'pointer',
}

const KPI_COLUMN = 'KPI Status'
const KPI_ORDER = { 'Within KPI': 0, Exception: 1, 'Major Exception': 2, 'Date Missing': 3 }

// Value used to sort a row by a column (null = empty, always sorted last).
function getSortValue(row, column) {
  if (column === DAYS_FILTER) return row._days
  if (column === KPI_COLUMN) return KPI_ORDER[row._kpi] ?? 99
  if (column.includes('Date')) return getDate(row[column])?.getTime() ?? null
  const text = clean(row[column])
  return text === '' ? null : text
}

function compareRows(a, b, column, direction) {
  const x = getSortValue(a, column)
  const y = getSortValue(b, column)
  if (x === null && y === null) return 0
  if (x === null) return 1
  if (y === null) return -1
  const result = typeof x === 'number'
    ? x - y
    : String(x).localeCompare(String(y), undefined, { numeric: true, sensitivity: 'base' })
  return direction === 'asc' ? result : -result
}

// Excel-style dropdown + sort button that sit under a table column title.
function HeaderFilter({ label, value, onChange, list = [], hasBlank = false, formatValue, sortColumn, sort, onSort }) {
  const items = value && value !== BLANK_OPTION && !list.includes(value) ? [...list, value] : list
  const direction = sort?.column === sortColumn ? sort.direction : null
  const SortIcon = direction === 'asc' ? ArrowUp : direction === 'desc' ? ArrowDown : ArrowUpDown
  const sortTitle = direction === 'asc'
    ? 'Sorted A → Z / smallest first (click for Z → A)'
    : direction === 'desc'
      ? 'Sorted Z → A / largest first (click to clear)'
      : 'Sort A → Z / smallest first'

  return (
    <div className="header-filter-row" style={{ display: 'flex', alignItems: 'center', gap: '3px', marginTop: '6px' }}>
      {onChange && (
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
      )}
      {onSort && (
        <button
          type="button"
          className="header-sort-button"
          title={sortTitle}
          aria-label={`Sort by ${label}`}
          onClick={() => onSort(sortColumn)}
          style={{
            flex: 'none', width: '24px', height: '24px', padding: 0, display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
            border: '1px solid #cbd5e1', borderRadius: '6px', cursor: 'pointer',
            background: direction ? '#e0e7ff' : '#ffffff', color: direction ? '#3730a3' : '#64748b',
          }}
        >
          <SortIcon size={14} />
        </button>
      )}
    </div>
  )
}

// Column widths (percent of the table) so all 13 columns fit on screen without scrolling.
// Order: Days Count, the 10 table columns, KPI Status, Details.
const TABLE_COLUMN_WIDTHS = ['6%', '9%', '7%', '12%', '8%', '8%', '7%', '7%', '7%', '7%', '12%', '7%', '3%']

const KPI_STATUS_LIST = ['Within KPI', 'Exception', 'Major Exception', 'Date Missing']

// Vertical stacked chart: one bar per pickup point, one colour block per transporter (driver).
// Only Major Exception trucks are counted, and it never changes with the filters.
const MAJOR_CHART_MAX_HEIGHT = 620 // the plot is never taller than this on the page
const MAJOR_CHART_MIN_HEIGHT = 380

// Short display name (hover shows the full one): drops LIMITED / LTD / (PTY) endings.
function shortDriverName(name) {
  let text = name
  text = text.replace(/\s+(LIMITED|LTD\.?)\s*$/i, '')
  text = text.replace(/\s*\(PTY\)\s*$/i, '')
  text = text.replace(/\s+(LIMITED|LTD\.?)\s*$/i, '')
  return text.trim() || name
}

function enterBrowserFullscreen() {
  const element = document.documentElement
  if (element.requestFullscreen) {
    const result = element.requestFullscreen()
    if (result && result.catch) result.catch(() => {})
  }
}

function leaveBrowserFullscreen() {
  if (document.fullscreenElement && document.exitFullscreen) {
    const result = document.exitFullscreen()
    if (result && result.catch) result.catch(() => {})
  }
}

function MajorExceptionPickupChart({ rows }) {
  const model = useMemo(() => {
    const byPickup = new Map()
    const driverTotals = new Map()
    let totalTrucks = 0

    rows.filter(row => row._kpi === 'Major Exception').forEach(row => {
      const pickup = clean(row['Pickup Point: Location Name']) || '(No pickup point)'
      const driver = clean(row['Transporter: Account Name']) || '(No transporter)'
      const reg = clean(row['Horse: Asset No [Reg]'])

      if (!byPickup.has(pickup)) byPickup.set(pickup, { pickup, total: 0, drivers: new Map() })
      const entry = byPickup.get(pickup)
      entry.total += 1
      if (!entry.drivers.has(driver)) entry.drivers.set(driver, { driver, count: 0, regs: [] })
      const segment = entry.drivers.get(driver)
      segment.count += 1
      if (reg) segment.regs.push(reg)

      driverTotals.set(driver, (driverTotals.get(driver) || 0) + 1)
      totalTrucks += 1
    })

    // Every driver keeps one colour everywhere.
    const driverList = [...driverTotals.entries()]
      .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
      .map(([name, total], index) => ({
        name,
        total,
        color: `hsl(${Math.round((index * 137.508) % 360)}, ${55 + (index % 2) * 10}%, ${30 + (index % 3) * 4}%)`,
      }))
    const colors = Object.fromEntries(driverList.map(item => [item.name, item.color]))

    const pickups = [...byPickup.values()]
      .map(entry => ({
        ...entry,
        segments: [...entry.drivers.values()].sort((a, b) => b.count - a.count || a.driver.localeCompare(b.driver)),
      }))
      .sort((a, b) => b.total - a.total || a.pickup.localeCompare(b.pickup))

    return { pickups, driverList, colors, totalTrucks, maxTotal: pickups[0]?.total || 0 }
  }, [rows])

  const { pickups, driverList, colors, totalTrucks, maxTotal } = model

  // Full-screen mode: the chart fills the whole screen, with no scrolling in either direction.
  const [full, setFull] = useState(false)
  const [viewport, setViewport] = useState(() => ({
    width: typeof window === 'undefined' ? 1280 : window.innerWidth,
    height: typeof window === 'undefined' ? 800 : window.innerHeight,
  }))

  useEffect(() => {
    if (!full) return undefined
    const onResize = () => setViewport({ width: window.innerWidth, height: window.innerHeight })
    const onKey = event => { if (event.key === 'Escape') setFull(false) }
    const onFullscreenChange = () => { if (!document.fullscreenElement) setFull(false) }
    const previousOverflow = document.body.style.overflow
    document.body.style.overflow = 'hidden'
    onResize()
    window.addEventListener('resize', onResize)
    window.addEventListener('keydown', onKey)
    document.addEventListener('fullscreenchange', onFullscreenChange)
    return () => {
      document.body.style.overflow = previousOverflow
      window.removeEventListener('resize', onResize)
      window.removeEventListener('keydown', onKey)
      document.removeEventListener('fullscreenchange', onFullscreenChange)
    }
  }, [full])

  function openFull() {
    setFull(true)
    enterBrowserFullscreen()
  }

  function closeFull() {
    setFull(false)
    leaveBrowserFullscreen()
  }

  // The scale always ends on a round number above the tallest bar.
  const step = maxTotal > 20 ? 5 : maxTotal > 10 ? 2 : 1
  const niceMax = (Math.floor(maxTotal / step) + 1) * step
  const ticks = []
  for (let value = 0; value <= niceMax; value += step) ticks.push(value)
  const percentOf = value => `${(value / niceMax) * 100}%`

  const pageHeight = Math.max(MAJOR_CHART_MIN_HEIGHT, Math.min(MAJOR_CHART_MAX_HEIGHT, niceMax * 28))
  const fullHeight = Math.max(240, viewport.height - 195)

  const barsRowStyle = { display: 'flex', justifyContent: 'center', gap: 'clamp(4px, 0.45vw, 9px)', padding: '0 4px' }
  const columnStyle = { flex: '1 1 0', minWidth: 0, maxWidth: '240px' }

  const renderHeader = isFull => (
    <div
      className={isFull ? undefined : 'panel-heading'}
      style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', gap: '12px', padding: '0 10px' }}
    >
      <div>
        <h2 style={isFull ? { margin: 0, fontSize: '20px' } : undefined}>Major Exception Trucks by Pickup Point</h2>
        {!isFull && <p>Each bar is a pickup point. Each colour block is one transporter (driver): the number is how many of its trucks are in Major Exception there.</p>}
        <div style={{ display: 'flex', flexWrap: 'wrap', gap: '8px', marginTop: '10px' }}>
          <span className="mpc-chip">{totalTrucks.toLocaleString()} Major Exception trucks</span>
          <span className="mpc-chip mpc-chip-neutral">{pickups.length.toLocaleString()} pickup points</span>
          <span className="mpc-chip mpc-chip-neutral">{driverList.length.toLocaleString()} transporters</span>
        </div>
      </div>
      {pickups.length > 0 && (
        isFull ? (
          <button className="secondary-button" type="button" onClick={closeFull} title="Close full screen (Esc)"><X size={16} /> Exit Full Screen</button>
        ) : (
          <button className="secondary-button" type="button" onClick={openFull} title="Show the whole chart on one screen"><Maximize2 size={16} /> Full Screen</button>
        )
      )}
    </div>
  )

  const renderPlot = (plotHeight, isFull) => (
    <div className={isFull ? undefined : 'mpc-scroll'} style={{ marginTop: '14px' }}>
      <div className={isFull ? undefined : 'mpc-inner'} style={{ display: 'flex', gap: '4px' }}>
        {/* Y axis */}
        <div style={{ position: 'relative', flex: 'none', width: '26px', height: `${plotHeight}px` }}>
          {ticks.map(tick => (
            <span key={tick} style={{ position: 'absolute', right: 0, bottom: percentOf(tick), transform: 'translateY(50%)', fontSize: '11px', color: '#64748b', fontVariantNumeric: 'tabular-nums' }}>{tick}</span>
          ))}
        </div>

        <div style={{ flex: 1, minWidth: 0 }}>
          {/* Plot */}
          <div style={{ position: 'relative', height: `${plotHeight}px`, borderBottom: '2px solid #94a3b8' }}>
            {ticks.slice(1).map(tick => (
              <div key={tick} style={{ position: 'absolute', left: 0, right: 0, bottom: percentOf(tick), borderTop: '1px dashed #e2e8f0' }} />
            ))}

            <div style={{ ...barsRowStyle, position: 'absolute', inset: 0, alignItems: 'flex-end' }}>
              {pickups.map(entry => (
                <div key={entry.pickup} style={{ ...columnStyle, height: '100%', display: 'flex', flexDirection: 'column', justifyContent: 'flex-end' }}>
                  <strong style={{ flex: 'none', textAlign: 'center', fontSize: '16px', lineHeight: '24px', color: '#172033' }}>{entry.total}</strong>
                  <div style={{ flex: 'none', height: percentOf(entry.total), display: 'flex', flexDirection: 'column-reverse', borderRadius: '8px 8px 0 0', overflow: 'hidden', boxShadow: '0 3px 10px rgba(15, 23, 42, 0.14)' }}>
                    {entry.segments.map(segment => {
                      const heightPx = (segment.count / niceMax) * plotHeight
                      const lines = Math.max(1, Math.floor((heightPx - 3) / 12))
                      const hint = `${segment.driver}\n${segment.count} truck${segment.count > 1 ? 's' : ''} at ${entry.pickup}${segment.regs.length ? `\nTrucks: ${segment.regs.join(', ')}` : ''}`
                      return (
                        <div
                          key={segment.driver}
                          className="mpc-seg"
                          title={hint}
                          style={{
                            flex: `${segment.count} 1 0`, minHeight: 0, boxSizing: 'border-box', padding: '0 3px', background: colors[segment.driver],
                            borderTop: '2px solid #ffffff', color: '#ffffff', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '5px', overflow: 'hidden',
                          }}
                        >
                          <span style={{ width: '100%', textAlign: 'center', fontSize: '11px', lineHeight: '12px', fontWeight: 600, textShadow: '0 1px 1px rgba(0, 0, 0, 0.35)', overflow: 'hidden', ...(lines === 1 ? { display: 'block', whiteSpace: 'nowrap', textOverflow: 'ellipsis' } : { display: '-webkit-box', WebkitBoxOrient: 'vertical', WebkitLineClamp: lines, wordBreak: 'break-word' }) }}>
                            <b style={{ fontSize: '12px' }}>{segment.count}</b> · {shortDriverName(segment.driver)}
                          </span>
                        </div>
                      )
                    })}
                  </div>
                </div>
              ))}
            </div>
          </div>

          {/* Pickup point names */}
          <div style={{ ...barsRowStyle, alignItems: 'flex-start' }}>
            {pickups.map(entry => (
              <div
                key={entry.pickup}
                title={entry.pickup}
                style={{ ...columnStyle, paddingTop: '10px', textAlign: 'center', fontSize: '11.5px', fontWeight: 600, lineHeight: 1.3, color: '#334155', wordBreak: 'break-word', display: '-webkit-box', WebkitBoxOrient: 'vertical', WebkitLineClamp: isFull ? 3 : 4, overflow: 'hidden' }}
              >
                {entry.pickup}
              </div>
            ))}
          </div>
        </div>
      </div>
    </div>
  )

  return (
    <article className="panel major-pickup-panel" style={{ padding: '16px 6px 14px', marginBottom: '20px', minWidth: 0 }}>
      <style>{`
        .mpc-seg { transition: filter 120ms ease; cursor: default; }
        .mpc-seg:hover { filter: brightness(1.18); }
        .mpc-chip { display: inline-flex; align-items: center; gap: 6px; padding: 5px 12px; border-radius: 999px; background: #fef2f2; color: #b91c1c; font-size: 13px; font-weight: 600; }
        .mpc-chip-neutral { background: #f1f5f9; color: #334155; }
        @media (max-width: 900px) {
          .mpc-scroll { overflow-x: auto; }
          .mpc-inner { min-width: 780px; }
        }
      `}</style>

      {renderHeader(false)}

      {pickups.length === 0 ? (
        <p style={{ margin: '24px 0 8px', textAlign: 'center', fontSize: '14px', color: 'var(--text-secondary, #64748b)' }}>No Major Exception trucks to show.</p>
      ) : renderPlot(pageHeight, false)}

      {full && pickups.length > 0 && createPortal(
        <div
          role="dialog"
          aria-modal="true"
          aria-label="Major Exception trucks by pickup point, full screen"
          style={{ position: 'fixed', inset: 0, zIndex: 3000, background: '#f5f7fb', padding: '10px 8px 8px', boxSizing: 'border-box', overflow: 'hidden', display: 'flex', flexDirection: 'column' }}
        >
          {renderHeader(true)}
          {renderPlot(fullHeight, true)}
        </div>,
        document.body
      )}
    </article>
  )
}

function RecordTable({ sectionId, title, description, rows, onDetails, options, columnFilters, onColumnFilter, statusFilter, onStatusFilter }) {
  const [expanded, setExpanded] = useState(false)
  const [fullScreen, setFullScreen] = useState(false)
  const [sort, setSort] = useState(null)

  useEffect(() => {
    if (!fullScreen) return undefined
    const previousOverflow = document.body.style.overflow
    const onKeyDown = event => {
      if (event.key === 'Escape') setFullScreen(false)
    }
    document.body.style.overflow = 'hidden'
    window.addEventListener('keydown', onKeyDown)
    return () => {
      document.body.style.overflow = previousOverflow
      window.removeEventListener('keydown', onKeyDown)
    }
  }, [fullScreen])
  const sortedRows = useMemo(
    () => (sort ? [...rows].sort((a, b) => compareRows(a, b, sort.column, sort.direction)) : rows),
    [rows, sort]
  )
  const visibleRows = expanded ? sortedRows : sortedRows.slice(0, 8)

  // Click cycle for a column: ascending -> descending -> no sorting.
  function toggleSort(column) {
    setSort(current => {
      if (!current || current.column !== column) return { column, direction: 'asc' }
      return current.direction === 'asc' ? { column, direction: 'desc' } : null
    })
  }

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
      tableClone.querySelectorAll('.header-filter-row').forEach(element => element.remove())
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
    <section
      className="records-section"
      id={sectionId}
      style={fullScreen ? {
        position: 'fixed',
        inset: 0,
        zIndex: 4000,
        width: '100vw',
        height: '100vh',
        maxHeight: '100vh',
        overflow: 'auto',
        boxSizing: 'border-box',
        padding: '18px',
        margin: 0,
        background: '#f5f7fb',
      } : undefined}
    >
      <div className="records-heading">
        <div>
          <div className="records-title-row"><h2>{title}</h2><span className="record-count">{rows.length.toLocaleString()} records</span></div>
          <p>{description}</p>
        </div>
        <div className="records-heading-actions">
          {fullScreen ? (
            <button className="secondary-button" onClick={() => setFullScreen(false)} type="button" title="Close full screen (Esc)">
              <X size={16} /> Exit Full Screen
            </button>
          ) : (
            <button className="secondary-button" onClick={() => setFullScreen(true)} type="button" title="Show this table on the full screen">
              <Maximize2 size={16} /> Full Screen
            </button>
          )}
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
          <colgroup>
            {TABLE_COLUMN_WIDTHS.map((width, index) => <col key={index} style={{ width }} />)}
          </colgroup>
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
                sortColumn={DAYS_FILTER}
                sort={sort}
                onSort={toggleSort}
              />
            </th>
            {TABLE_COLUMNS.map(column => (
              <th key={column}>
                {column}
                <HeaderFilter
                  label={column}
                  value={columnFilters[column]}
                  onChange={value => onColumnFilter(column, value)}
                  list={options[column]?.list}
                  hasBlank={options[column]?.hasBlank}
                  sortColumn={column}
                  sort={sort}
                  onSort={toggleSort}
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
                sortColumn={KPI_COLUMN}
                sort={sort}
                onSort={toggleSort}
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
                {TABLE_COLUMNS.map(column => (
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

  // Overall numbers (KPI cards + Overall Performance chart) always cover ALL records, whatever filters are set.
  const counts = useMemo(() => ({
    total: currentRows.length,
    within: currentRows.filter(row => row._kpi === 'Within KPI').length,
    exception: currentRows.filter(row => row._kpi === 'Exception').length,
    major: currentRows.filter(row => row._kpi === 'Major Exception').length,
    missingDate: currentRows.filter(row => row._kpi === 'Date Missing').length,
  }), [currentRows])

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
            <SummaryCard icon={Truck} title="All Truck Performance" value={counts.total} note="All operational records" color="navy" />
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
                {pieData.map(item => <div className="legend-item" key={item.name}><span className="legend-dot" style={{ background: COLORS[item.name], width: '14px', height: '14px' }} /><span style={{ fontSize: '17px' }}>{item.name}</span><strong style={{ fontSize: '24px' }}>{counts.total ? ((item.value / counts.total) * 100).toFixed(1) : '0.0'}%</strong></div>)}
              </div>
            </article>
          </section>

          <section aria-label="KPI performance by pickup group" style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(min(100%, 280px), 1fr))', gap: '18px', marginTop: '20px', marginBottom: '24px' }}>
            {GROUPS.map(group => (
              <GroupKpiChart key={group.id} group={group.id} rows={currentRows.filter(row => row._group === group.id)} />
            ))}
          </section>

          <MajorExceptionPickupChart rows={currentRows} />

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
