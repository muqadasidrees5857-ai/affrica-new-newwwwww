import { Fragment, useEffect, useMemo, useState } from 'react'
import HomePage from './Home.jsx'
import { createPortal } from 'react-dom'
import html2canvas from 'html2canvas'
import * as XLSX from 'xlsx'
import { PieChart, Pie, Cell, Tooltip, ResponsiveContainer } from 'recharts'
import {
  LayoutDashboard, Truck, Upload, Search, FileSpreadsheet,
  AlertTriangle, CircleCheck, Clock3, Eye, Printer, Camera, ChevronDown,
  Menu, X, RefreshCw, ArrowUp, ArrowDown, ArrowUpDown, Maximize2, Home, PackageCheck,
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
  { id: 'reload', label: 'Reload Warehouses' },
  { id: 'third', label: 'Third-Party Other Pickup Points' },
]

const EXCEL_COLUMNS = [
  'Client: Account Name',
  'Depart from Country',
  'Job Pickup Point',
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

// Loading KPI: days = Today - Loading Date. Change the numbers below if the Loading KPI changes.
KPI_RULES.loading = {
  drc:    { within: 5, exception: 10 },   // 0–5 | 6–10 | >10
  reload: { within: 2, exception: 3 },    // 0–2 | 3    | >3
  third:  { within: 3, exception: 5 },    // 0–3 | 4–5  | >5
}

// The two dashboard pages that open from the Home page.
const PAGES = {
  waiting: {
    title: 'Waiting to Load',
    categories: ['waiting to load', 'waiting for documents'],
    subtitle: 'Monitor Waiting To Load aging and KPI exceptions across pickup points.',
    scope: 'Waiting To Load + Waiting for Documents',
    unit: 'Waiting To Load / Documents',
  },
  loading: {
    title: 'Loading',
    categories: ['loading'],
    subtitle: 'Monitor Loading trucks: days since Loading Date (Today − Loading Date) with KPI exceptions across pickup points.',
    scope: 'Loading (Today − Loading Date)',
    unit: 'Loading',
  },
}
const ALL_CATEGORIES = [...PAGES.waiting.categories, ...PAGES.loading.categories]

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

// Loading rows: days = Today - Loading Date (Document Received Date is NOT used).
function getLoadingDays(row) {
  const start = getDate(row['Loading Date'])
  if (!start) return null

  const today = new Date()
  const startUtc = Date.UTC(start.getFullYear(), start.getMonth(), start.getDate())
  const todayUtc = Date.UTC(today.getFullYear(), today.getMonth(), today.getDate())
  return Math.max(0, Math.round((todayUtc - startUtc) / 86400000))
}

function classifyGroup(row) {
  const country = clean(row['Depart from Country']).toLowerCase()
  const pickup = clean(row['Job Pickup Point'])
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
  const operationalCategory = clean(row['Delivery Status'])
  const isLoading = operationalCategory.toLowerCase() === 'loading'
  const days = isLoading ? getLoadingDays(row) : getWaitingDays(row['Reporting Date'])

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

/* ------------------------------------------------------------------
   MAJOR EXCEPTION — CLIENT-WISE BAR CHART (horizontal bars)
   Left: the pickup point. In front of it: one bar per client account that
   has Major Exception trucks at that pickup point (count inside the bar,
   client name on the left, shared "Record Count" axis on top).
   The same client keeps the same colour everywhere in a chart.
   Only Major Exception trucks are counted; it never changes with filters.
   Works on every screen size (phones stack the pickup name above its bars).
------------------------------------------------------------------ */
const NORMAL_ROW_BUDGET = 24 // normal view shows pickup points until this many client rows are used

const CLIENT_COLORS = [
  '#2563eb', '#f59e0b', '#10b981', '#8b5cf6', '#ec4899', '#06b6d4',
  '#f97316', '#84cc16', '#6366f1', '#14b8a6', '#e11d48', '#a16207',
  '#0ea5e9', '#7c3aed', '#65a30d', '#db2777',
]

// Standing-days ageing for one client's trucks: min–max range + 5-day bands (1–5, 6–10, 11–15 ... 31+).
function ageStats(days) {
  const valid = days.filter(day => day !== null && day !== undefined)
  if (valid.length === 0) return { minDays: null, maxDays: null, bands: [] }
  const counts = new Map()
  valid.forEach(day => {
    const lo = day > 30 ? 31 : Math.floor((Math.max(day, 1) - 1) / 5) * 5 + 1
    counts.set(lo, (counts.get(lo) || 0) + 1)
  })
  const bands = [...counts.entries()]
    .sort((a, b) => a[0] - b[0])
    .map(([lo, count]) => ({ label: lo === 31 ? '31+' : `${lo}–${lo + 4}`, count }))
  return { minDays: Math.min(...valid), maxDays: Math.max(...valid), bands }
}

function formatAgeRange(minDays, maxDays) {
  if (minDays === null) return '—'
  return minDays === maxDays ? `${minDays} ${minDays === 1 ? 'day' : 'days'}` : `${minDays}–${maxDays} days`
}

function buildMajorModel(rows, group) {
  const byPickup = new Map()
  const clientTotals = new Map()

  rows
    .filter(row => row._group === group && row._kpi === 'Major Exception')
    .forEach(row => {
      const pickup = clean(row['Job Pickup Point']) || '(No pickup point)'
      const client = clean(row['Client: Account Name']) || '(No client account name)'
      if (!byPickup.has(pickup)) byPickup.set(pickup, { pickup, total: 0, clients: new Map() })
      const pickupEntry = byPickup.get(pickup)
      pickupEntry.total += 1
      if (!pickupEntry.clients.has(client)) pickupEntry.clients.set(client, { client, count: 0, days: [] })
      const clientEntry = pickupEntry.clients.get(client)
      clientEntry.count += 1
      clientEntry.days.push(row._days)
      clientTotals.set(client, (clientTotals.get(client) || 0) + 1)
    })

  // The same client keeps the same colour in every pickup block (biggest client first).
  const colorByClient = new Map(
    [...clientTotals.entries()]
      .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
      .map(([client], index) => [client, CLIENT_COLORS[index % CLIENT_COLORS.length]])
  )

  const pickups = [...byPickup.values()]
    .map(item => ({
      name: item.pickup,
      total: item.total,
      clients: [...item.clients.values()]
        .sort((a, b) => b.count - a.count || a.client.localeCompare(b.client))
        .map(entry => ({ ...entry, color: colorByClient.get(entry.client), ...ageStats(entry.days) })),
    }))
    .sort((a, b) => b.total - a.total || a.name.localeCompare(b.name))

  const totalTrucks = pickups.reduce((sum, item) => sum + item.total, 0)
  return { pickups, totalTrucks, clientCount: clientTotals.size }
}

// Round axis (0, 10, 20 ... ) with at most ~6 gridlines.
function niceAxis(max) {
  const raw = Math.max(1, max) / 5
  const magnitude = 10 ** Math.floor(Math.log10(raw))
  const normal = raw / magnitude
  const step = Math.max(1, (normal <= 1 ? 1 : normal <= 2 ? 2 : normal <= 5 ? 5 : 10) * magnitude)
  const ticks = Math.max(1, Math.ceil(max / step))
  return { step, ticks }
}

function pickNormalPickups(pickups) {
  const shown = []
  let rowsUsed = 0
  for (const pickup of pickups) {
    if (shown.length > 0 && rowsUsed + pickup.clients.length > NORMAL_ROW_BUDGET) break
    shown.push(pickup)
    rowsUsed += pickup.clients.length
  }
  return shown
}

function ClientBarChart({ pickups, rowHeight, fontSize }) {
  const maxCount = Math.max(1, ...pickups.flatMap(pickup => pickup.clients.map(client => client.count)))
  const axis = niceAxis(maxCount)
  const tickValues = Array.from({ length: axis.ticks + 1 }, (_, index) => index * axis.step)
  const axisTop = axis.ticks * axis.step

  return (
    <div className="mxc-chart" style={{ '--rh': `${rowHeight}px`, '--fs': `${fontSize}px`, '--ticks': axis.ticks }}>
      <div className="mxc-axis">
        <div className="mxc-axis-title">Record Count</div>
        <div className="mxc-axis-ticks">
          {tickValues.map((value, index) => (
            <span key={value} style={{ left: `${(index / axis.ticks) * 100}%` }}>{value}</span>
          ))}
        </div>
        <div className="mxc-axis-age">Standing days</div>
      </div>

      {pickups.map(pickup => (
        <div className="mxc-block" key={pickup.name}>
          <div className="mxc-label" style={{ gridRow: `2 / span ${pickup.clients.length}` }} title={pickup.name}>
            <span className="mxc-label-name">{pickup.name}</span>
            <span className="mxc-label-total">{pickup.total} {pickup.total === 1 ? 'truck' : 'trucks'}</span>
          </div>

          {pickup.clients.map((client, index) => {
            const percent = (client.count / axisTop) * 100
            const inside = percent >= 8
            const range = formatAgeRange(client.minDays, client.maxDays)
            const bandText = client.bands.map(band => `${band.count} truck(s): ${band.label} days`).join(' | ')
            const tip = `${pickup.name} · ${client.client}: ${client.count} ${client.count === 1 ? 'truck' : 'trucks'} standing ${range}`
            return (
              <Fragment key={client.client}>
                <div className="mxc-name" style={{ gridRow: index + 2 }} title={client.client}>{client.client}</div>
                <div className="mxc-track" style={{ gridRow: index + 2 }}>
                  <div className="mxc-bar" title={tip} style={{ width: `${percent}%`, background: client.color }} />
                  <span
                    className={`mxc-count ${inside ? 'mxc-count-in' : 'mxc-count-out'}`}
                    style={inside ? { right: `calc(${100 - percent}% + 4px)` } : { left: `calc(${percent}% + 4px)` }}
                  >
                    {client.count}
                  </span>
                </div>
                <div className="mxc-age" style={{ gridRow: index + 2 }} title={`${tip}${bandText ? ` — ${bandText}` : ''}`}>
                  <span className="mxc-range">{range}</span>
                  {client.bands.map(band => (
                    <span className="mxc-band" key={band.label}>{band.label}d<b>{band.count}</b></span>
                  ))}
                </div>
              </Fragment>
            )
          })}
        </div>
      ))}
    </div>
  )
}

// One normal-view panel (top pickup points only) with its OWN Full Screen button.
function MajorExceptionPickupChart({ model, title, pageTitle, onFullScreen }) {
  const { pickups, totalTrucks, clientCount } = model
  const shown = pickNormalPickups(pickups)
  return (
    <article className="panel major-pickup-panel" style={{ padding: '10px 12px', marginBottom: '10px', minWidth: 0 }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', gap: '8px', flexWrap: 'wrap' }}>
        <div style={{ minWidth: 0 }}>
          <h2 style={{ margin: 0, fontSize: 'clamp(17px, 2.4vw, 12px)', fontWeight: 500, lineHeight: 1.2, color: '#172033' }}>{title} Major Exceptions- {pageTitle}</h2>
          <p style={{ margin: '2px 0 6px', color: '#64748b', fontSize: '11px' }}>
            Each bar is one client account. Tags on the right show how many days the trucks have been standing.
          </p>
          <div style={{ display: 'flex', gap: '5px', flexWrap: 'wrap' }}>
            <span className="mpc-chip">{totalTrucks.toLocaleString()} Major Exception trucks</span>
            <span className="mpc-chip mpc-chip-neutral">{pickups.length.toLocaleString()} pickup points</span>
            <span className="mpc-chip mpc-chip-neutral">{clientCount.toLocaleString()} clients</span>
          </div>
        </div>
        {pickups.length > 0 && (
          <button className="secondary-button" type="button" onClick={onFullScreen} title={`Open ${title} in full screen`}>
            <Maximize2 size={13} /> Full Screen
          </button>
        )}
      </div>

      {pickups.length === 0 ? (
        <p style={{ margin: '10px 0 4px', textAlign: 'center', fontSize: '12px', color: '#64748b' }}>
          No Major Exception trucks in this group.
        </p>
      ) : (
        <div style={{ marginTop: '10px' }}>
          <ClientBarChart pickups={shown} rowHeight={20} fontSize={11} />
          {pickups.length > shown.length && (
            <p style={{ margin: '6px 0 0', textAlign: 'center', fontSize: '10px', color: '#64748b' }}>
              Top {shown.length} pickup point(s) shown · +{pickups.length - shown.length} more in Full Screen
            </p>
          )}
        </div>
      )}
    </article>
  )
}

const MAJOR_CHART_CSS = `
  .mpc-chip { display: inline-flex; align-items: center; gap: 4px; padding: 2px 8px; border-radius: 999px; background: #fef2f2; color: #b91c1c; font-size: 10px; font-weight: 600; }
  .mpc-chip-neutral { background: #f1f5f9; color: #334155; }

  .mxc-chart { --lw: clamp(84px, 14%, 180px); --nw: clamp(110px, 20%, 240px); --aw: clamp(170px, 28%, 380px); width: 100%; min-width: 0; }
  .mxc-axis, .mxc-block { display: grid; grid-template-columns: var(--lw) var(--nw) minmax(0, 1fr) var(--aw); }
  .mxc-axis { margin-bottom: 2px; }
  .mxc-axis-title { grid-column: 3; text-align: center; font-size: 10px; font-weight: 600; color: #64748b; line-height: 1.2; }
  .mxc-axis-ticks { grid-column: 3; position: relative; height: 14px; font-size: 10px; color: #64748b; }
  .mxc-axis-ticks span { position: absolute; top: 0; transform: translateX(-50%); line-height: 14px; }

  .mxc-block { grid-template-rows: auto; grid-auto-rows: minmax(var(--rh), auto); align-items: center; border-top: 1px solid #dfe6ef; padding: 4px 0; }
  .mxc-label { grid-column: 1; align-self: start; display: flex; flex-direction: column; gap: 2px; min-width: 0; padding-right: 8px; }
  .mxc-label-name { font-size: calc(var(--fs) + 1px); font-weight: 700; color: #172033; line-height: 1.2; overflow-wrap: anywhere; display: -webkit-box; -webkit-line-clamp: 3; -webkit-box-orient: vertical; overflow: hidden; }
  .mxc-label-total { font-size: calc(var(--fs) - 1px); font-weight: 600; color: #b91c1c; }
  .mxc-name { grid-column: 2; min-width: 0; padding-right: 8px; text-align: right; font-size: var(--fs); font-weight: 500; color: #334155; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
  .mxc-track { grid-column: 3; position: relative; height: 100%; border-right: 1px solid #e2e8f0; background-image: linear-gradient(to right, #e2e8f0 1px, transparent 1px); background-size: calc(100% / var(--ticks)) 100%; background-repeat: repeat-x; }
  .mxc-bar { position: absolute; left: 0; top: 50%; height: calc(var(--rh) - 4px); transform: translateY(-50%); border-radius: 0 3px 3px 0; }
  .mxc-count { position: absolute; top: 50%; transform: translateY(-50%); font-size: var(--fs); font-weight: 700; line-height: 1; }
  .mxc-count-in { color: #ffffff; text-shadow: 0 0 2px rgba(0, 0, 0, 0.4); }
  .mxc-count-out { color: #172033; }
  .mxc-axis-age { grid-column: 4; padding-left: 10px; font-size: 10px; font-weight: 600; color: #64748b; line-height: 1.2; text-align: left; }
  .mxc-age { grid-column: 4; min-width: 0; display: flex; align-items: center; gap: 4px; padding-left: 10px; overflow: hidden; white-space: nowrap; font-size: calc(var(--fs) - 1px); }
  .mxc-range { flex: none; padding: 1px 8px; border-radius: 999px; background: #fef2f2; color: #b91c1c; font-weight: 700; }
  .mxc-band { flex: none; display: inline-flex; align-items: center; gap: 4px; padding: 1px 6px; border-radius: 6px; background: #f1f5f9; color: #475569; font-weight: 500; }
  .mxc-band b { color: #172033; font-weight: 700; }

  @media (max-width: 640px) {
    .mxc-chart { --lw: 0px; --nw: clamp(90px, 32%, 160px); --aw: 96px; }
    .mxc-band { display: none; }
    .mxc-label { grid-column: 1 / -1 !important; grid-row: 1 !important; flex-direction: row; align-items: baseline; flex-wrap: wrap; gap: 8px; padding: 0 0 3px; }
  }
`

// The three charts. Every chart has its own Full Screen button (shows only that group),
// and there is also one button on top that opens all three groups together.
function MajorExceptionSection({ rows, pageTitle }) {
  const models = useMemo(
    () => GROUPS.map(group => ({ ...group, ...buildMajorModel(rows, group.id) })),
    [rows]
  )
  const grandTotal = models.reduce((sum, item) => sum + item.totalTrucks, 0)
  // fullMode: null = closed, 'all' = all three groups, or one group id ('drc' / 'reload' / 'third').
  const [fullMode, setFullMode] = useState(null)
  const full = fullMode !== null
  const single = full && fullMode !== 'all'
  const [viewport, setViewport] = useState(() => ({ w: window.innerWidth, h: window.innerHeight }))

  useEffect(() => {
    if (!full) return undefined
    const previousOverflow = document.body.style.overflow
    const onKey = event => { if (event.key === 'Escape') setFullMode(null) }
    const onResize = () => setViewport({ w: window.innerWidth, h: window.innerHeight })
    onResize()
    document.body.style.overflow = 'hidden'
    window.addEventListener('keydown', onKey)
    window.addEventListener('resize', onResize)
    return () => {
      document.body.style.overflow = previousOverflow
      window.removeEventListener('keydown', onKey)
      window.removeEventListener('resize', onResize)
    }
  }, [full])

  const closeFull = () => {
    setFullMode(null)
    leaveBrowserFullscreen()
  }
  const openFull = mode => {
    setFullMode(mode)
    enterBrowserFullscreen()
  }

  const fullModels = useMemo(
    () => (single ? models.filter(item => item.id === fullMode) : models),
    [models, single, fullMode]
  )
  // Wide screens: the three groups sit side by side. Small screens: one under the other.
  const fullCols = fullMode === 'all' && viewport.w >= 1100 ? 3 : 1

  // Row height / font size for the Full Screen view, calculated from the screen so nothing scrolls.
  const fullLayout = useMemo(() => {
    const availableHeight = (viewport.h - 4 - 46) * 0.97 // 1px margins + top bar
    const axisHeight = 34
    const blockPad = 10
    const rowsOf = item => item.pickups.reduce((sum, pickup) => sum + Math.max(3, pickup.clients.length), 0)
    const nonEmpty = fullModels.filter(item => item.pickups.length > 0)
    let rowHeight = 34

    if (nonEmpty.length > 0) {
      if (single) {
        const item = nonEmpty[0]
        rowHeight = (availableHeight - axisHeight - item.pickups.length * blockPad) / rowsOf(item)
      } else if (fullCols === 3) {
        rowHeight = Math.min(...nonEmpty.map(item => (availableHeight - 40 - axisHeight - item.pickups.length * blockPad - 6) / rowsOf(item)))
      } else {
        const pickupCount = nonEmpty.reduce((sum, item) => sum + item.pickups.length, 0)
        const totalRows = nonEmpty.reduce((sum, item) => sum + rowsOf(item), 0)
        const emptyCount = fullModels.length - nonEmpty.length
        const fixed = nonEmpty.length * (30 + axisHeight + 10) + pickupCount * blockPad + emptyCount * 48
        rowHeight = (availableHeight - fixed) / totalRows
      }
    }

    rowHeight = Math.max(10, Math.min(single ? 44 : 34, Math.floor(rowHeight)))
    const fontSize = Math.max(8, Math.min(17, Math.round(rowHeight * 5.5) / 10))
    return { rowHeight, fontSize }
  }, [fullModels, single, fullCols, viewport])

  const activeModel = single ? fullModels[0] : null

  return (
    <section aria-label="Major exceptions by pickup group">
      <style>{MAJOR_CHART_CSS}</style>

      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: '8px', flexWrap: 'wrap', margin: '0 0 8px' }}>
        <div style={{ minWidth: 0 }}>
<section className="page-heading">

  <div>
    <h2 className='new'>Major Exceptions by Pickup Group</h2>
  </div>
</section>
          <p style={{ margin: '2px 0 0', fontSize: '11px', color: '#64748b' }}>{grandTotal.toLocaleString()} Major Exception trucks across DRC, Reload and Third Party.</p>
        </div>
        
        {grandTotal > 0 && (
          <button className="secondary-button" type="button" onClick={() => openFull('all')}><Maximize2 size={13} /> All 3 Full Screen</button>
        )}
      </div>

      {models.map(item => (
        <MajorExceptionPickupChart key={item.id} model={item} title={item.label} pageTitle={pageTitle} onFullScreen={() => openFull(item.id)} />
      ))}

      {full && createPortal(
        <div role="dialog" aria-modal="true" aria-label="Major Exceptions full screen"
          style={{ position: 'fixed', inset: 0, zIndex: 3000, background: '#f5f7fb', padding: '1px', boxSizing: 'border-box', overflow: 'hidden', display: 'flex', flexDirection: 'column' }}>
          <style>{MAJOR_CHART_CSS}</style>
          <div style={{ background: '#ffffff', borderRadius: 0, border: 'none', padding: '1px', height: '100%', minHeight: 0, boxSizing: 'border-box', display: 'flex', flexDirection: 'column', gap: '2px' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: '8px', minHeight: '34px', flexWrap: 'wrap', padding: '2px 8px' }}>
              <h2 style={{ margin: 0, fontSize: 'clamp(16px, 2.4vw, 24px)', fontWeight: 800, lineHeight: 1.2, color: '#172033' }}>
                {single ? `${activeModel.label} — Major Exceptions` : 'Major Exceptions — All Pickup Groups'}
                <span className="mpc-chip" style={{ marginLeft: '12px', fontSize: '12px', verticalAlign: 'middle' }}>
                  {(single ? activeModel.totalTrucks : grandTotal).toLocaleString()} trucks
                </span>
                {single && <span className="mpc-chip mpc-chip-neutral" style={{ marginLeft: '6px', fontSize: '12px', verticalAlign: 'middle' }}>{activeModel.pickups.length.toLocaleString()} pickup points</span>}
                {single && <span className="mpc-chip mpc-chip-neutral" style={{ marginLeft: '6px', fontSize: '12px', verticalAlign: 'middle' }}>{activeModel.clientCount.toLocaleString()} clients</span>}
              </h2>
              <button className="secondary-button" type="button" onClick={closeFull}><X size={13} /> Exit Full Screen</button>
            </div>

            <div style={{
              flex: '1 1 auto', minHeight: 0, overflowY: 'auto', alignContent: 'start',
              display: fullCols === 3 ? 'grid' : 'flex', flexDirection: 'column',
              gridTemplateColumns: 'repeat(3, minmax(0, 1fr))', gap: fullCols === 3 ? '10px' : '6px', padding: '0 4px',
            }}>
              {fullModels.map(item => (
                <div key={item.id} style={{ minWidth: 0 }}>
                  {!single && (
                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap', minHeight: fullCols === 3 ? '40px' : '30px' }}>
                      <h3 style={{ margin: 0, fontSize: fullCols === 3 ? '16px' : '20px', fontWeight: 800, lineHeight: 1.2, color: '#172033' }}>{item.label}</h3>
                      <span className="mpc-chip">{item.totalTrucks.toLocaleString()} trucks</span>
                      {fullCols !== 3 && <span className="mpc-chip mpc-chip-neutral">{item.pickups.length.toLocaleString()} pickup points</span>}
                      {fullCols !== 3 && <span className="mpc-chip mpc-chip-neutral">{item.clientCount.toLocaleString()} clients</span>}
                    </div>
                  )}
                  {item.pickups.length === 0
                    ? <p style={{ margin: 0, fontSize: '12px', color: '#64748b' }}>No Major Exception trucks in this group.</p>
                    : <ClientBarChart pickups={item.pickups} rowHeight={fullLayout.rowHeight} fontSize={fullLayout.fontSize} />}
                </div>
              ))}
            </div>
          </div>
        </div>,
        document.body
      )}
    </section>
  )
}


function RecordTable({ sectionId, title, description, rows, onDetails, options, columnFilters, onColumnFilter, statusFilter, onStatusFilter }) {
  const [expanded, setExpanded] = useState(false)
  const [sort, setSort] = useState(null)
  const sortedRows = useMemo(
    () => (sort ? [...rows].sort((a, b) => compareRows(a, b, sort.column, sort.direction)) : rows),
    [rows, sort]
  )
  const visibleRows = expanded ? sortedRows : sortedRows.slice(0, 8)
  const [full, setFull] = useState(false)

  const closeFull = () => {
    setFull(false)
    leaveBrowserFullscreen()
  }
  const openFull = () => {
    setFull(true)
    enterBrowserFullscreen()
  }

  useEffect(() => {
    if (!full) return undefined
    const previousBodyOverflow = document.body.style.overflow
    const previousHtmlOverflow = document.documentElement.style.overflow
    const onKey = event => { if (event.key === 'Escape') closeFull() }
    const onFsChange = () => {
      if (!document.fullscreenElement && full) {
        setFull(false)
      }
    }
    document.body.style.overflow = 'hidden'
    document.documentElement.style.overflow = 'hidden'
    window.addEventListener('keydown', onKey)
    document.addEventListener('fullscreenchange', onFsChange)
    return () => {
      document.body.style.overflow = previousBodyOverflow
      document.documentElement.style.overflow = previousHtmlOverflow
      window.removeEventListener('keydown', onKey)
      document.removeEventListener('fullscreenchange', onFsChange)
    }
  }, [full])

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

  const renderTable = isFull => (
      <div className={`table-wrap full-columns-table-wrap ${isFull ? 'fs-table-wrap' : ''}`}>
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
                {column === 'Job Pickup Point' ? 'Job Pickup Point' : column}
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
            {(isFull ? sortedRows : visibleRows).map(row => (
              <tr key={row._id}>
                <td className="days-count-cell" style={{ fontWeight: 700, textAlign: 'center', fontVariantNumeric: 'tabular-nums' }}>
                  {row._days === null ? '—' : `${row._days} ${row._days === 1 ? 'day' : 'days'}`}
                </td>
                {TABLE_COLUMNS.map(column => (
                  <td key={column} className={column === 'Delivery Number' ? 'delivery-cell' : column === 'Job Pickup Point' ? 'pickup-cell' : ''} title={clean(row[column])}>
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
  )

  return (
    <section className="records-section" id={sectionId}>
      <div className="records-heading">
        <div>
          <div className="records-title-row"><h2>{title}</h2><span className="record-count">{rows.length.toLocaleString()} records</span></div>
          <p>{description}</p>
        </div>
        <div className="records-heading-actions">
          <button className="secondary-button" onClick={openFull} type="button" title="Open this table in full screen"><Maximize2 size={15} /> Full Screen</button>
          <button className="table-screenshot-button" onClick={downloadFullTableScreenshot} type="button" title="Download only this table as a PNG, including every row">
            <Camera size={15} /> Table Only PNG
          </button>
          <button className="secondary-button" onClick={() => setExpanded(!expanded)} type="button">
            {expanded ? 'Show Less' : 'Expand Table'}
            <ChevronDown size={16} className={expanded ? 'rotate-icon' : ''} />
          </button>
        </div>
      </div>

      {renderTable(false)}

      {rows.length > 8 && <button className="show-all-button" onClick={() => setExpanded(!expanded)}>{expanded ? 'Show fewer records' : `Show all ${rows.length} records`}</button>}

      {full && createPortal(
        <div role="dialog" aria-modal="true" aria-label={`${title} table full screen`}
          className="table-fullscreen-modal">
          <div className="table-fullscreen-container">
            <div className="table-fullscreen-header">
              <div>
                <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
                  <h2 style={{ margin: 0, fontSize: '20px', fontWeight: 800, lineHeight: 1.2, color: '#172033' }}>{title}</h2>
                  <span className="record-count">{rows.length.toLocaleString()} records</span>
                </div>
                <p style={{ margin: '3px 0 0', fontSize: '12px', color: '#64748b' }}>{description}</p>
              </div>
              <button className="secondary-button" type="button" onClick={closeFull}><X size={14} /> Exit Full Screen</button>
            </div>
            <div className="table-fullscreen-body">
              {renderTable(true)}
            </div>
          </div>
        </div>,
        document.body
      )}
    </section>
  )
}

function RecordModal({ row, onClose }) {
  if (!row) return null
  return (
    <div className="modal-backdrop" style={{ zIndex: 4000 }} onClick={onClose}>
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

// Page-wide responsive rules (phones, tablets, laptops, big screens).
const PAGE_RESPONSIVE_CSS = `
  .app-shell, .main-area, .dashboard-content { min-width: 0; }
  html body .app-shell .main-area main.dashboard-content { width: 100% !important; max-width: none !important; margin-left: 0 !important; margin-right: 0 !important; padding-left: clamp(14px, 2vw, 28px) !important; padding-right: clamp(14px, 2vw, 28px) !important; box-sizing: border-box; }
  .summary-grid { grid-template-columns: repeat(auto-fit, minmax(210px, 1fr)); }

  @media (max-width: 1100px) {
    .page-heading { flex-direction: column; align-items: stretch !important; gap: 12px; }
    .tables-intro { flex-direction: column; align-items: flex-start !important; gap: 12px; }
    .dashboard-export-buttons { flex-wrap: wrap; }
    .records-heading { flex-wrap: wrap; gap: 10px; }
    .records-heading-actions { flex-wrap: wrap; }
    .full-columns-table-wrap { overflow-x: auto !important; -webkit-overflow-scrolling: touch; }
    .records-section .wide-operations-table { min-width: 1100px; }
    .chart-legend { flex-wrap: wrap; }
  }

  @media (max-width: 640px) {
    .summary-grid { grid-template-columns: 1fr !important; }
    .topbar { gap: 8px; }
    .topbar-title { font-size: 13px; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    .upload-button { padding-left: 10px !important; padding-right: 10px !important; font-size: 13px !important; }
    .page-heading h1 { font-size: 24px !important; }
    .file-status { width: 100%; box-sizing: border-box; }
    .filter-toolbar > * { flex: 1 1 100% !important; min-width: 0 !important; width: 100%; }
    .filter-toolbar .filter-select select { flex: 1; min-width: 0; }
    .details-modal { width: calc(100vw - 20px) !important; max-width: none !important; max-height: 92vh; overflow-y: auto; }
    .details-grid { grid-template-columns: 1fr !important; }
    .overall-action-summary { flex-wrap: wrap; }
  }
`

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
  // 'home' (Home page) | 'waiting' | 'loading'. The file stays in memory while you switch pages.
  const [page, setPage] = useState('home')

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
  // allRows = Waiting To Load + Waiting for Documents + Loading (used by the Home page).
  const allRows = useMemo(
    () => rows.map((row, index) => prepareRow(row, index)),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [rows, todayKey]
  )

  const pageConfig = PAGES[page] || PAGES.waiting

  // Rows of the dashboard page that is open (Waiting To Load page or Loading page).
  const currentRows = useMemo(
    () => allRows.filter(row => pageConfig.categories.includes(clean(row._category).toLowerCase())),
    [allRows, pageConfig]
  )

  // Truck counts shown on the two buttons of the Home page.
  const homeStats = useMemo(() => ({
    waiting: { total: allRows.filter(row => PAGES.waiting.categories.includes(clean(row._category).toLowerCase())).length },
    loading: { total: allRows.filter(row => PAGES.loading.categories.includes(clean(row._category).toLowerCase())).length },
  }), [allRows])

  async function handleFileUpload(event) {
    const file = event.target.files?.[0]
    if (!file) return
    setLoading(true)
    setError('')
    try {
      const buffer = await file.arrayBuffer()
      // raw: true keeps CSV dates as text (DD/MM/YYYY) so SheetJS cannot swap day and month.
      const workbook = XLSX.read(buffer, { type: 'array', raw: true })
      const sheet = workbook.Sheets[workbook.SheetNames[0]]
      if (!sheet) throw new Error('The uploaded workbook has no readable worksheet.')
      const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: '' })
      if (rawRows.length === 0) throw new Error('The selected worksheet contains no records.')

      const headers = Object.keys(rawRows[0])
      const required = ['Delivery Status', 'Depart from Country', 'Job Pickup Point', 'Reporting Date']
      const missing = required.filter(header => !headers.includes(header))
      if (missing.length) throw new Error(`Missing required Excel columns: ${missing.join(', ')}`)

      // Keep raw rows only (Waiting To Load, Waiting for Documents and Loading); prepareRow runs in allRows.
      const keptRows = rawRows.filter(row =>
        ALL_CATEGORIES.includes(clean(row['Delivery Status']).toLowerCase())
      )
      setRows(keptRows)
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

  // Switch between Home / Waiting To Load / Loading. The uploaded data is NOT cleared.
  function openPage(next) {
    resetFilters()
    setSidebarOpen(false)
    setPage(next)
    window.scrollTo({ top: 0 })
  }

  // Home page (shown first). All hooks are above, so this early return is safe.
  if (page === 'home') {
    return (
      <HomePage
        fileName={fileName}
        loading={loading}
        error={error}
        allRows={allRows}
        stats={homeStats}
        onUpload={handleFileUpload}
        onOpen={openPage}
      />
    )
  }

  return (
    <div className="app-shell">
      <style>{PAGE_RESPONSIVE_CSS}</style>
      {sidebarOpen && <button className="sidebar-scrim" aria-label="Close navigation" onClick={() => setSidebarOpen(false)} />}
      <aside className={`sidebar ${sidebarOpen ? 'sidebar-open' : ''}`}>
        <div className="sidebar-brand">
          <div className="brand-mark">A</div>
          <div><strong>Africa Operations</strong><span>KPI MONITORING</span></div>
          <button className="mobile-close" aria-label="Close navigation" onClick={() => setSidebarOpen(false)}><X size={18} /></button>
        </div>
        <p className="sidebar-label">PAGES</p>
        <button className="nav-item" onClick={() => openPage('home')}><Home size={18} /> Home</button>
        <button className={`nav-item ${page === 'waiting' ? 'active' : ''}`} onClick={() => openPage('waiting')}><Clock3 size={18} /> Waiting To Load</button>
        <button className={`nav-item ${page === 'loading' ? 'active' : ''}`} onClick={() => openPage('loading')}><PackageCheck size={18} /> Loading</button>
        <p className="sidebar-label">THIS PAGE</p>
        <button className="nav-item" onClick={() => scrollTo('overview')}><LayoutDashboard size={18} /> Overview</button>
        <button className="nav-item" onClick={() => scrollTo('table-drc')}><Truck size={18} /> DRC</button>
        <button className="nav-item" onClick={() => scrollTo('table-reload')}><RefreshCw size={18} /> Reload Warehouses</button>
        <button className="nav-item" onClick={() => scrollTo('table-third')}><FileSpreadsheet size={18} /> Third Party</button>
        <button className="nav-item" onClick={() => scrollTo('action-summary')}><AlertTriangle size={18} /> Action Needed</button>
        <div className="sidebar-bottom"><span className="sidebar-status-dot" /><div><strong>Dashboard Ready</strong><small>{pageConfig.scope}</small></div></div>
      </aside>

      <div className="main-area">
        <header className="topbar">
          <button className="mobile-menu" aria-label="Toggle navigation" aria-expanded={sidebarOpen} onClick={() => setSidebarOpen(open => !open)} type="button"><Menu size={21} /></button>
          <div className="topbar-title"><span>Operations /</span> {pageConfig.title}</div>
          <button className="home-back-button" type="button" onClick={() => openPage('home')}><Home size={16} /> Home</button>
          <label className="upload-button"><Upload size={17} />{loading ? 'Reading file...' : 'Upload Excel File'}<input type="file" accept=".xlsx,.xls,.csv" hidden disabled={loading} onChange={handleFileUpload} /></label>
        </header>

        <main className="dashboard-content" id="overview">
          <section className="page-heading">
            <div><p className="eyebrow">OPERATIONS OVERVIEW</p><h1>{pageConfig.title}</h1><p className="subtitle">{pageConfig.subtitle}</p></div>
            <div className="file-status"><span className={`status-dot ${fileName ? 'status-dot-ready' : ''}`} /><div><strong>{fileName || 'No file uploaded'}</strong><p>{fileName ? `${currentRows.length.toLocaleString()} ${pageConfig.unit} records on this page` : 'Upload an Excel file to begin'}</p></div></div>
          </section>

          {error && <div className="error-banner"><AlertTriangle size={18} /><span>{error}</span></div>}
          <div className="scope-note"><Clock3 size={17} /><span>Current view: <strong>{pageConfig.scope}</strong>.</span></div>

          <section className="summary-grid">
            <SummaryCard icon={Truck} title="All Truck Performance" value={counts.total} note="All operational records" color="navy" />
            <SummaryCard icon={CircleCheck} title="Within KPI" value={counts.within} note="Within target" color="green" />
            <SummaryCard icon={Clock3} title="Exception" value={counts.exception} note="Needs attention" color="orange" />
            <SummaryCard icon={AlertTriangle} title="Major Exception" value={counts.major} note="Action required" color="red" />
          </section>

          <section className="charts-grid">
            <article className="panel chart-panel pie-panel" style={{ gridColumn: '1 / -1' }}>
              <div className="panel-heading"><div><h2>Overall KPI - {pageConfig.title}</h2><p>KPI status distribution</p></div><span className="panel-icon"><LayoutDashboard size={18} /></span></div>
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
                {pieData.map(item => <div className="legend-item" key={item.name}><span className="legend-dot" style={{ background: COLORS[item.name], width: '14px', height: '14px' }} /><span style={{ fontSize: '12px' }}>{item.name}</span><strong style={{ fontSize: '12px' }}>{counts.total ? ((item.value / counts.total) * 100).toFixed(1) : '0.0'}%</strong></div>)}
              </div>
            </article>
          </section>

          <section aria-label="KPI performance by pickup group" style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(min(100%, 280px), 1fr))', gap: '18px', marginTop: '20px', marginBottom: '24px' }}>
            {GROUPS.map(group => (
              <GroupKpiChart key={group.id} group={group.id} rows={currentRows.filter(row => row._group === group.id)} />
            ))}
          </section>

          <MajorExceptionSection rows={currentRows} pageTitle={pageConfig.title} />

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
                ? 'Non-DRC records where Job Pickup Point contains the word Reload.'
                : 'Non-DRC records where the pickup-point name does not contain Reload.'
            return (
              <RecordTable
                key={group.id}
                sectionId={`table-${group.id}`}
                title={`${group.label} - ${pageConfig.title}`}
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
