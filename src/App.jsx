import React, { useMemo, useState } from 'react'
import * as XLSX from 'xlsx'

const DEFAULT_BOX_PCT = 0.0325
const DEFAULT_REVERSAL = 3
const DEFAULT_LOOKBACK = 520

function tryParseDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value
  if (typeof value === 'number') {
    const parsed = XLSX.SSF.parse_date_code(value)
    if (parsed) return new Date(parsed.y, parsed.m - 1, parsed.d)
  }
  if (typeof value === 'string') {
    const trimmed = value.trim()
    if (!trimmed) return null
    const iso = trimmed.match(/^(\d{4})-(\d{2})-(\d{2})/)
    if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]))
    const eu = trimmed.match(/^(\d{2})\.(\d{2})\.(\d{4})/)
    if (eu) return new Date(Number(eu[3]), Number(eu[2]) - 1, Number(eu[1]))
    const native = new Date(trimmed)
    if (!Number.isNaN(native.getTime())) return native
  }
  return null
}

function formatDate(date) {
  if (!date) return '—'
  const y = date.getFullYear()
  const m = `${date.getMonth() + 1}`.padStart(2, '0')
  const d = `${date.getDate()}`.padStart(2, '0')
  return `${y}-${m}-${d}`
}

function weekEndingFriday(date) {
  const d = new Date(date)
  const jsDay = d.getDay()
  const mondayBased = jsDay === 0 ? 7 : jsDay
  d.setDate(d.getDate() + (5 - mondayBased))
  d.setHours(0, 0, 0, 0)
  return d
}

function boxIdxFloor(v, stepLn) {
  if (!(v > 0) || !(stepLn > 0)) return 0
  return Math.floor(Math.log(v) / stepLn)
}

function boxIdxCeil(v, stepLn) {
  if (!(v > 0) || !(stepLn > 0)) return 0
  const x = Math.log(v) / stepLn
  return Number.isInteger(x) ? x : Math.floor(x) + 1
}

function invertCode(code) {
  const c = String(code || '').trim().toUpperCase()
  if (c === 'BX') return 'SO'
  if (c === 'SO') return 'BX'
  if (c === 'BO') return 'SX'
  if (c === 'SX') return 'BO'
  return '--'
}

function isValidCode(code) {
  return ['BO', 'BX', 'SO', 'SX'].includes(String(code || '').trim().toUpperCase())
}

function codePoints(code) {
  const c = String(code || '').trim().toUpperCase()
  if (c === 'BX') return 1
  if (c === 'BO') return 0.5
  if (c === 'SO') return -0.5
  if (c === 'SX') return -1
  return 0
}

function parseTickerName(header) {
  const s = String(header || '').trim()
  if (!s) return { ticker: '—', name: '—' }
  if (s.includes(' - ')) {
    const [ticker, ...rest] = s.split(' - ')
    return { ticker: ticker.trim(), name: rest.join(' - ').trim() || ticker.trim() }
  }
  if (s.includes(' — ')) {
    const [ticker, ...rest] = s.split(' — ')
    return { ticker: ticker.trim(), name: rest.join(' — ').trim() || ticker.trim() }
  }
  const [first, ...rest] = s.split(' ')
  if (/^[A-Za-z0-9.-]{1,12}$/.test(first)) {
    return { ticker: first.trim(), name: rest.join(' ').trim() || first.trim() }
  }
  return { ticker: s, name: s }
}

function rsPnfCode(weeklyPrices, w1, w2, a, b, boxPct, reversal) {
  try {
    if (!weeklyPrices[a] || !weeklyPrices[b]) return '--'
    if (w2 - w1 < 12) return '--'

    const stepLn = Math.log(1 + boxPct)
    if (!(stepLn > 0)) return '--'

    let firstW = -1
    let firstRs = 0
    for (let w = w1; w <= w2; w += 1) {
      const pA = weeklyPrices[a][w]
      const pB = weeklyPrices[b][w]
      if (pA > 0 && pB > 0) {
        firstW = w
        firstRs = pA / pB
        break
      }
    }
    if (firstW === -1 || w2 - firstW < 12) return '--'

    const b0 = boxIdxFloor(firstRs, stepLn)
    const colType = []
    const colHigh = []
    const colLow = []
    let curType = 0
    let curHigh = b0
    let curLow = b0

    for (let w = firstW + 1; w <= w2; w += 1) {
      const pA = weeklyPrices[a][w]
      const pB = weeklyPrices[b][w]
      if (!(pA > 0) || !(pB > 0)) continue
      const rs = pA / pB
      const biUp = boxIdxFloor(rs, stepLn)
      const biDown = boxIdxCeil(rs, stepLn)

      if (curType === 0) {
        if (biUp >= b0 + 1) {
          curType = 1
          curHigh = biUp
          curLow = b0
          colType.push(1)
          colHigh.push(curHigh)
          colLow.push(curLow)
        } else if (biDown <= b0 - 1) {
          curType = -1
          curLow = biDown
          curHigh = b0
          colType.push(-1)
          colHigh.push(curHigh)
          colLow.push(curLow)
        }
      } else if (curType === 1) {
        if (biUp > curHigh) {
          curHigh = biUp
          colHigh[colHigh.length - 1] = curHigh
        } else if (biDown <= curHigh - reversal) {
          curType = -1
          curLow = biDown
          colType.push(-1)
          colHigh.push(curHigh - 1)
          colLow.push(curLow)
        }
      } else if (curType === -1) {
        if (biDown < curLow) {
          curLow = biDown
          colLow[colLow.length - 1] = curLow
        } else if (biUp >= curLow + reversal) {
          curType = 1
          curHigh = biUp
          colType.push(1)
          colHigh.push(curHigh)
          colLow.push(curLow + 1)
        }
      }
    }

    if (!colType.length) return '--'

    let lastSignal = 'S'
    let hasPrevX = false
    let hasPrevO = false
    let prevXHigh = 0
    let prevOLow = 0

    for (let i = 0; i < colType.length; i += 1) {
      if (colType[i] === 1) {
        if (!hasPrevX) {
          hasPrevX = true
          prevXHigh = colHigh[i]
        } else {
          if (colHigh[i] > prevXHigh) lastSignal = 'B'
          prevXHigh = colHigh[i]
        }
      } else if (colType[i] === -1) {
        if (!hasPrevO) {
          hasPrevO = true
          prevOLow = colLow[i]
        } else {
          if (colLow[i] < prevOLow) lastSignal = 'S'
          prevOLow = colLow[i]
        }
      }
    }

    return `${lastSignal}${colType[colType.length - 1] === 1 ? 'X' : 'O'}`
  } catch {
    return '--'
  }
}

function buildWeeklyCloses(rows, headers, lastDataIndex) {
  const weekMap = new Map()
  for (let i = 0; i <= lastDataIndex; i += 1) {
    const row = rows[i]
    const date = tryParseDate(row[0])
    if (!date) continue
    const friday = weekEndingFriday(date)
    weekMap.set(friday.getTime(), i)
  }

  const sortedKeys = [...weekMap.keys()].sort((a, b) => a - b)
  const weekDates = []
  const weeklyPrices = headers.map(() => [])

  sortedKeys.forEach((key, weekIndex) => {
    const rowIndex = weekMap.get(key)
    const row = rows[rowIndex]
    weekDates[weekIndex] = new Date(Number(key))
    for (let asset = 0; asset < headers.length; asset += 1) {
      const value = Number(row[asset + 1])
      weeklyPrices[asset][weekIndex] = Number.isFinite(value) ? value : 0
    }
  })

  return { weekDates, weeklyPrices }
}

function computeMatrix(dataset, selectedDate, boxPct, reversal, lookbackWeeks) {
  const { headers, rows } = dataset
  let lastIdx = rows.length - 1

  if (selectedDate) {
    const target = formatDate(new Date(selectedDate))
    const found = rows.findIndex((row) => {
      const d = tryParseDate(row[0])
      return d && formatDate(d) === target
    })
    if (found >= 0) lastIdx = found
  }

  const { weekDates, weeklyPrices } = buildWeeklyCloses(rows, headers, lastIdx)
  if (weekDates.length < 20) {
    throw new Error(`Недостаточно weekly точек: ${weekDates.length}`)
  }

  let w1 = 0
  const w2 = weekDates.length - 1
  if (lookbackWeeks > 0 && weekDates.length > lookbackWeeks) {
    w1 = weekDates.length - lookbackWeeks
    if (w2 - w1 < 12) w1 = 0
  }

  const n = headers.length
  const codes = Array.from({ length: n }, () => Array.from({ length: n }, () => '--'))
  for (let i = 0; i < n; i += 1) codes[i][i] = '—'

  for (let i = 0; i < n - 1; i += 1) {
    for (let j = i + 1; j < n; j += 1) {
      const code = rsPnfCode(weeklyPrices, w1, w2, i, j, boxPct, reversal)
      codes[i][j] = code
      codes[j][i] = invertCode(code)
    }
  }

  for (let i = 0; i < n - 1; i += 1) {
    for (let j = i + 1; j < n; j += 1) {
      const a = String(codes[i][j]).trim().toUpperCase()
      const b = String(codes[j][i]).trim().toUpperCase()
      if (isValidCode(a)) codes[j][i] = invertCode(a)
      else if (isValidCode(b)) codes[i][j] = invertCode(b)
    }
  }

  const stats = headers.map((header, i) => {
    let buys = 0
    let xs = 0
    let pts = 0
    for (let j = 0; j < n; j += 1) {
      if (i === j) continue
      const code = String(codes[i][j]).trim().toUpperCase()
      if (code.startsWith('B')) buys += 1
      if (code.endsWith('X')) xs += 1
      pts += codePoints(code)
    }
    const denom = Math.max(1, 2 * (n - 1))
    const tech = Math.max(0, Math.min(5, 5 * ((pts + (n - 1)) / denom)))
    const parsed = parseTickerName(header)
    return {
      index: i,
      header,
      ticker: parsed.ticker,
      name: parsed.name,
      buys,
      xs,
      total: n - 1,
      tech: Number(tech.toFixed(2)),
      pts,
    }
  })

  const order = [...stats].sort((a, b) => {
    if (a.pts !== b.pts) return a.pts - b.pts
    if (a.buys !== b.buys) return a.buys - b.buys
    if (a.xs !== b.xs) return a.xs - b.xs
    return a.ticker.localeCompare(b.ticker)
  })

  return {
    weekDates,
    asOfDate: weekDates[w2],
    weeksUsed: w2 - w1 + 1,
    codes,
    order,
    stats,
  }
}

function readWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = (event) => {
      try {
        const workbook = XLSX.read(event.target.result, { type: 'array', cellDates: true })
        const sheetName = workbook.SheetNames[0]
        const sheet = workbook.Sheets[sheetName]
        const json = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null })
        if (!json.length || json[0].length < 3) {
          reject(new Error('В Excel нужен первый лист: DateTime в колонке A и далее активы в колонках B..'))
          return
        }
        const headers = json[0].slice(1).map((v, idx) => String(v || `Asset${idx + 1}`).trim())
        const rows = json.slice(1).filter((row) => row && row.some((cell) => cell !== null && cell !== ''))
        resolve({ headers, rows, sheetName, fileName: file.name })
      } catch (error) {
        reject(error)
      }
    }
    reader.onerror = () => reject(new Error('Не удалось прочитать файл'))
    reader.readAsArrayBuffer(file)
  })
}

function codeClass(code) {
  const c = String(code || '').toUpperCase()
  if (c === 'BX') return 'bx'
  if (c === 'BO') return 'bo'
  if (c === 'SO') return 'so'
  if (c === 'SX') return 'sx'
  return 'neutral'
}

function downloadCsv(result) {
  const ordered = result.order
  const rows = []
  rows.push([
    'RANK', 'TICKER', 'NAME', 'BUYS', "X'S", 'TOTAL', 'TECH SCORE',
    ...ordered.map((item) => item.ticker),
  ])
  ordered.forEach((item, index) => {
    rows.push([
      index + 1,
      item.ticker,
      item.name,
      item.buys,
      item.xs,
      item.total,
      item.tech,
      ...ordered.map((col) => result.codes[item.index][col.index]),
    ])
  })

  const csv = rows
    .map((row) => row.map((cell) => `"${String(cell ?? '').replaceAll('"', '""')}"`).join(','))
    .join('\n')

  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = `rs-matrix-${formatDate(result.asOfDate)}.csv`
  a.click()
  URL.revokeObjectURL(url)
}

export default function App() {
  const [dataset, setDataset] = useState(null)
  const [selectedDate, setSelectedDate] = useState('')
  const [boxPct, setBoxPct] = useState(DEFAULT_BOX_PCT)
  const [reversal, setReversal] = useState(DEFAULT_REVERSAL)
  const [lookbackWeeks, setLookbackWeeks] = useState(DEFAULT_LOOKBACK)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState('')

  const result = useMemo(() => {
    if (!dataset) return null
    try {
      setError('')
      return computeMatrix(dataset, selectedDate, Number(boxPct), Number(reversal), Number(lookbackWeeks))
    } catch (e) {
      setError(e.message || 'Ошибка расчёта')
      return null
    }
  }, [dataset, selectedDate, boxPct, reversal, lookbackWeeks])

  const handleFile = async (event) => {
    const file = event.target.files?.[0]
    if (!file) return
    setLoading(true)
    setError('')
    try {
      const parsed = await readWorkbook(file)
      setDataset(parsed)
      setSelectedDate('')
    } catch (e) {
      setError(e.message || 'Ошибка загрузки файла')
      setDataset(null)
    } finally {
      setLoading(false)
    }
  }

  return (
    <div className="page">
      <div className="topbar">
        <div>
          <div className="eyebrow">RS MATRIX</div>
          <h1>Оценочная матрица активов в стиле Nasdaq</h1>
          <p>
            Загрузка Excel, weekly Friday closes, P&amp;F RS-коды BX / BO / SO / SX,
            ранжирование от слабых к сильным.
          </p>
        </div>
        <a className="sample-link" href="/sample-data.xlsx" download>
          Скачать пример Excel
        </a>
      </div>

      <div className="panel controls-panel">
        <div className="control-grid">
          <label className="field upload-field">
            <span>Excel файл</span>
            <input type="file" accept=".xlsx,.xls" onChange={handleFile} />
          </label>

          <label className="field">
            <span>Дата расчёта</span>
            <input type="date" value={selectedDate} onChange={(e) => setSelectedDate(e.target.value)} />
          </label>

          <label className="field">
            <span>Box %</span>
            <input
              type="number"
              step="0.0001"
              value={boxPct}
              onChange={(e) => setBoxPct(e.target.value)}
            />
          </label>

          <label className="field">
            <span>Reversal</span>
            <input
              type="number"
              step="1"
              min="1"
              value={reversal}
              onChange={(e) => setReversal(e.target.value)}
            />
          </label>

          <label className="field">
            <span>Weeks lookback</span>
            <input
              type="number"
              step="1"
              min="0"
              value={lookbackWeeks}
              onChange={(e) => setLookbackWeeks(e.target.value)}
            />
          </label>
        </div>

        <div className="meta-row">
          <div className="meta-card">
            <span>Файл</span>
            <strong>{dataset?.fileName || 'не загружен'}</strong>
          </div>
          <div className="meta-card">
            <span>Лист</span>
            <strong>{dataset?.sheetName || '—'}</strong>
          </div>
          <div className="meta-card">
            <span>Активов</span>
            <strong>{dataset?.headers?.length || 0}</strong>
          </div>
          <div className="meta-card">
            <span>Строк данных</span>
            <strong>{dataset?.rows?.length || 0}</strong>
          </div>
          <div className="meta-card action-card">
            <button type="button" disabled={!result} onClick={() => downloadCsv(result)}>
              Скачать CSV матрицы
            </button>
          </div>
        </div>
      </div>

      {loading && <div className="panel status">Загрузка Excel...</div>}
      {error && <div className="panel status error">{error}</div>}

      {result && (
        <>
          <div className="panel summary-panel">
            <div className="summary-card">
              <span>As of</span>
              <strong>{formatDate(result.asOfDate)}</strong>
            </div>
            <div className="summary-card">
              <span>Weeks used</span>
              <strong>{result.weeksUsed}</strong>
            </div>
            <div className="summary-card">
              <span>Strongest</span>
              <strong>{result.order[result.order.length - 1]?.ticker || '—'}</strong>
            </div>
            <div className="summary-card">
              <span>Weakest</span>
              <strong>{result.order[0]?.ticker || '—'}</strong>
            </div>
          </div>

          <div className="panel matrix-wrap">
            <table className="matrix-table">
              <thead>
                <tr>
                  <th>RANK</th>
                  <th>TICKER</th>
                  <th>NAME</th>
                  <th>BUYS</th>
                  <th>X&apos;S</th>
                  <th>TOTAL</th>
                  <th>TECH SCORE</th>
                  {result.order.map((item) => (
                    <th key={`col-${item.index}`} className="col-code">{item.ticker}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {result.order.map((rowItem, rowIdx) => (
                  <tr key={`row-${rowItem.index}`}>
                    <td>{rowIdx + 1}</td>
                    <td className="ticker-cell">{rowItem.ticker}</td>
                    <td className="name-cell">{rowItem.name}</td>
                    <td>{rowItem.buys}</td>
                    <td>{rowItem.xs}</td>
                    <td>{rowItem.total}</td>
                    <td>{rowItem.tech.toFixed(2)}</td>
                    {result.order.map((colItem) => {
                      const code = result.codes[rowItem.index][colItem.index]
                      return (
                        <td
                          key={`code-${rowItem.index}-${colItem.index}`}
                          className={`matrix-code ${codeClass(code)} ${rowItem.index === colItem.index ? 'diag' : ''}`}
                        >
                          {code}
                        </td>
                      )
                    })}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </>
      )}
    </div>
  )
}
