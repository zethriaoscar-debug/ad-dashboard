import { useState, useCallback, useMemo, useRef } from 'react'
import Papa from 'papaparse'
import ReactMarkdown from 'react-markdown'
import remarkGfm from 'remark-gfm'
import * as XLSX from 'xlsx'

// ---------------------------------------------------------------------------
// Free models available on OpenRouter (fetched 2026-06-03)
// ---------------------------------------------------------------------------
const FREE_MODELS = [
  { id: 'openai/gpt-oss-120b:free',                           label: 'OpenAI GPT OSS 120B',            ctx: '131K' },
  { id: 'nvidia/nemotron-3-super-120b-a12b:free',             label: 'NVIDIA Nemotron 3 Super 120B',   ctx: '1M'   },
  { id: 'qwen/qwen3-coder:free',                              label: 'Qwen3 Coder 480B',               ctx: '1M'   },
  { id: 'nvidia/nemotron-3-nano-30b-a3b:free',                label: 'NVIDIA Nemotron 3 Nano 30B',     ctx: '256K' },
  { id: 'moonshotai/kimi-k2.6:free',                         label: 'Kimi K2.6',                      ctx: '262K' },
  { id: 'google/gemma-4-31b-it:free',                        label: 'Google Gemma 4 31B',             ctx: '262K' },
  { id: 'meta-llama/llama-3.3-70b-instruct:free',            label: 'Meta Llama 3.3 70B',             ctx: '131K' },
  { id: 'openai/gpt-oss-20b:free',                           label: 'OpenAI GPT OSS 20B',             ctx: '131K' },
  { id: 'nousresearch/hermes-3-llama-3.1-405b:free',         label: 'Nous Hermes 3 405B',             ctx: '131K' },
  { id: 'deepseek/deepseek-r1:free',                         label: 'DeepSeek R1',                    ctx: '164K' },
  { id: 'deepseek/deepseek-chat:free',                       label: 'DeepSeek Chat V3',               ctx: '64K'  },
  { id: 'openrouter/free',                                   label: 'Auto (best available)',          ctx: '200K' },
]

// ---------------------------------------------------------------------------
// Column aliases — handles case differences across Meta Ads CSV exports
// ---------------------------------------------------------------------------
const COLS = {
  campaign:    ['Nama kampanye',  'Nama Kampanye'],
  adset:       ['Nama set iklan', 'Nama Set Iklan'],
  ad:          ['Nama iklan',     'Nama Iklan'],
  status:      ['Status Penayangan'],
  spend:       ['Jumlah yang dibelanjakan (IDR)'],
  impressions: ['Impresi'],
  reach:       ['Jangkauan'],
  frequency:   ['Frekuensi'],
  results:     ['Hasil'],
  cpl_col:     ['Biaya per hasil'],
  cpm:         ['CPM (Biaya Per 1.000 Tayangan)'],
  linkClicks:  ['Klik tautan', 'Klik Tautan'],
  ctr:         ['CTR (rasio klik tayang tautan)', 'P1 (CTR)', 'P1'],
  cpc:         ['CPC (biaya per klik tautan)', 'CPC (Biaya per Klik Tautan)'],
  lpv:         ['Tayangan halaman tujuan situs web', 'Tayangan halaman tujuan'],
  lpvRate:     ['P2 (CTLPV)', 'LPV/Click'],
  resultType:  ['Jenis hasil', 'Jenis Keuntungan'],
  qualityRank: ['Peringkat kualitas'],
  engRank:     ['Peringkat nilai interaksi'],
  convRank:    ['Peringkat nilai konversi'],
  dateStart:   ['Awal pelaporan'],
  dateEnd:     ['Akhir pelaporan'],
  // Donation
  donations:        ['Donasi'],
  roas:             ['ROAS Sumbangan (Imbal hasil belanja iklan)'],
  donationValue:    ['Nilai Konversi Donasi'],
  aov:              ['AOV Donation (IDR)'],
  initiateCheckout: ['Proses pembayaran yang dimulai'],
  cvrIC:            ['P3 (CVR - IC)'],
  // Sales
  purchases:        ['Pembelian Situs Web', 'Pembelian'],
  salesCVR:         ['P3 (sales)'],
  // Awareness / CTWA
  igVisits:         ['Kunjungan profil Instagram'],
  conversations:    ['Percakapan pesan dimulai'],
}

// Build case-insensitive lookup: lowercased header → original header
const buildColMap = headers => {
  const m = {}
  headers.forEach(h => { m[h.toLowerCase().trim()] = h })
  return m
}

// Get value from row using colMap + alias array fallback
const gv = (row, cm, aliases) => {
  for (const a of aliases) {
    const key = cm[a.toLowerCase().trim()]
    if (key !== undefined && row[key] !== undefined && row[key] !== '') return row[key]
  }
  return ''
}

// ---------------------------------------------------------------------------
// Objective detection & config
// ---------------------------------------------------------------------------
const RESULT_TYPE_MAP = {
  'donasi situs web':                  'donation',
  'memulai checkout situs web':        'donation',
  'pembelian situs web':               'sales',
  'prospek situs web':                 'leads',
  'percakapan pesan dimulai':          'ctwa',
  'tayangan halaman tujuan situs web': 'traffic',
  'tayangan halaman tujuan':           'traffic',
  'kunjungan profil instagram':        'awareness',
  'jangkauan':                         'awareness',
}
const detectRowObj = rt => RESULT_TYPE_MAP[(rt || '').toLowerCase().trim()] || 'leads'

const OBJ_CFG = {
  leads:      { label: 'Lead Generation', icon: '👤', color: 'var(--accent)',   resultLabel: 'Leads',      costLabel: 'CPL'          },
  donation:   { label: 'Donation',        icon: '❤️', color: 'var(--success)',  resultLabel: 'Donasi',     costLabel: 'CPA'          },
  sales:      { label: 'Sales',           icon: '🛒', color: 'var(--warning)',  resultLabel: 'Pembelian',  costLabel: 'CPA'          },
  traffic:    { label: 'Traffic',         icon: '🌐', color: '#60a5fa',         resultLabel: 'LPV',        costLabel: 'CPC'          },
  awareness:  { label: 'Awareness',       icon: '👁',  color: 'var(--tertiary)', resultLabel: 'Reach',      costLabel: 'CPM'          },
  ctwa:       { label: 'CTWA / WhatsApp', icon: '💬', color: '#25D366',         resultLabel: 'Percakapan', costLabel: 'Cost/Chat'    },
  fullFunnel: { label: 'Full Funnel',     icon: '🔄', color: 'var(--accent)',   resultLabel: 'Results',    costLabel: 'Cost/Result'  },
}

const OBJ_DEFAULTS = {
  leads:      { cpm: 80000,  cpl: 150000, ctrMin: 1.5, freqMax: 3.0 },
  donation:   { cpm: 50000,  cpl: 30000,  ctrMin: 1.0, freqMax: 4.0 },
  sales:      { cpm: 50000,  cpl: 200000, ctrMin: 1.0, freqMax: 4.0 },
  traffic:    { cpm: 30000,  cpl: 5000,   ctrMin: 1.5, freqMax: 4.0 },
  awareness:  { cpm: 20000,  cpl: 0,      ctrMin: 0.5, freqMax: 5.0 },
  ctwa:       { cpm: 50000,  cpl: 50000,  ctrMin: 1.0, freqMax: 3.0 },
  fullFunnel: { cpm: 80000,  cpl: 150000, ctrMin: 1.0, freqMax: 4.0 },
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------
const n        = v => parseFloat(String(v || '').replace(/,/g, '')) || 0
const idr      = v => `Rp ${Math.round(v).toLocaleString('id-ID')}`
const pct      = v => `${parseFloat(v).toFixed(2)}%`
const topRank  = arr => arr.filter(Boolean).slice(-1)[0] || '—'

const shortName = name => {
  if (!name) return '—'
  const m = name.match(/KA(\d+)_[^_]+_[^_]+_([^_]+)/)
  return m ? `${m[2]} · KA${m[1]}` : name.substring(0, 24)
}

// ---------------------------------------------------------------------------
// Styles (inline object helpers)
// ---------------------------------------------------------------------------
const S = {
  card: {
    background: 'var(--card)',
    border: '1px solid var(--border)',
    borderRadius: 'var(--radius-lg)',
    padding: '1.25rem',
  },
  badge: (color, bg) => ({
    background: bg,
    color,
    padding: '2px 9px',
    borderRadius: 'var(--radius-sm)',
    fontSize: '11px',
    whiteSpace: 'nowrap',
  }),
  th: {
    padding: '9px 13px',
    textAlign: 'left',
    fontWeight: 500,
    color: 'var(--muted)',
    fontSize: '11px',
    borderBottom: '1px solid var(--border)',
    whiteSpace: 'nowrap',
    textTransform: 'uppercase',
    letterSpacing: '0.05em',
  },
  td: { padding: '10px 13px', borderBottom: '1px solid var(--border)' },
}

// ---------------------------------------------------------------------------
// Sub-components
// ---------------------------------------------------------------------------
function KPICard({ label, value, color, sub }) {
  return (
    <div style={{ background: 'var(--card)', border: '1px solid var(--border)', borderRadius: 'var(--radius-md)', padding: '1rem', flex: 1, minWidth: 0 }}>
      <div style={{ fontSize: '11px', color: 'var(--muted)', textTransform: 'uppercase', letterSpacing: '0.07em', marginBottom: 6 }}>{label}</div>
      <div style={{ fontSize: '20px', fontWeight: 500, fontFamily: 'var(--font-mono)', color: color || 'var(--text)' }}>{value}</div>
      {sub && <div style={{ fontSize: '11px', color: 'var(--tertiary)', marginTop: 3 }}>{sub}</div>}
    </div>
  )
}

function TabBtn({ id, active, onClick, label }) {
  return (
    <button
      onClick={() => onClick(id)}
      style={{
        background: 'none',
        border: 'none',
        borderBottom: active ? '2px solid var(--accent)' : '2px solid transparent',
        padding: '9px 18px',
        color: active ? 'var(--text)' : 'var(--muted)',
        fontWeight: active ? 500 : 400,
        fontSize: '14px',
        borderRadius: 0,
        marginBottom: '-1px',
      }}
    >
      {label}
    </button>
  )
}

// ---------------------------------------------------------------------------
// Main component
// ---------------------------------------------------------------------------
export default function App() {
  const [data,        setData]       = useState(null)
  const [rawRows,     setRawRows]    = useState([])
  const [apiKey,      setApiKey]     = useState(localStorage.getItem('or_key') || '')
  const [modelId,     setModelId]    = useState(localStorage.getItem('or_model') || FREE_MODELS[0].id)
  const [insight,     setInsight]    = useState('')
  const [loading,     setLoading]    = useState(false)
  const [dragOver,    setDragOver]   = useState(false)
  const [tab,         setTab]        = useState('campaign')
  const [objective,   setObjective]  = useState('leads')
  const [isFullFunnel,setIsFF]       = useState(false)
  const [thr,         setThr]        = useState(OBJ_DEFAULTS.leads)
  const [copySuccess, setCopySuccess]= useState(false)
  const fileRef = useRef()

  const handleCopyInsight = () => {
    if (insight) {
      navigator.clipboard.writeText(insight)
      setCopySuccess(true)
      setTimeout(() => setCopySuccess(false), 2000)
    }
  }

  const exportToExcel = () => {
    if (!data) return
    const overviewData = [
      { Metric: 'Total Spend', Value: idr(data.totals.spend) },
      { Metric: 'Total Hasil', Value: data.totals.results },
      { Metric: 'CPL', Value: idr(data.totals.cpl) },
      { Metric: 'CPA', Value: idr(data.totals.cpa) },
      { Metric: 'CPM', Value: idr(data.totals.cpm) },
      { Metric: 'CTR', Value: pct(data.totals.ctr) },
      { Metric: 'ROAS', Value: data.totals.roas > 0 ? `${data.totals.roas.toFixed(2)}x` : '—' },
      { Metric: 'Nilai Donasi', Value: idr(data.totals.donationValue) },
      { Metric: 'AOV', Value: idr(data.totals.donations > 0 ? data.totals.donationValue / data.totals.donations : 0) },
      { Metric: 'Reach', Value: data.totals.reach },
      { Metric: 'LPV', Value: data.totals.lpv },
      { Metric: 'Cost per Chat', Value: idr(data.totals.costPerChat) },
    ]
    const campaignData = data.campaigns.map(c => {
      const row = { Campaign: c.name, Objective: c.objective, Status: c.status, Spend: c.spend, Hasil: c.results, 'CPL / CPA': c.cpa > 0 ? c.cpa : c.cpl, CPM: c.cpm, CTR: c.ctr, 'Average Freq': c.avgFreq }
      if (objective === 'donation' || isFullFunnel) {
        row['Donasi'] = c.donations; row['Nilai Donasi'] = c.donationValue; row['AOV'] = c.donations > 0 ? c.donationValue / c.donations : 0; row['ROAS'] = c.roas
      }
      return row
    })
    const cm2 = rawRows.length > 0 ? buildColMap(Object.keys(rawRows[0])) : {}
    const gr  = (r, col) => gv(r, cm2, COLS[col])
    const adData = rawRows.map(r => {
      const impr = n(gr(r, 'impressions')), clicks = n(gr(r, 'linkClicks')), don = n(gr(r, 'donations')), donVal = n(gr(r, 'donationValue')), sp = n(gr(r, 'spend'))
      const rowObj = detectRowObj(gr(r, 'resultType'))
      const row = { 'Nama Iklan': gr(r, 'ad'), Campaign: gr(r, 'campaign'), Objective: OBJ_CFG[rowObj]?.label || rowObj, Status: gr(r, 'status'), Spend: sp, Hasil: n(gr(r, 'results')), 'CPL / CPA': (don || n(gr(r,'purchases')) || n(gr(r,'results'))) > 0 ? sp / (don || n(gr(r,'purchases')) || n(gr(r,'results'))) : 0, CPM: n(gr(r, 'cpm')), CTR: impr > 0 ? (clicks/impr)*100 : 0 }
      if (objective === 'donation' || isFullFunnel) {
        row['Donasi'] = don; row['Nilai Donasi'] = donVal; row['AOV'] = don > 0 ? donVal / don : 0; row['ROAS'] = donVal > 0 && sp > 0 ? donVal / sp : 0
      }
      return row
    })
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(overviewData), 'Overview')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(campaignData), 'Campaigns')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(adData), 'Detail Iklan')
    XLSX.writeFile(wb, `Laporan-Iklan-${data.dateEnd.replace(/-/g, '')}.xlsx`)
  }


  const handleModelChange = id => {
    setModelId(id)
    localStorage.setItem('or_model', id)
  }

  const activeModel = FREE_MODELS.find(m => m.id === modelId) || FREE_MODELS[0]

  // -------------------------------------------------------------------------
  // Save API key to localStorage on change
  // -------------------------------------------------------------------------
  const handleApiKeyChange = v => {
    setApiKey(v)
    localStorage.setItem('or_key', v)
  }

  // -------------------------------------------------------------------------
  // Parse & aggregate CSV
  // -------------------------------------------------------------------------
  const processFile = file => {
    if (!file || !file.name.endsWith('.csv')) return
    Papa.parse(file, {
      header: true,
      skipEmptyLines: true,
      complete: ({ data: rows }) => {
        if (!rows.length) return
        setRawRows(rows)
        const cm = buildColMap(Object.keys(rows[0]))
        const g  = (r, col) => gv(r, cm, COLS[col])

        // Detect dominant objective
        const objSet = new Set(rows.map(r => detectRowObj(g(r, 'resultType'))))
        const isFF   = objSet.size > 1
        const dominant = isFF ? 'fullFunnel' : [...objSet][0] || 'leads'
        setObjective(dominant)
        setIsFF(isFF)
        setThr(OBJ_DEFAULTS[dominant] || OBJ_DEFAULTS.leads)

        const map = {}
        rows.forEach(r => {
          const key = g(r, 'campaign')
          if (!key) return
          if (!map[key]) map[key] = {
            name: key, spend: 0, results: 0, impressions: 0,
            reach: 0, linkClicks: 0, lpv: 0, freqArr: [],
            status: g(r, 'status'), qRanks: [], eRanks: [], cRanks: [],
            donations: 0, donationValue: 0, initiateCheckout: 0,
            purchases: 0, conversations: 0, igVisits: 0,
            objTypes: new Set(),
          }
          const c = map[key]
          c.spend            += n(g(r, 'spend'))
          c.results          += n(g(r, 'results'))
          c.impressions      += n(g(r, 'impressions'))
          c.reach            += n(g(r, 'reach'))
          c.linkClicks       += n(g(r, 'linkClicks'))
          c.lpv              += n(g(r, 'lpv'))
          c.donations        += n(g(r, 'donations'))
          c.donationValue    += n(g(r, 'donationValue'))
          c.initiateCheckout += n(g(r, 'initiateCheckout'))
          c.purchases        += n(g(r, 'purchases'))
          c.conversations    += n(g(r, 'conversations'))
          c.igVisits         += n(g(r, 'igVisits'))
          const freq = n(g(r, 'frequency'))
          if (freq > 0) c.freqArr.push(freq)
          c.qRanks.push(g(r, 'qualityRank'))
          c.eRanks.push(g(r, 'engRank'))
          c.cRanks.push(g(r, 'convRank'))
          c.objTypes.add(detectRowObj(g(r, 'resultType')))
        })

        const campaigns = Object.values(map).map(c => ({
          ...c,
          cpm:            c.impressions > 0 ? (c.spend / c.impressions) * 1000 : 0,
          ctr:            c.impressions > 0 ? (c.linkClicks / c.impressions) * 100 : 0,
          cpl:            c.results > 0 ? c.spend / c.results : 0,
          cpa:            (c.donations + c.purchases) > 0 ? c.spend / (c.donations + c.purchases) : 0,
          roas:           c.donationValue > 0 ? c.donationValue / c.spend : 0,
          cpc:            c.linkClicks > 0 ? c.spend / c.linkClicks : 0,
          avgFreq:        c.freqArr.length > 0 ? c.freqArr.reduce((a, b) => a + b, 0) / c.freqArr.length : 0,
          qualityRank:    topRank(c.qRanks),
          engagementRank: topRank(c.eRanks),
          conversionRank: topRank(c.cRanks),
          objective:      c.objTypes.size > 1 ? 'fullFunnel' : [...c.objTypes][0] || 'leads',
        }))

        const tot = {
          spend:         campaigns.reduce((s, c) => s + c.spend,         0),
          results:       campaigns.reduce((s, c) => s + c.results,       0),
          impressions:   campaigns.reduce((s, c) => s + c.impressions,   0),
          linkClicks:    campaigns.reduce((s, c) => s + c.linkClicks,    0),
          reach:         campaigns.reduce((s, c) => s + c.reach,         0),
          lpv:           campaigns.reduce((s, c) => s + c.lpv,           0),
          donations:     campaigns.reduce((s, c) => s + c.donations,     0),
          donationValue: campaigns.reduce((s, c) => s + c.donationValue, 0),
          purchases:     campaigns.reduce((s, c) => s + c.purchases,     0),
          conversations: campaigns.reduce((s, c) => s + c.conversations, 0),
          igVisits:      campaigns.reduce((s, c) => s + c.igVisits,      0),
        }
        tot.cpl  = tot.results > 0  ? tot.spend / tot.results  : 0
        tot.cpm  = tot.impressions > 0 ? (tot.spend / tot.impressions) * 1000 : 0
        tot.ctr  = tot.impressions > 0 ? (tot.linkClicks / tot.impressions) * 100 : 0
        tot.cpc  = tot.linkClicks > 0 ? tot.spend / tot.linkClicks : 0
        tot.roas = tot.donationValue > 0 ? tot.donationValue / tot.spend : 0
        tot.cpa  = (tot.donations + tot.purchases) > 0
          ? tot.spend / (tot.donations + tot.purchases) : 0
        tot.costPerChat = tot.conversations > 0 ? tot.spend / tot.conversations : 0

        const firstRow = rows[0]
        setData({
          campaigns, totals: tot,
          dateStart: g(firstRow, 'dateStart'),
          dateEnd:   g(firstRow, 'dateEnd'),
        })
        setInsight('')
        setTab('campaign')
      },
    })
  }

  const handleDrop = useCallback(e => {
    e.preventDefault()
    setDragOver(false)
    processFile(e.dataTransfer.files[0])
  }, [])

  // -------------------------------------------------------------------------
  // KPI alerts
  // -------------------------------------------------------------------------
  const alerts = useMemo(() => {
    if (!data) return []
    const list = []
    const cfg = OBJ_CFG[objective] || OBJ_CFG.leads
    data.campaigns.forEach(c => {
      const s = shortName(c.name)
      if (c.cpm > thr.cpm && c.cpm > 0)
        list.push({ level: 'danger', msg: `${s}: CPM ${idr(c.cpm)} melebihi batas ${idr(thr.cpm)}` })
      // Cost per result — objective-aware
      const costPerResult = (objective === 'donation' || objective === 'sales') ? c.cpa : c.cpl
      if (thr.cpl > 0 && costPerResult > thr.cpl && costPerResult > 0)
        list.push({ level: 'danger', msg: `${s}: ${cfg.costLabel} ${idr(costPerResult)} melebihi batas ${idr(thr.cpl)}` })
      if (c.ctr < thr.ctrMin && c.impressions > 0)
        list.push({ level: 'warning', msg: `${s}: CTR ${pct(c.ctr)} di bawah minimum ${pct(thr.ctrMin)}` })
      if (c.avgFreq > 0 && c.avgFreq > thr.freqMax)
        list.push({ level: 'warning', msg: `${s}: Frekuensi ${c.avgFreq.toFixed(2)} — potensi creative fatigue` })
    })
    return list
  }, [data, thr, objective])

  // -------------------------------------------------------------------------
  // Metric color coding
  // -------------------------------------------------------------------------
  const mc = (metric, val) => {
    if (metric === 'cpm') {
      if (val > thr.cpm)        return 'var(--danger)'
      if (val > thr.cpm * 0.8) return 'var(--warning)'
      return 'var(--success)'
    }
    if (metric === 'cpl' || metric === 'cpa') {
      if (val <= 0)             return 'var(--muted)'
      if (val > thr.cpl)        return 'var(--danger)'
      if (val > thr.cpl * 0.8) return 'var(--warning)'
      return 'var(--success)'
    }
    if (metric === 'ctr') {
      if (val < thr.ctrMin)         return 'var(--danger)'
      if (val < thr.ctrMin * 1.5)   return 'var(--warning)'
      return 'var(--success)'
    }
    if (metric === 'freq') {
      if (val <= 0)                 return 'var(--muted)'
      if (val > thr.freqMax)        return 'var(--danger)'
      if (val > thr.freqMax * 0.8)  return 'var(--warning)'
      return 'var(--text)'
    }
    if (metric === 'roas') {
      if (val <= 0)  return 'var(--muted)'
      if (val < 1.0) return 'var(--danger)'
      if (val < 2.0) return 'var(--warning)'
      return 'var(--success)'
    }
    return 'var(--text)'
  }

  // -------------------------------------------------------------------------
  // AI Insight via OpenRouter
  // -------------------------------------------------------------------------
  const generateInsight = async () => {
    if (!apiKey || !data) return
    setLoading(true)
    setInsight('')

    // ── Build colMap from raw rows for gv() ──
    const cm = rawRows.length > 0 ? buildColMap(Object.keys(rawRows[0])) : {}
    const g  = (r, col) => gv(r, cm, COLS[col])

    // ── Totals ──
    const totSpend   = data.totals.spend
    const totClicks  = data.totals.linkClicks
    const totImpr    = data.totals.impressions
    const totLpv     = data.totals.lpv
    const totCPM     = data.totals.cpm
    const totCTR     = data.totals.ctr
    const totCPC     = data.totals.cpc
    const totCPL     = data.totals.cpl
    const totLeads   = data.totals.results
    const totCVR     = totClicks > 0 ? (totLeads / totClicks) * 100 : 0
    const totLpvRate = totClicks > 0 ? (totLpv   / totClicks) * 100 : 0

    // ── Per-AD rows — include all relevant fields per objective ──
    const adRows = rawRows.map(r => {
      const spend    = n(g(r, 'spend'))
      const results  = n(g(r, 'results'))
      const clicks   = n(g(r, 'linkClicks'))
      const impr     = n(g(r, 'impressions'))
      const lpv      = n(g(r, 'lpv'))
      const cpm      = n(g(r, 'cpm'))
      const ctr      = impr > 0 ? (clicks / impr) * 100 : 0
      const cpc      = n(g(r, 'cpc')) || (clicks > 0 ? spend / clicks : 0)
      const freq     = n(g(r, 'frequency'))
      const donations= n(g(r, 'donations'))
      const donVal   = n(g(r, 'donationValue'))
      const ic       = n(g(r, 'initiateCheckout'))
      const purchases= n(g(r, 'purchases'))
      const convos   = n(g(r, 'conversations'))
      const rt       = g(r, 'resultType')
      const rowObj   = detectRowObj(rt)

      const mainConv = donations || purchases || results || convos
      const cpa      = mainConv > 0 ? spend / mainConv : 0
      const cvr      = clicks > 0 ? (results / clicks) * 100 : 0
      const lpvRate  = clicks > 0 ? (lpv / clicks) * 100 : 0
      const roas     = donVal > 0 && spend > 0 ? donVal / spend : 0

      const parts = [
        `Campaign: ${g(r, 'campaign')} | Adset: ${g(r, 'adset')}`,
        `Ad: ${g(r, 'ad')} | Objective: ${OBJ_CFG[rowObj]?.label || rowObj} | Status: ${g(r, 'status')}`,
        `Spend: ${idr(spend)} | CPM: ${idr(cpm)} | CTR: ${pct(ctr)} | CPC: ${idr(cpc)}`,
        `LPV: ${lpv} | LPV/Klik: ${pct(lpvRate)}`,
      ]
      if (results > 0)    parts.push(`Hasil: ${results} | CVR: ${pct(cvr)} | CPA: ${idr(cpa)}`)
      if (ic > 0)         parts.push(`IC: ${ic} | CVR-IC: ${pct(clicks > 0 ? (ic / clicks) * 100 : 0)}`)
      if (donations > 0)  parts.push(`Donasi: ${donations} | Nilai: ${idr(donVal)} | ROAS: ${roas.toFixed(2)}x`)
      if (purchases > 0)  parts.push(`Pembelian: ${purchases} | CPA: ${idr(cpa)}`)
      if (convos > 0)     parts.push(`Percakapan WA: ${convos} | Cost/Chat: ${idr(spend / convos)}`)
      if (freq > 0)       parts.push(`Frekuensi: ${freq.toFixed(2)}`)
      const qr = g(r, 'qualityRank'), er = g(r, 'engRank'), cr = g(r, 'convRank')
      if (qr || er || cr) parts.push(`Quality: ${qr||'—'} | Engagement: ${er||'—'} | Conversion: ${cr||'—'}`)
      return parts.join('\n')
    }).join('\n\n---\n\n')

    const alertsText = alerts.length
      ? alerts.map(a => `[${a.level.toUpperCase()}] ${a.msg}`).join('\n')
      : 'Semua KPI dalam batas normal.'

    // ── System prompt — objective-aware ──
    const sysPrompts = {
      leads:
`Kamu adalah analis Meta Ads sekaligus konsultan digital marketing berpengalaman di Indonesia, spesialis lead generation untuk bisnis high-consideration (interior design, properti, renovation, pendidikan).

Gaya penulisan: profesional tapi conversational, campur Bahasa Indonesia dengan istilah marketing (CVR, CTR, LP, angle, hook, funnel, trigger) secara natural. Selalu sebut angka spesifik dari data. Jangan pakai pernyataan umum tanpa angka pendukung.
Gunakan Markdown: ## heading, ### sub-heading, **bold**, - bullet.

ATURAN ANALISIS: CTR rendah+CPM normal → masalah creative. CTR tinggi+CVR rendah → LP/offer bermasalah. LPV/Klik <40% → mismatch ad-LP. Frekuensi >3 → creative fatigue. Spend besar+0 leads → budget drain → Gagal.
Baca nama creative untuk angle. Gunakan "kita" untuk rekomendasi. Tidak ada basa-basi.`,

      donation:
`Kamu adalah analis Meta Ads spesialis donation campaign dan social fundraising di Indonesia. Funnel: Ad → LP → Initiate Checkout (IC) → Donasi.

Gaya: profesional tapi conversational, pakai istilah marketing secara natural. Selalu sebut angka spesifik.
Gunakan Markdown: ## heading, ### sub-heading, **bold**, - bullet.

ATURAN ANALISIS: Metrik kunci = ROAS, CPA per donasi, IC→Donasi rate, AOV, LPV/Klik. ROAS >1x = break even, >2x = profitable. CTR tinggi+CVR-IC rendah → LP checkout bermasalah. Spend besar+0 donasi → budget drain → Gagal.
Baca nama creative untuk angle/trigger (emotional, social proof, urgency). Gunakan "kita" untuk rekomendasi.`,

      sales:
`Kamu adalah analis Meta Ads spesialis e-commerce dan sales conversion di Indonesia. Funnel: Ad → LP/PDP → Add to Cart → Purchase.

Gaya: profesional tapi conversational, pakai istilah marketing. Selalu sebut angka spesifik.
Gunakan Markdown: ## heading, ### sub-heading, **bold**, - bullet.

ATURAN ANALISIS: Metrik kunci = ROAS, CPA, Purchase CVR. ROAS <1 = merugi. CTR tinggi+CVR rendah → LP atau checkout friction. Spend besar+0 purchase → budget drain → Gagal.
Gunakan "kita" untuk rekomendasi. Tidak ada basa-basi.`,

      traffic:
`Kamu adalah analis Meta Ads spesialis traffic dan awareness campaign di Indonesia. Tujuan: mendatangkan traffic berkualitas ke website/LP.

Gaya: profesional tapi conversational. Selalu sebut angka spesifik.
Gunakan Markdown: ## heading, ### sub-heading, **bold**, - bullet.

ATURAN ANALISIS: Metrik kunci = CTR, LPV rate, CPC, CPM. LPV/Klik <40% → bounce tinggi atau landing page tidak relevan. CTR rendah+CPM normal → creative problem. CTR tinggi+LPV rendah → LP loading lambat atau mismatch.
Gunakan "kita" untuk rekomendasi.`,

      awareness:
`Kamu adalah analis Meta Ads spesialis brand awareness dan reach campaign di Indonesia.

Gaya: profesional tapi conversational. Selalu sebut angka spesifik.
Gunakan Markdown: ## heading, ### sub-heading, **bold**, - bullet.

ATURAN ANALISIS: Metrik kunci = CPM, Reach, Frekuensi, CTR (sebagai engagement signal). Frekuensi >5 → oversaturation. CPM tinggi → audience terlalu sempit. CTR rendah di awareness = normal tapi perlu dimonitor sebagai brand signal.
Gunakan "kita" untuk rekomendasi.`,

      ctwa:
`Kamu adalah analis Meta Ads spesialis Click-to-WhatsApp (CTWA) campaign di Indonesia.

Gaya: profesional tapi conversational. Selalu sebut angka spesifik.
Gunakan Markdown: ## heading, ### sub-heading, **bold**, - bullet.

ATURAN ANALISIS: Metrik kunci = Cost per Conversation, CTR, CPM, LPV/Klik. Cost/Chat tinggi → creative atau audience problem. CTR tinggi+conversation rendah → WhatsApp landing mismatch atau respons lambat dari tim sales.
Gunakan "kita" untuk rekomendasi.`,

      fullFunnel:
`Kamu adalah analis Meta Ads spesialis full-funnel campaign di Indonesia. Data ini mencakup beberapa objective sekaligus (Awareness, Traffic, CTWA, Lead Gen, Sales/Donation) dalam satu CSV.

Gaya: profesional tapi conversational, pakai istilah marketing. Selalu sebut angka spesifik.
Gunakan Markdown: ## heading, ### sub-heading, **bold**, - bullet.

ATURAN ANALISIS PER FUNNEL STAGE:
- Awareness (Reach/IG Visit): optimasi CPM, Frekuensi, CTR sebagai signal
- Traffic (LPV): optimasi LPV rate, CPC, CTR
- CTWA: optimasi Cost/Chat, CTR
- Lead Gen: optimasi CPL, CVR, LPV/Klik
- Sales/Donation: optimasi ROAS, CPA, IC→Konversi rate

Identifikasi juga cross-funnel efficiency: apakah traffic dari Awareness mengalir ke Conversion? Gunakan "kita" untuk rekomendasi.`,
    }

    const systemPrompt = sysPrompts[objective] || sysPrompts.leads

    // ── Objective-aware totals summary ──
    let totalsSummary = `Spend ${idr(totSpend)} | CPM ${idr(totCPM)} | CTR ${pct(totCTR)} | CPC ${idr(totCPC)} | LPV ${totLpv.toLocaleString('id-ID')} | LPV/Klik ${pct(totLpvRate)}`
    if (objective === 'leads' || objective === 'fullFunnel')
      totalsSummary += ` | Leads ${totLeads} | CPL ${idr(totCPL)} | CVR ${pct(totCVR)}`
    if (objective === 'donation' || objective === 'fullFunnel')
      totalsSummary += ` | Donasi ${data.totals.donations} | CPA ${idr(data.totals.cpa)} | ROAS ${data.totals.roas.toFixed(2)}x | Nilai ${idr(data.totals.donationValue)}`
    if (objective === 'sales' || objective === 'fullFunnel')
      totalsSummary += ` | Pembelian ${data.totals.purchases}`
    if (objective === 'ctwa' || objective === 'fullFunnel')
      totalsSummary += ` | Percakapan ${data.totals.conversations} | Cost/Chat ${data.totals.costPerChat > 0 ? idr(data.totals.costPerChat) : '—'}`

    // ── User prompt ──
    const userPrompt =
`Analisis data Meta Ads berikut dan berikan output PERSIS sesuai struktur ini. Jangan tambahkan intro atau penutup di luar struktur.

---

## 📊 ANALISA KINERJA — ${data.dateStart} s/d ${data.dateEnd}

### Performa Keseluruhan

Tulis 4-5 poin dengan format:
- **[Metrik]:** [nilai] → [interpretasi vs benchmark industri jika relevan]

Tutup dengan 1 kalimat **Kesimpulan:** yang menyatakan kondisi keseluruhan dan di mana bottleneck utama berada.

---

## 🎯 KATEGORISASI IKLAN

Kelompokkan SEMUA iklan ke dalam 4 kategori. Jangan analisis per-iklan satu per satu.

### 🔥 Winning Ad
> Iklan dengan performa terbaik pada metrik utama objective ini, layak di-scale
- **[Nama creative/angle]** — [1 kalimat: metrik kunci yang membuatnya unggul]

### ✅ Potential
> Performa solid tapi belum optimal, perlu 1 perbaikan spesifik
- **[Nama creative/angle]** — [1 kalimat: apa yang perlu diperbaiki]

### ⚠️ Underperform
> Split menjadi dua kelompok:

**Masih bisa dioptimasi** (ada sinyal positif, worth fixing):
- **[Nama creative/angle]** — [root cause singkat + apa yang harus diubah]

**Sebaiknya di-takedown** (tidak ada sinyal positif, buang budget):
- **[Nama creative/angle]** — [alasan 1 kalimat mengapa tidak layak dilanjutkan]

### 🔴 Gagal
> Konversi = 0 atau ROAS < 0.5, tidak ada sinyal positif
- **[Nama creative/angle]** — [root cause: mismatch di mana — creative, audience, atau LP]

**Pola yang terlihat:**
- Winning angles: [apa kesamaan iklan yang perform?]
- Pola kegagalan: [apa kesamaan iklan yang tidak perform?]

---

## 🧪 ANALISA TESTING & EKSPERIMEN

Identifikasi semua eksperimen dari nama ad/adset/campaign. Untuk setiap eksperimen:

### [Nomor]. Testing [Apa yang diuji]

| Variabel | Metrik Utama | CPL/CPA | CPM | CTR |
|----------|-------------|---------|-----|-----|
| [A]      | ...         | ...     | ... | ... |
| [B]      | ...         | ...     | ... | ... |

- **Winner:** [nama] — alasan spesifik berdasarkan angka
- **Karakteristik winner vs loser:** [pola yang bisa dijelaskan]

**Insight Testing:** [Kesimpulan konkret yang bisa diterapkan ke campaign berikutnya]

Cek eksperimen: angle/pesan, audience, format, region/segmen.

---

## 💼 BUSINESS INSIGHTS

Berikan 3-4 insight. Setiap insight wajib format:

### 💡 [Nomor]. [Judul — pernyataan tentang market atau customer behavior]

[2-3 kalimat: apa yang data tunjukkan tentang behavior customer, bukan hanya tentang iklan. Sertakan angka pendukung.]
- [Bullet evidence dari data]
- [Bullet evidence dari data]

**Aksi bisnis:** [Langkah konkret dan spesifik yang bisa langsung dieksekusi — bukan saran umum]

---

## 📌 RANGKUMAN SINGKAT

### Analisa
1. **[Temuan utama dengan angka]** — [konteks mengapa ini penting dan implikasinya]
2. **[Temuan kedua dengan angka]** — [konteks dan implikasi]
3. **[Temuan ketiga dengan angka]** — [konteks dan implikasi]

### Rekomendasi
1. **[Action spesifik]** → [expected outcome dengan angka estimasi jika memungkinkan]
2. **[Action spesifik]** → [expected outcome]
3. **[Action spesifik]** → [expected outcome]

---

DATA INPUT

PERIODE: ${data.dateStart} – ${data.dateEnd}
TOTAL: ${totalsSummary}

STATUS KPI:
${alertsText}

PER IKLAN:
${adRows}`


    try {
      const res = await fetch('https://openrouter.ai/api/v1/chat/completions', {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${apiKey}`,
          'Content-Type': 'application/json',
          'HTTP-Referer': window.location.origin,
          'X-Title': 'Ad Performance Monitor',
        },
        body: JSON.stringify({
          model: modelId,
          messages: [
            { role: 'system', content: systemPrompt },
            { role: 'user',   content: userPrompt   },
          ],
          max_tokens: 2500,
        }),
      })
      const json = await res.json()
      if (json.choices?.[0]?.message?.content) {
        setInsight(json.choices[0].message.content)
      } else {
        const err = json.error || json
        const code = err.code || res.status
        let friendly = ''
        if (code === 429) {
          friendly = `⚠️ Model "${activeModel.label}" sedang rate-limited (terlalu banyak pengguna).\n\nCoba lagi dalam beberapa menit, atau ganti model di tab Settings.`
        } else if (code === 404) {
          friendly = `⚠️ Model "${activeModel.label}" tidak tersedia saat ini (endpoint down).\n\nSilakan ganti model di tab Settings.`
        } else if (code === 401 || code === 403) {
          friendly = `⚠️ API key tidak valid atau tidak punya akses.\n\nCek kembali API key di tab Settings → openrouter.ai/keys`
        } else {
          friendly = `⚠️ Error ${code}: ${err.message || JSON.stringify(err)}`
        }
        setInsight('__error__' + friendly)
      }
    } catch (e) {
      setInsight('__error__' + `⚠️ Koneksi gagal: ${e.message}\n\nCek koneksi internet kamu.`)
    }
    setLoading(false)
  }

  // -------------------------------------------------------------------------
  // Upload screen
  // -------------------------------------------------------------------------
  if (!data) {
    return (
      <div style={{ minHeight: '100vh', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', padding: '2rem', gap: '1.5rem' }}>
        <div style={{ textAlign: 'center' }}>
          <div style={{ fontFamily: 'var(--font-mono)', color: 'var(--accent)', fontSize: '11px', letterSpacing: '0.15em', marginBottom: 8 }}>AD PERFORMANCE MONITOR</div>
          <div style={{ fontSize: '24px', fontWeight: 600 }}>Dashboard Iklan</div>
          <div style={{ color: 'var(--muted)', marginTop: 6, fontSize: '14px' }}>Upload export CSV dari Meta Ads untuk mulai</div>
        </div>
        <div
          onDrop={handleDrop}
          onDragOver={e => { e.preventDefault(); setDragOver(true) }}
          onDragLeave={() => setDragOver(false)}
          onClick={() => fileRef.current?.click()}
          style={{
            border: `2px dashed ${dragOver ? 'var(--accent)' : 'var(--border)'}`,
            borderRadius: 'var(--radius-lg)',
            padding: '3rem 2.5rem',
            textAlign: 'center',
            cursor: 'pointer',
            background: dragOver ? 'rgba(139,124,248,0.05)' : 'var(--card)',
            width: '100%',
            maxWidth: 400,
            transition: 'all 0.15s',
          }}
        >
          <i className="ti ti-upload" style={{ fontSize: 36, color: 'var(--muted)', display: 'block', marginBottom: 12 }} />
          <div style={{ fontWeight: 500, marginBottom: 4 }}>Drop file CSV di sini</div>
          <div style={{ color: 'var(--muted)', fontSize: '13px' }}>atau klik untuk browse</div>
          <div style={{ marginTop: 12, fontFamily: 'var(--font-mono)', fontSize: '11px', color: 'var(--tertiary)' }}>
            Meta Ads export (.csv) · Bahasa Indonesia
          </div>
        </div>
        <input ref={fileRef} type="file" accept=".csv" style={{ display: 'none' }} onChange={e => processFile(e.target.files[0])} />
      </div>
    )
  }

  // -------------------------------------------------------------------------
  // Dashboard screen
  // -------------------------------------------------------------------------
  const dangerCount  = alerts.filter(a => a.level === 'danger').length
  const warningCount = alerts.filter(a => a.level === 'warning').length

  return (
    <div style={{ minHeight: '100vh', display: 'flex', flexDirection: 'column' }}>

      {/* ── Header ── */}
      <div style={{ background: 'var(--card)', borderBottom: '1px solid var(--border)', padding: '0.875rem 1.5rem', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 8 }}>
        <div>
          <div style={{ fontFamily: 'var(--font-mono)', color: 'var(--accent)', fontSize: '11px', letterSpacing: '0.12em' }}>AD PERFORMANCE MONITOR</div>
          <div style={{ color: 'var(--muted)', fontSize: '12px', marginTop: 2, fontFamily: 'var(--font-mono)' }}>{data.dateStart} – {data.dateEnd}</div>
        </div>
        <div style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>
          {/* Objective badge + override */}
          <div style={{ display: 'flex', alignItems: 'center', gap: 5, background: 'var(--card-hover)', border: '1px solid var(--border)', borderRadius: 'var(--radius-sm)', padding: '4px 8px 4px 10px', fontSize: '12px' }}>
            <span style={{ fontSize: 14 }}>{OBJ_CFG[objective]?.icon}</span>
            <span style={{ color: OBJ_CFG[objective]?.color, fontWeight: 500 }}>{OBJ_CFG[objective]?.label}</span>
            <select
              value={objective}
              onChange={e => { setObjective(e.target.value); setThr(OBJ_DEFAULTS[e.target.value] || OBJ_DEFAULTS.leads) }}
              style={{ background: 'transparent', border: 'none', color: 'var(--muted)', fontSize: '11px', cursor: 'pointer', outline: 'none', marginLeft: 2 }}
            >
              {Object.entries(OBJ_CFG).map(([k, v]) => (
                <option key={k} value={k} style={{ background: 'var(--card)' }}>{v.icon} {v.label}</option>
              ))}
            </select>
          </div>
          {dangerCount > 0 && <span style={S.badge('var(--danger)',  'var(--danger-bg)')} >{dangerCount} masalah</span>}
                    {warningCount > 0 && <span style={S.badge('var(--warning)', 'var(--warning-bg)')}>{warningCount} peringatan</span>}
          {data && (
            <button onClick={exportToExcel} style={{ background: 'var(--card-hover)', border: '1px solid var(--border)' }}>
              <i className="ti ti-download" style={{ marginRight: 4, fontSize: 13, verticalAlign: '-2px' }} />
              Export
            </button>
          )}
          <button onClick={() => { setData(null); setRawRows([]); setInsight('') }}>
            <i className="ti ti-upload" style={{ marginRight: 4, fontSize: 13, verticalAlign: '-2px' }} />
            Upload baru
          </button>
        </div>
      </div>

      {/* ── Content ── */}
      <div style={{ padding: '1.25rem 1.5rem', maxWidth: 1200, margin: '0 auto', width: '100%', display: 'flex', flexDirection: 'column', gap: '1rem' }}>

        {/* KPI Cards — dynamic per objective */}
        <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
          {/* Always: Spend */}
          <KPICard
            label="Total Spend"
            value={idr(data.totals.spend)}
            sub={
              objective === 'donation' ? `${data.totals.donations} donasi` :
              objective === 'sales'    ? `${data.totals.purchases} pembelian` :
              objective === 'traffic'  ? `${data.totals.lpv.toLocaleString('id-ID')} LPV` :
              objective === 'ctwa'     ? `${data.totals.conversations} percakapan` :
              objective === 'awareness'? `${data.totals.reach.toLocaleString('id-ID')} reach` :
                                        `${data.totals.results} leads`
            }
          />
          {/* Cost per primary result */}
          {(objective === 'leads' || objective === 'fullFunnel') &&
            <KPICard label="CPL" value={data.totals.cpl > 0 ? idr(data.totals.cpl) : '—'} color={mc('cpl', data.totals.cpl)} />
          }
          {(objective === 'donation' || objective === 'sales') &&
            <KPICard label="CPA" value={data.totals.cpa > 0 ? idr(data.totals.cpa) : '—'} color={mc('cpa', data.totals.cpa)} />
          }
          {objective === 'traffic' &&
            <KPICard label="CPC" value={data.totals.cpc > 0 ? idr(data.totals.cpc) : '—'} />
          }
          {objective === 'ctwa' &&
            <KPICard label="Cost/Chat" value={data.totals.costPerChat > 0 ? idr(data.totals.costPerChat) : '—'} />
          }
          {/* ROAS for donation/sales */}
          {(objective === 'donation' || objective === 'sales') &&
            <KPICard
              label="ROAS"
              value={data.totals.roas > 0 ? `${data.totals.roas.toFixed(2)}x` : '—'}
              color={mc('roas', data.totals.roas)}
              sub={data.totals.donationValue > 0 ? idr(data.totals.donationValue) : undefined}
            />
          }
          {/* Always: CPM + CTR */}
          <KPICard label="CPM" value={idr(data.totals.cpm)} color={mc('cpm', data.totals.cpm)} />
          <KPICard label="CTR" value={pct(data.totals.ctr)} color={mc('ctr', data.totals.ctr)} />
          {/* Reach for non-donation/sales */}
          {objective !== 'donation' && objective !== 'sales' &&
            <KPICard label="Reach" value={data.totals.reach.toLocaleString('id-ID')} sub={`${data.totals.impressions.toLocaleString('id-ID')} impresi`} />
          }
        </div>

        {/* Alerts */}
        {alerts.length > 0 && (
          <div style={{ background: 'var(--danger-bg)', border: '1px solid rgba(248,113,113,0.25)', borderRadius: 'var(--radius-md)', padding: '0.875rem 1.125rem' }}>
            <div style={{ fontSize: '11px', color: 'var(--danger)', fontWeight: 600, letterSpacing: '0.07em', marginBottom: 8 }}>KPI ALERTS</div>
            {alerts.map((a, i) => (
              <div key={i} style={{ fontSize: '13px', color: a.level === 'danger' ? 'var(--danger)' : 'var(--warning)', padding: '2px 0', display: 'flex', gap: 7, alignItems: 'flex-start' }}>
                <i className={`ti ${a.level === 'danger' ? 'ti-circle-x' : 'ti-alert-triangle'}`} style={{ fontSize: 14, marginTop: 1, flexShrink: 0 }} />
                {a.msg}
              </div>
            ))}
          </div>
        )}

        {/* Tabs */}
        <div style={{ borderBottom: '1px solid var(--border)', display: 'flex' }}>
          <TabBtn id="campaign" active={tab==='campaign'} onClick={setTab} label="Campaign" />
          <TabBtn id="ads"      active={tab==='ads'}      onClick={setTab} label="Detail iklan" />
          <TabBtn id="settings" active={tab==='settings'} onClick={setTab} label="Settings" />
        </div>

        {/* ── Tab: Campaign ── */}
        {tab === 'campaign' && (
          <>
            <div style={{ ...S.card, padding: 0, overflow: 'hidden' }}>
              <div style={{ overflowX: 'auto' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px' }}>
                  <thead>
                    <tr style={{ background: 'var(--card-hover)' }}>
                      {isFullFunnel && <th style={S.th}>Objective</th>}
                      <th style={S.th}>Campaign</th>
                      <th style={S.th}>Status</th>
                      <th style={S.th}>Spend</th>
                      <th style={S.th}>{OBJ_CFG[objective]?.resultLabel || 'Hasil'}</th>
                      <th style={S.th}>{OBJ_CFG[objective]?.costLabel || 'CPL'}</th>
                      {(objective === 'donation' || objective === 'sales' || isFullFunnel) && <th style={S.th}>ROAS</th>}
                      {(objective === 'donation' || isFullFunnel) && <th style={S.th}>Nilai Donasi</th>}
                      {(objective === 'donation' || isFullFunnel) && <th style={S.th}>AOV</th>}
                      <th style={S.th}>CPM</th>
                      <th style={S.th}>CTR</th>
                      {objective !== 'awareness' && objective !== 'traffic' && <th style={S.th}>Freq</th>}
                    </tr>
                  </thead>
                  <tbody>
                    {data.campaigns.map((c, i) => {
                      const objLabel = OBJ_CFG[c.objective] || OBJ_CFG.leads
                      const primaryVal =
                        c.objective === 'donation'  ? c.donations :
                        c.objective === 'sales'     ? c.purchases :
                        c.objective === 'traffic'   ? c.lpv :
                        c.objective === 'ctwa'      ? c.conversations :
                        c.objective === 'awareness' ? c.reach :
                        c.results
                      const costVal =
                        (c.objective === 'donation' || c.objective === 'sales') ? c.cpa :
                        c.objective === 'traffic'   ? c.cpc :
                        c.objective === 'ctwa'      ? (c.conversations > 0 ? c.spend / c.conversations : 0) :
                        c.cpl
                      const costMetric =
                        (c.objective === 'donation' || c.objective === 'sales') ? 'cpa' :
                        c.objective === 'traffic'   ? 'cpl' : 'cpl'
                      return (
                        <tr key={i} style={{ background: i % 2 === 0 ? 'transparent' : 'rgba(255,255,255,0.01)' }}>
                          {isFullFunnel && (
                            <td style={S.td}>
                              <span style={{ ...S.badge(objLabel.color, `${objLabel.color}22`), border: `1px solid ${objLabel.color}44` }}>
                                {objLabel.icon} {objLabel.label}
                              </span>
                            </td>
                          )}
                          <td style={{ ...S.td, fontWeight: 500, maxWidth: 180, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={c.name}>{shortName(c.name)}</td>
                          <td style={S.td}>
                            <span style={S.badge(c.status === 'active' ? 'var(--success)' : 'var(--danger)', c.status === 'active' ? 'var(--success-bg)' : 'var(--danger-bg)')}>
                              {c.status === 'active' ? 'Aktif' : 'Nonaktif'}
                            </span>
                          </td>
                          <td style={{ ...S.td, fontFamily: 'var(--font-mono)', fontSize: '12px' }}>{idr(c.spend)}</td>
                          <td style={{ ...S.td, fontWeight: 600, textAlign: 'center' }}>{primaryVal || '—'}</td>
                          <td style={{ ...S.td, fontFamily: 'var(--font-mono)', fontSize: '12px', color: mc(costMetric, costVal) }}>
                            {costVal > 0 ? idr(costVal) : '—'}
                          </td>
                          {(objective === 'donation' || objective === 'sales' || isFullFunnel) && (
                            <td style={{ ...S.td, fontFamily: 'var(--font-mono)', fontSize: '12px', color: mc('roas', c.roas) }}>
                              {c.roas > 0 ? `${c.roas.toFixed(2)}x` : '—'}
                            </td>
                          )}
                          {(objective === 'donation' || isFullFunnel) && (
                            <td style={{ ...S.td, fontFamily: 'var(--font-mono)', fontSize: '12px' }}>
                              {c.donationValue > 0 ? idr(c.donationValue) : '—'}
                            </td>
                          )}
                          {(objective === 'donation' || isFullFunnel) && (
                            <td style={{ ...S.td, fontFamily: 'var(--font-mono)', fontSize: '12px' }}>
                              {c.donations > 0 ? idr(c.donationValue / c.donations) : '—'}
                            </td>
                          )}
                          <td style={{ ...S.td, fontFamily: 'var(--font-mono)', fontSize: '12px', color: mc('cpm', c.cpm) }}>{idr(c.cpm)}</td>
                          <td style={{ ...S.td, fontFamily: 'var(--font-mono)', fontSize: '12px', color: mc('ctr', c.ctr) }}>{pct(c.ctr)}</td>
                          {objective !== 'awareness' && objective !== 'traffic' && (
                            <td style={{ ...S.td, fontFamily: 'var(--font-mono)', fontSize: '12px', color: mc('freq', c.avgFreq) }}>
                              {c.avgFreq > 0 ? c.avgFreq.toFixed(2) : '—'}
                            </td>
                          )}
                        </tr>
                      )
                    })}
                  </tbody>
                  <tfoot>
                    <tr style={{ background: 'var(--card-hover)', borderTop: '2px solid var(--border)' }}>
                      <td colSpan={isFullFunnel ? 3 : 2} style={{ ...S.td, fontWeight: 700, color: 'var(--text)' }}>TOTAL KESELURUHAN</td>
                      <td style={{ ...S.td, fontFamily: 'var(--font-mono)', fontSize: '12px', fontWeight: 700 }}>{idr(data.totals.spend)}</td>
                      <td style={{ ...S.td, fontWeight: 700, textAlign: 'center' }}>
                        {objective === 'donation' ? data.totals.donations :
                         objective === 'sales'    ? data.totals.purchases :
                         objective === 'traffic'  ? data.totals.lpv :
                         objective === 'ctwa'     ? data.totals.conversations :
                         objective === 'awareness'? data.totals.reach :
                         data.totals.results || '—'}
                      </td>
                      <td style={{ ...S.td, fontFamily: 'var(--font-mono)', fontSize: '12px', fontWeight: 700, color: mc(objective === 'traffic' ? 'cpl' : (objective === 'donation' || objective === 'sales') ? 'cpa' : 'cpl', objective === 'traffic' ? data.totals.cpc : (objective === 'donation' || objective === 'sales') ? data.totals.cpa : objective === 'ctwa' ? data.totals.costPerChat : data.totals.cpl) }}>
                        {(objective === 'traffic' ? data.totals.cpc : (objective === 'donation' || objective === 'sales') ? data.totals.cpa : objective === 'ctwa' ? data.totals.costPerChat : data.totals.cpl) > 0 ? idr(objective === 'traffic' ? data.totals.cpc : (objective === 'donation' || objective === 'sales') ? data.totals.cpa : objective === 'ctwa' ? data.totals.costPerChat : data.totals.cpl) : '—'}
                      </td>
                      {(objective === 'donation' || objective === 'sales' || isFullFunnel) && (
                        <td style={{ ...S.td, fontFamily: 'var(--font-mono)', fontSize: '12px', fontWeight: 700, color: mc('roas', data.totals.roas) }}>
                          {data.totals.roas > 0 ? `${data.totals.roas.toFixed(2)}x` : '—'}
                        </td>
                      )}
                      {(objective === 'donation' || isFullFunnel) && (
                        <td style={{ ...S.td, fontFamily: 'var(--font-mono)', fontSize: '12px', fontWeight: 700 }}>
                          {data.totals.donationValue > 0 ? idr(data.totals.donationValue) : '—'}
                        </td>
                      )}
                      {(objective === 'donation' || isFullFunnel) && (
                        <td style={{ ...S.td, fontFamily: 'var(--font-mono)', fontSize: '12px', fontWeight: 700 }}>
                          {data.totals.donations > 0 ? idr(data.totals.donationValue / data.totals.donations) : '—'}
                        </td>
                      )}
                      <td style={{ ...S.td, fontFamily: 'var(--font-mono)', fontSize: '12px', fontWeight: 700, color: mc('cpm', data.totals.cpm) }}>{idr(data.totals.cpm)}</td>
                      <td style={{ ...S.td, fontFamily: 'var(--font-mono)', fontSize: '12px', fontWeight: 700, color: mc('ctr', data.totals.ctr) }}>{pct(data.totals.ctr)}</td>
                      {objective !== 'awareness' && objective !== 'traffic' && (
                        <td style={{ ...S.td, fontFamily: 'var(--font-mono)', fontSize: '12px', fontWeight: 700, color: mc('freq', data.totals.reach > 0 ? data.totals.impressions / data.totals.reach : 0) }}>
                          {data.totals.reach > 0 ? (data.totals.impressions / data.totals.reach).toFixed(2) : '—'}
                        </td>
                      )}
                    </tr>
                  </tfoot>
                </table>
              </div>
            </div>

            {/* AI Insight */}
            <div style={S.card}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 8, marginBottom: '0.875rem' }}>
                <div>
                  <div style={{ fontWeight: 500, fontSize: '14px' }}>AI Insight</div>
                  <div style={{ color: 'var(--muted)', fontSize: '12px', marginTop: 2 }}>OpenRouter · <span style={{ color: 'var(--accent)' }}>{activeModel.label}</span></div>
                </div>
                <button onClick={generateInsight} disabled={!apiKey || loading} style={{ padding: '6px 16px' }}>
                  <i className="ti ti-sparkles" style={{ marginRight: 5, fontSize: 13, verticalAlign: '-2px' }} />
                  {loading ? 'Generating...' : 'Generate insight'}
                </button>
              </div>
              {!apiKey && (
                <div style={{ background: 'var(--warning-bg)', color: 'var(--warning)', padding: '8px 12px', borderRadius: 'var(--radius-sm)', fontSize: '13px', display: 'flex', gap: 7, alignItems: 'center' }}>
                  <i className="ti ti-key" style={{ fontSize: 14 }} />
                  Set OpenRouter API key di tab Settings untuk mengaktifkan
                </div>
              )}
              {insight && (
                insight.startsWith('__error__') ? (
                  <div style={{ background: 'var(--danger-bg)', border: '1px solid rgba(248,113,113,0.2)', borderRadius: 'var(--radius-md)', padding: '1rem', marginTop: '0.875rem', display: 'flex', gap: 10, alignItems: 'flex-start' }}>
                    <i className="ti ti-alert-circle" style={{ color: 'var(--danger)', fontSize: 18, flexShrink: 0, marginTop: 1 }} />
                    <div>
                      <div style={{ color: 'var(--danger)', fontWeight: 500, fontSize: '13px', marginBottom: 6 }}>Gagal generate insight</div>
                      <div style={{ color: 'var(--text)', fontSize: '13px', lineHeight: 1.75, whiteSpace: 'pre-wrap' }}>{insight.replace('__error__', '').replace('⚠️ ', '')}</div>
                      <button onClick={() => setTab('settings')} style={{ marginTop: 10, background: 'var(--danger-bg)', border: '1px solid rgba(248,113,113,0.3)', color: 'var(--danger)', padding: '5px 12px', fontSize: '12px', borderRadius: 'var(--radius-sm)' }}>
                        <i className="ti ti-settings" style={{ marginRight: 5, fontSize: 12, verticalAlign: '-1px' }} />
                        Ganti model di Settings
                      </button>
                    </div>
                  </div>
                ) : (
                  <div style={{ background: 'var(--card-hover)', borderRadius: 'var(--radius-md)', padding: '1rem 1.25rem', marginTop: '0.875rem', color: 'var(--text)' }}>
                    <ReactMarkdown
                      remarkPlugins={[remarkGfm]}
                      components={{
                        h1: ({ children }) => <div style={{ fontSize: '16px', fontWeight: 600, color: 'var(--text)', margin: '1.25rem 0 0.5rem', letterSpacing: '-0.01em' }}>{children}</div>,
                        h2: ({ children }) => <div style={{ fontSize: '15px', fontWeight: 600, color: 'var(--text)', margin: '1.25rem 0 0.5rem', letterSpacing: '-0.01em' }}>{children}</div>,
                        h3: ({ children }) => <div style={{ fontSize: '13px', fontWeight: 600, color: 'var(--accent)', margin: '1rem 0 0.35rem', textTransform: 'uppercase', letterSpacing: '0.06em' }}>{children}</div>,
                        p:  ({ children }) => <p style={{ fontSize: '14px', lineHeight: 1.75, margin: '0.4rem 0', color: 'var(--text)' }}>{children}</p>,
                        ul: ({ children }) => <ul style={{ paddingLeft: '1.25rem', margin: '0.4rem 0', display: 'flex', flexDirection: 'column', gap: '0.25rem' }}>{children}</ul>,
                        ol: ({ children }) => <ol style={{ paddingLeft: '1.25rem', margin: '0.4rem 0', display: 'flex', flexDirection: 'column', gap: '0.25rem' }}>{children}</ol>,
                        li: ({ children }) => <li style={{ fontSize: '14px', lineHeight: 1.7, color: 'var(--text)' }}>{children}</li>,
                        strong: ({ children }) => <strong style={{ fontWeight: 600, color: 'var(--text)' }}>{children}</strong>,
                        em: ({ children }) => <em style={{ color: 'var(--muted)', fontStyle: 'italic' }}>{children}</em>,
                        hr: () => <hr style={{ border: 'none', borderTop: '1px solid var(--border)', margin: '1rem 0' }} />,
                        code: ({ inline, children }) => inline
                          ? <code style={{ fontFamily: 'var(--font-mono)', fontSize: '12px', background: 'var(--bg)', color: 'var(--accent)', padding: '1px 6px', borderRadius: '4px' }}>{children}</code>
                          : <pre style={{ fontFamily: 'var(--font-mono)', fontSize: '12px', background: 'var(--bg)', color: 'var(--text)', padding: '0.75rem 1rem', borderRadius: 'var(--radius-sm)', overflowX: 'auto', margin: '0.5rem 0' }}><code>{children}</code></pre>,
                        blockquote: ({ children }) => <blockquote style={{ borderLeft: '3px solid var(--accent)', paddingLeft: '1rem', margin: '0.5rem 0', color: 'var(--muted)', fontStyle: 'italic' }}>{children}</blockquote>,
                        table: ({ children }) => (
                          <div style={{ overflowX: 'auto', margin: '0.75rem 0' }}>
                            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px' }}>{children}</table>
                          </div>
                        ),
                        thead: ({ children }) => <thead style={{ background: 'var(--bg)' }}>{children}</thead>,
                        tbody: ({ children }) => <tbody>{children}</tbody>,
                        tr:   ({ children }) => <tr style={{ borderBottom: '1px solid var(--border)' }}>{children}</tr>,
                        th:   ({ children }) => <th style={{ padding: '7px 12px', textAlign: 'left', fontWeight: 600, color: 'var(--muted)', fontSize: '11px', textTransform: 'uppercase', letterSpacing: '0.05em', whiteSpace: 'nowrap', borderBottom: '2px solid var(--border)' }}>{children}</th>,
                        td:   ({ children }) => <td style={{ padding: '7px 12px', color: 'var(--text)', fontSize: '13px', fontFamily: 'var(--font-mono)', whiteSpace: 'nowrap' }}>{children}</td>,
                      }}
                    >
                      {insight}
                    </ReactMarkdown>
                  </div>
                )
              )}
            </div>
          </>
        )}

        {/* ── Tab: Ad Detail ── */}
        {tab === 'ads' && (() => {
          const cm2 = rawRows.length > 0 ? buildColMap(Object.keys(rawRows[0])) : {}
          const gr  = (r, col) => gv(r, cm2, COLS[col])
          return (
            <div style={{ ...S.card, padding: 0, overflow: 'hidden' }}>
              <div style={{ overflowX: 'auto' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px', minWidth: 750 }}>
                  <thead>
                    <tr style={{ background: 'var(--card-hover)' }}>
                      {['Nama iklan','Campaign','Objective','Status','Spend','Hasil','CPL/CPA'].map(h => (
                        <th key={h} style={S.th}>{h}</th>
                      ))}
                      {(objective === 'donation' || isFullFunnel) && <th style={S.th}>ROAS</th>}
                      {(objective === 'donation' || isFullFunnel) && <th style={S.th}>Nilai Donasi</th>}
                      {(objective === 'donation' || isFullFunnel) && <th style={S.th}>AOV</th>}
                      {['CPM','CTR'].map(h => (
                        <th key={h} style={S.th}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {rawRows.map((r, i) => {
                      const rowObj = detectRowObj(gr(r, 'resultType'))
                      const objCfg = OBJ_CFG[rowObj] || OBJ_CFG.leads
                      const sp  = n(gr(r, 'spend'))
                      const res = n(gr(r, 'results'))
                      const impr = n(gr(r, 'impressions'))
                      const clicks = n(gr(r, 'linkClicks'))
                      const don = n(gr(r, 'donations'))
                      const pur = n(gr(r, 'purchases'))
                      const mainConv = don || pur || res
                      const cpa = mainConv > 0 ? sp / mainConv : 0
                      const cplCol = n(gr(r, 'cpl_col'))
                      const donVal = n(gr(r, 'donationValue'))
                      const roas = donVal > 0 && sp > 0 ? donVal / sp : 0
                      const costDisplay = cpa > 0 ? idr(cpa) : (cplCol > 0 ? idr(cplCol) : '—')
                      return (
                        <tr key={i} style={{ background: i % 2 === 0 ? 'transparent' : 'rgba(255,255,255,0.01)' }}>
                          <td style={{ ...S.td, maxWidth: 180, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={gr(r, 'ad')}>{gr(r, 'ad')}</td>
                          <td style={{ ...S.td, color: 'var(--muted)', maxWidth: 120, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{shortName(gr(r, 'campaign'))}</td>
                          <td style={S.td}>
                            <span style={{ fontSize: 12, color: objCfg.color }}>{objCfg.icon} {objCfg.label}</span>
                          </td>
                          <td style={S.td}>
                            <span style={S.badge(gr(r,'status')==='active'?'var(--success)':'var(--danger)', gr(r,'status')==='active'?'var(--success-bg)':'var(--danger-bg)')}>
                              {gr(r,'status')==='active'?'Aktif':'Nonaktif'}
                            </span>
                          </td>
                          <td style={{ ...S.td, fontFamily: 'var(--font-mono)' }}>{idr(sp)}</td>
                          <td style={{ ...S.td, textAlign: 'center', fontWeight: 500 }}>{mainConv || '—'}</td>
                          <td style={{ ...S.td, fontFamily: 'var(--font-mono)', color: mainConv > 0 ? mc('cpa', cpa) : 'var(--muted)' }}>{costDisplay}</td>
                          {(objective === 'donation' || isFullFunnel) && (
                            <td style={{ ...S.td, fontFamily: 'var(--font-mono)', color: mc('roas', roas) }}>
                              {roas > 0 ? `${roas.toFixed(2)}x` : '—'}
                            </td>
                          )}
                          {(objective === 'donation' || isFullFunnel) && (
                            <td style={{ ...S.td, fontFamily: 'var(--font-mono)' }}>
                              {donVal > 0 ? idr(donVal) : '—'}
                            </td>
                          )}
                          {(objective === 'donation' || isFullFunnel) && (
                            <td style={{ ...S.td, fontFamily: 'var(--font-mono)' }}>
                              {don > 0 ? idr(donVal / don) : '—'}
                            </td>
                          )}
                          <td style={{ ...S.td, fontFamily: 'var(--font-mono)', color: mc('cpm', n(gr(r,'cpm'))) }}>{idr(n(gr(r,'cpm')))}</td>
                          <td style={{ ...S.td, fontFamily: 'var(--font-mono)', color: mc('ctr', impr > 0 ? (clicks/impr)*100 : 0) }}>{pct(impr > 0 ? (clicks/impr)*100 : 0)}</td>
                        </tr>
                      )
                    })}
                  </tbody>
                  <tfoot>
                    <tr style={{ background: 'var(--card-hover)', borderTop: '2px solid var(--border)' }}>
                      <td colSpan={4} style={{ ...S.td, fontWeight: 700, color: 'var(--text)' }}>TOTAL KESELURUHAN</td>
                      <td style={{ ...S.td, fontFamily: 'var(--font-mono)', fontSize: '12px', fontWeight: 700 }}>{idr(data.totals.spend)}</td>
                      <td style={{ ...S.td, fontWeight: 700, textAlign: 'center' }}>
                        {objective === 'donation' ? data.totals.donations :
                         objective === 'sales'    ? data.totals.purchases :
                         objective === 'traffic'  ? data.totals.lpv :
                         objective === 'ctwa'     ? data.totals.conversations :
                         objective === 'awareness'? data.totals.reach :
                         data.totals.results || '—'}
                      </td>
                      <td style={{ ...S.td, fontFamily: 'var(--font-mono)', fontSize: '12px', fontWeight: 700, color: mc(objective === 'traffic' ? 'cpl' : (objective === 'donation' || objective === 'sales') ? 'cpa' : 'cpl', objective === 'traffic' ? data.totals.cpc : (objective === 'donation' || objective === 'sales') ? data.totals.cpa : objective === 'ctwa' ? data.totals.costPerChat : data.totals.cpl) }}>
                        {(objective === 'traffic' ? data.totals.cpc : (objective === 'donation' || objective === 'sales') ? data.totals.cpa : objective === 'ctwa' ? data.totals.costPerChat : data.totals.cpl) > 0 ? idr(objective === 'traffic' ? data.totals.cpc : (objective === 'donation' || objective === 'sales') ? data.totals.cpa : objective === 'ctwa' ? data.totals.costPerChat : data.totals.cpl) : '—'}
                      </td>
                      {(objective === 'donation' || isFullFunnel) && (
                        <td style={{ ...S.td, fontFamily: 'var(--font-mono)', fontSize: '12px', fontWeight: 700, color: mc('roas', data.totals.roas) }}>
                          {data.totals.roas > 0 ? `${data.totals.roas.toFixed(2)}x` : '—'}
                        </td>
                      )}
                      {(objective === 'donation' || isFullFunnel) && (
                        <td style={{ ...S.td, fontFamily: 'var(--font-mono)', fontSize: '12px', fontWeight: 700 }}>
                          {data.totals.donationValue > 0 ? idr(data.totals.donationValue) : '—'}
                        </td>
                      )}
                      {(objective === 'donation' || isFullFunnel) && (
                        <td style={{ ...S.td, fontFamily: 'var(--font-mono)', fontSize: '12px', fontWeight: 700 }}>
                          {data.totals.donations > 0 ? idr(data.totals.donationValue / data.totals.donations) : '—'}
                        </td>
                      )}
                      <td style={{ ...S.td, fontFamily: 'var(--font-mono)', fontSize: '12px', fontWeight: 700, color: mc('cpm', data.totals.cpm) }}>{idr(data.totals.cpm)}</td>
                      <td style={{ ...S.td, fontFamily: 'var(--font-mono)', fontSize: '12px', fontWeight: 700, color: mc('ctr', data.totals.ctr) }}>{pct(data.totals.ctr)}</td>
                    </tr>
                  </tfoot>
                </table>
              </div>
            </div>
          )
        })()}

        {/* ── Tab: Settings ── */}
        {tab === 'settings' && (
          <div style={{ display: 'flex', flexDirection: 'column', gap: '1rem', maxWidth: 480 }}>
            <div style={S.card}>
              <div style={{ fontWeight: 500, marginBottom: '0.875rem' }}>OpenRouter API key</div>
              <input
                type="password"
                value={apiKey}
                onChange={e => handleApiKeyChange(e.target.value)}
                placeholder="sk-or-v1-..."
                style={{ width: '100%', fontFamily: 'var(--font-mono)', fontSize: '13px' }}
              />
              <div style={{ fontSize: '12px', color: 'var(--muted)', marginTop: 8 }}>
                Daftar gratis di{' '}
                <a href="https://openrouter.ai" target="_blank" rel="noreferrer" style={{ color: 'var(--accent)' }}>openrouter.ai</a>
              </div>
            </div>

            <div style={S.card}>
              <div style={{ fontWeight: 500, marginBottom: '0.875rem' }}>AI Model</div>
              <div style={{ position: 'relative' }}>
                <select
                  value={modelId}
                  onChange={e => handleModelChange(e.target.value)}
                  style={{
                    width: '100%',
                    background: 'var(--card-hover)',
                    border: '1px solid var(--border)',
                    color: 'var(--text)',
                    padding: '8px 32px 8px 10px',
                    borderRadius: 'var(--radius-sm)',
                    fontSize: '13px',
                    fontFamily: 'var(--font-mono)',
                    appearance: 'none',
                    cursor: 'pointer',
                    outline: 'none',
                  }}
                >
                  {FREE_MODELS.map(m => (
                    <option key={m.id} value={m.id}>
                      {m.label} · ctx {m.ctx}
                    </option>
                  ))}
                </select>
                <i className="ti ti-chevron-down" style={{
                  position: 'absolute', right: 10, top: '50%', transform: 'translateY(-50%)',
                  fontSize: 14, color: 'var(--muted)', pointerEvents: 'none',
                }} />
              </div>
              <div style={{ fontSize: '11px', color: 'var(--tertiary)', marginTop: 8, fontFamily: 'var(--font-mono)' }}>
                {modelId}
              </div>
              <div style={{ fontSize: '12px', color: 'var(--muted)', marginTop: 6 }}>
                Semua model di atas gratis. Jika satu down, ganti ke model lain.
              </div>
            </div>

            <div style={S.card}>
              <div style={{ fontWeight: 500, marginBottom: '0.875rem' }}>KPI thresholds</div>
              <div style={{ display: 'flex', flexDirection: 'column', gap: '0.75rem' }}>
                {[
                  { key: 'cpm',     label: 'CPM maksimum (IDR)',   ph: '80000' },
                  { key: 'cpl',     label: 'CPL maksimum (IDR)',   ph: '150000' },
                  { key: 'ctrMin',  label: 'CTR minimum (%)',      ph: '1.5' },
                  { key: 'freqMax', label: 'Frekuensi maksimum',   ph: '3.0' },
                ].map(({ key, label, ph }) => (
                  <div key={key}>
                    <label style={{ display: 'block', fontSize: '12px', color: 'var(--muted)', marginBottom: 4 }}>{label}</label>
                    <input
                      type="number"
                      value={thr[key]}
                      onChange={e => setThr(prev => ({ ...prev, [key]: parseFloat(e.target.value) || 0 }))}
                      placeholder={ph}
                      style={{ width: '100%' }}
                    />
                  </div>
                ))}
              </div>
            </div>

            <div style={{ ...S.card, color: 'var(--muted)', fontSize: '13px', lineHeight: 1.7 }}>
              <div style={{ fontWeight: 500, color: 'var(--text)', marginBottom: 6 }}>Cara deploy ke Vercel</div>
              <ol style={{ paddingLeft: '1.25rem', display: 'flex', flexDirection: 'column', gap: 4 }}>
                <li>Ekstrak zip → buka folder di terminal</li>
                <li><code style={{ fontFamily: 'var(--font-mono)', background: 'var(--card-hover)', padding: '1px 6px', borderRadius: 4 }}>npm install</code></li>
                <li><code style={{ fontFamily: 'var(--font-mono)', background: 'var(--card-hover)', padding: '1px 6px', borderRadius: 4 }}>npm run dev</code> untuk test lokal</li>
                <li>Push ke GitHub → connect ke Vercel → deploy</li>
              </ol>
            </div>
          </div>
        )}
      </div>
    </div>
  )
}
