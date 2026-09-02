import React, { useState, useEffect, useLayoutEffect, useCallback, useMemo, useRef } from 'react'
import jsPDF from 'jspdf'
import html2canvas from 'html2canvas'

function useIsMobile() {
  const [isMobile, setIsMobile] = useState(window.innerWidth < 768)
  useEffect(() => {
    const h = () => setIsMobile(window.innerWidth < 768)
    window.addEventListener('resize', h)
    return () => window.removeEventListener('resize', h)
  }, [])
  return isMobile
}
import { ChevronRight, FileDown, ChevronsLeft, ChevronsRight, X } from 'lucide-react'
import { runSimulation } from './engine'
import {
  ComposedChart, Bar, Line, XAxis, YAxis, CartesianGrid, Tooltip,
  ReferenceLine, ResponsiveContainer, Cell,
  PieChart as RPieChart, Pie, Legend,
} from 'recharts'
import ExcelJS from 'exceljs'
import pptxgen from 'pptxgenjs'
import { simulate } from './api'
import {
  CHARGER_TYPES, ELEC_PRESET, DEFAULT_PARAMS,
  type SimParams, type SimResult, type ChargerConfig, type MonthRecord,
} from './types'

const C = {
  sidebar1: '#1E1B4B',
  sidebar2: '#3730A3',
  primary:  '#6D28D9',
  accent:   '#A78BFA',
  light:    '#EDE9FE',
  navy:     '#2D2878',
  red:      '#DC2626',
  green:    '#6366F1',
}
const PIE_COLORS = ['#6D28D9','#2563EB','#0891B2','#16a34a','#DC2626']
const TYPE_COLORS = ['#A78BFA','#6D28D9','#60A5FA']

// ── 공통 스타일 ───────────────────────────────────────────────
const inp: React.CSSProperties = {
  width: '100%', background: 'rgba(255,255,255,0.10)',
  border: '1px solid rgba(255,255,255,0.22)', borderRadius: 5,
  color: 'white', fontSize: 16, padding: '6px 10px', height: 34,
  outline: 'none', boxSizing: 'border-box',
}
const lbl: React.CSSProperties = {
  color: 'rgba(255,255,255,0.60)', fontSize: 13, display: 'block', marginBottom: 6, marginTop: 4,
}
const row2: React.CSSProperties = {
  display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10, marginBottom: 14,
}

function SLabel({ ch }: { ch: string }) {
  return <label style={lbl}>{ch}</label>
}
function SText({ value, onChange, style: extraStyle }: { value: string; onChange: (v: string) => void; style?: React.CSSProperties }) {
  const [local, setLocal] = React.useState(value)
  const composing = React.useRef(false)
  React.useEffect(() => { if (!composing.current) setLocal(value) }, [value])
  return (
    <input type="text" value={local}
      onCompositionStart={() => { composing.current = true }}
      onCompositionEnd={e => { composing.current = false; const v = (e.target as HTMLInputElement).value; setLocal(v); onChange(v) }}
      onChange={e => { setLocal(e.target.value); if (!composing.current) onChange(e.target.value) }}
      style={{ ...inp, ...extraStyle }}/>
  )
}

function fmtComma(n: number) { return n === 0 ? '0' : n.toLocaleString('ko-KR') }
function parseComma(s: string) { return Number(s.replace(/,/g, '')) }
function SNum({ value, onChange, step = 1, min }: {
  value: number; onChange: (v: number) => void; step?: number; min?: number
}) {
  const [local, setLocal] = React.useState(fmtComma(value))
  const [focused, setFocused] = React.useState(false)
  React.useEffect(() => { if (!focused) setLocal(fmtComma(value)) }, [value, focused])
  return (
    <input type="text" inputMode="numeric" value={local}
      onChange={e => {
        const raw = e.target.value.replace(/[^0-9\-]/g, '')
        setLocal(raw)
        const n = Number(raw)
        if (!isNaN(n)) onChange(n)
      }}
      onFocus={() => { setFocused(true); setLocal(String(value === 0 ? '' : value)) }}
      onBlur={() => {
        setFocused(false)
        const n = parseComma(local)
        const safe = isNaN(n) ? 0 : (min !== undefined ? Math.max(min, n) : n)
        onChange(safe)
        setLocal(fmtComma(safe))
      }}
      style={inp}/>
  )
}
function SReadonly({ value }: { value: number | string }) {
  return (
    <div style={{ ...inp, display: 'flex', alignItems: 'center', opacity: 0.55, cursor: 'not-allowed', userSelect: 'none' }}>
      {value}
    </div>
  )
}
function SSelect({ value, onChange, options }: { value: string; onChange: (v: string) => void; options: string[] }) {
  return (
    <div style={{ position: 'relative', width: '100%' }}>
      <select value={value} onChange={e => onChange(e.target.value)}
        style={{ ...inp, cursor: 'pointer', appearance: 'none', WebkitAppearance: 'none', paddingRight: 28, width: '100%' } as React.CSSProperties}>
        {options.map(o => <option key={o} value={o} style={{ color: '#111', background: '#fff' }}>{o}</option>)}
      </select>
      <span style={{ position: 'absolute', right: 9, top: '50%', transform: 'translateY(-50%)', pointerEvents: 'none', color: 'rgba(255,255,255,0.45)', fontSize: 11, lineHeight: 1 }}>▼</span>
    </div>
  )
}

// ── 차량 개별 주행량 설정 모달 ────────────────────────────────
function getVkwPresets() {
  const saved = loadVkwPresets()
  return [...saved, { label: '직접입력', kwh: 0 }]
}
const VKW_PRESETS = getVkwPresets()

function kwhToSel(kwh: number): string {
  const p = VKW_PRESETS.find(x => x.kwh === kwh && x.label !== '직접입력')
  return p ? p.label : '직접입력'
}

function VehicleKwhModal({ total, avgKwh, initArr, onSave, onReset, onClose, presets }: {
  total: number
  avgKwh: number
  initArr: number[]
  onSave: (arr: number[]) => void
  onReset: () => void
  onClose: () => void
  presets?: { label: string; kwh: number }[]
}) {
  const vkwPresets = presets ?? VKW_PRESETS
  const kwhToSelLocal = (kwh: number) => {
    const p = vkwPresets.find(x => x.kwh === kwh && x.label !== '직접입력')
    return p ? p.label : '직접입력'
  }
  type Group = { sel: string; kwh: number; count: number }
  const initGroups = (): Group[] => {
    if (initArr.length === total && initArr.some(v => v !== avgKwh)) {
      const groups: Group[] = []
      for (const v of initArr) {
        const last = groups[groups.length - 1]
        if (last && last.kwh === v) { last.count++ }
        else { groups.push({ sel: kwhToSelLocal(v), kwh: v, count: 1 }) }
      }
      return groups
    }
    return [{ sel: kwhToSelLocal(avgKwh), kwh: avgKwh, count: total }]
  }
  const [groups, setGroups] = useState<Group[]>(initGroups)

  const assigned = groups.reduce((s, g) => s + g.count, 0)
  const remaining = total - assigned

  const updGroup = (gi: number, patch: Partial<Group>) =>
    setGroups(gs => gs.map((g, i) => i === gi ? { ...g, ...patch } : g))

  const handleSave = () => {
    const flat: number[] = []
    for (const g of groups) for (let i = 0; i < g.count; i++) flat.push(g.kwh)
    onSave(flat)
  }

  const selStyle: React.CSSProperties = {
    ...inp, cursor: 'pointer', appearance: 'none', WebkitAppearance: 'none',
    paddingRight: 28, width: '100%',
  }

  return (
    <Modal title="🚗 차량 개별 주행량 설정" onClose={onClose}>
      <div style={{ padding: '10px 14px 11px', marginBottom: 14, borderRadius: 8, background: 'rgba(99,102,241,0.10)', border: '1px solid rgba(99,102,241,0.22)' }}>
        <p style={{ color: 'rgba(196,191,239,0.55)', fontSize: 10, fontWeight: 700, letterSpacing: '0.07em', textTransform: 'uppercase', marginBottom: 4 }}>항목 안내</p>
        <p style={{ color: 'rgba(255,255,255,0.75)', fontSize: 12, lineHeight: 1.65 }}>차종과 해당 차량 수를 설정합니다. 항목 추가로 여러 종류를 혼합할 수 있습니다.</p>
      </div>

      {/* 총 차량 현황 */}
      <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: 10 }}>
        <span style={{ fontSize: 11, color: remaining === 0 ? '#4ade80' : remaining < 0 ? '#f87171' : 'rgba(255,255,255,0.45)' }}>
          설정 {assigned}대 / 전체 {total}대{remaining > 0 ? ` (${remaining}대 미설정)` : remaining < 0 ? ` (${-remaining}대 초과)` : ' ✓'}
        </span>
      </div>

      <div style={{ display: 'flex', flexDirection: 'column', gap: 10, marginBottom: 14 }}>
        {groups.map((g, gi) => (
          <div key={gi} style={{ padding: '12px 14px', borderRadius: 10, background: 'rgba(255,255,255,0.05)', border: '1px solid rgba(255,255,255,0.09)' }}>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8, alignItems: 'end' }}>
              <div>
                <SLabel ch="차종"/>
                <div style={{ position: 'relative' }}>
                  <select value={g.sel} onChange={e => {
                    const sel = e.target.value
                    const preset = vkwPresets.find(p => p.label === sel)
                    if (sel === '직접입력') updGroup(gi, { sel, kwh: g.kwh || avgKwh })
                    else if (preset) updGroup(gi, { sel, kwh: preset.kwh })
                  }} style={selStyle}>
                    {vkwPresets.map(p => (
                      <option key={p.label} value={p.label} style={{ color: '#111', background: '#fff' }}>
                        {p.label}{p.label !== '직접입력' ? ` (${p.kwh}kWh)` : ''}
                      </option>
                    ))}
                  </select>
                  <span style={{ position: 'absolute', right: 9, top: '50%', transform: 'translateY(-50%)', pointerEvents: 'none', color: 'rgba(255,255,255,0.45)', fontSize: 11 }}>▼</span>
                </div>
              </div>
              <div>
                <SLabel ch="차량 수 (대)"/>
                <SNum value={g.count} onChange={v => updGroup(gi, { count: Math.max(1, Math.floor(v)) })} min={1} step={1}/>
              </div>
            </div>
            {g.sel === '직접입력' && (
              <div style={{ marginTop: 8 }}>
                <SLabel ch="충전량 (kWh)"/>
                <SNum value={g.kwh || avgKwh} onChange={v => updGroup(gi, { kwh: v })} min={1}/>
              </div>
            )}
            {groups.length > 1 && (
              <button onClick={() => setGroups(gs => gs.filter((_, i) => i !== gi))} style={{
                marginTop: 10, width: '100%', padding: '5px 0', borderRadius: 6, fontSize: 11, cursor: 'pointer',
                background: 'none', border: '1px solid rgba(255,255,255,0.12)', color: 'rgba(255,255,255,0.30)',
              }}>이 항목 삭제</button>
            )}
          </div>
        ))}
      </div>

      {/* 항목 추가 */}
      <button onClick={() => setGroups(gs => [...gs, { sel: kwhToSel(avgKwh), kwh: avgKwh, count: 1 }])} style={{
        width: '100%', padding: '8px 0', borderRadius: 8, fontSize: 12, fontWeight: 700, cursor: 'pointer',
        background: 'rgba(99,102,241,0.12)', border: '1px dashed rgba(99,102,241,0.40)', color: 'rgba(167,139,250,0.80)',
        marginBottom: 10,
      }}>+ 항목 추가</button>

      {/* 적용 */}
      <button onClick={handleSave} disabled={remaining !== 0} style={{
        width: '100%', padding: '11px 0', borderRadius: 8, fontSize: 13, fontWeight: 700, cursor: remaining !== 0 ? 'not-allowed' : 'pointer',
        background: remaining !== 0 ? 'rgba(255,255,255,0.08)' : C.accent,
        color: remaining !== 0 ? 'rgba(255,255,255,0.30)' : C.sidebar1,
        border: 'none', marginBottom: 8,
      }}>적용</button>

      <button onClick={onReset} style={{
        width: '100%', padding: '8px 0', borderRadius: 8, fontSize: 11, cursor: 'pointer',
        background: 'none', border: '1px solid rgba(255,255,255,0.15)', color: 'rgba(255,255,255,0.40)',
      }}>초기화 (전체 동일 적용)</button>
    </Modal>
  )
}

// ── 모달 ──────────────────────────────────────────────────────
function Modal({ title, onClose, onReset, children, width }: {
  title: string; onClose: () => void; onReset?: () => void; children: React.ReactNode; width?: number
}) {
  const isMobile = useIsMobile()
  return (
    <div style={{
      position: 'fixed', inset: 0, zIndex: 1000,
      display: 'flex',
      alignItems: isMobile ? 'flex-end' : 'center',
      justifyContent: 'center',
    }}>
      {/* 배경 오버레이 — X 버튼으로만 닫기 */}
      <div style={{
        position: 'absolute', inset: 0,
        background: 'rgba(10,8,30,0.65)', backdropFilter: 'blur(3px)',
      }}/>
      {/* 모달 창 */}
      <div style={{
        position: 'relative', zIndex: 1,
        width: isMobile ? '100%' : (width ?? 460),
        maxWidth: isMobile ? '100%' : '96vw',
        maxHeight: isMobile ? '92vh' : '85vh',
        borderRadius: isMobile ? '20px 20px 0 0' : 16,
        overflow: 'hidden',
        boxShadow: '0 24px 60px rgba(0,0,0,0.5)',
        display: 'flex', flexDirection: 'column',
        background: `linear-gradient(160deg, ${C.sidebar1} 0%, #2D2875 100%)`,
        border: '1px solid rgba(167,139,250,0.25)',
      }}>
        {/* 모달 헤더 */}
        <div style={{
          display: 'flex', alignItems: 'center', justifyContent: 'space-between',
          padding: '20px 24px 16px',
          borderBottom: '1px solid rgba(255,255,255,0.10)',
        }}>
          <h2 style={{ color: 'white', fontSize: 17, fontWeight: 700, letterSpacing: '-0.4px' }}>{title}</h2>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
            {onReset && (
              <button onClick={onReset} style={{
                background: 'none', border: '1px solid rgba(255,255,255,0.20)', borderRadius: 7,
                color: 'rgba(255,255,255,0.45)', cursor: 'pointer', padding: '5px 10px', fontSize: 11, fontWeight: 600,
              }}>초기화</button>
            )}
            <button onClick={onClose} style={{
              background: 'rgba(255,255,255,0.10)', border: 'none', borderRadius: 8,
              color: 'white', cursor: 'pointer', padding: '6px 10px', fontSize: 16, lineHeight: 1,
            }}>✕</button>
          </div>
        </div>
        {/* 모달 내용 */}
        <div style={{ overflowY: 'auto', padding: '20px 24px 24px' }}>
          {children}
        </div>
      </div>
    </div>
  )
}

// ── ⑤ 추가 발생 비용 모달 (아코디언) ────────────────────────
function ExtraCostsModal({ params, setParams, onClose, onReset, totalKw, totalCount, isLowVoltage, estKepco, estKepcoBase, estKepcoDistance, estSafety, estElecSafety, estInsurance, insuranceNote }: {
  params: SimParams; setParams: (p: Partial<SimParams>) => void
  onClose: () => void; onReset: () => void
  totalKw: number; totalCount: number; isLowVoltage: boolean
  estKepco: number; estKepcoBase: number; estKepcoDistance: number
  estSafety: number; estElecSafety: number; estInsurance: number; insuranceNote: string
}) {
  const [openItems, setOpenItems] = useState<Record<string, boolean>>({})
  const toggle = (key: string) => setOpenItems(prev => ({ ...prev, [key]: !prev[key] }))

  const acItems = [
    {
      key: 'kepco',
      label: '한전 시설부담금',
      badge: <span style={{ fontSize: 11, color: 'rgba(255,200,100,0.75)', background: 'rgba(255,200,100,0.12)', borderRadius: 4, padding: '2px 7px' }}>저압 신청 시</span>,
      summary: estKepco > 0 ? `추산: ${estKepco.toLocaleString()}원` : '초기 1회 발생 · 초기 투자비 합산',
      inputArea: (<>
        <p style={{ color: 'rgba(255,255,255,0.45)', fontSize: 13, display: 'block', marginBottom: 6 }}>한전 시설부담금 (원)</p>
        <SNum value={params.cost_kepco_burden ?? 0} onChange={v => setParams({ cost_kepco_burden: v })} step={100000} min={0}/>
        {(params.cost_kepco_burden ?? 0) > 0 && <p style={{ fontSize: 12, color: 'rgba(255,255,255,0.45)', marginTop: 4, fontVariantNumeric: 'tabular-nums' }}>{(params.cost_kepco_burden ?? 0).toLocaleString()}원</p>}
      </>),
      detail: (<p style={{ fontSize: 12, color: 'rgba(255,255,255,0.55)', lineHeight: 1.75 }}>
        한전에 전기공급 신청 시 발생하는 비용(기본시설부담금 + 거리시설부담금). 저압 5kW까지 336,600원 + 초과분 133,100원/kW, 고압 26,400원/kW (공중 포설 기준). 인입거리가 기본거리(공중 200m)를 넘으면 초과 1m당 58,300원이 가산됩니다. 지중 포설·단상 계약은 별도 산정. 세금계산서 발행. <strong style={{ color: 'rgba(52,211,153,0.85)' }}>부가세(10%) 포함 금액입니다.</strong>
      </p>),
    },
    {
      key: 'safety',
      label: '사용전검사 · 안전관리자 · 감리비',
      badge: <span style={{ fontSize: 11, color: 'rgba(248,113,113,0.80)', background: 'rgba(248,113,113,0.10)', borderRadius: 4, padding: '2px 7px' }}>75kW 이상</span>,
      summary: estSafety > 0 ? `추산: ${estSafety.toLocaleString()}원` : '초기 1회 발생 · 75kW 이상 의무',
      inputArea: (<>
        <p style={{ color: 'rgba(255,255,255,0.45)', fontSize: 13, display: 'block', marginBottom: 6 }}>사용전검사·안전관리자·감리비 합계 (원)</p>
        <SNum value={params.cost_safety_inspection ?? 0} onChange={v => setParams({ cost_safety_inspection: v })} step={100000} min={0}/>
        {(params.cost_safety_inspection ?? 0) > 0 && <p style={{ fontSize: 12, color: 'rgba(255,255,255,0.45)', marginTop: 4, fontVariantNumeric: 'tabular-nums' }}>{(params.cost_safety_inspection ?? 0).toLocaleString()}원</p>}
      </>),
      detail: (<p style={{ fontSize: 12, color: 'rgba(255,255,255,0.55)', lineHeight: 1.75 }}>
        총 설치 용량 75kW 이상 시 한전 사용전검사 및 전기감리 의무. 이에스앤에이치(ES&H) 등 대행업체에서 일괄 처리. <strong style={{ color: 'rgba(52,211,153,0.85)' }}>부가세(10%) 포함 금액입니다.</strong><br/>
        <span style={{ color: 'rgba(255,255,255,0.65)' }}>사용전검사 수수료 (한국전기안전공사 2026년도 요율, VAT 10% 포함):</span><br/>
        · 저압: 기본료 103,400원 + {totalKw}kW × 430원 = <strong style={{ color: 'rgba(255,255,255,0.75)' }}>{Math.round((94_000 + totalKw * 391) * 1.1).toLocaleString()}원</strong><br/>
        · 고압 300kW이하: 기본료 327,800원 + kW당 726원<br/>
        · 고압 300kW초과: 기본료 614,900원 + kW당 347원<br/>
        <span style={{ color: 'rgba(255,255,255,0.65)' }}>감리비 추산 (부가세 포함):</span><br/>
        · 75~100kW: 약 150만원 &nbsp;· 100~500kW: 약 300만원 &nbsp;· 500kW 초과: 약 500만원<br/>
        세금계산서 발행.
      </p>),
    },
    {
      key: 'elec_safety',
      label: '전기안전관리대행비',
      badge: <span style={{ fontSize: 11, color: 'rgba(248,113,113,0.80)', background: 'rgba(248,113,113,0.10)', borderRadius: 4, padding: '2px 7px' }}>75kW 이상</span>,
      summary: estElecSafety > 0 ? `추산: ${estElecSafety.toLocaleString()}원/월` : '월 발생 · 75kW 이상 의무',
      inputArea: (<>
        <p style={{ color: 'rgba(255,255,255,0.45)', fontSize: 13, display: 'block', marginBottom: 6 }}>전기안전관리대행비 (원/월)</p>
        <SNum value={params.monthly_elec_safety ?? 0} onChange={v => setParams({ monthly_elec_safety: v })} step={10000} min={0}/>
        {(params.monthly_elec_safety ?? 0) > 0 && <p style={{ fontSize: 12, color: 'rgba(255,255,255,0.45)', marginTop: 4, fontVariantNumeric: 'tabular-nums' }}>{(params.monthly_elec_safety ?? 0).toLocaleString()}원/월</p>}
      </>),
      detail: (<p style={{ fontSize: 12, color: 'rgba(255,255,255,0.55)', lineHeight: 1.75 }}>
        전기안전관리대행업체에 매월 납부하는 의무 계약 수수료. 세금계산서 발행. <strong style={{ color: 'rgba(52,211,153,0.85)' }}>부가세(10%) 포함 금액입니다.</strong><br/>
        <span style={{ color: 'rgba(255,255,255,0.65)' }}>월 수수료 (한국전기안전공사 2026년도 요율, VAT 10% 포함):</span><br/>
        {isLowVoltage ? <>
          · 저압 50kW이하: 104,390원 &nbsp;· 50~100kW: 119,350원<br/>
          · 100~200kW: 136,070원 &nbsp;· 200~300kW: 155,650원<br/>
          · 300~400kW: 237,710원 &nbsp;· 400kW초과: 285,120원
        </> : <>
          · 고압 100kW이하: 147,510원 &nbsp;· 100~200kW: 182,710원<br/>
          · 200~300kW: 202,730원 &nbsp;· 300~400kW: 296,780원<br/>
          · 400~500kW: 351,120원 &nbsp;· 500~600kW: 464,090원 이상
        </>}<br/>
        · 현재 {totalKw}kW({isLowVoltage ? '저압' : '고압'}) → <strong style={{ color: 'rgba(255,255,255,0.80)' }}>월 {(totalKw < 75 ? 0 : estElecSafety).toLocaleString()}원</strong>
      </p>),
    },
    {
      key: 'insurance',
      label: '충전시설 사고배상책임보험',
      badge: <span style={{ fontSize: 11, color: 'rgba(52,211,153,0.80)', background: 'rgba(52,211,153,0.10)', borderRadius: 4, padding: '2px 7px' }}>연간 입력</span>,
      summary: insuranceNote ? `⚠ ${insuranceNote}` : `추산: ${estInsurance.toLocaleString()}원/년 (${totalCount}대 기준)`,
      inputArea: (<>
        {insuranceNote && <p style={{ fontSize: 12, color: 'rgba(248,113,113,0.80)', marginBottom: 8, fontWeight: 600 }}>⚠ {insuranceNote}</p>}
        <p style={{ color: 'rgba(255,255,255,0.45)', fontSize: 13, display: 'block', marginBottom: 6 }}>연간 보험료 (원/년)</p>
        <SNum value={params.insurance_yearly ?? 0} onChange={v => setParams({ insurance_yearly: v })} step={10000} min={0}/>
        {(params.insurance_yearly ?? 0) > 0 && <p style={{ fontSize: 12, color: 'rgba(255,255,255,0.45)', marginTop: 4, fontVariantNumeric: 'tabular-nums' }}>{(params.insurance_yearly ?? 0).toLocaleString()}원/년</p>}
      </>),
      detail: (<p style={{ fontSize: 12, color: 'rgba(255,255,255,0.55)', lineHeight: 1.75 }}>
        EV 충전시설 화재·감전 등 사고 발생 시 제3자 손해를 보상하는 배상책임보험. 1년 단위 갱신. 연간 보험료를 입력하세요. <strong style={{ color: 'rgba(255,200,100,0.80)' }}>보험료는 부가세 면세 항목입니다.</strong><br/>
        <span style={{ color: 'rgba(255,255,255,0.65)' }}>충전기 대수별 연간 보험료 기준:</span><br/>
        · 1~5대: 20,000원 &nbsp;· 6~8대: 30,000원 &nbsp;· 9~10대: 40,000원<br/>
        · 11~15대: 50,000원 &nbsp;· 16~18대: 60,000원 &nbsp;· 19~20대: 70,000원<br/>
        · 21대 이상: 별도 확인 필요
      </p>),
    },
    {
      key: 'extra_notice',
      label: '⚠ 추가 비용 안내',
      badge: undefined as React.ReactNode,
      summary: '자가용설비 정기검사 · 법령 변경에 따른 비용',
      inputArea: null as React.ReactNode,
      detail: (<div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
        <p style={{ fontSize: 13, color: 'rgba(255,255,255,0.65)', lineHeight: 1.7 }}>
          · <strong style={{ color: 'rgba(255,200,100,0.85)' }}>자가용설비 정기검사</strong> — 전기안전관리법 제65조에 따라 자가용 전기설비는 주기적으로 정기검사를 받아야 합니다. 검사 주기 및 수수료는 설비 용량과 종류에 따라 다르며, 별도 비용이 발생할 수 있습니다.
        </p>
        <p style={{ fontSize: 13, color: 'rgba(255,255,255,0.65)', lineHeight: 1.7 }}>
          · <strong style={{ color: 'rgba(255,200,100,0.85)' }}>법령 변경에 따른 추가 비용</strong> — 전기안전관리법, 전기사업법, 환경부 충전인프라 관련 고시 등 관련 법령의 개정에 따라 안전관리 의무, 수수료 요율, 검사 기준 등이 변경될 수 있으며, 이에 따른 추가 비용이 발생할 수 있습니다.
        </p>
        <p style={{ fontSize: 11, color: 'rgba(255,255,255,0.35)', marginTop: 2 }}>
          위 항목들은 현장 상황 및 법령 시행 시기에 따라 달라질 수 있으므로 설치 전 담당 기관에 확인하시기 바랍니다.
        </p>
      </div>),
    },
  ]

  return (
    <Modal title="⑤ 추가 발생 비용" onClose={onClose} onReset={onReset}>
      {/* 안내 배너 */}
      <div style={{ padding: '10px 14px 11px', marginBottom: 14, borderRadius: 8, background: 'rgba(99,102,241,0.10)', border: '1px solid rgba(99,102,241,0.22)' }}>
        <p style={{ color: 'rgba(255,255,255,0.75)', fontSize: 13, lineHeight: 1.65 }}>
          충전기 설치 외에 계약 조건·용량에 따라 추가로 발생하는 비용입니다. <strong style={{ color: 'rgba(255,200,100,0.80)' }}>고객 안내용 참고 정보이며 시뮬레이션 비용에는 포함되지 않습니다.</strong>
          <br/><span style={{ color: 'rgba(52,211,153,0.75)', fontSize: 12, fontWeight: 600 }}>※ 모든 금액은 부가세(VAT 10%) 포함 기준입니다.</span>
        </p>
      </div>

      {/* 추산 요약 */}
      <div style={{ marginBottom: 14, padding: '11px 14px 13px', borderRadius: 8, background: 'rgba(52,211,153,0.07)', border: '1px solid rgba(52,211,153,0.22)' }}>
        <p style={{ fontSize: 13, fontWeight: 700, color: 'rgba(52,211,153,0.90)', marginBottom: 6 }}>
          총 계약 용량 {totalKw}kW · {isLowVoltage ? '저압' : '고압'} 기준 추산 (안내용)
        </p>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 3 }}>
          {estKepco > 0 && <p style={{ fontSize: 12, color: 'rgba(255,255,255,0.60)' }}>한전 시설부담금: {estKepco.toLocaleString()}원 <span style={{ color: 'rgba(52,211,153,0.60)' }}>(부가세 포함)</span>{estKepcoDistance > 0 && ` (기본 ${estKepcoBase.toLocaleString()} + 거리 ${estKepcoDistance.toLocaleString()})`}</p>}
          {estSafety > 0 && <p style={{ fontSize: 12, color: 'rgba(255,255,255,0.60)' }}>사용전검사·안전관리·감리: {estSafety.toLocaleString()}원 <span style={{ color: 'rgba(52,211,153,0.60)' }}>(부가세 포함)</span></p>}
          {estElecSafety > 0 && <p style={{ fontSize: 12, color: 'rgba(255,255,255,0.60)' }}>전기안전관리대행비: {estElecSafety.toLocaleString()}원/월 <span style={{ color: 'rgba(52,211,153,0.60)' }}>(부가세 포함)</span></p>}
          {insuranceNote ? <p style={{ fontSize: 12, color: 'rgba(248,113,113,0.70)' }}>보험료: {insuranceNote}</p>
            : <p style={{ fontSize: 12, color: 'rgba(255,255,255,0.60)' }}>충전시설 배상책임보험: {estInsurance.toLocaleString()}원/년 ({totalCount}대 기준) <span style={{ color: 'rgba(255,255,255,0.35)' }}>(부가세 면세)</span></p>}
          {estKepco === 0 && estSafety === 0 && estElecSafety === 0 && <p style={{ fontSize: 12, color: 'rgba(255,255,255,0.35)' }}>충전기 구성을 입력하면 추산이 표시됩니다</p>}
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginTop: 10 }}>
          <span style={{ fontSize: 12, color: 'rgba(255,255,255,0.55)', whiteSpace: 'nowrap' }}>한전 신설 인입거리</span>
          <input type="number" min={0} step={10} value={params.kepco_line_distance_m || 0}
            onChange={e => setParams({ kepco_line_distance_m: Math.max(0, Number(e.target.value) || 0) })}
            style={{ width: 70, padding: '4px 7px', borderRadius: 5, fontSize: 12, background: 'rgba(0,0,0,0.25)', border: '1px solid rgba(255,255,255,0.15)', color: '#fff' }}/>
          <span style={{ fontSize: 12, color: 'rgba(255,255,255,0.55)' }}>m (기본거리 200m 초과분만 부담금 발생)</span>
        </div>
        <p style={{ fontSize: 11, color: 'rgba(255,255,255,0.30)', marginTop: 7 }}>
          한전 기본시설부담금·거리시설부담금(공중 포설·3상 기준)·사용전검사·전기감리·전기안전관리대행 요율표를 반영한 추산값. 계약 용량·전압 종별·인입거리가 바뀌면 자동 반영됩니다.
          <strong style={{ color: 'rgba(255,255,255,0.42)' }}> 현장 상황에 따라 실제 금액은 상이할 수 있으니 참고용으로만 활용하고 직접 수정해 주세요.</strong>
        </p>
      </div>

      {/* 아코디언 */}
      <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
        {acItems.map(item => {
          const isOpen = !!openItems[item.key]
          return (
            <div key={item.key} style={{ borderRadius: 10, border: `1px solid ${isOpen ? 'rgba(167,139,250,0.35)' : 'rgba(255,255,255,0.10)'}`, overflow: 'hidden', background: isOpen ? 'rgba(109,40,217,0.08)' : 'rgba(255,255,255,0.04)' }}>
              <button onClick={() => toggle(item.key)} style={{ width: '100%', display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '13px 14px', background: 'none', border: 'none', cursor: 'pointer', textAlign: 'left', gap: 10 }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap', flex: 1, minWidth: 0 }}>
                  <span style={{ fontSize: 14, fontWeight: 700, color: isOpen ? 'rgba(167,139,250,0.95)' : 'rgba(255,255,255,0.85)' }}>{item.label}</span>
                  {item.badge}
                  {!isOpen && <span style={{ fontSize: 11, color: 'rgba(255,255,255,0.35)' }}>{item.summary}</span>}
                </div>
                <span style={{ fontSize: 14, color: 'rgba(255,255,255,0.40)', flexShrink: 0 }}>{isOpen ? '▲' : '▼'}</span>
              </button>
              {isOpen && (
                <div style={{ padding: '0 14px 16px' }}>
                  <div style={{ marginBottom: item.inputArea ? 12 : 0 }}>{item.detail}</div>
                  {item.inputArea}
                </div>
              )}
            </div>
          )
        })}
      </div>
    </Modal>
  )
}

// ── 관리자 기본값 업데이트 모달 ──────────────────────────────
const DEFAULTS_KEY = 'sim_global_defaults'
type ChargerTypeEntry = { label: string; kw: number; default_rate: number; default_cost_unit: number; default_cost_install: number; default_monthly_ops: number }
type CompanyEntry = { id: string; name: string; managers: string[] }
type GlobalDefaults = Partial<SimParams> & {
  vkw_presets?: { label: string; kwh: number }[]
  custom_charger_types?: ChargerTypeEntry[]
  companies?: CompanyEntry[]
}
function loadCompanies(): CompanyEntry[] {
  return loadGlobalDefaults().companies ?? []
}
function isRegisteredManager(name: string): boolean {
  const companies = loadCompanies()
  return companies.some(c => c.managers.some(m => m.trim() === name.trim()))
}
const BASE_CHARGER_TYPES: ChargerTypeEntry[] = [
  { label: '완속 7kW',     kw: 7,   default_rate: 320, default_cost_unit: 1_100_000,  default_cost_install: 3_300_000,  default_monthly_ops: 33_000 },
  { label: '중속 30kW',    kw: 30,  default_rate: 350, default_cost_unit: 3_300_000,  default_cost_install: 5_500_000,  default_monthly_ops: 88_000 },
  { label: '급속 50kW',    kw: 50,  default_rate: 400, default_cost_unit: 11_000_000, default_cost_install: 7_700_000,  default_monthly_ops: 110_000 },
  { label: '급속 100kW',   kw: 100, default_rate: 450, default_cost_unit: 22_000_000, default_cost_install: 16_500_000, default_monthly_ops: 330_000 },
  { label: '초급속 200kW', kw: 200, default_rate: 500, default_cost_unit: 55_000_000, default_cost_install: 27_500_000, default_monthly_ops: 550_000 },
]
function loadGlobalDefaults(): GlobalDefaults {
  try { const s = localStorage.getItem(DEFAULTS_KEY); if (s) return JSON.parse(s) } catch {}
  return {}
}
function saveGlobalDefaults(d: GlobalDefaults) {
  localStorage.setItem(DEFAULTS_KEY, JSON.stringify(d))
}
function loadVkwPresets() {
  const d = loadGlobalDefaults()
  return d.vkw_presets ?? [{ label: '일반', kwh: 15 }, { label: '택시', kwh: 47 }, { label: '트럭', kwh: 58 }]
}
function loadChargerTypes(): ChargerTypeEntry[] {
  return loadGlobalDefaults().custom_charger_types ?? BASE_CHARGER_TYPES
}

function GlobalUpdateModal({ params, setParams, pwd, setPwd, pwdErr, setPwdErr, authed, setAuthed, correctPwd, onClose, onSaved }: {
  params: SimParams; setParams: (p: Partial<SimParams>) => void
  pwd: string; setPwd: (v: string) => void
  pwdErr: boolean; setPwdErr: (v: boolean) => void
  authed: boolean; setAuthed: (v: boolean) => void
  correctPwd: string; onClose: () => void
  onSaved: (chargerTypes: ChargerTypeEntry[], vkwPresets: { label: string; kwh: number }[]) => void
}) {
  const [tab, setTab] = useState<'charger' | 'cost' | 'period' | 'elec' | 'managers'>('charger')
  const [draft, setDraft] = useState<GlobalDefaults>({})
  const [vkwDraft, setVkwDraft] = useState<{ label: string; kwh: number }[]>(loadVkwPresets())
  const [ctDraft, setCtDraft] = useState<ChargerTypeEntry[]>(loadChargerTypes)
  const [companiesDraft, setCompaniesDraft] = useState<CompanyEntry[]>(loadCompanies)
  const [newCompanyName, setNewCompanyName] = useState('')
  const [newManagerInputs, setNewManagerInputs] = useState<Record<string, string>>({})
  const updDraft = (p: Partial<SimParams>) => setDraft(prev => ({ ...prev, ...p }))

  const merged: SimParams = { ...params, ...draft }
  const sectionLabel = { charger: '① 충전기 구성', cost: '② 비용 설정', period: '③ 기간·성장률', elec: '④ 고객 직접 납부', managers: '⑤ 담당자 관리' }
  const tabs = Object.keys(sectionLabel) as (keyof typeof sectionLabel)[]

  const handleSave = () => {
    const existing = loadGlobalDefaults()
    const prevTypes = existing.custom_charger_types ?? BASE_CHARGER_TYPES
    // index 기준으로 old label → new type 매핑 (rename 대응)
    const labelMap = new Map<string, ChargerTypeEntry>()
    prevTypes.forEach((old, i) => { if (ctDraft[i]) labelMap.set(old.label, ctDraft[i]) })
    const baseConfigs = (draft.charger_configs ?? params.charger_configs) as ChargerConfig[]
    const applyType = (c: ChargerConfig, t: ChargerTypeEntry): ChargerConfig => ({
      ...c, label: t.label, kw: t.kw, rate: c.rate, // rate는 사용자 설정 유지
      cost_unit: t.default_cost_unit, cost_install: t.default_cost_install,
      monthly_ops_unit: t.default_monthly_ops, monthly_ops_bulk: undefined,
    })
    const syncedConfigs = baseConfigs.map((c: ChargerConfig) => {
      // 1) 현재 label이 ctDraft에 있으면 해당 타입 기본값 동기화
      const exact = ctDraft.find(x => x.label === c.label)
      if (exact) return applyType(c, exact)
      // 2) prevTypes에서 현재 label을 찾아 index 기준 rename 대응
      const prevIdx = prevTypes.findIndex(x => x.label === c.label)
      if (prevIdx >= 0 && ctDraft[prevIdx]) return applyType(c, ctDraft[prevIdx])
      // 3) 그래도 없으면 첫 번째 타입으로 대체
      return ctDraft[0] ? applyType(c, ctDraft[0]) : c
    })
    saveGlobalDefaults({ ...existing, ...draft, charger_configs: syncedConfigs, vkw_presets: vkwDraft, custom_charger_types: ctDraft, companies: companiesDraft })
    setParams({ ...draft, charger_configs: syncedConfigs })
    onSaved(ctDraft, vkwDraft)
    onClose()
  }

  if (!authed) {
    return (
      <Modal title="⚙ 기본 설정값 업데이트" onClose={onClose}>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>
          <p style={{ fontSize: 13, color: 'rgba(255,255,255,0.65)', lineHeight: 1.7 }}>전체 기본 설정값을 변경합니다.<br/>관리자 비밀번호를 입력하세요.</p>
          <input type="password" value={pwd} autoFocus
            onChange={e => { setPwd(e.target.value); setPwdErr(false) }}
            onKeyDown={e => { if (e.key === 'Enter') { if (pwd === correctPwd) { setAuthed(true); setDraft({}) } else setPwdErr(true) } }}
            placeholder="비밀번호"
            style={{ padding: '10px 12px', borderRadius: 8, background: 'rgba(255,255,255,0.07)', border: `1px solid ${pwdErr ? '#f87171' : 'rgba(255,255,255,0.20)'}`, color: 'white', fontSize: 14 }}
          />
          {pwdErr && <p style={{ fontSize: 12, color: '#f87171' }}>비밀번호가 올바르지 않습니다.</p>}
          <button onClick={() => { if (pwd === correctPwd) { setAuthed(true); setDraft({}) } else setPwdErr(true) }} style={{
            padding: '11px 0', borderRadius: 8, fontSize: 13, fontWeight: 700, cursor: 'pointer',
            background: C.accent, border: 'none', color: C.sidebar1,
          }}>확인</button>
        </div>
      </Modal>
    )
  }

  const row2: React.CSSProperties = { display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 10 }
  const SLabel = ({ ch }: { ch: string }) => <p style={{ color: 'rgba(255,255,255,0.45)', fontSize: 10, marginBottom: 4 }}>{ch}</p>
  const SNum = ({ value, onChange, step = 1, min }: { value: number; onChange: (v: number) => void; step?: number; min?: number }) => (
    <input type="number" value={value} step={step} min={min}
      onChange={e => onChange(Number(e.target.value))}
      style={{ width: '100%', padding: '7px 10px', borderRadius: 7, background: 'rgba(255,255,255,0.07)', border: '1px solid rgba(255,255,255,0.15)', color: 'white', fontSize: 13 }}
    />
  )
  const totalCount = merged.charger_configs.reduce((s, c) => s + c.count, 0)

  return (
    <Modal title="⚙ 기본 설정값 업데이트" onClose={onClose} width={500}>
      <p style={{ fontSize: 11, color: 'rgba(255,255,255,0.40)', marginBottom: 14 }}>변경한 값이 새 기본값으로 저장됩니다. 초기화 버튼 클릭 시 이 값으로 복원됩니다.</p>
      {/* 탭 */}
      <div style={{ display: 'flex', gap: 6, marginBottom: 16 }}>
        {tabs.map(t => (
          <button key={t} onClick={() => setTab(t)} style={{
            flex: 1, padding: '7px 0', borderRadius: 7, fontSize: 10, fontWeight: 700, cursor: 'pointer',
            background: tab === t ? C.accent : 'rgba(255,255,255,0.07)',
            color: tab === t ? C.sidebar1 : 'rgba(255,255,255,0.55)',
            border: tab === t ? 'none' : '1px solid rgba(255,255,255,0.16)',
          }}>{sectionLabel[t]}</button>
        ))}
      </div>

      {tab === 'charger' && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
          {merged.charger_configs.map((cfg, ci) => {
            const updCfg = (patch: Partial<ChargerConfig>) => updDraft({ charger_configs: merged.charger_configs.map((c, i) => i === ci ? { ...c, ...patch } : c) })
            return (
              <div key={ci} style={{ padding: '12px', borderRadius: 8, background: 'rgba(255,255,255,0.05)', border: '1px solid rgba(255,255,255,0.09)' }}>
                <p style={{ fontSize: 11, fontWeight: 700, color: 'rgba(255,255,255,0.50)', textTransform: 'uppercase', letterSpacing: '0.06em', marginBottom: 10 }}>타입 {ci + 1}</p>
                <div style={{ marginBottom: 10 }}>
                  <SLabel ch="충전기 종류"/>
                  <select value={cfg.label} onChange={e => { const t = ctDraft.find(x => x.label === e.target.value); if (t) updCfg({ label: t.label, kw: t.kw, rate: t.default_rate, cost_unit: t.default_cost_unit, cost_install: t.default_cost_install, monthly_ops_unit: t.default_monthly_ops }) }}
                    style={{ width: '100%', padding: '7px 10px', borderRadius: 7, background: 'rgba(255,255,255,0.07)', border: '1px solid rgba(255,255,255,0.15)', color: 'white', fontSize: 13 }}>
                    {ctDraft.map(ct => <option key={ct.label} value={ct.label} style={{ background: '#1e1b4b' }}>{ct.label}</option>)}
                  </select>
                </div>
                <div style={row2}>
                  <div><SLabel ch="수량 (기)"/><SNum value={cfg.count} min={1} onChange={v => updCfg({ count: v })}/></div>
                  <div><SLabel ch="일 차량 (대)"/><SNum value={cfg.daily_ev} onChange={v => updCfg({ daily_ev: v })}/></div>
                </div>
                <div style={{ ...row2, marginTop: 8 }}>
                  <div><SLabel ch="요금 (원/kWh)"/><SNum value={cfg.rate} onChange={v => updCfg({ rate: v })}/></div>
                  <div><SLabel ch="1회 평균 충전량 (kWh)"/><SNum value={cfg.avg_kwh} step={0.1} onChange={v => updCfg({ avg_kwh: v })}/></div>
                </div>
              </div>
            )
          })}
          <div style={{ padding: '12px', borderRadius: 8, background: 'rgba(255,255,255,0.05)', border: '1px solid rgba(255,255,255,0.09)' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 10 }}>
              <p style={{ fontSize: 10, fontWeight: 700, color: 'rgba(255,255,255,0.40)', textTransform: 'uppercase', letterSpacing: '0.07em' }}>충전기 종류</p>
              <button onClick={() => setCtDraft(prev => [...prev, { label: '신규 충전기', kw: 300, default_rate: 500, default_cost_unit: 60_000_000, default_cost_install: 30_000_000, default_monthly_ops: 600_000 }])} style={{
                padding: '3px 10px', borderRadius: 5, fontSize: 10, fontWeight: 700, cursor: 'pointer',
                background: 'rgba(167,139,250,0.20)', border: '1px solid rgba(167,139,250,0.40)', color: 'rgba(167,139,250,0.90)',
              }}>+ 추가</button>
            </div>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
              {ctDraft.map((ct, ci) => (
                <div key={ci} style={{ padding: '10px', borderRadius: 7, background: 'rgba(255,255,255,0.04)', border: '1px solid rgba(255,255,255,0.08)' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 8 }}>
                    <SText value={ct.label} onChange={v => setCtDraft(prev => prev.map((x, i) => i === ci ? { ...x, label: v } : x))}
                      style={{ flex: 1, padding: '5px 8px', borderRadius: 6, background: 'rgba(255,255,255,0.10)', border: '1px solid rgba(255,255,255,0.20)', color: 'white', fontSize: 12, fontWeight: 700 }}/>
                    {ctDraft.length > 1 && (
                      <button onClick={() => setCtDraft(prev => prev.filter((_, i) => i !== ci))} style={{
                        padding: '3px 8px', borderRadius: 5, fontSize: 10, cursor: 'pointer',
                        background: 'rgba(248,113,113,0.15)', border: '1px solid rgba(248,113,113,0.35)', color: '#f87171',
                      }}>삭제</button>
                    )}
                  </div>
                  <div style={row2}>
                    <div><SLabel ch="kW"/><SNum value={ct.kw} onChange={v => setCtDraft(prev => prev.map((x, i) => i === ci ? { ...x, kw: v } : x))}/></div>
                    <div><SLabel ch="요금 (원/kWh)"/><SNum value={ct.default_rate} onChange={v => setCtDraft(prev => prev.map((x, i) => i === ci ? { ...x, default_rate: v } : x))}/></div>
                  </div>
                </div>
              ))}
            </div>
          </div>
          <div style={{ padding: '12px', borderRadius: 8, background: 'rgba(255,255,255,0.05)', border: '1px solid rgba(255,255,255,0.09)' }}>
            <p style={{ fontSize: 10, fontWeight: 700, color: 'rgba(255,255,255,0.40)', textTransform: 'uppercase', letterSpacing: '0.07em', marginBottom: 10 }}>차량 유형별 1회 충전량 기본값</p>
            <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
              {vkwDraft.map((p, pi) => (
                <div key={pi} style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                  <SText value={p.label} onChange={v => setVkwDraft(prev => prev.map((x, i) => i === pi ? { ...x, label: v } : x))}
                    style={{ width: 70, padding: '7px 8px', borderRadius: 7, background: 'rgba(255,255,255,0.07)', border: '1px solid rgba(255,255,255,0.15)', color: 'white', fontSize: 12 }}/>
                  <SNum value={p.kwh} step={1} onChange={v => setVkwDraft(prev => prev.map((x, i) => i === pi ? { ...x, kwh: v } : x))}/>
                  <span style={{ color: 'rgba(255,255,255,0.35)', fontSize: 11 }}>kWh</span>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}

      {tab === 'cost' && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
          {ctDraft.map((ct, ci) => (
            <div key={ci} style={{ padding: '12px', borderRadius: 8, background: 'rgba(255,255,255,0.05)', border: '1px solid rgba(255,255,255,0.09)' }}>
              <p style={{ fontSize: 11, fontWeight: 700, color: 'rgba(255,255,255,0.70)', marginBottom: 10 }}>{ct.label}</p>
              <div style={row2}>
                <div><SLabel ch="충전기 단가 (원)"/><SNum value={ct.default_cost_unit} step={100000} onChange={v => setCtDraft(prev => prev.map((x, i) => i === ci ? { ...x, default_cost_unit: v } : x))}/></div>
                <div><SLabel ch="설치비 (원)"/><SNum value={ct.default_cost_install} step={100000} onChange={v => setCtDraft(prev => prev.map((x, i) => i === ci ? { ...x, default_cost_install: v } : x))}/></div>
              </div>
              <div style={{ marginTop: 8 }}>
                <SLabel ch="월 운영비 (원/기)"/>
                <SNum value={ct.default_monthly_ops} step={10000} onChange={v => setCtDraft(prev => prev.map((x, i) => i === ci ? { ...x, default_monthly_ops: v } : x))}/>
              </div>
            </div>
          ))}
          <div style={row2}>
            <div><SLabel ch="PG 수수료 (%)"/><SReadonly value={merged.pg_fee_pct}/></div>
            <div><SLabel ch="수익 배분율 (%)"/><SNum value={merged.revenue_share_pct} onChange={v => updDraft({ revenue_share_pct: v })}/></div>
          </div>
        </div>
      )}

      {tab === 'period' && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
          <div><SLabel ch="운영 기간 (개월)"/><SNum value={merged.operation_months} onChange={v => updDraft({ operation_months: v })}/></div>
          <div><SLabel ch="연 EV 성장률 (%)"/><SNum value={merged.ev_growth_rate} step={0.1} onChange={v => updDraft({ ev_growth_rate: v })}/></div>
          <div><SLabel ch="연 요금 인상률 (%)"/><SNum value={merged.rate_increase} step={0.1} onChange={v => updDraft({ rate_increase: v })}/></div>
        </div>
      )}

      {tab === 'elec' && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
          <div style={{ display: 'flex', gap: 8 }}>
            {(['저압', '고압'] as const).map(t => (
              <button key={t} onClick={() => { const p = ELEC_PRESET[t]; updDraft({ elec_type: t, elec_basic_rate: p.basic, elec_kwh_rate: p.kwh, elec_climate_rate: p.climate, elec_fuel_rate: p.fuel, elec_fund_pct: p.fund }) }} style={{
                flex: 1, padding: '8px 0', borderRadius: 7, fontSize: 13, fontWeight: 700, cursor: 'pointer',
                background: merged.elec_type === t ? C.accent : 'rgba(255,255,255,0.07)',
                color: merged.elec_type === t ? C.sidebar1 : 'rgba(255,255,255,0.55)',
                border: merged.elec_type === t ? 'none' : '1px solid rgba(255,255,255,0.16)',
              }}>{t}</button>
            ))}
          </div>
          <div style={row2}>
            <div><SLabel ch="기본요금 (원/kW)"/><SReadonly value={merged.elec_basic_rate}/></div>
            <div><SLabel ch="전력량요금 (원/kWh)"/><SReadonly value={merged.elec_kwh_rate}/></div>
          </div>
          <div style={row2}>
            <div><SLabel ch="기후환경요금 (원/kWh)"/><SReadonly value={merged.elec_climate_rate}/></div>
            <div><SLabel ch="연료비조정액 (원/kWh)"/><SReadonly value={merged.elec_fuel_rate}/></div>
          </div>
          <div><SLabel ch="전력기금 (%)"/><SReadonly value={merged.elec_fund_pct}/></div>
        </div>
      )}

      {tab === 'managers' && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>
          <p style={{ fontSize: 11, color: 'rgba(255,255,255,0.40)', lineHeight: 1.6 }}>업체를 등록하고 각 업체별 담당자를 추가하세요. 등록된 담당자만 시뮬레이터의 담당자 확인 기능을 사용할 수 있습니다.</p>
          {/* 업체 추가 */}
          <div style={{ display: 'flex', gap: 8 }}>
            <input
              type="text"
              value={newCompanyName}
              onChange={e => setNewCompanyName(e.target.value)}
              onKeyDown={e => {
                if (e.key === 'Enter' && newCompanyName.trim()) {
                  setCompaniesDraft(prev => [...prev, { id: Date.now().toString(), name: newCompanyName.trim(), managers: [] }])
                  setNewCompanyName('')
                }
              }}
              placeholder="업체명 입력"
              style={{ flex: 1, padding: '8px 10px', borderRadius: 7, background: 'rgba(255,255,255,0.07)', border: '1px solid rgba(255,255,255,0.18)', color: 'white', fontSize: 13, outline: 'none' }}
            />
            <button
              onClick={() => {
                if (!newCompanyName.trim()) return
                setCompaniesDraft(prev => [...prev, { id: Date.now().toString(), name: newCompanyName.trim(), managers: [] }])
                setNewCompanyName('')
              }}
              style={{ padding: '8px 14px', borderRadius: 7, border: 'none', cursor: 'pointer', background: 'rgba(167,139,250,0.25)', color: 'rgba(167,139,250,0.95)', fontWeight: 700, fontSize: 12 }}
            >+ 업체 추가</button>
          </div>
          {companiesDraft.length === 0 && (
            <p style={{ fontSize: 12, color: 'rgba(255,255,255,0.25)', textAlign: 'center', padding: '20px 0' }}>등록된 업체가 없습니다.</p>
          )}
          {companiesDraft.map(company => (
            <div key={company.id} style={{ padding: '12px', borderRadius: 8, background: 'rgba(255,255,255,0.05)', border: '1px solid rgba(255,255,255,0.10)' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 10 }}>
                <p style={{ fontSize: 13, fontWeight: 700, color: 'rgba(255,255,255,0.85)' }}>🏢 {company.name}</p>
                <button
                  onClick={() => setCompaniesDraft(prev => prev.filter(c => c.id !== company.id))}
                  style={{ background: 'rgba(248,113,113,0.15)', border: '1px solid rgba(248,113,113,0.35)', color: '#f87171', fontSize: 10, padding: '3px 8px', borderRadius: 5, cursor: 'pointer' }}
                >삭제</button>
              </div>
              {/* 담당자 목록 */}
              <div style={{ display: 'flex', flexDirection: 'column', gap: 6, marginBottom: 8 }}>
                {company.managers.length === 0 && (
                  <p style={{ fontSize: 11, color: 'rgba(255,255,255,0.25)' }}>담당자가 없습니다.</p>
                )}
                {company.managers.map((mgr, mi) => (
                  <div key={mi} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '5px 10px', borderRadius: 6, background: 'rgba(255,200,100,0.07)', border: '1px solid rgba(255,200,100,0.18)' }}>
                    <span style={{ fontSize: 12, color: 'rgba(255,200,100,0.85)' }}>👤 {mgr}</span>
                    <button
                      onClick={() => setCompaniesDraft(prev => prev.map(c => c.id === company.id ? { ...c, managers: c.managers.filter((_, i) => i !== mi) } : c))}
                      style={{ background: 'none', border: 'none', color: 'rgba(255,255,255,0.35)', fontSize: 11, cursor: 'pointer', padding: '0 4px' }}
                    >✕</button>
                  </div>
                ))}
              </div>
              {/* 담당자 추가 */}
              <div style={{ display: 'flex', gap: 6 }}>
                <input
                  type="text"
                  value={newManagerInputs[company.id] ?? ''}
                  onChange={e => setNewManagerInputs(prev => ({ ...prev, [company.id]: e.target.value }))}
                  onKeyDown={e => {
                    if (e.key === 'Enter') {
                      const name = (newManagerInputs[company.id] ?? '').trim()
                      if (!name) return
                      setCompaniesDraft(prev => prev.map(c => c.id === company.id ? { ...c, managers: [...c.managers, name] } : c))
                      setNewManagerInputs(prev => ({ ...prev, [company.id]: '' }))
                    }
                  }}
                  placeholder="담당자 이름"
                  style={{ flex: 1, padding: '6px 9px', borderRadius: 6, background: 'rgba(255,255,255,0.07)', border: '1px solid rgba(255,255,255,0.15)', color: 'white', fontSize: 12, outline: 'none' }}
                />
                <button
                  onClick={() => {
                    const name = (newManagerInputs[company.id] ?? '').trim()
                    if (!name) return
                    setCompaniesDraft(prev => prev.map(c => c.id === company.id ? { ...c, managers: [...c.managers, name] } : c))
                    setNewManagerInputs(prev => ({ ...prev, [company.id]: '' }))
                  }}
                  style={{ padding: '6px 12px', borderRadius: 6, border: 'none', cursor: 'pointer', background: 'rgba(255,200,100,0.20)', color: 'rgba(255,200,100,0.90)', fontWeight: 700, fontSize: 11 }}
                >+ 추가</button>
              </div>
            </div>
          ))}
        </div>
      )}

      <div style={{ display: 'flex', gap: 8, marginTop: 20 }}>
        <button onClick={onClose} style={{
          flex: 1, padding: '10px 0', borderRadius: 8, fontSize: 13, cursor: 'pointer',
          background: 'none', border: '1px solid rgba(255,255,255,0.18)', color: 'rgba(255,255,255,0.50)',
        }}>취소</button>
        <button onClick={handleSave} style={{
          flex: 2, padding: '10px 0', borderRadius: 8, fontSize: 13, fontWeight: 700, cursor: 'pointer',
          background: C.accent, border: 'none', color: C.sidebar1,
        }}>기본값으로 저장</button>
      </div>
    </Modal>
  )
}

// ── 사이드바 메뉴 항목 ────────────────────────────────────────
function SideMenuItem({ num, title, desc, onClick, mobile }: {
  num: string; title: string; desc: string; onClick: () => void; mobile?: boolean
}) {
  const [hover, setHover] = useState(false)
  return (
    <button onClick={onClick}
      onMouseEnter={() => setHover(true)} onMouseLeave={() => setHover(false)}
      style={{
        width: '100%', display: 'flex', alignItems: 'center', gap: 14,
        padding: '16px 18px', cursor: 'pointer', border: 'none',
        textAlign: mobile ? 'center' : 'left' as const,
        justifyContent: mobile ? 'center' : undefined,
        flexDirection: mobile ? 'column' : 'row',
        background: hover ? 'rgba(255,255,255,0.10)' : 'rgba(255,255,255,0.03)',
        borderLeft: `3px solid ${hover ? C.accent : 'rgba(255,255,255,0.12)'}`,
        transition: 'all 0.15s',
        marginBottom: 2,
      }}>
      <span style={{
        color: C.accent, fontSize: 15, fontWeight: 800,
        minWidth: 24, textAlign: 'center' as const,
      }}>{num}</span>
      <div style={{ flex: mobile ? undefined : 1 }}>
        <p style={{ color: 'white', fontSize: 15, fontWeight: 700, marginBottom: 2, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{title}</p>
        <p style={{ color: 'rgba(255,255,255,0.42)', fontSize: 11, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{desc}</p>
      </div>
      {!mobile && <ChevronRight size={14} color="rgba(255,255,255,0.35)"/>}
    </button>
  )
}

// ── 사이드바 ──────────────────────────────────────────────────
function SettingsSidebar({ params, setParams, collapsed, setCollapsed, firstRec, onClose }: {
  params: SimParams
  setParams: (p: Partial<SimParams>) => void
  collapsed: boolean
  setCollapsed: (v: boolean) => void
  firstRec: MonthRecord | null
  onClose?: () => void
}) {
  const isMobile = useIsMobile()
  const [customChargerTypes, setCustomChargerTypes] = useState<ChargerTypeEntry[]>(loadChargerTypes)
  const [vkwPresets, setVkwPresets] = useState(loadVkwPresets)
  const typeNames = customChargerTypes.map(t => t.label)

  // 앱 시작 시 localStorage에 저장된 기본값을 params에 자동 적용
  useEffect(() => {
    const d = loadGlobalDefaults()
    const { vkw_presets: _v, custom_charger_types: _c, ...simDefaults } = d as GlobalDefaults & Record<string, unknown>
    if (Object.keys(simDefaults).length > 0) {
      setParams(simDefaults as Partial<import('./types').SimParams>)
      const saved = simDefaults as { charger_configs?: unknown[] }
      if (saved.charger_configs && Array.isArray(saved.charger_configs)) {
        setNumTypes(saved.charger_configs.length)
      }
    }
  }, []) // eslint-disable-line react-hooks/exhaustive-deps

  // 용량 기준 자동 추산값 — 한전 표준시설부담금(공중 기준)·사용전검사·전기감리·전기안전관리대행 요율표 반영
  const totalKw = params.charger_configs.reduce((s, c) => s + c.kw * c.count, 0)
  const isLowVoltage = params.elec_type === '저압'
  // 한전 기본시설부담금 단가표(공중공급 기준, 부가가치세 불포함 단가 × 1.1로 VAT 포함액 산정)
  // 저압 5kW까지 306,000원 / 초과분 121,000원·kW · 고압·특고압 24,000원/kW (부가세 불포함 단가)
  const estKepcoBase = isLowVoltage
    ? Math.round(totalKw <= 5 ? 306_000 * 1.1 : (306_000 + (totalKw - 5) * 121_000) * 1.1)
    : Math.round(totalKw * 24_000 * 1.1)
  // 한전 거리시설부담금(신설거리, 공중공급 기준·부가세 불포함 단가 × 1.1): 기본거리(공중 200m) 초과분 × 1m당 단가
  // 저압 삼상 53,000원/m · 고압·특고압 53,000원/m 가정(단상 저압은 48,000원/m — 대부분 3상 계약 기준으로 산정)
  const kepcoDistanceM = params.kepco_line_distance_m || 0
  const kepcoExcessM = Math.floor(Math.max(0, kepcoDistanceM - 200))
  const estKepcoDistance = Math.round(kepcoExcessM * 53_000 * 1.1)
  const estKepco = estKepcoBase + estKepcoDistance
  // 사용전검사 + 전기감리(초기 1회, 75kW 이상 대상): 용량 구간별 권장 예산 합산
  // 사용전검사 수수료 (한국전기안전공사 2026년도, VAT별도 × 1.1)
  const estSafetyInspection = isLowVoltage
    ? Math.round((94_000 + totalKw * 391) * 1.1)
    : totalKw <= 300
      ? Math.round((298_000 + totalKw * 660) * 1.1)
      : totalKw <= 1_000
        ? Math.round((559_000 + totalKw * 315) * 1.1)
        : Math.round((678_000 + Math.min(totalKw - 1_000, 99_000) * 152 + Math.max(0, totalKw - 100_000) * 32) * 1.1)
  // 감리비 추산 (감리대가 기준, 부가세 포함)
  const estSupervision = totalKw < 75 ? 0
    : totalKw <= 100 ? 1_500_000 : totalKw <= 500 ? 3_000_000 : 5_000_000
  const estSafety = totalKw < 75 ? 0 : estSafetyInspection + estSupervision
  // 전기안전관리대행 수수료 (한국전기안전공사 2026년도, VAT별도 × 1.1, 월 수수료)
  const estElecSafetyBase = isLowVoltage
    ? totalKw <= 50  ? 94_900
    : totalKw <= 100 ? 108_500
    : totalKw <= 200 ? 123_700
    : totalKw <= 300 ? 141_500
    : totalKw <= 400 ? 216_100
    :                  259_200
    : /* 고압 */
      totalKw <= 100  ? 134_100
    : totalKw <= 200  ? 166_100
    : totalKw <= 300  ? 184_300
    : totalKw <= 400  ? 269_800
    : totalKw <= 500  ? 319_200
    : totalKw <= 600  ? 421_900
    : totalKw <= 700  ? 542_900
    : totalKw <= 800  ? 656_300
    : totalKw <= 900  ? 814_200
    : totalKw <= 1_000  ? 945_200
    : totalKw <= 1_250  ? 1_236_400
    : totalKw <= 1_500  ? 1_487_600
    : totalKw <= 2_000  ? 2_032_400
    : totalKw <= 2_500  ? 2_705_900
    : totalKw <= 3_500  ? 3_096_900
    :                     3_716_300
  const estElecSafety = totalKw < 75 ? 0 : Math.round(estElecSafetyBase * 1.1)
  // 충전시설 사고배상책임보험(월): 총 충전기 대수 구간별
  const totalCount = params.charger_configs.reduce((s, c) => s + c.count, 0)
  // 연간 보험료 기준 (월 단가 × 12)
  const estInsurance = totalCount >= 21 ? 0
    : totalCount >= 19 ? 70_000
    : totalCount >= 16 ? 60_000
    : totalCount >= 11 ? 50_000
    : totalCount >= 9  ? 40_000
    : totalCount >= 6  ? 30_000
    : 20_000
  const insuranceNote = totalCount >= 21 ? '21대 이상: 별도 확인 필요' : null

  // 계약 용량·전압 종별·대수가 바뀔 때마다 추산값을 자동 반영 (최초 로드 시 불러온 값은 보존)
  const didAutoEstRef = useRef(false)
  useEffect(() => {
    if (!didAutoEstRef.current) { didAutoEstRef.current = true; return }
    setParams({
      cost_kepco_burden: estKepco,
      cost_safety_inspection: estSafety,
      monthly_elec_safety: estElecSafety,
      insurance_yearly: estInsurance,
    })
  }, [estKepco, estSafety, estElecSafety, estInsurance]) // eslint-disable-line react-hooks/exhaustive-deps

  const [numTypes, setNumTypes] = useState(params.charger_configs.length)
  const [activeModal, setActiveModal] = useState<'charger' | 'cost' | 'period' | 'elec' | 'installment' | 'extra_costs' | 'kwh_ref' | 'global_update' | null>(null)
  const [vehicleKwhOpen, setVehicleKwhOpen] = useState<number | null>(null)
  const [vehicleKwhModal, setVehicleKwhModal] = useState<number | null>(null)
  const [elecPwdOpen, setElecPwdOpen] = useState(false)
  const [elecPwdVal, setElecPwdVal] = useState('')
  const [elecPwdErr, setElecPwdErr] = useState(false)
  const ELEC_UPDATE_PWD = 'ego1234'
  const gd = () => loadGlobalDefaults()
  const managerInputRef = useRef<HTMLInputElement>(null)
  const [managerConfirmed, setManagerConfirmed] = useState(false)
  const [managerNotFound, setManagerNotFound] = useState(false)

  const confirmManager = () => {
    const name = managerInputRef.current?.value.trim()
    if (!name) return
    const companies = loadCompanies()
    const hasAnyManagers = companies.some(c => c.managers.length > 0)
    if (hasAnyManagers && !isRegisteredManager(name)) {
      setManagerNotFound(true)
      return
    }
    setManagerNotFound(false)
    setParams({ manager_name: name, manager_discount: 0 })
    setManagerConfirmed(true)
  }

  const changeNumTypes = (n: number) => {
    setNumTypes(n)
    const cur = params.charger_configs
    if (n > cur.length) {
      const def = customChargerTypes.find(x => x.label === '급속 50kW') ?? customChargerTypes[0]
      const added: ChargerConfig[] = Array.from({ length: n - cur.length }, () =>
        ({ label: def.label, kw: def.kw, count: 1, daily_ev: 5, avg_kwh: 30, rate: def.default_rate, cost_unit: def.default_cost_unit, cost_install: def.default_cost_install, monthly_ops_unit: def.default_monthly_ops }))
      setParams({ charger_configs: [...cur, ...added] })
    } else {
      setParams({ charger_configs: cur.slice(0, n) })
    }
  }

  const upd = (i: number, patch: Partial<ChargerConfig>) =>
    setParams({ charger_configs: params.charger_configs.map((c, idx) => idx === i ? { ...c, ...patch } : c) })

  return (
    <>
      <aside style={{
        width: isMobile ? '100%' : (collapsed ? 27 : 225),
        flexShrink: 0, display: 'flex', flexDirection: 'column', height: '100%',
        background: `linear-gradient(180deg, ${C.sidebar1} 0%, ${C.sidebar2} 100%)`,
        transition: 'width 0.35s cubic-bezier(0.4, 0, 0.2, 1)',
        overflow: 'hidden',
      }}>
        {collapsed && (
          <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', paddingTop: 10 }}>
            <button onClick={() => setCollapsed(false)} title="사이드바 열기" style={{
              background: 'rgba(255,255,255,0.10)', border: 'none', borderRadius: 6,
              color: 'white', cursor: 'pointer', padding: '5px 6px', marginTop: 2,
            }}>
              <ChevronsRight size={14}/>
            </button>
          </div>
        )}
        {!collapsed && (
          <>
          {/* 헤더 */}
          <div style={{ padding: '36px 14px 18px', flexShrink: 0, position: 'relative' }}>
            <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 10 }}>
              <p style={{ color: 'rgba(255,255,255,0.92)', fontWeight: 700, fontSize: 16, letterSpacing: '-0.3px' }}>시뮬레이터 설정</p>
              <p style={{ color: 'rgba(255,255,255,0.40)', fontSize: 11, textAlign: 'center', lineHeight: 1.6 }}>결제 방식 선택 후 아래 ①②③④ 항목을 순서대로 설정하세요.</p>
            </div>
            <button onClick={() => setCollapsed(true)} title="사이드바 접기" style={{
              position: 'absolute', top: 14, right: 12,
              background: 'none', border: 'none', color: 'rgba(255,255,255,0.40)',
              cursor: 'pointer', padding: '3px 4px', display: 'flex', alignItems: 'center',
            }}>
              {isMobile ? <X size={18}/> : <ChevronsLeft size={16}/>}
            </button>
          </div>
          <div style={{ height: 1, background: 'rgba(255,255,255,0.10)', margin: '0 14px', flexShrink: 0 }}/>

          {/* 메뉴 목록 */}
          <div style={{ flex: 1, overflowY: 'auto', paddingTop: 10, paddingBottom: 10 }}>
            {/* 결제방식 토글 */}
            <div style={{ padding: '10px 14px 14px', marginBottom: 4 }}>
              <p style={{ color: 'rgba(255,255,255,0.40)', fontSize: 10, fontWeight: 700, letterSpacing: '0.07em', textTransform: 'uppercase', marginBottom: 8 }}>결제 방식</p>
              <div style={{ display: 'flex', gap: 6 }}>
                {(['일시불', '할부'] as const).map(t => (
                  <button key={t} onClick={() => setParams({ payment_type: t })} className={params.payment_type === t ? 'payment-selected' : 'payment-unselected'} style={{
                    flex: 1, padding: '8px 0', borderRadius: 7, fontSize: 13, fontWeight: 700, cursor: 'pointer',
                    background: params.payment_type === t ? C.accent : 'rgba(255,255,255,0.07)',
                    color: params.payment_type === t ? C.sidebar1 : 'rgba(255,255,255,0.55)',
                    border: params.payment_type === t ? 'none' : '1px solid rgba(255,255,255,0.16)',
                    transition: 'all 0.15s',
                  }}>{t}</button>
                ))}
              </div>
              {params.payment_type === '할부' && (
                <button onClick={() => setActiveModal('installment')} style={{
                  width: '100%', marginTop: 8, padding: '8px 12px',
                  borderRadius: 7, background: 'rgba(167,139,250,0.12)',
                  border: '1px solid rgba(167,139,250,0.30)', cursor: 'pointer',
                  display: 'flex', justifyContent: 'space-between', alignItems: 'center',
                }}>
                  <span style={{ color: 'rgba(255,255,255,0.60)', fontSize: 11 }}>할부 조건</span>
                  <span style={{ color: C.accent, fontSize: 12, fontWeight: 700 }}>
                    {params.operation_months}개월
                  </span>
                </button>
              )}
              {params.payment_type === '일시불' && (
                <button onClick={() => setActiveModal('installment')} style={{
                  width: '100%', marginTop: 8, padding: '8px 12px',
                  borderRadius: 7, background: 'rgba(167,139,250,0.12)',
                  border: '1px solid rgba(167,139,250,0.30)', cursor: 'pointer',
                  display: 'flex', justifyContent: 'space-between', alignItems: 'center',
                }}>
                  <span style={{ color: 'rgba(255,255,255,0.60)', fontSize: 11 }}>일시불 조건</span>
                  <span style={{ color: C.accent, fontSize: 12, fontWeight: 700 }}>
                    {params.discount_rate}%
                  </span>
                </button>
              )}
            </div>
            <div style={{ height: 1, background: 'rgba(255,255,255,0.08)', margin: '0 14px 6px' }}/>
            <SideMenuItem num="①" title="충전기 구성" desc={`${params.charger_configs.length}종 · ${params.charger_configs.reduce((s,c)=>s+c.count,0)}대`} onClick={() => setActiveModal('charger')} mobile={isMobile}/>
            <SideMenuItem num="②" title="비용 설정" desc="초기투자 · 운영비 · 수수료" onClick={() => setActiveModal('cost')} mobile={isMobile}/>
            <SideMenuItem num="③" title="기간 · 성장률" desc={`${params.operation_months}개월 · 성장률 ${params.ev_growth_rate}%`} onClick={() => setActiveModal('period')} mobile={isMobile}/>
            <SideMenuItem num="④" title="고객 직접 납부" desc={`전기요금 · 통신비`} onClick={() => setActiveModal('elec')} mobile={isMobile}/>
            <SideMenuItem num="⑤" title="추가 발생 비용" desc="한전부담금 · 안전관리 · 감리" onClick={() => setActiveModal('extra_costs')} mobile={isMobile}/>
            <SideMenuItem num="⑥" title="충전량 참고" desc="차종별 일평균 추정 충전량" onClick={() => setActiveModal('kwh_ref')} mobile={isMobile}/>
            <div style={{ margin: '10px 14px 4px' }}>
              <button onClick={() => { setElecPwdVal(''); setElecPwdErr(false); setElecPwdOpen(false); setActiveModal('global_update') }} style={{
                width: '100%', padding: '9px 0', borderRadius: 8, fontSize: 11, fontWeight: 700, cursor: 'pointer',
                background: 'rgba(99,102,241,0.12)', border: '1px dashed rgba(99,102,241,0.40)', color: 'rgba(167,139,250,0.80)',
              }}>⚙ 기본 설정값 업데이트 (관리자)</button>
            </div>
          </div>
          </>
        )}
      </aside>

      {/* ── 모달들 ──────────────────────────────────────────── */}

      {/* 할부 조건 */}
      {activeModal === 'installment' && (() => {
        const initCost = params.cost_charger_unit * params.charger_configs.reduce((s,c)=>s+c.count,0) + params.cost_installation
        const isInst = params.payment_type === '할부'
        const n = params.operation_months
        const pmt = isInst && initCost > 0 && n > 0 ? initCost / n : 0
        const discounted = !isInst ? initCost * (1 - (params.discount_rate ?? 0) / 100) : 0
        const saved = initCost - discounted
        return (
          <Modal title={isInst ? '할부 조건' : '일시불 조건'} onClose={() => setActiveModal(null)}>
            {isInst && pmt > 0 && (
              <div style={{ borderRadius: 10, overflow: 'hidden', border: '1px solid rgba(167,139,250,0.25)', marginBottom: 20 }}>
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: 1, background: 'rgba(167,139,250,0.15)' }}>
                  {[
                    { label: '초기 투자액', value: `${Math.round(initCost).toLocaleString('ko-KR')}원` },
                    { label: '월 할부금',   value: `${Math.round(pmt).toLocaleString('ko-KR')}원` },
                    { label: '총 납부액',   value: `${Math.round(pmt * n).toLocaleString('ko-KR')}원` },
                  ].map(k => (
                    <div key={k.label} style={{ background: 'rgba(30,27,75,0.90)', padding: '12px 10px', textAlign: 'center' }}>
                      <p style={{ fontSize: 10, color: 'rgba(255,255,255,0.40)', marginBottom: 5 }}>{k.label}</p>
                      <p style={{ fontSize: 13, fontWeight: 800, color: C.accent }}>{k.value}</p>
                    </div>
                  ))}
                </div>
              </div>
            )}
            {!isInst && (
              <div style={{ borderRadius: 10, overflow: 'hidden', border: '1px solid rgba(167,139,250,0.25)', marginBottom: 20 }}>
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr 1fr', gap: 1, background: 'rgba(167,139,250,0.15)' }}>
                  {[
                    { label: '정상 투자액',   value: `${Math.round(initCost).toLocaleString('ko-KR')}원` },
                    { label: '할인 후 투자액', value: `${Math.round(discounted).toLocaleString('ko-KR')}원` },
                    { label: '할인 금액',      value: `${Math.round(saved).toLocaleString('ko-KR')}원` },
                  ].map(k => (
                    <div key={k.label} style={{ background: 'rgba(30,27,75,0.90)', padding: '12px 10px', textAlign: 'center' }}>
                      <p style={{ fontSize: 10, color: 'rgba(255,255,255,0.40)', marginBottom: 5 }}>{k.label}</p>
                      <p style={{ fontSize: 13, fontWeight: 800, color: C.accent }}>{k.value}</p>
                    </div>
                  ))}
                </div>
              </div>
            )}
            {isInst && (
              <div style={{ padding: '14px', borderRadius: 10, background: 'rgba(255,255,255,0.06)', marginBottom: 12, display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <span style={{ color: 'rgba(255,255,255,0.55)', fontSize: 13 }}>할부 기간</span>
                <span style={{ color: C.accent, fontSize: 15, fontWeight: 800 }}>{params.operation_months}개월</span>
              </div>
            )}
            {!isInst && (
              <div style={{ padding: '16px 14px 18px', borderRadius: 10, background: 'rgba(255,255,255,0.06)', marginBottom: 12 }}>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 10 }}>
                  <label style={{ ...lbl, marginBottom: 0, marginTop: 0 }}>할인율 (%)</label>
                  <span style={{ fontSize: 16, fontWeight: 800, color: C.accent }}>{params.discount_rate}</span>
                </div>
                <input type="range" min={0} max={50} step={1} value={params.discount_rate}
                  readOnly
                  style={{ width: '100%', cursor: 'default', accentColor: C.accent, display: 'block', height: 5, opacity: 0.5, pointerEvents: 'none' }}/>
              </div>
            )}
            <p style={{ color: 'rgba(255,255,255,0.35)', fontSize: 11, textAlign: 'center', marginTop: 4 }}>
              {isInst ? '할부 기간은 ③ 기간·성장률의 시뮬레이션 기간과 동일합니다' : '할인율은 ② 비용 설정에서 변경할 수 있습니다'}
            </p>
          </Modal>
        )
      })()}

      {/* ⑤ 추가 발생 비용 */}
      {activeModal === 'extra_costs' && (
        <ExtraCostsModal
          params={params}
          setParams={setParams}
          onClose={() => setActiveModal(null)}
          onReset={() => { const d = gd(); setParams({ cost_kepco_burden: d.cost_kepco_burden ?? 0, cost_safety_inspection: d.cost_safety_inspection ?? 0, monthly_elec_safety: d.monthly_elec_safety ?? 0, insurance_yearly: d.insurance_yearly ?? 0 }) }}
          totalKw={totalKw}
          totalCount={totalCount}
          isLowVoltage={isLowVoltage}
          estKepco={estKepco}
          estKepcoBase={estKepcoBase}
          estKepcoDistance={estKepcoDistance}
          estSafety={estSafety}
          estElecSafety={estElecSafety}
          estInsurance={estInsurance}
          insuranceNote={insuranceNote}
        />
      )}

      {/* ⑥ 충전량 참고 */}
      {activeModal === 'kwh_ref' && (
        <Modal title="⑥ 충전량 참고 (2025년 기준)" onClose={() => setActiveModal(null)} width={760}>
          {/* 메인 테이블 */}
          <div style={{ borderRadius: 10, overflow: 'hidden', border: '1px solid rgba(255,255,255,0.12)', marginBottom: 14 }}>
            <div style={{ padding: '10px 14px', background: 'rgba(255,255,255,0.06)', borderBottom: '1px solid rgba(255,255,255,0.07)' }}>
              <p style={{ color: 'rgba(255,255,255,0.50)', fontSize: 12, fontWeight: 700, letterSpacing: '0.06em', textTransform: 'uppercase' }}>⚡ 2025년 용도별·차종별 일평균 추정 충전량</p>
            </div>
            <div style={{ position: 'relative' }}>
              {isMobile && <div style={{ position: 'absolute', top: 0, right: 0, bottom: 0, width: 44, background: 'linear-gradient(to right, transparent, rgba(20,10,50,0.9))', pointerEvents: 'none', zIndex: 2, display: 'flex', alignItems: 'center', justifyContent: 'flex-end', paddingRight: 8 }}>
                <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 2 }}>
                  <span style={{ color: 'rgba(255,255,255,0.9)', fontSize: 18, lineHeight: 1 }}>›</span>
                  <span style={{ color: 'rgba(255,255,255,0.55)', fontSize: 8 }}>scroll</span>
                </div>
              </div>}
            <div style={{ overflowX: 'auto' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12, minWidth: 520 }}>
                <thead>
                  <tr style={{ background: 'rgba(109,40,217,0.45)' }}>
                    {[
                      { h: '차종',     w: 58 },
                      { h: '구분',     w: 100 },
                      { h: '적용 전비', w: 88 },
                      { h: '사업용(영업용)\n일 충전량', w: 80 },
                      { h: '비사업용(일반)\n일 충전량', w: 80 },
                      { h: '전체 평균\n일 충전량', w: 72 },
                      ...(!isMobile ? [{ h: '비고 및 설명', w: undefined as number | undefined }] : []),
                    ].map(({ h, w }) => (
                      <th key={h} style={{ padding: '9px 10px', color: 'rgba(255,255,255,0.85)', fontWeight: 700, textAlign: 'center', fontSize: 11, whiteSpace: 'pre-line', width: w, minWidth: w }}>
                        {h}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {[
                    {
                      type: '승합차', desc: '대형/중형 버스 등', rate: '1.08 km/kWh',
                      biz: '149.8', prv: '30.5', avg: '51.3', color: '#60A5FA',
                      note: ['노선 버스 등 운행량이 많아 완속이 아닌 고출력 급속충전기 필수 구역'],
                    },
                    {
                      type: '화물차', desc: '1톤 트럭(포터EV 등), 택배차', rate: '2.3 km/kWh',
                      biz: '58.6', prv: '14.8', avg: '20.9', color: '#34D399',
                      note: ['1톤 영업용 트럭 기준일 약 58~60kWh 충전 필요'],
                    },
                    {
                      type: '승용차', desc: '택시 / 개인용 승용차', rate: '4.2 km/kWh',
                      biz: '16.5', prv: '7.1', avg: '8.0', color: '#A78BFA',
                      note: ['영업용은 매일 약 16.5kWh 이상 소모 (렌트 포함)', '개인택시는 일 200km 이상 운행 (일 47kWh 이상 필요)'],
                    },
                    {
                      type: '특수차', desc: '청소차, 견인차 등', rate: '2.3 km/kWh',
                      biz: '61.1', prv: '10.8', avg: '37.3', color: '#FBBF24',
                      note: ['영업용 장비 운행에 따라 사업용 충전 수요가 화물차 수준으로 높음'],
                    },
                  ].map((row, i) => (
                    <React.Fragment key={row.type}>
                      <tr style={{ background: i % 2 === 0 ? 'rgba(255,255,255,0.03)' : 'transparent', borderTop: '1px solid rgba(255,255,255,0.07)', verticalAlign: 'top' }}>
                        <td style={{ padding: '11px 10px', fontWeight: 800, color: row.color, fontSize: 13, textAlign: 'center' }}>{row.type}</td>
                        <td style={{ padding: '11px 10px', color: 'rgba(255,255,255,0.65)', fontSize: 12 }}>{row.desc}</td>
                        <td style={{ padding: '11px 10px', color: 'rgba(255,255,255,0.50)', fontSize: 12, textAlign: 'center', fontWeight: 600 }}>{row.rate}</td>
                        <td style={{ padding: '11px 10px', textAlign: 'center', fontWeight: 800, color: C.accent, fontSize: 14 }}>{row.biz} <span style={{ fontSize: 10, fontWeight: 400 }}>kWh</span></td>
                        <td style={{ padding: '11px 10px', textAlign: 'center', color: 'rgba(255,255,255,0.70)', fontSize: 13 }}>{row.prv} <span style={{ fontSize: 10 }}>kWh</span></td>
                        <td style={{ padding: '11px 10px', textAlign: 'center', color: 'rgba(255,255,255,0.55)', fontSize: 13 }}>{row.avg} <span style={{ fontSize: 10 }}>kWh</span></td>
                        {!isMobile && (
                          <td style={{ padding: '11px 10px', color: 'rgba(255,255,255,0.55)', fontSize: 11, lineHeight: 1.6 }}>
                            {row.note.map((n, ni) => <div key={ni}>• {n}</div>)}
                          </td>
                        )}
                      </tr>
                      {isMobile && (
                        <tr style={{ background: i % 2 === 0 ? 'rgba(255,255,255,0.03)' : 'transparent' }}>
                          <td colSpan={6} style={{ padding: '6px 12px 12px', color: 'rgba(255,255,255,0.55)', fontSize: 11, lineHeight: 1.7, borderBottom: '1px solid rgba(255,255,255,0.07)' }}>
                            {row.note.map((n, ni) => <div key={ni}>• {n}</div>)}
                          </td>
                        </tr>
                      )}
                    </React.Fragment>
                  ))}
                </tbody>
              </table>
            </div>
            </div>
          </div>

          {/* 주요 수치 요약 */}
          <div style={{ padding: '12px 14px', borderRadius: 10, background: 'rgba(255,255,255,0.05)', marginBottom: 14 }}>
            <p style={{ color: 'rgba(255,255,255,0.40)', fontSize: 12, fontWeight: 700, letterSpacing: '0.07em', textTransform: 'uppercase', marginBottom: 10 }}>📌 주요 수치 요약</p>
            {[
              { label: '영업용 화물차 (1톤 포터/봉고 EV)', value: '58.6 kWh/일', sub: '배터리 1회 완충 분량' },
              { label: '영업용 전기택시 (렌트 제외)',        value: '47.6 kWh/일', sub: '일 200km 이상 운행 기준' },
              { label: '화물/승용 평균',                    value: '14.5 kWh/일', sub: '비사업용 일반 차량 기준' },
            ].map((item, i) => (
              <div key={i} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '8px 0', borderBottom: i < 2 ? '1px solid rgba(255,255,255,0.06)' : 'none' }}>
                <div>
                  <p style={{ color: 'rgba(255,255,255,0.80)', fontSize: 13, fontWeight: 600 }}>{item.label}</p>
                  <p style={{ color: 'rgba(255,255,255,0.35)', fontSize: 11, marginTop: 2 }}>{item.sub}</p>
                </div>
                <span style={{ color: C.accent, fontWeight: 800, fontSize: 15, whiteSpace: 'nowrap', marginLeft: 10 }}>{item.value}</span>
              </div>
            ))}
          </div>

          {/* 비고 */}
          <div style={{ padding: '10px 14px', borderRadius: 10, background: 'rgba(255,255,255,0.04)', marginBottom: 10 }}>
            <p style={{ color: 'rgba(255,255,255,0.40)', fontSize: 12, fontWeight: 700, letterSpacing: '0.07em', textTransform: 'uppercase', marginBottom: 8 }}>비고</p>
            <p style={{ color: 'rgba(255,255,255,0.55)', fontSize: 12, lineHeight: 1.7 }}>
              전비 산정 기준: 승용 4.2 km/kWh, 화물 2.3 km/kWh 적용.<br/>
              주행 환경(히터/에어컨, 적재량, 도심 vs 고속도로)에 따라 실제 충전량은 10~20% 이상 증가할 수 있습니다.
            </p>
          </div>

          {/* 출처 */}
          <div style={{ padding: '8px 2px' }}>
            <p style={{ color: 'rgba(255,255,255,0.60)', fontSize: 11, lineHeight: 1.7 }}>
              출처 1. 한국교통안전공단(TS) 『2025년 기준 자동차주행거리 분석 결과 발표』<br/>
              출처 2. 한국에너지공단 수송에너지 『전기자동차 에너지소비효율 및 등급표시에 관한 규정』
            </p>
          </div>
        </Modal>
      )}

      {/* ① 충전기 구성 */}
      {activeModal === 'charger' && (
        <Modal title="① 충전기 구성" onClose={() => setActiveModal(null)} onReset={() => { const d = gd(); setParams({ charger_configs: (d.charger_configs ?? DEFAULT_PARAMS.charger_configs).map((c: ChargerConfig) => ({ ...c })) }) }}>
          <div style={{ padding: '10px 14px 11px', marginBottom: 14, borderRadius: 8, background: 'rgba(99,102,241,0.10)', border: '1px solid rgba(99,102,241,0.22)' }}>
            <p style={{ color: 'rgba(196,191,239,0.55)', fontSize: 10, fontWeight: 700, letterSpacing: '0.07em', textTransform: 'uppercase', marginBottom: 4 }}>항목 안내</p>
            <p style={{ color: 'rgba(255,255,255,0.75)', fontSize: 12, lineHeight: 1.65 }}>설치할 충전기 종류와 대수, 하루 평균 이용 차량수, 충전 단가를 입력합니다. 여러 종류의 충전기를 동시에 추가해 복합 구성도 시뮬레이션할 수 있습니다.</p>
          </div>
          <div style={{ marginBottom: 20, padding: '14px 14px 16px', borderRadius: 10, background: 'rgba(255,255,255,0.06)' }}>
            <SLabel ch="충전기 종류 수"/>
            <div style={{ display: 'flex', gap: 22, marginTop: 8 }}>
              {[1,2,3].map(n => (
                <label key={n} style={{ display: 'flex', alignItems: 'center', gap: 7, cursor: 'pointer' }}>
                  <input type="radio" name="numTypes" checked={numTypes === n} onChange={() => changeNumTypes(n)}
                    style={{ accentColor: C.accent, cursor: 'pointer', width: 16, height: 16 }}/>
                  <span style={{ color: 'rgba(255,255,255,0.90)', fontSize: 15, fontWeight: 600 }}>{n}종</span>
                </label>
              ))}
            </div>
          </div>
          {params.charger_configs.map((cfg, i) => (
            <div key={i} style={{
              padding: '16px 14px 18px', borderRadius: 10,
              background: 'rgba(255,255,255,0.05)',
              marginTop: i > 0 ? 14 : 0,
              border: `1px solid rgba(${i === 0 ? '167,139,250' : i === 1 ? '52,211,153' : '251,146,60'},0.30)`,
            }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 16 }}>
                <span style={{ width: 12, height: 12, borderRadius: 3, background: TYPE_COLORS[i], flexShrink: 0 }}/>
                <span style={{ color: 'white', fontSize: 15, fontWeight: 700 }}>타입 {i+1}</span>
              </div>
              <div style={{ display: 'grid', gridTemplateColumns: '3fr 1.3fr', gap: 10, marginBottom: 14 }}>
                <SSelect value={cfg.label} onChange={v => {
                  const t = customChargerTypes.find(x => x.label === v); if (t) upd(i, { label: v, kw: t.kw, rate: t.default_rate, cost_unit: t.default_cost_unit, cost_install: t.default_cost_install, monthly_ops_unit: t.default_monthly_ops, monthly_ops_bulk: undefined })
                }} options={typeNames}/>
                <SNum value={cfg.count} onChange={v => upd(i, { count: v })} min={1}/>
              </div>
              <div style={{ height: 1, background: 'rgba(255,255,255,0.07)', margin: '0 0 14px' }}/>
              <div style={row2}>
                <div><SLabel ch="일 차량 (대)"/><SNum value={cfg.daily_ev} onChange={v => upd(i, { daily_ev: v })} step={0.5} min={0.5}/></div>
                <div><SLabel ch="요금 (원/kWh)"/><SNum value={cfg.rate} onChange={v => upd(i, { rate: v })} step={10} min={100}/></div>
              </div>
              <div>
                <SLabel ch="1기당 일평균 주행 충전량 (kWh)"/>
                <div style={{ display: 'flex', gap: 6, marginBottom: 6 }}>
                  {vkwPresets.map(p => (
                    <button key={p.label} onClick={() => upd(i, { avg_kwh: p.kwh })} style={{
                      flex: 1, padding: '5px 0', borderRadius: 6, fontSize: 12, fontWeight: 700, cursor: 'pointer',
                      background: cfg.avg_kwh === p.kwh ? C.accent : 'rgba(255,255,255,0.07)',
                      color: cfg.avg_kwh === p.kwh ? C.sidebar1 : 'rgba(255,255,255,0.55)',
                      border: cfg.avg_kwh === p.kwh ? 'none' : '1px solid rgba(255,255,255,0.16)',
                      transition: 'all 0.15s',
                    }}>
                      {p.label}<span style={{ fontSize: 10, fontWeight: 400, marginLeft: 3, opacity: 0.75 }}>{p.kwh}kWh</span>
                    </button>
                  ))}
                </div>
                {cfg.vehicle_kwh ? (
                  <div style={{ ...inp, display: 'flex', alignItems: 'center', justifyContent: 'space-between', opacity: 0.6, pointerEvents: 'none' } as React.CSSProperties}>
                    <span>{Math.round(cfg.vehicle_kwh.reduce((s, v) => s + v, 0) / cfg.vehicle_kwh.length * 10) / 10}</span>
                    <span style={{ fontSize: 10, color: 'rgba(255,255,255,0.45)' }}>평균 (개별 설정됨)</span>
                  </div>
                ) : (
                  <SNum value={cfg.avg_kwh} onChange={v => upd(i, { avg_kwh: v })} min={1}/>
                )}
                {Math.floor(cfg.daily_ev) >= 2 && (
                  <div style={{ marginTop: 8 }}>
                    <button onClick={() => setVehicleKwhModal(i)} style={{
                      width: '100%', padding: '7px 10px', borderRadius: 6, fontSize: 11, fontWeight: 700, cursor: 'pointer',
                      background: cfg.vehicle_kwh ? 'rgba(99,102,241,0.20)' : 'rgba(255,255,255,0.06)',
                      color: cfg.vehicle_kwh ? '#A78BFA' : 'rgba(255,255,255,0.45)',
                      border: `1px solid ${cfg.vehicle_kwh ? 'rgba(99,102,241,0.40)' : 'rgba(255,255,255,0.12)'}`,
                      textAlign: 'left', display: 'flex', justifyContent: 'space-between',
                    } as React.CSSProperties}>
                      <span>🚗 차량 개별 주행량 설정</span>
                      <span>{cfg.vehicle_kwh ? '개별 설정됨 ✓' : '→'}</span>
                    </button>
                    {!cfg.vehicle_kwh && (
                      <p style={{ marginTop: 5, fontSize: 10, color: 'rgba(255,255,255,0.30)', lineHeight: 1.5 }}>
                        미설정 시 위 충전량이 전체 차량에 동일 적용됩니다.
                      </p>
                    )}
                  </div>
                )}
              </div>
            </div>
          ))}
        </Modal>
      )}

      {/* 차량 개별 주행량 설정 모달 */}
      {vehicleKwhModal !== null && (() => {
        const ci = vehicleKwhModal
        const cfg = params.charger_configs[ci]
        if (!cfg) return null
        const n = Math.floor(cfg.daily_ev)
        const arr = cfg.vehicle_kwh && cfg.vehicle_kwh.length === n
          ? cfg.vehicle_kwh : Array(n).fill(cfg.avg_kwh)
        return (
          <VehicleKwhModal
            key={ci}
            total={n}
            avgKwh={cfg.avg_kwh}
            initArr={arr}
            presets={[...vkwPresets, { label: '직접입력', kwh: 0 }]}
            onSave={flat => {
              setParams({ charger_configs: params.charger_configs.map((c, idx) => idx === ci ? { ...c, vehicle_kwh: flat } : c) })
              setVehicleKwhModal(null)
            }}
            onReset={() => {
              setParams({ charger_configs: params.charger_configs.map((c, idx) => idx === ci ? { ...c, vehicle_kwh: undefined } : c) })
              setVehicleKwhModal(null)
            }}
            onClose={() => setVehicleKwhModal(null)}
          />
        )
      })()}

      {/* ② 비용 설정 */}
      {activeModal === 'cost' && (
        <Modal title="② 비용 설정" onClose={() => setActiveModal(null)} onReset={() => { const d = gd(); setParams({ cost_charger_unit: d.cost_charger_unit ?? DEFAULT_PARAMS.cost_charger_unit, cost_installation: d.cost_installation ?? DEFAULT_PARAMS.cost_installation, cost_other_init: d.cost_other_init ?? DEFAULT_PARAMS.cost_other_init, monthly_ops: d.monthly_ops ?? DEFAULT_PARAMS.monthly_ops, monthly_as: d.monthly_as ?? DEFAULT_PARAMS.monthly_as, monthly_comm: d.monthly_comm ?? DEFAULT_PARAMS.monthly_comm, monthly_other: d.monthly_other ?? DEFAULT_PARAMS.monthly_other, pg_fee_pct: d.pg_fee_pct ?? DEFAULT_PARAMS.pg_fee_pct, revenue_share_pct: d.revenue_share_pct ?? DEFAULT_PARAMS.revenue_share_pct, discount_rate: d.discount_rate ?? DEFAULT_PARAMS.discount_rate, charger_configs: (d.charger_configs ?? params.charger_configs).map((c: ChargerConfig) => ({ ...c })) }) }}>
          <div style={{ padding: '10px 14px 11px', marginBottom: 14, borderRadius: 8, background: 'rgba(99,102,241,0.10)', border: '1px solid rgba(99,102,241,0.22)' }}>
            <p style={{ color: 'rgba(196,191,239,0.55)', fontSize: 10, fontWeight: 700, letterSpacing: '0.07em', textTransform: 'uppercase', marginBottom: 4 }}>항목 안내</p>
            <p style={{ color: 'rgba(255,255,255,0.75)', fontSize: 12, lineHeight: 1.65 }}>충전기 단가·설치비 등 초기 투자 비용과 월 운영비를 입력합니다.<br/><span style={{ color: 'rgba(52,211,153,0.75)', fontSize: 11, fontWeight: 600 }}>※ 모든 금액은 부가세(VAT 10%) 포함 기준입니다.</span></p>
          </div>
          <div style={{ padding: '14px 14px 16px', borderRadius: 10, background: 'rgba(255,255,255,0.06)', marginBottom: 14 }}>
            <p style={{ color: 'rgba(255,255,255,0.40)', fontSize: 10, fontWeight: 700, letterSpacing: '0.07em', textTransform: 'uppercase', marginBottom: 12 }}>초기 투자</p>
            {params.charger_configs.length === 1 ? (
              <div style={row2}>
                <div><SLabel ch="충전기 단가 (원)"/><SNum value={params.charger_configs[0].cost_unit ?? params.cost_charger_unit} onChange={v => setParams({ charger_configs: params.charger_configs.map((c, i) => i === 0 ? { ...c, cost_unit: v } : c), cost_charger_unit: v })} step={100000}/></div>
                <div><SLabel ch="설치비 (원)"/><SNum value={params.charger_configs[0].cost_install ?? params.cost_installation} onChange={v => setParams({ charger_configs: params.charger_configs.map((c, i) => i === 0 ? { ...c, cost_install: v } : c), cost_installation: v })} step={100000}/></div>
              </div>
            ) : (
              <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
                {params.charger_configs.map((cfg, ci) => (
                  <div key={ci} style={{ padding: '10px 12px', borderRadius: 8, background: 'rgba(255,255,255,0.04)', border: '1px solid rgba(255,255,255,0.09)' }}>
                    <p style={{ fontSize: 11, fontWeight: 700, color: 'rgba(255,255,255,0.55)', marginBottom: 8 }}>{cfg.label} ×{cfg.count} <span style={{ fontWeight: 400, color: 'rgba(255,255,255,0.35)' }}>(1기당 입력)</span></p>
                    <div style={row2}>
                      <div>
                        <SLabel ch="충전기 단가 (원)"/>
                        <SNum value={cfg.cost_unit ?? params.cost_charger_unit} onChange={v => setParams({ charger_configs: params.charger_configs.map((c, i) => i === ci ? { ...c, cost_unit: v } : c) })} step={100000}/>
                      </div>
                      <div>
                        <SLabel ch="설치비 (원)"/>
                        <SNum value={cfg.cost_install ?? Math.round(params.cost_installation / params.charger_configs.reduce((s, c) => s + c.count, 0))} onChange={v => setParams({ charger_configs: params.charger_configs.map((c, i) => i === ci ? { ...c, cost_install: v } : c) })} step={100000}/>
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            )}
            <div style={{ marginTop: 10 }}>
              <SLabel ch="부속 시설물 (원)"/>
              <SNum value={params.cost_other_init ?? 0} onChange={v => setParams({ cost_other_init: v })} step={100000} min={0}/>
              <p style={{ fontSize: 10, color: 'rgba(255,255,255,0.35)', marginTop: 4, lineHeight: 1.5 }}>
                스탠드, 캐노피, 안전시설 등 부속 시설물 비용 (미입력 시 0)
              </p>
            </div>
            {params.payment_type === '일시불' && (
              <div style={{ marginTop: 10 }}>
                <SLabel ch="할인율 (%)"/>
                <SNum value={params.discount_rate} onChange={v => setParams({ discount_rate: v })} step={1} min={0}/>
              </div>
            )}
          </div>
          <div style={{ padding: '14px 14px 16px', borderRadius: 10, background: 'rgba(255,255,255,0.06)', marginBottom: 14 }}>
            <p style={{ color: 'rgba(255,255,255,0.40)', fontSize: 10, fontWeight: 700, letterSpacing: '0.07em', textTransform: 'uppercase', marginBottom: 12 }}>월 운영비</p>
            {params.charger_configs.length === 1 ? (
              <div style={row2}>
                <div><SLabel ch="운영비 (원/월)"/><SNum value={params.charger_configs[0].monthly_ops_unit ?? params.monthly_ops} onChange={v => setParams({ charger_configs: params.charger_configs.map((c, i) => i === 0 ? { ...c, monthly_ops_unit: v } : c), monthly_ops: v })} step={10000}/></div>
                <div><SLabel ch="AS비 (원/월)"/><SNum value={params.monthly_as} onChange={v => setParams({ monthly_as: v })} step={10000}/></div>
              </div>
            ) : (
              <>
                <div style={{ display: 'flex', flexDirection: 'column', gap: 10, marginBottom: 10 }}>
                  {params.charger_configs.map((cfg, ci) => (
                    <div key={ci} style={{ padding: '10px 12px', borderRadius: 8, background: 'rgba(255,255,255,0.04)', border: '1px solid rgba(255,255,255,0.09)' }}>
                      <p style={{ fontSize: 11, fontWeight: 700, color: 'rgba(255,255,255,0.55)', marginBottom: 8 }}>{cfg.label} ×{cfg.count} <span style={{ fontWeight: 400, color: 'rgba(255,255,255,0.35)' }}>(1기당 입력)</span></p>
                      <div style={row2}>
                        <div>
                          <SLabel ch="운영비 단가 (원/기)"/>
                          <SNum value={cfg.monthly_ops_unit ?? Math.round(params.monthly_ops / params.charger_configs.reduce((s, c) => s + c.count, 0))} onChange={v => setParams({ charger_configs: params.charger_configs.map((c, i) => i === ci ? { ...c, monthly_ops_unit: v } : c) })} step={10000}/>
                        </div>
                        {cfg.count >= 2 && (
                          <div>
                            <SLabel ch="2기 이상 단가 (원/기)"/>
                            <SNum value={cfg.monthly_ops_bulk ?? cfg.monthly_ops_unit ?? Math.round(params.monthly_ops / params.charger_configs.reduce((s, c) => s + c.count, 0))} onChange={v => setParams({ charger_configs: params.charger_configs.map((c, i) => i === ci ? { ...c, monthly_ops_bulk: v } : c) })} step={10000}/>
                          </div>
                        )}
                      </div>
                    </div>
                  ))}
                </div>
                <div><SLabel ch="AS비 (원/월)"/><SNum value={params.monthly_as} onChange={v => setParams({ monthly_as: v })} step={10000}/></div>
              </>
            )}
          </div>
          <div style={{ padding: '14px 14px 16px', borderRadius: 10, background: 'rgba(255,255,255,0.06)', marginBottom: 14 }}>
            <p style={{ color: 'rgba(255,255,255,0.40)', fontSize: 10, fontWeight: 700, letterSpacing: '0.07em', textTransform: 'uppercase', marginBottom: 12 }}>수수료 / 배분</p>
            <div style={row2}>
              <div><SLabel ch="PG 수수료 (%)"/><SReadonly value={params.pg_fee_pct}/></div>
              <div><SLabel ch="수익 배분율 (%)"/><SNum value={params.revenue_share_pct} onChange={v => setParams({ revenue_share_pct: v })}/></div>
            </div>
          </div>
          {/* 담당자 확인 — 일시불일 때만 표시 */}
          {/* 담당자 미등록 팝업 */}
          {managerNotFound && (
            <div style={{ position: 'fixed', inset: 0, zIndex: 9999, display: 'flex', alignItems: 'center', justifyContent: 'center', background: 'rgba(0,0,0,0.55)' }}
              onClick={() => setManagerNotFound(false)}>
              <div onClick={e => e.stopPropagation()} style={{ background: '#1e1b4b', border: '1px solid rgba(248,113,113,0.45)', borderRadius: 14, padding: '28px 32px', maxWidth: 320, textAlign: 'center', boxShadow: '0 8px 40px rgba(0,0,0,0.6)' }}>
                <p style={{ fontSize: 28, marginBottom: 10 }}>⚠️</p>
                <p style={{ fontSize: 15, fontWeight: 700, color: 'white', marginBottom: 8 }}>등록되지 않은 담당자입니다</p>
                <p style={{ fontSize: 12, color: 'rgba(255,255,255,0.55)', lineHeight: 1.6, marginBottom: 20 }}>입력하신 이름이 담당자 목록에 없습니다.<br/>관리자에게 담당자 등록을 요청하세요.</p>
                <button onClick={() => setManagerNotFound(false)} style={{ padding: '9px 28px', borderRadius: 8, border: 'none', cursor: 'pointer', background: '#f87171', color: 'white', fontWeight: 700, fontSize: 13 }}>확인</button>
              </div>
            </div>
          )}
          {params.payment_type === '일시불' && <div style={{ padding: '14px 14px 16px', borderRadius: 10, background: 'rgba(255,200,100,0.08)', border: `1px solid ${managerConfirmed ? 'rgba(255,200,100,0.50)' : 'rgba(255,200,100,0.20)'}` }}>
            <p style={{ color: 'rgba(255,200,100,0.70)', fontSize: 10, fontWeight: 700, letterSpacing: '0.07em', textTransform: 'uppercase', marginBottom: 12 }}>🔑 담당자 확인</p>
            {!managerConfirmed ? (
              <>
                <SLabel ch="담당자 이름"/>
                <div style={{ display: 'flex', gap: 8 }}>
                  <input
                    ref={managerInputRef}
                    defaultValue=""
                    onChange={() => setManagerNotFound(false)}
                    onKeyDown={e => {
                      if (e.key === 'Enter' && !e.nativeEvent.isComposing) confirmManager()
                    }}
                    placeholder="담당자 이름 입력"
                    autoComplete="off"
                    spellCheck={false}
                    lang="ko"
                    style={{
                      flex: 1, boxSizing: 'border-box',
                      background: 'rgba(255,255,255,0.10)',
                      border: `1px solid ${managerNotFound ? 'rgba(248,113,113,0.60)' : 'rgba(255,200,100,0.30)'}`,
                      borderRadius: 8, padding: '10px 12px', color: 'white', fontSize: 16, outline: 'none',
                    }}
                  />
                  <button
                    onClick={confirmManager}
                    style={{
                      padding: '10px 18px', borderRadius: 8, border: 'none', cursor: 'pointer',
                      background: 'rgba(255,200,100,0.80)', color: '#1e1560',
                      fontWeight: 700, fontSize: 13, whiteSpace: 'nowrap',
                    }}
                  >확인</button>
                </div>
              </>
            ) : (
              <>
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 14 }}>
                  <span style={{ color: 'rgba(255,200,100,0.90)', fontWeight: 700, fontSize: 15 }}>👤 {params.manager_name}</span>
                  <button onClick={() => { setManagerConfirmed(false); setManagerNotFound(false); if (managerInputRef.current) managerInputRef.current.value = ''; setParams({ manager_name: '', manager_discount: 0 }) }}
                    style={{ background: 'none', border: 'none', color: 'rgba(255,255,255,0.40)', fontSize: 12, cursor: 'pointer', padding: '2px 8px', textDecoration: 'underline' }}>변경</button>
                </div>
                <SLabel ch="충전기당 할인 금액 (원/대)"/>
                <SNum value={params.manager_discount} onChange={v => setParams({ manager_discount: v })} step={10000}/>
                {(() => {
                  const cnt = params.charger_configs.reduce((s,c) => s + c.count, 0)
                  const total = params.manager_discount * cnt
                  return total > 0 ? (
                    <p style={{ color: '#FB923C', fontSize: 11, marginTop: 8, fontWeight: 600 }}>
                      {params.manager_discount.toLocaleString()}원/대 × {cnt}대 = 총 <strong>{total.toLocaleString()}원</strong> 할인 적용
                    </p>
                  ) : null
                })()}
              </>
            )}
          </div>}
        </Modal>
      )}

      {/* ③ 기간·성장률 */}
      {activeModal === 'period' && (
        <Modal title="③ 기간 · 성장률" onClose={() => setActiveModal(null)} onReset={() => { const d = gd(); setParams({ operation_months: d.operation_months ?? DEFAULT_PARAMS.operation_months, ev_growth_rate: d.ev_growth_rate ?? DEFAULT_PARAMS.ev_growth_rate, rate_increase: d.rate_increase ?? DEFAULT_PARAMS.rate_increase }) }}>
          <div style={{ padding: '10px 14px 11px', marginBottom: 14, borderRadius: 8, background: 'rgba(99,102,241,0.10)', border: '1px solid rgba(99,102,241,0.22)' }}>
            <p style={{ color: 'rgba(196,191,239,0.55)', fontSize: 10, fontWeight: 700, letterSpacing: '0.07em', textTransform: 'uppercase', marginBottom: 4 }}>항목 안내</p>
            <p style={{ color: 'rgba(255,255,255,0.75)', fontSize: 12, lineHeight: 1.65 }}>시뮬레이션 운영 기간(개월)과 연간 EV 이용 증가율, 충전 요금 인상률을 설정합니다. 성장률이 높을수록 후반부 수익이 빠르게 늘어납니다.</p>
          </div>
          {[
            { label: '시뮬레이션 기간 (개월)', key: 'operation_months' as const, val: params.operation_months, min: 12, max: 120, step: 12 },
            { label: '연간 차량 증가율 (%)',    key: 'ev_growth_rate'   as const, val: params.ev_growth_rate,   min: 0,  max: 50,  step: 1 },
            { label: '연간 충전요금 인상률 (%)', key: 'rate_increase'   as const, val: params.rate_increase,   min: 0,  max: 20,  step: 1 },
          ].map(s => (
            <div key={s.key} style={{ padding: '16px 14px 18px', borderRadius: 10, background: 'rgba(255,255,255,0.06)', marginBottom: 12 }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 10 }}>
                <label style={{ ...lbl, marginBottom: 0, marginTop: 0 }}>{s.label}</label>
                <span style={{ fontSize: 16, fontWeight: 800, color: C.accent }}>{s.val}</span>
              </div>
              <input type="range" min={s.min} max={s.max} step={s.step} value={s.val}
                onChange={e => setParams({ [s.key]: Number(e.target.value) })}
                style={{ width: '100%', cursor: 'pointer', accentColor: C.accent, display: 'block', height: 5 }}/>
            </div>
          ))}
        </Modal>
      )}

      {/* ④ 고객 직접 납부 */}
      {activeModal === 'elec' && (
        <Modal title="④ 고객 직접 납부" onClose={() => setActiveModal(null)} onReset={() => { const d = gd(); setParams({ elec_type: d.elec_type ?? DEFAULT_PARAMS.elec_type, elec_basic_rate: d.elec_basic_rate ?? DEFAULT_PARAMS.elec_basic_rate, elec_kwh_rate: d.elec_kwh_rate ?? DEFAULT_PARAMS.elec_kwh_rate, elec_climate_rate: d.elec_climate_rate ?? DEFAULT_PARAMS.elec_climate_rate, elec_fuel_rate: d.elec_fuel_rate ?? DEFAULT_PARAMS.elec_fuel_rate, elec_fund_pct: d.elec_fund_pct ?? DEFAULT_PARAMS.elec_fund_pct, elec_vat_pct: d.elec_vat_pct ?? DEFAULT_PARAMS.elec_vat_pct, monthly_comm: d.monthly_comm ?? DEFAULT_PARAMS.monthly_comm }) }}>
          <div style={{ padding: '10px 14px 11px', marginBottom: 14, borderRadius: 8, background: 'rgba(99,102,241,0.10)', border: '1px solid rgba(99,102,241,0.22)' }}>
            <p style={{ color: 'rgba(196,191,239,0.55)', fontSize: 10, fontWeight: 700, letterSpacing: '0.07em', textTransform: 'uppercase', marginBottom: 4 }}>항목 안내</p>
            <p style={{ color: 'rgba(255,255,255,0.75)', fontSize: 12, lineHeight: 1.65 }}>전기요금·통신비 등 고객이 직접 납부하는 비용을 설정합니다. 운영사 비용(운영비·AS비 등)과 분리하여 관리합니다.</p>
          </div>
          {/* 상세 내역 카드 */}
          {firstRec && (
            <div style={{ borderRadius: 10, overflow: 'hidden', border: '1px solid rgba(255,255,255,0.12)', marginBottom: 20 }}>
              <div style={{ padding: '9px 14px', background: 'rgba(255,255,255,0.06)', borderBottom: '1px solid rgba(255,255,255,0.07)' }}>
                <p style={{ color: 'rgba(255,255,255,0.50)', fontSize: 10, fontWeight: 700, letterSpacing: '0.06em', textTransform: 'uppercase' }}>전기요금 상세 (1개월 기준)</p>
              </div>
              <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 1, background: 'rgba(255,255,255,0.07)' }}>
                {[
                  { label: '기본요금',   value: firstRec.elec_basic },
                  { label: '전력량요금', value: firstRec.elec_usage },
                  { label: '전력기금',   value: firstRec.elec_fund  },
                  { label: '부가세(VAT)', value: firstRec.elec_vat   },
                ].map(k => (
                  <div key={k.label} style={{ background: 'rgba(30,27,75,0.85)', padding: '10px 12px' }}>
                    <p style={{ fontSize: 10, color: 'rgba(255,255,255,0.40)', marginBottom: 4 }}>{k.label}</p>
                    <p style={{ fontSize: 14, fontWeight: 700, color: C.accent }}>{k.value.toLocaleString('ko-KR')} 원</p>
                  </div>
                ))}
              </div>
              <div style={{ padding: '11px 14px', background: 'rgba(109,40,217,0.35)', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <p style={{ fontSize: 12, color: 'rgba(255,255,255,0.65)', fontWeight: 600 }}>합계</p>
                <p style={{ fontSize: 17, fontWeight: 800, color: 'white' }}>{firstRec.elec_cost.toLocaleString('ko-KR')} 원</p>
              </div>
            </div>
          )}
          <div style={{ display: 'flex', gap: 8, marginBottom: 20 }}>
            {(['저압', '고압'] as const).map(t => (
              <button key={t} onClick={() => {
                const p = ELEC_PRESET[t]
                setParams({ elec_type: t, elec_basic_rate: p.basic, elec_kwh_rate: p.kwh, elec_climate_rate: p.climate, elec_fuel_rate: p.fuel, elec_fund_pct: p.fund })
              }} style={{
                flex: 1, padding: '10px 0', borderRadius: 8, fontSize: 15, fontWeight: 700, cursor: 'pointer',
                background: params.elec_type === t ? C.accent : 'rgba(255,255,255,0.07)',
                color: params.elec_type === t ? C.sidebar1 : 'rgba(255,255,255,0.55)',
                border: params.elec_type === t ? 'none' : '1px solid rgba(255,255,255,0.16)',
              }}>{t}</button>
            ))}
          </div>
          <div style={{ padding: '14px 14px 16px', borderRadius: 10, background: 'rgba(255,255,255,0.06)', marginBottom: 14 }}>
            <p style={{ color: 'rgba(255,255,255,0.40)', fontSize: 10, fontWeight: 700, letterSpacing: '0.07em', textTransform: 'uppercase', marginBottom: 12 }}>기본 요금</p>
            <div style={row2}>
              <div><SLabel ch="기본요금 (원/kW)"/><SReadonly value={params.elec_basic_rate}/></div>
              <div><SLabel ch="전력량요금 (원/kWh)"/><SReadonly value={params.elec_kwh_rate}/></div>
            </div>
          </div>
          <div style={{ padding: '14px 14px 16px', borderRadius: 10, background: 'rgba(255,255,255,0.06)', marginBottom: 14 }}>
            <p style={{ color: 'rgba(255,255,255,0.40)', fontSize: 10, fontWeight: 700, letterSpacing: '0.07em', textTransform: 'uppercase', marginBottom: 12 }}>부가 요금</p>
            <div style={row2}>
              <div><SLabel ch="기후환경요금 (원/kWh)"/><SReadonly value={params.elec_climate_rate}/></div>
              <div><SLabel ch="연료비조정액 (원/kWh)"/><SReadonly value={params.elec_fuel_rate}/></div>
            </div>
            <div><SLabel ch="전력기금 (%)"/><SReadonly value={params.elec_fund_pct}/></div>
          </div>
          <div style={{ padding: '14px 14px 16px', borderRadius: 10, background: 'rgba(255,255,255,0.06)' }}>
            <p style={{ color: 'rgba(255,255,255,0.40)', fontSize: 10, fontWeight: 700, letterSpacing: '0.07em', textTransform: 'uppercase', marginBottom: 12 }}>통신비</p>
            <SLabel ch="통신비 (원/월)"/>
            <SNum value={params.monthly_comm ?? 5000} onChange={v => setParams({ monthly_comm: v })} step={1000} min={0}/>
            <p style={{ fontSize: 10, color: 'rgba(255,255,255,0.35)', marginTop: 6 }}>SIM 통신 요금 등 고객이 직접 납부하는 월 통신 비용</p>
          </div>
        </Modal>
      )}

      {/* 관리자 기본값 업데이트 */}
      {activeModal === 'global_update' && <GlobalUpdateModal
        params={params}
        setParams={setParams}
        pwd={elecPwdVal} setPwd={setElecPwdVal}
        pwdErr={elecPwdErr} setPwdErr={setElecPwdErr}
        authed={elecPwdOpen} setAuthed={setElecPwdOpen}
        correctPwd={ELEC_UPDATE_PWD}
        onClose={() => { setActiveModal(null); setElecPwdOpen(false); setElecPwdVal(''); setElecPwdErr(false) }}
        onSaved={(cts, vkw) => { setCustomChargerTypes(cts); setVkwPresets(vkw) }}
      />}
    </>
  )
}

// ── 탭 정의 ──────────────────────────────────────────────────
const TABS = [
  { id: 'profit',      label: '📊 손익 추이',   mLabel: '📊 손익'   },
  { id: 'cost',        label: '📉 비용 분석',   mLabel: '📉 비용'   },
  { id: 'detail',      label: '📅 월별 데이터', mLabel: '📅 월별'   },
  { id: 'scenario',    label: '🔁 시나리오',    mLabel: '🔁 시나리오' },
  { id: 'compare',     label: '⚖️ 일시불·할부 비교', mLabel: '⚖️ 비교' },
  { id: 'bep_kwh',     label: '⚡ 손익분기 kWh', mLabel: '⚡ BEP'   },
  { id: 'report',      label: '📋 리포트',      mLabel: '📋 리포트' },
]

const SC_DEFAULTS = [
  { label: '낙관', dm: 1.5, cm: 1.2, gm: 1.25, km: 1.2, color: '#6366F1' },
  { label: '기본', dm: 1.0, cm: 1.0, gm: 1.0,  km: 1.0, color: '#1E3A5F' },
  { label: '보수', dm: 0.6, cm: 0.85, gm: 0.75, km: 0.8, color: C.red },
]

// ── 메인 콘텐츠 ──────────────────────────────────────────────
function MainContent({ params, setParams, onResult, isMobile = false, scrollContainerRef }: {
  params: SimParams; setParams: (p: Partial<SimParams>) => void
  onResult?: (r: SimResult | null) => void
  isMobile?: boolean
  scrollContainerRef?: React.RefObject<HTMLDivElement | null>
}) {
  const [result, setResult] = useState<SimResult | null>(null)
  const [loading, setLoading] = useState(false)
  const [tab, setTab] = useState('profit')
  const [kpiCollapsed, setKpiCollapsed] = useState(false)
  const stickyBarRef = useRef<HTMLDivElement>(null)
  const [stickyBarH, setStickyBarH] = useState(280)
  const [rTitle, setRTitle] = useState('EV충전소 수익 시뮬레이션 분석')
  const [rCo, setRCo] = useState('')
  const [rManager, setRManager] = useState('')
  const [scResults, setScResults] = useState<{ label: string; result: SimResult }[]>([])
  const [scRan, setScRan] = useState(false)
  const [toast, setToast] = useState<{ msg: string; type: 'success'|'info' } | null>(null)
  const showToast = (msg: string, type: 'success'|'info' = 'success') => {
    setToast({ msg, type })
    setTimeout(() => setToast(null), 3500)
  }
  useLayoutEffect(() => {
    const el = stickyBarRef.current
    if (!el) return
    const measure = () => setStickyBarH(el.offsetHeight)
    measure()
    const ro = new ResizeObserver(measure)
    ro.observe(el)
    return () => ro.disconnect()
  }, [result])

  const triggerDownload = async (blob: Blob, filename: string) => {
    // iOS Safari: Web Share API로 파일 공유 (다운로드 속성 미지원)
    const isIOS = /iPad|iPhone|iPod/.test(navigator.userAgent) ||
      (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1)
    if (isIOS) {
      try {
        const file = new File([blob], filename, { type: blob.type })
        if (navigator.canShare && navigator.canShare({ files: [file] })) {
          await navigator.share({ files: [file], title: filename })
          return
        }
      } catch (e) {
        // 사용자가 취소한 경우 무시
        if ((e as DOMException).name === 'AbortError') return
      }
      // Web Share API 미지원 시 새 탭으로 열기
      const url = URL.createObjectURL(blob)
      window.open(url, '_blank')
      showToast('📂 파일을 길게 눌러 "파일에 저장"을 선택하세요', 'info')
      setTimeout(() => URL.revokeObjectURL(url), 10000)
      return
    }
    // Android / Desktop
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = filename
    document.body.appendChild(a)
    a.click()
    setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url) }, 300)
  }
  const [scDefs, setScDefs] = useState(SC_DEFAULTS.map(s => ({ ...s })))
  const [scLocked, setScLocked] = useState<Record<string,boolean>>({ dm: false, cm: false, gm: false, km: false })
  const pdfAreaRef = useRef<HTMLDivElement>(null)
  const [pdfPreview, setPdfPreview] = useState<{ images: string[]; blob: Blob } | null>(null)
  const [xlsxPreview, setXlsxPreview] = useState(false)
  const [pdfCaptureSlide, setPdfCaptureSlide] = useState<number | null>(null)
  const pdfSlideRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    const style = document.createElement('style')
    style.textContent = `
      @keyframes fadeInUp { from { opacity: 0; transform: translateX(-50%) translateY(12px); } to { opacity: 1; transform: translateX(-50%) translateY(0); } }
      @keyframes blink { 0%, 100% { opacity: 1; } 50% { opacity: 0.2; } }
      button:not(.payment-selected):hover { filter: brightness(1.18); background-color: rgba(236,72,153,0.18) !important; border-color: rgba(236,72,153,0.55) !important; transition: filter 0.15s, background-color 0.15s, border-color 0.15s; }
      button.payment-selected:hover { filter: none !important; opacity: 1 !important; cursor: default; }
      button.payment-unselected:hover { filter: none !important; background-color: rgba(236,72,153,0.22) !important; border-color: rgba(236,72,153,0.70) !important; color: rgba(255,255,255,0.95) !important; }
      input[type="number"]:hover, input[type="text"]:hover { border-color: rgba(236,72,153,0.7) !important; background: rgba(236,72,153,0.08) !important; transition: border-color 0.15s, background 0.15s; }
      input[type="number"]:focus, input[type="text"]:focus { border-color: rgba(236,72,153,0.9) !important; outline: none; }
      select:hover { border-color: rgba(236,72,153,0.7) !important; background: rgba(236,72,153,0.08) !important; transition: border-color 0.15s; }
      @media print {
        body > * { display: none !important; }
        body > #print-area { display: block !important; }
        #print-area { display: block; width: 100%; }
        .no-print { display: none !important; }
        .recharts-wrapper, .recharts-surface { overflow: visible !important; }
        * { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
      }
    `
    document.head.appendChild(style)
    return () => { document.head.removeChild(style) }
  }, [])

  const runSim = useCallback(async () => {
    setLoading(true)
    try { const r = await simulate(params); setResult(r); onResult?.(r) } finally { setLoading(false) }
  }, [params])

  useEffect(() => { runSim() }, [runSim])

  const runScenario = useCallback(async () => {
    const res = await Promise.all(scDefs.map(async sc => {
      const p: SimParams = {
        ...params,
        ev_growth_rate: params.ev_growth_rate * sc.gm,
        charger_configs: params.charger_configs.map(c => ({
          ...c,
          daily_ev: c.daily_ev * sc.dm,
          rate: c.rate * sc.cm,
          avg_kwh: c.avg_kwh * sc.km,
        })),
      }
      return { label: sc.label, result: await simulate(p) }
    }))
    setScResults(res); setScRan(true)
  }, [scDefs, params])

  useEffect(() => {
    const timer = setTimeout(() => { runScenario() }, 200)
    return () => clearTimeout(timer)
  }, [runScenario])

  const scChartData = useMemo(() => {
    if (!scResults.length) return []
    return Array.from({ length: params.operation_months }, (_, i) => i + 1).map(m => {
      const row: Record<string, number> = { month: m }
      scResults.forEach(sc => { const rec = sc.result.records[m - 1]; row[sc.label] = rec ? Math.round(rec.cumulative / 10000) : 0 })
      return row
    })
  }, [scResults, params.operation_months])

  const fmt  = (n: number) => Math.round(n).toLocaleString('ko-KR')
  const fmtM = (n: number) => (n / 10000).toLocaleString('ko-KR', { minimumFractionDigits: 1, maximumFractionDigits: 1 })

  if (loading && !result) return (
    <div style={{ flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center', background: '#F5F3FF' }}>
      <div style={{ textAlign: 'center' }}>
        <div style={{ width: 34, height: 34, borderRadius: '50%', border: `3px solid ${C.accent}`, borderTopColor: 'transparent', margin: '0 auto 10px', animation: 'spin 0.8s linear infinite' }}/>
        <p style={{ fontSize: 12, color: '#9ca3af' }}>계산 중...</p>
      </div>
    </div>
  )
  // 비용 파이 데이터 (숨겨진 도넛 차트 & PPTX 공용) — result가 null이면 빈 배열
  const costPieData = useMemo(() => {
    if (!result) return []
    const rr = result
    const elec  = Math.round(rr.records.reduce((s, rec) => s + rec.elec_cost, 0) / 10000)
    const ops   = Math.round(rr.records.reduce((s, rec) => s + rec.ops, 0) / 10000)
    const as_   = Math.round(rr.records.reduce((s, rec) => s + rec.as_cost, 0) / 10000)
    const other = Math.round(rr.records.reduce((s, rec) => s + rec.other, 0) / 10000)
    const pg    = Math.round(rr.records.reduce((s, rec) => s + rec.pg_fee, 0) / 10000)
    const inst  = Math.round(rr.records.reduce((s, rec) => s + rec.installment, 0) / 10000)
    const total = elec + ops + as_ + other + pg + inst || 1
    return [
      { name: '전기요금', value: elec,  fill: '#0284C7' },
      { name: '운영비',   value: ops,   fill: '#F59E0B' },
      { name: 'AS비',     value: as_,   fill: '#10B981' },
      { name: '기타',     value: other, fill: '#F43F5E' },
      { name: 'PG수수료', value: pg,    fill: '#8B5CF6' },
      { name: '할부금',   value: inst,  fill: '#EC4899' },
    ].filter(e => e.value / total >= 0.02)
  }, [result])

  if (!result) return null

  const r = result
  const lastRec  = r.records[r.records.length - 1]
  const firstRec = r.records[0]

  const chartData = r.records.map(rec => ({
    month: rec.month,
    net: Math.round(rec.net_profit / 10000),
    cum: Math.round(rec.cumulative / 10000),
  }))
  const minCum = Math.min(...chartData.map(d => d.cum))
  const maxCum = Math.max(...chartData.map(d => d.cum))

  const pieData = [
    { name: '전기요금', value: r.records.reduce((s, x) => s + x.elec_cost, 0),    fill: '#0284C7' },
    { name: '운영비',   value: r.records.reduce((s, x) => s + x.ops, 0),          fill: '#6366F1' },
    { name: 'AS비',     value: r.records.reduce((s, x) => s + x.as_cost, 0),      fill: '#10B981' },
    { name: '기타',     value: r.records.reduce((s, x) => s + x.other, 0),        fill: '#F43F5E' },
    { name: 'PG수수료', value: r.records.reduce((s, x) => s + x.pg_fee, 0),       fill: '#8B5CF6' },
    { name: '할부금',   value: r.records.reduce((s, x) => s + x.installment, 0),  fill: '#EC4899' },
  ].filter(d => d.value > 0)
  const totalCostAll = pieData.reduce((s, d) => s + d.value, 0)

  const yearGroups: Record<string, typeof r.records> = {}
  r.records.forEach(rec => { if (!yearGroups[rec.year_label]) yearGroups[rec.year_label] = []; yearGroups[rec.year_label].push(rec) })
  const yearlyData = Object.entries(yearGroups).map(([yr, rows]) => {
    const netProfit = rows.reduce((s, x) => s + x.net_profit, 0)
    return {
      yr,
      myRevenue: rows.reduce((s, x) => s + x.my_revenue, 0),
      totalCost: rows.reduce((s, x) => s + x.total_cost, 0),
      netProfit,
      avgMonthly: Math.round(netProfit / rows.length),
      cumulative: rows[rows.length - 1].cumulative,
    }
  })

  const exportExcel = async () => {
    const wb = new ExcelJS.Workbook()
    wb.creator = 'ego 시뮬레이터'

    // ── 색상 팔레트 ──
    const COL_HEADER_BG = '4C1D95'  // 진한 보라
    const COL_SUBHEADER = '7C3AED'  // 중간 보라
    const COL_ROW_ODD   = 'F5F3FF'  // 연보라 줄
    const COL_PROFIT_POS = '065F46' // 흑자 글씨 (짙은 초록)
    const COL_PROFIT_NEG = '991B1B' // 적자 글씨 (짙은 빨강)
    const NUM_FMT = '#,##0'

    // ── 헬퍼: 셀 스타일 적용 ──
    const applyStyle = (cell: ExcelJS.Cell, opts: {
      bold?: boolean; color?: string; bg?: string; align?: ExcelJS.Alignment['horizontal']
      numFmt?: string; border?: boolean; size?: number
    }) => {
      if (opts.bold !== undefined) cell.font = { ...cell.font, bold: opts.bold, size: opts.size ?? 10, color: opts.color ? { argb: 'FF' + opts.color } : cell.font?.color }
      else if (opts.color || opts.size) cell.font = { ...cell.font, size: opts.size ?? 10, color: opts.color ? { argb: 'FF' + opts.color } : cell.font?.color }
      if (opts.bg) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + opts.bg } }
      if (opts.align) cell.alignment = { horizontal: opts.align, vertical: 'middle' }
      if (opts.numFmt) cell.numFmt = opts.numFmt
      if (opts.border) cell.border = {
        top: { style: 'thin', color: { argb: 'FFD1D5DB' } },
        bottom: { style: 'thin', color: { argb: 'FFD1D5DB' } },
        left: { style: 'thin', color: { argb: 'FFD1D5DB' } },
        right: { style: 'thin', color: { argb: 'FFD1D5DB' } },
      }
    }

    // ════════════════════════════════════════
    // 시트 1: 월별 상세 데이터
    // ════════════════════════════════════════
    const ws1 = wb.addWorksheet('월별 상세')
    ws1.views = [{ state: 'frozen', xSplit: 0, ySplit: 3 }]

    // 제목 행
    ws1.mergeCells('A1:M1')
    const titleCell = ws1.getCell('A1')
    titleCell.value = rTitle || 'EV 충전소 수익 시뮬레이션 – 월별 상세'
    applyStyle(titleCell, { bold: true, color: 'FFFFFF', bg: COL_HEADER_BG, align: 'center', size: 13 })
    ws1.getRow(1).height = 28

    // 부제목 행
    ws1.mergeCells('A2:M2')
    const subCell = ws1.getCell('A2')
    subCell.value = `${r.charger_summary} | 운영 ${params.operation_months}개월 | 초기투자 ${r.total_init_cost.toLocaleString()}원${rCo ? '  |  ' + rCo : ''}`
    applyStyle(subCell, { color: 'FFFFFF', bg: COL_SUBHEADER, align: 'center', size: 10 })
    ws1.getRow(2).height = 20

    // 헤더
    const hasInst = params.payment_type === '할부'
    const headers = ['월차','충전량(kWh)','총매출(원)','PG수수료(원)','PG차감매출(원)','운영사매출(원)','전기요금(원)','운영비(원)','AS비(원)','기타(원)',...(hasInst?['할부금(원)']:[]),'총비용(원)','순이익(원)','누적손익(원)']
    const colWidths = [8, 13, 14, 13, 14, 13, 13, 11, 11, 11, ...(hasInst?[11]:[]), 13, 13, 14]
    ws1.addRow(headers)
    const hRow = ws1.getRow(3)
    hRow.height = 22
    hRow.eachCell(cell => {
      applyStyle(cell, { bold: true, color: 'FFFFFF', bg: COL_HEADER_BG, align: 'center', border: true })
    })
    colWidths.forEach((w, i) => { ws1.getColumn(i + 1).width = w })

    // 데이터 행
    r.records.forEach((rec, idx) => {
      const bg = idx % 2 === 0 ? 'FFFFFF' : COL_ROW_ODD
      const row = ws1.addRow([
        rec.month, rec.kwh, rec.gross_revenue, rec.pg_fee, rec.net_revenue,
        rec.my_revenue, rec.elec_cost, rec.ops, rec.as_cost, rec.other,
        ...(hasInst ? [rec.installment] : []),
        rec.total_cost, rec.net_profit, rec.cumulative,
      ])
      row.height = 18
      row.eachCell((cell, colNum) => {
        applyStyle(cell, { bg, align: colNum === 1 ? 'center' : 'right', border: true })
        if (colNum >= 3) cell.numFmt = NUM_FMT  // 숫자 열에 천단위 서식
      })
      // 순이익/누적손익 색상
      const profitCell = row.getCell(12)
      const cumCell = row.getCell(13)
      applyStyle(profitCell, { bold: true, color: rec.net_profit >= 0 ? COL_PROFIT_POS : COL_PROFIT_NEG })
      applyStyle(cumCell, { bold: true, color: rec.cumulative >= 0 ? COL_PROFIT_POS : COL_PROFIT_NEG })
    })

    // ════════════════════════════════════════
    // 시트 2: 연도별 요약
    // ════════════════════════════════════════
    const ws2 = wb.addWorksheet('연도별 요약')
    ws2.mergeCells('A1:E1')
    const t2 = ws2.getCell('A1')
    t2.value = '연도별 요약'
    applyStyle(t2, { bold: true, color: 'FFFFFF', bg: COL_HEADER_BG, align: 'center', size: 13 })
    ws2.getRow(1).height = 28

    ws2.addRow(['연도', '연간 운영사매출(원)', '연간 총비용(원)', '연간 순이익(원)', '월평균 순이익(원)'])
    const h2 = ws2.getRow(2)
    h2.height = 22
    h2.eachCell(cell => applyStyle(cell, { bold: true, color: 'FFFFFF', bg: COL_HEADER_BG, align: 'center', border: true }))
    ;[10, 20, 20, 20, 20].forEach((w, i) => { ws2.getColumn(i + 1).width = w })

    yearlyData.forEach((y, idx) => {
      const bg = idx % 2 === 0 ? 'FFFFFF' : COL_ROW_ODD
      const row = ws2.addRow([y.yr, y.myRevenue, y.totalCost, y.netProfit, y.avgMonthly])
      row.height = 18
      row.eachCell((cell, col) => {
        applyStyle(cell, { bg, align: col === 1 ? 'center' : 'right', border: true })
        if (col >= 2) cell.numFmt = NUM_FMT
      })
      const netCell = row.getCell(4)
      applyStyle(netCell, { bold: true, color: y.netProfit >= 0 ? COL_PROFIT_POS : COL_PROFIT_NEG })
    })
    // 합계 행
    const totals = yearlyData.reduce((a, y) => ({ rev: a.rev + y.myRevenue, cost: a.cost + y.totalCost, net: a.net + y.netProfit }), { rev: 0, cost: 0, net: 0 })
    const sumRow = ws2.addRow(['합계', totals.rev, totals.cost, totals.net, yearlyData.length ? Math.round(totals.net / (yearlyData.length * 12)) : 0])
    sumRow.height = 20
    sumRow.eachCell((cell, col) => {
      applyStyle(cell, { bold: true, bg: 'E0E7FF', align: col === 1 ? 'center' : 'right', border: true })
      if (col >= 2) cell.numFmt = NUM_FMT
      if (col === 4) applyStyle(cell, { color: totals.net >= 0 ? COL_PROFIT_POS : COL_PROFIT_NEG })
    })
    sumRow.getCell(1).border = { top: { style: 'medium', color: { argb: 'FF7C3AED' } }, bottom: { style: 'medium', color: { argb: 'FF7C3AED' } }, left: { style: 'thin', color: { argb: 'FFD1D5DB' } }, right: { style: 'thin', color: { argb: 'FFD1D5DB' } } }

    // 월별 합계 행
    const mTotals = r.records.reduce((a, rec) => ({
      kwh: a.kwh + rec.kwh, gross: a.gross + rec.gross_revenue, pg: a.pg + rec.pg_fee,
      net_rev: a.net_rev + rec.net_revenue, my: a.my + rec.my_revenue,
      elec: a.elec + rec.elec_cost, ops: a.ops + rec.ops, as_c: a.as_c + rec.as_cost,
      other: a.other + rec.other, total: a.total + rec.total_cost, net: a.net + rec.net_profit,
    }), { kwh: 0, gross: 0, pg: 0, net_rev: 0, my: 0, elec: 0, ops: 0, as_c: 0, other: 0, total: 0, net: 0 })
    const mSumRow = ws1.addRow([
      '합계', Math.round(mTotals.kwh), mTotals.gross, mTotals.pg, mTotals.net_rev,
      mTotals.my, mTotals.elec, mTotals.ops, mTotals.as_c, mTotals.other,
      mTotals.total, mTotals.net, r.records[r.records.length - 1].cumulative,
    ])
    mSumRow.height = 20
    mSumRow.eachCell((cell, col) => {
      applyStyle(cell, { bold: true, bg: 'E0E7FF', align: col === 1 ? 'center' : 'right', border: true })
      if (col >= 2) cell.numFmt = NUM_FMT
      if (col === 12) applyStyle(cell, { color: mTotals.net >= 0 ? COL_PROFIT_POS : COL_PROFIT_NEG })
      if (col === 13) applyStyle(cell, { color: r.records[r.records.length - 1].cumulative >= 0 ? COL_PROFIT_POS : COL_PROFIT_NEG })
    })

    // 다운로드
    const buf = await wb.xlsx.writeBuffer()
    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
    triggerDownload(blob, `${rTitle || 'ego_시뮬레이션'}.xlsx`)
    showToast('✅ Excel 파일 다운로드가 시작됐습니다')
  }

  const exportPdf = async () => {
    if (!result) return
    const totalSlides = 5 + (scResults.length > 0 ? 2 : 0) + 1 + 1
    const previewImages: string[] = []
    const canvasList: HTMLCanvasElement[] = []
    try {
      for (let i = 0; i < totalSlides; i++) {
        setPdfCaptureSlide(i)
        await new Promise(res => setTimeout(res, 1400))
        if (!pdfSlideRef.current) continue
        const canvas = await html2canvas(pdfSlideRef.current, {
          scale: 2, useCORS: true, backgroundColor: '#2D2878', logging: false,
        })
        canvasList.push(canvas)
        previewImages.push(canvas.toDataURL('image/jpeg', 0.90))
      }
      setPdfCaptureSlide(null)
      const W = 960, H = 540
      const pdf = new jsPDF({ orientation: 'landscape', unit: 'px', format: [W, H] })
      canvasList.forEach((canvas, i) => {
        if (i > 0) pdf.addPage()
        pdf.addImage(canvas.toDataURL('image/jpeg', 0.90), 'JPEG', 0, 0, W, H)
      })
      setPdfPreview({ images: previewImages, blob: pdf.output('blob') })
    } catch (e) {
      setPdfCaptureSlide(null)
      showToast('❌ PDF 생성 실패', 'info')
    }
  }

  const exportPptx = async () => {

    const prs = new pptxgen()
    prs.layout = 'LAYOUT_16x9'
    const pptAvgMyRevenue = r.records.length > 0 ? r.records.reduce((s, rec) => s + rec.my_revenue, 0) / r.records.length : 0
    const pptAvgTotalCost = r.records.length > 0 ? r.records.reduce((s, rec) => s + rec.total_cost, 0) / r.records.length : 0

    // ── 팔레트 (딥 네이비 × 일렉트릭 틸 × 화이트) ──
    const NAVY  = '2D2878'   // 딥 인디고 퍼플 (primary dark)
    const NAVY2 = '3D3799'   // 미드 인디고
    const TEAL  = '7C6FCD'   // 라이트 퍼플 (accent)
    const TEAL2 = 'C4BFEF'   // 페일 퍼플 (tint)
    const GR    = '6366F1'   // 인디고 (흑자)
    const RD    = 'B91C1C'   // 딥 레드 (적자)
    const WH    = 'FFFFFF'
    const GY    = '64748B'
    const GY2   = 'F1F5F9'   // 연회색 배경
    const SLATE = '334155'   // 바디 텍스트

    // ── 테이블 공통 보더 ──
    const BD: [pptxgen.BorderProps, pptxgen.BorderProps, pptxgen.BorderProps, pptxgen.BorderProps] = [
      { type: 'solid', color: 'E2E8F0', pt: 0.4 },
      { type: 'solid', color: 'E2E8F0', pt: 0.4 },
      { type: 'solid', color: 'E2E8F0', pt: 0.4 },
      { type: 'solid', color: 'E2E8F0', pt: 0.4 },
    ]

    // ── 헬퍼: 콘텐츠 슬라이드 배경 ──
    const addBg = (slide: pptxgen.Slide) => {
      slide.background = { color: WH }
      // 왼쪽 세로 네이비 사이드바
      slide.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 0.18, h: 5.625, fill: { color: NAVY } })
      // 상단 얇은 틸 라인
      slide.addShape(prs.ShapeType.rect, { x: 0.18, y: 0, w: 9.82, h: 0.055, fill: { color: TEAL } })
    }

    // ── 헬퍼: 섹션 헤더 ──
    const addHeader = (slide: pptxgen.Slide, title: string, sub?: string) => {
      slide.addText(title.replace(/^[^\w가-힣]+/, ''), {
        x: 0.35, y: 0.12, w: 8.5, h: 0.48,
        fontSize: 22, bold: true, color: NAVY, fontFace: 'Calibri',
      })
      if (sub) slide.addText(sub, {
        x: 0.35, y: 0.6, w: 9.3, h: 0.24,
        fontSize: 9, color: GY, fontFace: 'Calibri', italic: true,
      })
      // 헤더 아래 구분선
      slide.addShape(prs.ShapeType.rect, { x: 0.35, y: 0.86, w: 9.47, h: 0.02, fill: { color: 'E2E8F0' } })
    }

    // ── 헬퍼: KPI 카드 ──
    const addKpiCard = (slide: pptxgen.Slide, x: number, y: number, w: number, label: string, value: string, sub?: string, valColor = TEAL) => {
      const H = 1.15
      slide.addShape(prs.ShapeType.rect, { x, y, w, h: H,
        fill: { color: GY2 }, line: { color: 'E2E8F0', width: 0.5 } })
      slide.addShape(prs.ShapeType.rect, { x, y, w, h: 0.05, fill: { color: valColor } })
      slide.addText(label,{ x: x + 0.12, y: y + 0.06, w: w - 0.24, h: 0.28, fontSize: 9, color: GY, fontFace: 'Calibri' })
      slide.addText(value, { x: x + 0.12, y: y + 0.3, w: w - 0.24, h: 0.42, fontSize: 14, bold: true, color: valColor, fontFace: 'Calibri' })
      if (sub) slide.addText(sub, { x: x + 0.12, y: y + 0.74, w: w - 0.24, h: 0.22, fontSize: 8, color: GY, fontFace: 'Calibri' })
    }

    // ════════════════════════════════
    // ════════════════════════════════
    // 슬라이드 1: 표지 — 다크 네이비 전면, 틸 액센트
    // ════════════════════════════════
    const s1 = prs.addSlide()
    s1.background = { color: NAVY }

    // ── 좌측 세로 틸 강조선 (슬림) ──
    s1.addShape(prs.ShapeType.rect, { x: 0, y: 0, w: 0.06, h: 5.625, fill: { color: TEAL } })

    // ── 우상단 브랜드 배지 ──
    s1.addShape(prs.ShapeType.rect, { x: 8.9, y: 0.32, w: 0.8, h: 0.34, fill: { color: TEAL } })
    s1.addText('ego', { x: 8.9, y: 0.32, w: 0.8, h: 0.34,
      fontSize: 13, bold: true, color: NAVY, fontFace: 'Calibri', align: 'center', valign: 'middle' })

    // ── 상단 레이블 ──
    s1.addText('EV충전소 수익성 시뮬레이션', { x: 0.42, y: 0.34, w: 8.2, h: 0.36,
      fontSize: 16, color: TEAL2, fontFace: 'Calibri', bold: false })

    // ── 메인 타이틀 ──
    s1.addText(rTitle || 'EV충전소 수익 분석', { x: 0.42, y: 0.72, w: 8.2, h: 1.2,
      fontSize: 40, bold: true, color: WH, fontFace: 'Calibri' })

    // ── 구분선 ──
    s1.addShape(prs.ShapeType.rect, { x: 0.42, y: 2.0, w: 2.0, h: 0.05, fill: { color: TEAL } })

    // ── 충전기 사양 ──
    s1.addText(r.charger_summary, { x: 0.42, y: 2.12, w: 8.2, h: 0.42,
      fontSize: 18, color: TEAL2, fontFace: 'Calibri', bold: true })

    // ── 운영 조건 ──
    s1.addText(`운영 기간 ${params.operation_months}개월  ·  초기 투자 ${r.total_init_cost.toLocaleString()}원${params.payment_type === '할부' ? ' (할부)' : ''}${rCo ? '  ·  ' + rCo : ''}`, {
      x: 0.42, y: 2.60, w: 8.2, h: 0.38, fontSize: 15, color: '94A3B8', fontFace: 'Calibri' })

    // ── 담당자 & 할인 (항상 표시, 입력값이 있을 때) ──
    const managerName = params.manager_name || ''
    const managerDiscount = params.manager_discount || 0
    if (managerName) {
      const discountText = managerDiscount > 0 ? `  ·  추가 할인 ${managerDiscount.toLocaleString()}원 적용` : ''
      s1.addText(`담당자: ${managerName}${discountText}`, {
        x: 0.42, y: 3.06, w: 8.2, h: 0.40, fontSize: 16, color: 'FB923C', fontFace: 'Calibri', bold: true })
    }

    s1.addText(`분석일: ${new Date().toLocaleDateString('ko-KR')}`, {
      x: 0.42, y: managerName ? 3.52 : 3.08, w: 9.16, h: 0.28, fontSize: 13, color: '475569', fontFace: 'Calibri', align: 'right' })

    // ── KPI 4개 — 하단 반투명 패널 ──
    s1.addShape(prs.ShapeType.rect, { x: 0, y: 3.88, w: 10, h: 1.745, fill: { color: '3D3799' } })
    s1.addShape(prs.ShapeType.rect, { x: 0, y: 3.88, w: 10, h: 0.04, fill: { color: 'A89EE8' } })

    const kpiDefs = [
      { label: '투자 회수 (BEP)', value: r.bep_month ? `${r.bep_month}개월차` : 'N/A',      accent: r.bep_month ? GR : RD },
      { label: '월 평균 순이익',  value: `${(r.avg_monthly_net/10000).toLocaleString('ko-KR', {minimumFractionDigits:1,maximumFractionDigits:1})}만원`, accent: TEAL },
      { label: `${params.operation_months}개월 총 순이익`, value: `${Math.round(r.records[r.records.length-1].cumulative/10000).toLocaleString('ko-KR')}만원`, accent: r.records[r.records.length-1].cumulative >= 0 ? GR : RD },
      { label: 'ROI',             value: `${r.roi.toFixed(1)}%`,                             accent: TEAL },
    ]
    kpiDefs.forEach((k, i) => {
      const kx = 0.25 + i * 2.44
      if (i > 0) s1.addShape(prs.ShapeType.rect, { x: kx - 0.04, y: 4.06, w: 0.01, h: 1.3, fill: { color: '1E3A5F' } })
      s1.addText(k.label, { x: kx, y: 4.0, w: 2.3, h: 0.32,
        fontSize: 13, color: 'C4BFEF', fontFace: 'Calibri', align: 'center' })
      s1.addText(k.value, { x: kx, y: 4.36, w: 2.3, h: 0.82,
        fontSize: 26, bold: true, color: WH, fontFace: 'Calibri', align: 'center' })
    })

    // ════════════════════════════════
    // 슬라이드 2: 손익 추이 – KPI 카드 4개 + 콤보 차트
    // ════════════════════════════════
    const s2 = prs.addSlide()
    addBg(s2); addHeader(s2, '투자 회수 분석', `본전까지 몇 개월? 매달 얼마를 버나? — 운영 ${params.operation_months}개월 시뮬레이션`)

    const lastCum = r.records[r.records.length-1].cumulative
    // KPI 3개: 월 운영사 매출 · 월 총비용 · 월 순이익
    addKpiCard(s2, 0.35, 0.98, 2.88, '월 운영사 매출', `${(pptAvgMyRevenue/10000).toFixed(1)}만원`,
      `PG ${params.pg_fee_pct}% 차감 후 ${params.revenue_share_pct}% 배분 · ${params.operation_months}개월 평균`, TEAL)
    addKpiCard(s2, 3.56, 0.98, 2.88, '월 총비용', `${(pptAvgTotalCost/10000).toFixed(1)}만원`,
      `전기요금·운영비·AS비·기타${params.payment_type === '할부' ? '·할부금' : ''} 합산 · ${params.operation_months}개월 평균`, RD)
    addKpiCard(s2, 6.77, 0.98, 2.88, '월 순이익', `${(r.avg_monthly_net/10000).toFixed(1)}만원`,
      `운영사 매출 - 총비용 · ${params.operation_months}개월 평균`, r.avg_monthly_net >= 0 ? GR : RD)

    const chartMonths = r.records.map(rec => `${rec.month}M`)
    const cumulatives = r.records.map(rec => Math.round(rec.cumulative / 10000))
    const bepIdx = r.bep_month ? r.bep_month - 1 : r.records.length
    const netBeforeBep = r.records.map((rec, i) => i < bepIdx ? Math.round(rec.net_profit / 1000) : 0)
    const netAfterBep  = r.records.map((rec, i) => i >= bepIdx ? Math.round(rec.net_profit / 1000) : 0)

    {
      ;(s2 as any).addChart([
        { type: prs.ChartType.bar,  data: [{ name: '회수 전(천원)',   labels: chartMonths, values: netBeforeBep }], options: { chartColors: ['F87171'] } },
        { type: prs.ChartType.bar,  data: [{ name: '수익 구간(천원)', labels: chartMonths, values: netAfterBep  }], options: { chartColors: ['6366F1'] } },
        { type: prs.ChartType.line, data: [{ name: '누적 손익(만원)', labels: chartMonths, values: cumulatives  }],
          options: { chartColors: ['F59E0B'], secondaryValAxis: true, secondaryCatAxis: true } },
      ], {
        x: 0.35, y: 2.16, w: 9.3, h: 3.18,
        barGrouping: 'stacked',
        valAxes: [
          { showValAxisTitle: true, valAxisTitle: '월 순이익 (천원)', valAxisTitleColor: GY, valAxisLabelColor: GY, valAxisLabelFontSize: 8 },
          { showValAxisTitle: true, valAxisTitle: '누적 손익 (만원)', valAxisTitleColor: 'F59E0B', valAxisLabelColor: 'F59E0B', valAxisLabelFontSize: 8 },
        ],
        catAxes: [{ catAxisLabelColor: GY, catAxisLabelFontSize: 7 }, { catAxisHidden: true }],
        showLegend: true, legendPos: 'b', legendFontSize: 9,
        plotAreaBorderColor: 'F1F5F9', valGridLine: { color: 'E2E8F0', size: 0.5 },
        showTitle: false,
      })
    }
    if (r.bep_month) {
      // 차트 플롯 영역 추정: x=0.35+0.72=1.07, w≈7.9 (우측 보조축 여백 0.69 제외)
      const chartLeft = 1.07
      const chartWidth = 7.9
      const chartTop = 2.28   // 차트 상단 y + 범례 높이
      const chartBottom = 5.22
      const N = r.records.length
      const bepX = chartLeft + ((r.bep_month - 1) / N) * chartWidth
      // 수직 BEP 선
      s2.addShape(prs.ShapeType.rect, {
        x: bepX, y: chartTop, w: 0.02, h: chartBottom - chartTop,
        fill: { color: 'FB923C' }, line: { color: 'FB923C', width: 0 },
      })
      // BEP 라벨 박스
      s2.addShape(prs.ShapeType.rect, {
        x: bepX - 0.5, y: chartTop, w: 1.05, h: 0.24,
        fill: { color: 'FB923C' },
      })
      s2.addText(`BEP (${r.bep_month}개월)`, {
        x: bepX - 0.5, y: chartTop, w: 1.05, h: 0.24,
        fontSize: 7.5, bold: true, color: WH, fontFace: 'Calibri', align: 'center', valign: 'middle',
      })
    }

    // ════════════════════════════════
    // 슬라이드 3: 손익 추이 – 연도별 요약 테이블
    // ════════════════════════════════
    const s3 = prs.addSlide()
    const s3Sub = params.manager_name
      ? `연도별 내 순이익 및 누적 투자 회수 현황  ·  담당자: ${params.manager_name}${params.manager_discount > 0 ? `  ·  추가 할인 ${params.manager_discount.toLocaleString()}원 적용` : ''}`
      : '연도별 내 순이익 및 누적 투자 회수 현황'
    addBg(s3); addHeader(s3, '연도별 수익 흐름', s3Sub)

    // 상단 3개 대형 스탯
    const totRev2 = yearlyData.reduce((s, y) => s + y.myRevenue, 0)
    const totCost2 = yearlyData.reduce((s, y) => s + y.totalCost, 0)
    const totNet2  = yearlyData.reduce((s, y) => s + y.netProfit, 0)
    ;[
      { label: '총 운영사 매출', value: `${Math.round(totRev2/10000).toLocaleString('ko-KR')}만원`, color: TEAL },
      { label: '총 비용',        value: `${Math.round(totCost2/10000).toLocaleString('ko-KR')}만원`, color: RD   },
      { label: '총 순이익',      value: `${Math.round(totNet2/10000).toLocaleString('ko-KR')}만원`,  color: totNet2 >= 0 ? GR : RD },
    ].forEach((s, i) => {
      const sx = 0.35 + i * 3.12
      s3.addShape(prs.ShapeType.rect, { x: sx, y: 0.98, w: 3.0, h: 0.96, fill: { color: GY2 } })
      s3.addShape(prs.ShapeType.rect, { x: sx, y: 0.98, w: 3.0, h: 0.04, fill: { color: s.color } })
      s3.addText(s.label, { x: sx+0.1, y: 1.05, w: 2.8, h: 0.24, fontSize: 9, color: GY, fontFace: 'Calibri', align: 'center' })
      s3.addText(s.value, { x: sx+0.1, y: 1.29, w: 2.8, h: 0.52, fontSize: 22, bold: true, color: s.color, fontFace: 'Calibri', align: 'center' })
    })

    // 연도별 테이블
    let yrCum = 0
    const yrHeaders = ['연도', '운영사 매출 (원)', '총비용 (원)', '순이익 (원)', '누적 손익 (원)']
    const yrDataRows = yearlyData.map(y => {
      yrCum += y.netProfit
      return { row: [y.yr, y.myRevenue.toLocaleString(), y.totalCost.toLocaleString(), y.netProfit.toLocaleString(), yrCum.toLocaleString()], net: y.netProfit, cum: yrCum }
    })
    const yrSumRow = ['합계', totRev2.toLocaleString(), totCost2.toLocaleString(), totNet2.toLocaleString(), totNet2.toLocaleString()]
    const yrAllRows = [yrHeaders, ...yrDataRows.map(d => d.row), yrSumRow]
    s3.addTable(yrAllRows.map((row, ri) => row.map((cell, ci) => {
      const isHead = ri === 0, isSum = ri === yrAllRows.length - 1
      const rowData = ri > 0 && !isSum ? yrDataRows[ri - 1] : null
      return {
        text: cell,
        options: {
          bold: isHead || isSum || ci === 0,
          color: isHead ? WH
            : isSum ? NAVY
            : ci === 3 ? (rowData && rowData.net >= 0 ? GR : RD)
            : ci === 4 ? (rowData && rowData.cum >= 0 ? GR : RD)
            : SLATE,
          fill: { color: isHead ? NAVY : isSum ? TEAL2 : ri % 2 === 1 ? GY2 : WH },
          align: ci === 0 ? 'center' : 'right',
          fontSize: isHead ? 9 : 10, fontFace: 'Calibri', valign: 'middle',
          border: BD,
        },
      }
    })), { x: 0.35, y: 2.06, w: 9.3, rowH: 0.36 })

    // ════════════════════════════════
    // 슬라이드 4: 흑자 전환 조건 — 민감도 분석 (소유자 핵심: 몇 대 와야 돈 되나?)
    // ════════════════════════════════
    const s4 = prs.addSlide()
    addBg(s4); addHeader(s4, '흑자 전환 조건', '충전요금 × 일 이용 차량수에 따른 월 손익 — 내 조건에서 몇 대가 와야 흑자인가?')

    // ── 슬라이드 4 내용: 민감도 히트맵 (흑자 전환 조건) ──
    const RATES_S = [200, 250, 300, 350, 400, 450, 500]
    const EVCTS_S = [1, 2, 3, 5, 7, 10, 15]
    const heatData: number[][] = EVCTS_S.map(ev =>
      RATES_S.map(rate => {
        const p2 = { ...params, charger_configs: params.charger_configs.map((c, idx) => idx === 0 ? { ...c, daily_ev: ev, rate } : c) }
        return Math.round(runSimulation(p2).records[0].net_profit / 1000)
      })
    )
    const allHeat = heatData.flat()
    const heatMin = Math.min(...allHeat), heatMax = Math.max(...allHeat)
    const heatColor = (v: number) => {
      const toHex = (n: number) => Math.round(Math.max(0, Math.min(255, n))).toString(16).padStart(2, '0')
      if (v < 0) {
        const t = heatMin < 0 ? Math.min(1, v / heatMin) : 0
        return `${toHex(220 - 70*t)}${toHex(38 + 10*t)}${toHex(38 + 10*t)}`
      }
      const t = heatMax > 0 ? Math.min(1, v / heatMax) : 0
      return `${toHex(255 - 210*t)}${toHex(255 - 140*t + 120*t)}${toHex(255 - 220*t)}`
    }
    const curEv0   = params.charger_configs[0]?.daily_ev ?? 0
    const curRate0 = params.charger_configs[0]?.rate ?? 0

    const heatTableRows: string[][] = [
      ['차량↓ / 요금→', ...RATES_S.map(r2 => `${r2}원`)],
      ...EVCTS_S.map((ev, ei) => [`${ev}대`, ...RATES_S.map((_r, ri2) => `${heatData[ei][ri2].toLocaleString()}`)]),
    ]
    s4.addTable(heatTableRows.map((row, ri) => row.map((cell, ci) => {
      const isHead = ri === 0 || ci === 0
      const ev   = ci > 0 && ri > 0 ? EVCTS_S[ri - 1] : 0
      const rate = ri > 0 && ci > 0 ? RATES_S[ci - 1] : 0
      const isCurrent = ev === curEv0 && rate === curRate0
      const val = ri > 0 && ci > 0 ? heatData[ri - 1][ci - 1] : 0
      const bg = isHead ? NAVY : heatColor(val)
      return {
        text: cell,
        options: {
          bold: isHead || isCurrent,
          color: isHead || val < 0 ? WH : val === 0 ? GY : '1E293B',
          fill: { color: bg },
          align: 'center', fontSize: 9, fontFace: 'Calibri', valign: 'middle',
          border: isCurrent
            ? [{ type: 'solid', color: 'F59E0B', pt: 2 }, { type: 'solid', color: 'F59E0B', pt: 2 }, { type: 'solid', color: 'F59E0B', pt: 2 }, { type: 'solid', color: 'F59E0B', pt: 2 }]
            : [{ type: 'solid', color: WH, pt: 0.6 }, { type: 'solid', color: WH, pt: 0.6 }, { type: 'solid', color: WH, pt: 0.6 }, { type: 'solid', color: WH, pt: 0.6 }],
        },
      }
    })), { x: 0.35, y: 1.0, w: 9.3, rowH: 0.42 })

    // 흑자 전환 최소 차량수 테이블
    const bepEvRow = RATES_S.map((_rate, ri) => {
      const idx = EVCTS_S.findIndex((_e, ei) => heatData[ei][ri] >= 0)
      return idx >= 0 ? `${EVCTS_S[idx]}대` : '불가'
    })
    s4.addText('충전요금별 흑자 전환 최소 차량수  (초기 투자비 제외 — 월 운영 기준)', {
      x: 0.35, y: 4.66, w: 9.3, h: 0.25, fontSize: 9.5, bold: true, color: NAVY, fontFace: 'Calibri' })
    const bepTable = [
      RATES_S.map(r2 => `${r2}원`),
      bepEvRow,
    ]
    s4.addTable(bepTable.map((row, ri) => row.map((cell, ci) => ({
      text: cell,
      options: {
        bold: true,
        color: ri === 0 ? WH : cell === '불가' ? RD : GR,
        fill: { color: ri === 0 ? NAVY2 : (RATES_S[ci] === curRate0 ? TEAL2 : (cell === '불가' ? 'FEE2E2' : 'EEF2FF')) },
        align: 'center', fontSize: ri === 0 ? 9 : 11, fontFace: 'Calibri', valign: 'middle',
        border: [{ type: 'solid', color: 'E2E8F0', pt: 0.5 }, { type: 'solid', color: 'E2E8F0', pt: 0.5 }, { type: 'solid', color: 'E2E8F0', pt: 0.5 }, { type: 'solid', color: 'E2E8F0', pt: 0.5 }],
      },
    }))), { x: 0.35, y: 4.92, w: 9.3, rowH: 0.34 })
    s4.addText('■ 주황 테두리 = 현재 설정값   ■ 적색 = 적자   ■ 녹색 = 흑자 (진할수록 수익 큼)', {
      x: 0.35, y: 5.48, w: 9.3, h: 0.18, fontSize: 7.5, color: GY, fontFace: 'Calibri', italic: true })

    // ════════════════════════════════
    // 슬라이드 5: 시나리오별 리스크 분석 (낙관/기본/보수 비교)
    // ════════════════════════════════
    const s5 = prs.addSlide()
    addBg(s5); addHeader(s5, '시나리오별 리스크 분석', '최선 · 기본 · 최악 시나리오별 내 수익 비교 — 투자 위험도 확인')

    const scColors: Record<string, string> = { '낙관': '6366F1', '기본': NAVY2, '보수': RD }
    const pptScResults = scDefs.map(sc => {
      const p2: SimParams = {
        ...params,
        ev_growth_rate: params.ev_growth_rate * sc.gm,
        charger_configs: params.charger_configs.map(c => ({
          ...c, daily_ev: c.daily_ev * sc.dm, rate: c.rate * sc.cm, avg_kwh: c.avg_kwh * sc.km,
        })),
      }
      return { label: sc.label, sc, result: runSimulation(p2) }
    })
    const c0 = params.charger_configs[0]

    // 시나리오 카드 3개 (상단)
    pptScResults.forEach((sc, i) => {
      const cx = 0.35 + i * 3.12
      const col = scColors[sc.label] || TEAL
      const cumV = sc.result.records[sc.result.records.length - 1].cumulative
      s5.addShape(prs.ShapeType.rect, { x: cx, y: 0.98, w: 3.0, h: 2.08,
        fill: { color: WH }, line: { color: col, width: 1.2 } })
      s5.addShape(prs.ShapeType.rect, { x: cx, y: 0.98, w: 3.0, h: 0.44, fill: { color: col } })
      s5.addText(`${sc.label} 시나리오`, { x: cx + 0.06, y: 0.99, w: 2.88, h: 0.38,
        fontSize: 13, bold: true, color: WH, fontFace: 'Calibri', align: 'center' })
      ;[
        { l: '투자 회수 (BEP)', v: sc.result.bep_month ? `${sc.result.bep_month}개월` : '미달성' },
        { l: '월 평균 순이익',  v: `${(sc.result.avg_monthly_net/10000).toLocaleString('ko-KR', {minimumFractionDigits:1,maximumFractionDigits:1})}만원` },
        { l: '최종 누적 순이익', v: `${Math.round(cumV/10000).toLocaleString('ko-KR')}만원` },
        { l: 'ROI',            v: `${sc.result.roi.toFixed(1)}%` },
      ].forEach((item, j) => {
        const iy = 1.47 + j * 0.36
        s5.addText(item.l, { x: cx + 0.14, y: iy, w: 1.3, h: 0.3, fontSize: 8, color: GY, fontFace: 'Calibri' })
        s5.addText(item.v, { x: cx + 1.4, y: iy, w: 1.48, h: 0.3, fontSize: 11, bold: true,
          color: item.l === '최종 누적 순이익' ? (cumV >= 0 ? GR : RD) : col, fontFace: 'Calibri', align: 'right' })
      })
    })

    // 수치 비교 테이블 (하단)
    s5.addShape(prs.ShapeType.rect, { x: 0.35, y: 3.18, w: 9.3, h: 0.02, fill: { color: 'E2E8F0' } })
    s5.addText('시나리오별 수치 비교', { x: 0.35, y: 3.24, w: 9.3, h: 0.28,
      fontSize: 10, bold: true, color: NAVY, fontFace: 'Calibri' })
    const scCH = ['시나리오', '일평균 차량', '충전요금', '충전량/회', '증가율', '총 순이익', '투자 회수', 'ROI']
    const scCRows = [
      scCH,
      ...pptScResults.map(sc => [
        sc.label,
        (c0 ? (c0.daily_ev * sc.sc.dm).toFixed(1) : '-') + '대',
        (c0 ? Math.round(c0.rate * sc.sc.cm) : '-') + '원/kWh',
        (c0 ? (c0.avg_kwh * sc.sc.km).toFixed(1) : '-') + 'kWh',
        (params.ev_growth_rate * sc.sc.gm).toFixed(1) + '%',
        (sc.result.records[sc.result.records.length-1].cumulative / 10000).toFixed(0) + '만',
        sc.result.bep_month ? `${sc.result.bep_month}개월` : 'N/A',
        sc.result.roi.toFixed(1) + '%',
      ]),
    ]
    s5.addTable(scCRows.map((row, ri) => row.map((cell, ci) => {
      const sc = ri > 0 ? pptScResults[ri - 1] : null
      const col = sc ? (scColors[sc.label] || TEAL) : NAVY
      const cumV2 = sc ? sc.result.records[sc.result.records.length - 1].cumulative : 0
      return {
        text: cell,
        options: {
          bold: ri === 0 || ci === 0,
          color: ri === 0 ? WH : ci === 0 ? col : ci === 5 ? (cumV2 >= 0 ? GR : RD) : SLATE,
          fill: { color: ri === 0 ? NAVY : ri % 2 === 1 ? GY2 : WH },
          align: 'center', fontSize: ri === 0 ? 8.5 : 9.5, fontFace: 'Calibri', valign: 'middle', border: BD,
        },
      }
    })), { x: 0.35, y: 3.56, w: 9.3, rowH: 0.38 })

    // ════════════════════════════════
    // 슬라이드 6: 시나리오 누적 손익 추이
    // ════════════════════════════════
    const s6 = prs.addSlide()
    addBg(s6); addHeader(s6, '시나리오별 누적 손익 추이', '낙관/기본/보수 시나리오별 내 수익 누적 비교 (만원)')

    const scChartLabels = r.records.map(rec => `${rec.month}M`)
    const scCumData = pptScResults.map(sc =>
      sc.result.records.map(rec => Math.round(rec.cumulative / 10000))
    )
    {
      s6.addChart(prs.ChartType.line, pptScResults.map((sc, si) => ({
        name: sc.label,
        labels: scChartLabels,
        values: scCumData[si],
      })), {
        x: 0.35, y: 1.0, w: 8.6, h: 4.3,
        chartColors: pptScResults.map(sc => scColors[sc.label] || TEAL),
        showLegend: true, legendPos: 'b', legendFontSize: 11,
        lineDataSymbol: 'none', lineSmooth: true,
        valAxisLabelColor: GY, valAxisLabelFontSize: 8,
        catAxisLabelColor: GY, catAxisLabelFontSize: 7,
        valGridLine: { color: 'E2E8F0', size: 0.5 },
        catGridLine: { style: 'none' },
        showTitle: false,
      })
    }

    // 최종값 레이블
    // 플롯 영역 추정: left=1.15, right=8.65, top=1.22, bottom=5.07 (범례 제외)
    const N8 = r.records.length
    const allCum8 = scCumData.flat()
    const yMin8 = Math.min(...allCum8), yMax8 = Math.max(...allCum8)
    const yRange8 = yMax8 - yMin8 || 1
    const plotL8 = 1.15, plotR8 = 8.62
    const plotT8 = 1.22, plotB8 = 4.55
    const plotW8 = plotR8 - plotL8, plotH8 = plotB8 - plotT8

    const toSlideX = (m: number) => plotL8 + ((m - 1) / (N8 - 1)) * plotW8
    const toSlideY = (v: number) => plotB8 - ((v - yMin8) / yRange8) * plotH8

    // 각 시나리오 끝값 레이블 — 겹치지 않도록 y 위치를 최종값 순으로 정렬 후 최소 간격 보장
    const bw8 = 1.35, bh8 = 0.22, minGap = 0.26
    const endItems = pptScResults.map((sc, si) => {
      const col = scColors[sc.label] || TEAL
      const vEnd = scCumData[si][N8 - 1]
      const sx = toSlideX(N8)
      const sy = toSlideY(vEnd)
      const lbl = `${sc.label}: ${vEnd >= 0 ? '+' : ''}${vEnd.toLocaleString()}만`
      return { col, vEnd, sx, sy, lbl }
    })
    // 이상적 y 위치(라인 끝점 기준), 위에서 아래로 정렬
    const sorted = [...endItems].sort((a, b) => a.sy - b.sy)
    // 겹침 방지: 위에서 아래로 최소 간격 보장
    for (let i = 1; i < sorted.length; i++) {
      if (sorted[i].sy - sorted[i - 1].sy < minGap) {
        sorted[i] = { ...sorted[i], sy: sorted[i - 1].sy + minGap }
      }
    }
    // 아래 경계 초과 시 위로 밀기 (x축 레이블 영역 충분히 피하기 위해 0.5 버퍼)
    for (let i = sorted.length - 1; i >= 0; i--) {
      const maxY = plotB8 - bh8 - 0.12
      if (sorted[i].sy > maxY) sorted[i] = { ...sorted[i], sy: maxY }
      if (i > 0 && sorted[i].sy - sorted[i - 1].sy < minGap) {
        sorted[i - 1] = { ...sorted[i - 1], sy: sorted[i].sy - minGap }
      }
    }
    endItems.forEach((item, si) => {
      const adjusted = sorted.find(s => s.lbl === item.lbl)!
      const sx = item.sx, sy = adjusted.sy
      const bx = Math.max(sx - bw8 - 0.12, 0.35)
      // 종료점 점 (원래 라인 위치)
      s6.addShape(prs.ShapeType.ellipse, { x: item.sx - 0.07, y: item.sy - 0.07, w: 0.14, h: 0.14,
        fill: { color: item.col }, line: { color: WH, width: 1.2 } })
      // 레이블
      s6.addText(item.lbl, { x: bx, y: sy, w: bw8, h: bh8,
        fontSize: 8.5, bold: true, color: item.col, fontFace: 'Calibri', align: 'right', valign: 'middle' })
    })

    // ════════════════════════════════
    // 슬라이드 7: 비용 구조 (참고)
    // ════════════════════════════════
    const s7 = prs.addSlide()
    addBg(s7); addHeader(s7, '비용 구조 (참고)', '전체 운영기간 비용 구성 및 항목별 합계')

    const firstRec2 = r.records[0]
    const pieLabels   = costPieData.map(e => e.name)
    const pieVals     = costPieData.map(e => e.value)
    const costSum     = pieVals.reduce((a, v) => a + v, 0)
    const COST_COLORS = costPieData.map(e => e.fill.replace('#', ''))

    {
      s7.addChart(prs.ChartType.doughnut, [{ name: '비용구성', labels: pieLabels, values: pieVals }], {
        x: 0.35, y: 0.95, w: 4.3, h: 3.8,
        chartColors: COST_COLORS,
        showLegend: true, legendPos: 'b', legendFontSize: 9,
        showLabel: false, showPercent: true, dataLabelColor: WH, dataLabelFontSize: 8,
        holeSize: 55, showTitle: false,
      })
    }

    const costTH = ['비용 항목', '합계 (만원)', '비율']
    const costTR = pieLabels.map((l, i) => [l, pieVals[i].toLocaleString(), `${((pieVals[i]/costSum)*100).toFixed(1)}%`])
    const costSum2Row = ['합  계', costSum.toLocaleString(), '100%']
    const costAllRows = [costTH, ...costTR, costSum2Row]
    s7.addTable(costAllRows.map((row, ri) => row.map((cell, ci) => {
      const isHead = ri === 0, isSum = ri === costAllRows.length - 1
      const dotColor = !isHead && !isSum ? COST_COLORS[ri - 1] : undefined
      return {
        text: ci === 0 && !isHead && !isSum ? `■ ${cell}` : cell,
        options: {
          bold: isHead || isSum,
          color: isHead ? WH : isSum ? NAVY : ci === 0 ? (dotColor || SLATE) : SLATE,
          fill: { color: isHead ? NAVY : isSum ? TEAL2 : ri % 2 === 1 ? GY2 : WH },
          align: ci === 0 ? 'left' : 'right',
          fontSize: isHead ? 9 : 10, fontFace: 'Calibri', valign: 'middle', border: BD,
        },
      }
    })), { x: 4.85, y: 1.0, w: 4.8, rowH: 0.36 })

    s7.addShape(prs.ShapeType.rect, { x: 4.85, y: 3.82, w: 4.8, h: 0.02, fill: { color: 'E2E8F0' } })
    s7.addText('전기요금 세부 내역  (1개월 기준)', { x: 4.85, y: 3.9, w: 4.8, h: 0.28,
      fontSize: 9, bold: true, color: NAVY2, fontFace: 'Calibri' })
    const elecRows = [
      ['항목', '월 금액 (원)'],
      ['기본료',        firstRec2.elec_basic.toLocaleString()],
      ['사용요금',      firstRec2.elec_usage.toLocaleString()],
      ['기금 · 부가세', firstRec2.elec_fund.toLocaleString()],
      ['합  계',        firstRec2.elec_cost.toLocaleString()],
    ]
    s7.addTable(elecRows.map((row, ri) => row.map((cell, ci) => ({
      text: cell,
      options: {
        bold: ri === 0 || ri === elecRows.length - 1,
        color: ri === 0 ? WH : ri === elecRows.length - 1 ? NAVY : SLATE,
        fill: { color: ri === 0 ? NAVY2 : ri === elecRows.length - 1 ? TEAL2 : ri % 2 === 1 ? GY2 : WH },
        align: ci === 0 ? 'left' : 'right',
        fontSize: 9, fontFace: 'Calibri', valign: 'middle', border: BD,
      },
    }))), { x: 4.85, y: 4.2, w: 4.8, rowH: 0.28 })

    // ════════════════════════════════
    // 슬라이드 8: 일시불·할부 비교
    // ════════════════════════════════
    const s8c = prs.addSlide()
    addBg(s8c); addHeader(s8c, '일시불 · 할부 비교', `할인율 ${params.discount_rate}% 일시불 vs 할부 (${params.operation_months}개월 분납) 주요 지표 비교`)

    const cmpLump = runSimulation({ ...params, payment_type: '일시불' })
    const cmpInst = runSimulation({ ...params, payment_type: '할부' })
    const LUMP_CLR = '6366F1'
    const INST_CLR = 'EC4899'

    // 상단 KPI 카드 2개
    ;[
      { label: '일시불', color: LUMP_CLR, res: cmpLump, desc: `할인율 ${params.discount_rate}%` },
      { label: '할부', color: INST_CLR, res: cmpInst, desc: `월 ${Math.round(cmpInst.monthly_installment).toLocaleString()}원 × ${params.operation_months}개월` },
    ].forEach((c, ci) => {
      const cx = 0.35 + ci * 4.75
      s8c.addShape(prs.ShapeType.rect, { x: cx, y: 0.98, w: 4.55, h: 0.36, fill: { color: c.color } })
      s8c.addText(`${c.label}  (${c.desc})`, { x: cx + 0.06, y: 0.99, w: 4.43, h: 0.30, fontSize: 11, bold: true, color: WH, fontFace: 'Calibri', align: 'center' })
      ;[
        { l: '실 투자액', v: c.res.total_init_cost.toLocaleString('ko-KR') + '원' },
        { l: 'BEP', v: c.res.bep_month ? `${c.res.bep_month}개월차` : 'N/A' },
        { l: 'ROI', v: `${c.res.roi.toFixed(1)}%` },
        { l: '월 평균 순이익', v: `${fmtM(c.res.avg_monthly_net)} 만원` },
        { l: '총 누적 순이익', v: `${fmtM(c.res.total_net)} 만원` },
      ].forEach((kv, ki) => {
        const ky = 1.42 + ki * 0.34
        s8c.addShape(prs.ShapeType.rect, { x: cx, y: ky, w: 4.55, h: 0.32, fill: { color: ki % 2 === 0 ? GY2 : WH }, line: { color: 'E2E8F0', width: 0.3 } })
        s8c.addText(kv.l, { x: cx + 0.1, y: ky + 0.01, w: 2.0, h: 0.30, fontSize: 9, color: GY, fontFace: 'Calibri', valign: 'middle' })
        s8c.addText(kv.v, { x: cx + 2.1, y: ky + 0.01, w: 2.35, h: 0.30, fontSize: 10, bold: true, color: c.color, fontFace: 'Calibri', align: 'right', valign: 'middle' })
      })
    })

    // 항목별 비교 테이블
    const cmpTableY = 3.18
    s8c.addShape(prs.ShapeType.rect, { x: 0.35, y: cmpTableY - 0.04, w: 9.3, h: 0.02, fill: { color: 'E2E8F0' } })
    s8c.addText('항목별 상세 비교', { x: 0.35, y: cmpTableY + 0.02, w: 9.3, h: 0.25, fontSize: 10, bold: true, color: NAVY, fontFace: 'Calibri' })
    const cmpBepDiff = (cmpInst.bep_month ?? 0) - (cmpLump.bep_month ?? 0)
    const cmpRoiDiff = cmpInst.roi - cmpLump.roi
    const cmpNetDiff = cmpInst.total_net - cmpLump.total_net
    const cmpAvgDiff = cmpInst.avg_monthly_net - cmpLump.avg_monthly_net
    const cmpHints = [
      `할인율 ${params.discount_rate}% 적용으로 일시불 납부액이 ${fmtM(cmpInst.total_init_cost - cmpLump.total_init_cost)}만원 절감`,
      `일시불 선택 시 매월 ${Math.round(cmpInst.monthly_installment).toLocaleString()}원 할부 부담 없음`,
      cmpBepDiff > 0 ? `일시불이 ${cmpBepDiff}개월 더 빨리 흑자 전환` : cmpBepDiff < 0 ? `할부가 ${Math.abs(cmpBepDiff)}개월 먼저 흑자 전환` : '두 방식 동일',
      cmpRoiDiff < 0 ? `일시불 ROI가 ${Math.abs(cmpRoiDiff).toFixed(1)}%p 높음` : `할부 ROI가 ${cmpRoiDiff.toFixed(1)}%p 높음`,
      cmpNetDiff < 0 ? `운영 기간 합산 일시불이 ${fmtM(Math.abs(cmpNetDiff))}만원 더 수익` : `할부가 ${fmtM(cmpNetDiff)}만원 더 수익`,
      cmpAvgDiff < 0 ? `매월 평균 일시불이 ${fmtM(Math.abs(cmpAvgDiff))}만원 더 수익` : `매월 평균 할부가 ${fmtM(cmpAvgDiff)}만원 더 수익`,
    ]
    const cmpRows: string[][] = [
      ['항목', '일시불', '할부', '차이 (할부-일시불)', '설명'],
      ['실 투자액 (원)', cmpLump.total_init_cost.toLocaleString(), cmpInst.total_init_cost.toLocaleString(), (cmpInst.total_init_cost - cmpLump.total_init_cost).toLocaleString(), cmpHints[0]],
      ['월 할부금 (원)', '-', Math.round(cmpInst.monthly_installment).toLocaleString(), '-', cmpHints[1]],
      ['손익분기 (BEP)', cmpLump.bep_month ? `${cmpLump.bep_month}개월` : 'N/A', cmpInst.bep_month ? `${cmpInst.bep_month}개월` : 'N/A',
        cmpLump.bep_month && cmpInst.bep_month ? `${cmpInst.bep_month - cmpLump.bep_month}개월` : '-', cmpHints[2]],
      ['ROI (%)', `${cmpLump.roi.toFixed(1)}%`, `${cmpInst.roi.toFixed(1)}%`, `${(cmpInst.roi - cmpLump.roi) >= 0 ? '+' : ''}${(cmpInst.roi - cmpLump.roi).toFixed(1)}%`, cmpHints[3]],
      ['총 누적 순이익 (만원)', `${fmtM(cmpLump.total_net)} 만`, `${fmtM(cmpInst.total_net)} 만`, `${cmpInst.total_net - cmpLump.total_net >= 0 ? '+' : ''}${fmtM(cmpInst.total_net - cmpLump.total_net)} 만`, cmpHints[4]],
      ['월 평균 순이익 (만원)', `${fmtM(cmpLump.avg_monthly_net)} 만`, `${fmtM(cmpInst.avg_monthly_net)} 만`, `${cmpInst.avg_monthly_net - cmpLump.avg_monthly_net >= 0 ? '+' : ''}${fmtM(cmpInst.avg_monthly_net - cmpLump.avg_monthly_net)} 만`, cmpHints[5]],
    ]
    s8c.addTable(cmpRows.map((row, ri) => row.map((cell, ci) => {
      const isHeader = ri === 0
      const isDiff = ci === 3 && ri > 0
      const isHint = ci === 4 && ri > 0
      const diffPos = isDiff && !cell.startsWith('-') && cell.startsWith('+')
      const diffNeg = isDiff && cell.startsWith('-') && cell !== '-'
      const hintFavLump = isHint && (cell.includes('일시불이') || cell.includes('일시불 선택') || cell.includes('일시불 납부') || cell.includes('일시불 ROI'))
      return {
        text: cell,
        options: {
          bold: isHeader || ci === 0,
          color: isHeader ? WH
            : isDiff ? (diffPos ? GR : diffNeg ? RD : GY)
            : isHint ? (hintFavLump ? '15803D' : cell.includes('동일') ? GY : 'B45309')
            : ci === 1 ? LUMP_CLR : ci === 2 ? INST_CLR : SLATE,
          fill: { color: isHeader ? NAVY : ri % 2 === 1 ? GY2 : WH },
          align: isHeader ? 'center' : ci === 0 ? 'left' : ci === 4 ? 'left' : 'right',
          fontSize: 8, fontFace: 'Calibri', valign: 'middle', border: BD,
        },
      }
    })), { x: 0.35, y: cmpTableY + 0.32, w: 9.3, rowH: 0.22 })

    // ════════════════════════════════
    // 슬라이드 9: 월별 상세 데이터 (부록)
    // ════════════════════════════════
    const s8b = prs.addSlide()
    addBg(s8b); addHeader(s8b, '월별 상세 데이터 (부록)', '월별 충전량 · 매출 · 비용 · 순이익 전체 내역')

    const hasInst = params.payment_type === '할부'
    const detailHeaders = ['월', '충전량\n(kWh)', '총매출\n(원)', 'PG\n수수료', 'PG차감\n매출', '운영사\n매출', '전기\n요금', '운영비', 'AS비', '기타', ...(hasInst ? ['할부금'] : []), '총비용', '순이익', '누적\n손익']
    const recToRow = (rec: typeof r.records[0]) => [
      String(rec.month),
      rec.kwh.toLocaleString(), rec.gross_revenue.toLocaleString(), rec.pg_fee.toLocaleString(),
      rec.net_revenue.toLocaleString(), rec.my_revenue.toLocaleString(), rec.elec_cost.toLocaleString(),
      rec.ops.toLocaleString(), rec.as_cost.toLocaleString(), rec.other.toLocaleString(),
      ...(hasInst ? [rec.installment.toLocaleString()] : []),
      rec.total_cost.toLocaleString(), rec.net_profit.toLocaleString(), rec.cumulative.toLocaleString(),
    ]
    const colCount = detailHeaders.length
    const HEAD_N = 10, TAIL_N = 5
    const total = r.records.length
    const needEllipsis = total > HEAD_N + TAIL_N
    const ellipsisRow = Array(colCount).fill('…')
    const displayRecs: (typeof r.records[0] | null)[] = [
      ...r.records.slice(0, HEAD_N),
      ...(needEllipsis ? [null] : []),
      ...(needEllipsis ? r.records.slice(total - TAIL_N) : r.records.slice(HEAD_N)),
    ]
    const detailRows = [detailHeaders, ...displayRecs.map(rec => rec ? recToRow(rec) : ellipsisRow)]
    s8b.addTable(detailRows.map((row, ri) => row.map((cell, ci) => {
      const isEllipsis = ri > 0 && row[0] === '…'
      const recIdx = ri === 0 ? -1
        : ri - 1 < HEAD_N ? ri - 1
        : needEllipsis && ri - 1 === HEAD_N ? -1
        : total - TAIL_N + (ri - 1 - HEAD_N - (needEllipsis ? 1 : 0))
      const rec = recIdx >= 0 ? r.records[recIdx] : null
      return {
        text: cell,
        options: {
          bold: ri === 0,
          color: ri === 0 ? WH : isEllipsis ? GY
            : ci === colCount - 2 ? (rec && rec.net_profit >= 0 ? GR : RD)
            : ci === colCount - 1 ? (rec && rec.cumulative >= 0 ? GR : RD)
            : SLATE,
          fill: { color: ri === 0 ? NAVY : isEllipsis ? 'F8FAFC' : ri % 2 === 1 ? GY2 : WH },
          align: ri === 0 || ci === 0 || isEllipsis ? 'center' : 'right',
          fontSize: 7, fontFace: 'Calibri', valign: 'middle', border: BD,
        },
      }
    })), { x: 0.18, y: 1.05, w: 9.65, rowH: 0.17 })
    s8b.addText(`처음 ${HEAD_N}개월 + 마지막 ${TAIL_N}개월 표시 | 전체 ${total}개월 데이터는 Excel 파일 참조`, {
      x: 0.35, y: 5.34, w: 9.3, h: 0.2, fontSize: 7.5, color: GY, fontFace: 'Calibri', italic: true })

    // 다운로드
    const arrayBuffer = await prs.write({ outputType: 'arraybuffer' }) as ArrayBuffer
    const pptxBlob = new Blob([arrayBuffer], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' })
    await triggerDownload(pptxBlob, `${rTitle || 'ego_시뮬레이션'}.pptx`)
    showToast('✅ PPTX 파일 다운로드가 시작됐습니다')
  }

  const avgMyRevenue = r.records.length > 0 ? r.records.reduce((s, rec) => s + rec.my_revenue, 0) / r.records.length : 0
  const avgTotalCost = r.records.length > 0 ? r.records.reduce((s, rec) => s + rec.total_cost, 0) / r.records.length : 0
  const kpiRow1 = [
    { label: '총 누적 순이익', value: `${fmtM(lastRec.cumulative)} 만원`, badge: `+${r.roi.toFixed(1)}% ROI`, badgeBg: C.primary, top: C.primary },
    { label: '손익분기점',     value: r.bep_month ? `${r.bep_month}개월차` : 'N/A', badge: r.bep_month ? 'BEP 달성' : '기간내 미달', badgeBg: r.bep_month ? C.green : C.red, top: r.bep_month ? C.green : C.red },
    { label: 'ROI',           value: `${r.roi.toFixed(1)}%`, sub: '초기투자 대비 수익율', top: C.accent },
    { label: '총 충전량',      value: `${(r.kwh_total / 1000).toFixed(1)} MWh`, sub: `${fmt(r.kwh_total)} kWh`, top: C.navy },
  ]
  const kpiRow2 = [
    { label: '월 운영사 매출', value: `${fmtM(avgMyRevenue)} 만원`, sub: `PG ${params.pg_fee_pct}% 차감 후 ${params.revenue_share_pct}% 배분 · ${params.operation_months}개월 평균`, top: C.primary },
    { label: '월 총비용',      value: `${fmtM(avgTotalCost)} 만원`, sub: `전기요금·운영비·AS비·기타${params.payment_type === '할부' ? '·할부금' : ''} 합산 · ${params.operation_months}개월 평균`, top: C.red },
    { label: '월 순이익',      value: `${fmtM(r.avg_monthly_net)} 만원`, sub: `운영사 매출 - 총비용 · ${params.operation_months}개월 평균`, top: avgMyRevenue - avgTotalCost >= 0 ? C.green : C.red },
  ]

  return (
    <div style={{ flex: 1, minHeight: 0, display: 'flex', flexDirection: 'column', background: '#F5F3FF' }}>

      {/* PDF 슬라이드 렌더링 (캡처용, 로딩 오버레이 뒤) */}
      {pdfCaptureSlide !== null && result && (() => {
        const r = result
        const NC = '#2D2878', AC = '#7C6FCD', LV = '#C4BFEF'
        const GR = '#16a34a', RD = '#DC2626', GY = '#64748B', GY2 = '#F1F5F9'
        const W = 960, H = 540
        const bepIdx = r.bep_month ? r.bep_month - 1 : chartData.length
        const SS: React.CSSProperties = { width: W, height: H, position: 'relative', overflow: 'hidden', fontFamily: '"Pretendard","Apple SD Gothic Neo",sans-serif', boxSizing: 'border-box', background: 'white' }
        const SB = <div style={{ position:'absolute', left:0, top:0, width:17, height:H, background:NC }}/>
        const TL = <div style={{ position:'absolute', top:0, left:17, right:0, height:5, background:AC }}/>
        const SH = (title: string, sub?: string) => (
          <div style={{ marginLeft:17, padding:'10px 24px 5px' }}>
            <div style={{ fontSize:22, fontWeight:800, color:NC }}>{title}</div>
            {sub && <div style={{ fontSize:9, color:GY, marginTop:1 }}>{sub}</div>}
            <div style={{ height:1, background:'#E2E8F0', marginTop:5 }}/>
          </div>
        )
        // 흑자전환 히트맵 데이터
        const RATES_S = [200,250,300,350,400,450,500]
        const EVCTS_S = [1,2,3,5,7,10,15]
        const heatData = EVCTS_S.map(ev => RATES_S.map(rate => {
          const p2 = { ...params, charger_configs: params.charger_configs.map((c,idx) => idx===0 ? {...c,daily_ev:ev,rate} : c) }
          return Math.round(runSimulation(p2).records[0].net_profit/1000)
        }))
        const curEv0 = params.charger_configs[0]?.daily_ev ?? 0
        const curRate0 = params.charger_configs[0]?.rate ?? 0
        const heatColor = (v: number) => {
          const h2 = (n: number) => Math.round(Math.max(0,Math.min(255,n))).toString(16).padStart(2,'0')
          const allFlat = heatData.flat()
          const mn = Math.min(...allFlat), mx = Math.max(...allFlat)
          if (v < 0) { const t = mn<0?Math.min(1,v/mn):0; return `#${h2(220-70*t)}${h2(38+10*t)}${h2(38+10*t)}` }
          const t = mx>0?Math.min(1,v/mx):0; return `#${h2(255-210*t)}${h2(255-140*t+120*t)}${h2(255-220*t)}`
        }
        const scColors: Record<string,string> = { '낙관':'#7C3AED', '기본':'#1E3A5F', '보수':'#B91C1C' }
        const totalSlides = 5 + (scResults.length > 0 ? 2 : 0) + 1 + 1
        const lastCum = r.records[r.records.length-1].cumulative
        const totRev = yearlyData.reduce((s,y)=>s+y.myRevenue,0)
        const totCost = yearlyData.reduce((s,y)=>s+y.totalCost,0)
        const totNet = yearlyData.reduce((s,y)=>s+y.netProfit,0)

        const slides: React.ReactNode[] = [
          // 0: 표지
          <div ref={pdfSlideRef} style={{...SS, background:NC}}>
            <div style={{position:'absolute',left:0,top:0,width:17,height:H,background:AC}}/>
            <div style={{position:'absolute',top:36,left:50,right:50}}>
              <div style={{fontSize:16,color:LV,marginBottom:12,fontWeight:500}}>EV충전소 수익성 시뮬레이션</div>
              <div style={{fontSize:52,fontWeight:900,color:'white',lineHeight:1.1,marginBottom:18}}>{rTitle||'EV충전소 수익 분석'}</div>
              <div style={{height:5,width:180,background:AC,marginBottom:16}}/>
              <div style={{fontSize:20,color:LV,fontWeight:700,marginBottom:10}}>{r.charger_summary}</div>
              <div style={{fontSize:16,color:'#94A3B8',marginBottom:8}}>운영 기간 {params.operation_months}개월 · 초기 투자 {r.total_init_cost.toLocaleString()}원{rCo?' · '+rCo:''}</div>
              {params.manager_name && (
                <div style={{fontSize:18,color:'#FB923C',fontWeight:700,marginBottom:6}}>
                  담당자: {params.manager_name}{params.manager_discount>0?`  ·  추가 할인 ${params.manager_discount.toLocaleString()}원 적용`:''}
                </div>
              )}
              <div style={{fontSize:14,color:'#475569',marginTop:4}}>분석일: {new Date().toLocaleDateString('ko-KR')}</div>
            </div>
            <div style={{position:'absolute',bottom:0,left:17,right:0,background:'rgba(0,0,0,0.40)',padding:'20px 30px',display:'grid',gridTemplateColumns:'repeat(4,1fr)'}}>
              {[
                {label:'손익분기점',value:r.bep_month?`${r.bep_month}개월`:'미달성',color:r.bep_month?'#34D399':'#F87171'},
                {label:'월 평균 순이익',value:`${(r.avg_monthly_net/10000).toFixed(1)}만원`,color:LV},
                {label:'총 순이익',value:`${(lastCum/10000).toFixed(0)}만원`,color:lastCum>=0?'#34D399':'#F87171'},
                {label:'ROI',value:`${r.roi.toFixed(1)}%`,color:LV},
              ].map((k,i)=>(
                <div key={i} style={{textAlign:'center',borderLeft:i>0?'1px solid rgba(255,255,255,0.15)':'none',paddingLeft:i>0?16:0}}>
                  <div style={{fontSize:14,color:LV,marginBottom:6}}>{k.label}</div>
                  <div style={{fontSize:30,fontWeight:900,color:k.color}}>{k.value}</div>
                </div>
              ))}
            </div>
          </div>,

          // 1: 투자 회수 분석
          <div ref={pdfSlideRef} style={SS}>
            {SB}{TL}
            {SH('투자 회수 분석',`손익분기점 (BEP) · 매달 수익 — 운영 ${params.operation_months}개월 시뮬레이션`)}
            {(() => {
              const pdfAvgMyRev = r.records.length > 0 ? r.records.reduce((s,rec)=>s+rec.my_revenue,0)/r.records.length : 0
              const pdfAvgCost  = r.records.length > 0 ? r.records.reduce((s,rec)=>s+rec.total_cost,0)/r.records.length : 0
              return (
                <div style={{marginLeft:17,padding:'5px 20px',display:'grid',gridTemplateColumns:'repeat(3,1fr)',gap:7}}>
                  {[
                    {label:'월 운영사 매출',value:`${(pdfAvgMyRev/10000).toFixed(1)}만원`,sub:`PG ${params.pg_fee_pct}% 차감 후 ${params.revenue_share_pct}% 배분 · ${params.operation_months}개월 평균`,color:AC},
                    {label:'월 총비용',value:`${(pdfAvgCost/10000).toFixed(1)}만원`,sub:`전기요금·운영비·AS비·기타${params.payment_type==='할부'?'·할부금':''} 합산 · ${params.operation_months}개월 평균`,color:RD},
                    {label:'월 순이익',value:`${(r.avg_monthly_net/10000).toFixed(1)}만원`,sub:`운영사 매출 - 총비용 · ${params.operation_months}개월 평균`,color:r.avg_monthly_net>=0?GR:RD},
                  ].map((k,i)=>(
                    <div key={i} style={{background:GY2,borderRadius:6,padding:'6px 9px',borderTop:`3px solid ${k.color}`}}>
                      <div style={{fontSize:7.5,color:GY}}>{k.label}</div>
                      <div style={{fontSize:16,fontWeight:800,color:k.color,marginTop:2}}>{k.value}</div>
                      <div style={{fontSize:7,color:GY,marginTop:2}}>{k.sub}</div>
                    </div>
                  ))}
                </div>
              )
            })()}
            <div style={{marginLeft:17,padding:'0 16px',height:265}}>
              <ResponsiveContainer width="100%" height={265}>
                <ComposedChart data={chartData} margin={{top:4,right:52,left:0,bottom:4}}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0"/>
                  <XAxis dataKey="month" tick={{fontSize:6}} interval={Math.floor(chartData.length/10)}/>
                  <YAxis yAxisId="bar" tick={{fontSize:6}} tickFormatter={v=>`${v}만`}/>
                  <YAxis yAxisId="line" orientation="right" tick={{fontSize:6}} tickFormatter={v=>`${v}만`}/>
                  <Legend wrapperStyle={{fontSize:9}} formatter={(value)=>value==='net'?'월 순이익 (막대, 좌축)':value==='cum'?'누적 손익 (선, 우축)':value}/>
                  <Bar yAxisId="bar" dataKey="net" name="net" radius={[2,2,0,0]} isAnimationActive={false}>
                    {chartData.map((_,i)=><Cell key={i} fill={i<bepIdx?'#F43F5E':'#6366F1'} fillOpacity={0.88}/>)}
                  </Bar>
                  <Line yAxisId="line" type="monotone" dataKey="cum" name="cum" stroke="#FBBF24" strokeWidth={2.5} dot={false} isAnimationActive={false}/>
                  {r.bep_month && <ReferenceLine yAxisId="bar" x={r.bep_month} stroke="#FB923C" strokeWidth={2} label={{value:`BEP ${r.bep_month}M`,position:'insideTopLeft',fontSize:8,fill:'#FB923C'}}/>}
                </ComposedChart>
              </ResponsiveContainer>
            </div>
            <div style={{marginLeft:34,padding:'0 16px',display:'flex',gap:16,fontSize:8,color:GY}}>
              <span><span style={{display:'inline-block',width:10,height:10,background:'#F43F5E',marginRight:4,verticalAlign:'middle'}}/>BEP 이전 (적자 구간)</span>
              <span><span style={{display:'inline-block',width:10,height:10,background:'#6366F1',marginRight:4,verticalAlign:'middle'}}/>BEP 이후 (흑자 구간)</span>
              <span><span style={{display:'inline-block',width:20,height:3,background:'#FBBF24',marginRight:4,verticalAlign:'middle'}}/>누적 손익 (우축)</span>
              {r.bep_month && <span><span style={{display:'inline-block',width:3,height:10,background:'#FB923C',marginRight:4,verticalAlign:'middle'}}/>손익분기점 (BEP)</span>}
            </div>
          </div>,

          // 2: 연도별 수익 흐름
          <div ref={pdfSlideRef} style={SS}>
            {SB}{TL}
            {SH('연도별 수익 흐름','연도별 내 순이익 및 누적 투자 회수 현황')}
            <div style={{marginLeft:17,padding:'6px 24px',display:'grid',gridTemplateColumns:'repeat(3,1fr)',gap:10,marginBottom:10}}>
              {[
                {label:'총 운영사 매출',value:`${Math.round(totRev).toLocaleString()}원`,color:AC},
                {label:'총 비용',value:`${Math.round(totCost).toLocaleString()}원`,color:RD},
                {label:'총 순이익',value:`${Math.round(totNet).toLocaleString()}원`,color:totNet>=0?GR:RD},
              ].map((s,i)=>(
                <div key={i} style={{background:GY2,borderRadius:6,padding:'8px 12px',borderTop:`3px solid ${s.color}`}}>
                  <div style={{fontSize:9,color:GY}}>{s.label}</div>
                  <div style={{fontSize:18,fontWeight:800,color:s.color,marginTop:3}}>{s.value}</div>
                </div>
              ))}
            </div>
            <div style={{marginLeft:17,padding:'0 24px'}}>
              <table style={{width:'100%',borderCollapse:'collapse',fontSize:11}}>
                <thead><tr style={{background:NC}}>
                  {['연도','운영사 매출 (원)','총비용 (원)','순이익 (원)','누적 손익 (원)'].map(h=>(
                    <th key={h} style={{padding:'6px 8px',color:'white',fontWeight:700,textAlign:h==='연도'?'left':'right',fontSize:9}}>{h}</th>
                  ))}
                </tr></thead>
                <tbody>
                  {(() => { let cum=0; return yearlyData.map((y,i)=>{cum+=y.netProfit; return (
                    <tr key={i} style={{background:i%2===0?GY2:'white'}}>
                      <td style={{padding:'5px 8px',fontWeight:600,color:AC,fontSize:10}}>{y.yr}</td>
                      <td style={{padding:'5px 8px',textAlign:'right',fontSize:10}}>{Math.round(y.myRevenue).toLocaleString()}</td>
                      <td style={{padding:'5px 8px',textAlign:'right',fontSize:10}}>{Math.round(y.totalCost).toLocaleString()}</td>
                      <td style={{padding:'5px 8px',textAlign:'right',fontWeight:700,fontSize:10,color:y.netProfit>=0?GR:RD}}>{Math.round(y.netProfit).toLocaleString()}</td>
                      <td style={{padding:'5px 8px',textAlign:'right',fontWeight:700,fontSize:10,color:cum>=0?GR:RD}}>{Math.round(cum).toLocaleString()}</td>
                    </tr>
                  )})})()}
                  <tr style={{background:LV}}>
                    <td style={{padding:'5px 8px',fontWeight:700,fontSize:10,color:NC}}>합계</td>
                    <td style={{padding:'5px 8px',textAlign:'right',fontWeight:700,fontSize:10}}>{Math.round(totRev).toLocaleString()}</td>
                    <td style={{padding:'5px 8px',textAlign:'right',fontWeight:700,fontSize:10}}>{Math.round(totCost).toLocaleString()}</td>
                    <td style={{padding:'5px 8px',textAlign:'right',fontWeight:700,fontSize:10,color:totNet>=0?GR:RD}}>{Math.round(totNet).toLocaleString()}</td>
                    <td style={{padding:'5px 8px',textAlign:'right',fontWeight:700,fontSize:10,color:totNet>=0?GR:RD}}>{Math.round(totNet).toLocaleString()}</td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>,

          // 3: 흑자 전환 조건 히트맵
          <div ref={pdfSlideRef} style={SS}>
            {SB}{TL}
            {SH('흑자 전환 조건','충전요금 × 일 이용 차량수에 따른 월 순이익 (천원) — 오렌지 테두리 = 현재 설정')}
            <div style={{marginLeft:17,padding:'4px 16px'}}>
              <table style={{width:'100%',borderCollapse:'collapse',fontSize:9.5}}>
                <thead><tr>
                  <th style={{padding:'3px 5px',background:NC,color:'white',fontSize:8}}>차량↓/요금→</th>
                  {RATES_S.map(rt=><th key={rt} style={{padding:'3px 5px',background:NC,color:'white',fontSize:8,textAlign:'center'}}>{rt}원</th>)}
                </tr></thead>
                <tbody>
                  {EVCTS_S.map((ev,ei)=>(
                    <tr key={ei}>
                      <td style={{padding:'3px 5px',background:NC,color:'white',fontWeight:700,fontSize:8,textAlign:'center'}}>{ev}대</td>
                      {RATES_S.map((rate,ri)=>{
                        const v=heatData[ei][ri]
                        const isCur=ev===curEv0&&rate===curRate0
                        return <td key={ri} style={{padding:'1px 3px',background:heatColor(v),color:v<0?'white':v===0?GY:'#1E293B',fontWeight:isCur?900:400,fontSize:8.5,textAlign:'center',border:isCur?'3px solid #F59E0B':'1px solid rgba(0,0,0,0.05)'}}>{v.toLocaleString()}</td>
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
              <div style={{marginTop:5,display:'flex',gap:16,fontSize:7.5,color:GY,marginBottom:10}}>
                <span>■ 적색=적자 ■ 녹색=흑자(진할수록 수익 큼)</span>
                <span style={{color:'#F59E0B',fontWeight:700}}>오렌지 테두리 = 현재 설정값</span>
              </div>
              {/* 충전요금별 흑자 전환 최소 차량수 */}
              <div style={{fontSize:9,fontWeight:700,color:NC,marginBottom:5}}>충전요금별 흑자 전환 최소 차량수  (초기 투자비 제외 — 월 운영 기준)</div>
              <table style={{width:'100%',borderCollapse:'collapse',fontSize:9}}>
                <thead><tr>
                  {RATES_S.map(rt=><th key={rt} style={{padding:'4px 6px',background:NC,color:'white',fontSize:8,textAlign:'center'}}>{rt}원</th>)}
                </tr></thead>
                <tbody><tr>
                  {RATES_S.map((_rate,ri)=>{
                    const idx=EVCTS_S.findIndex((_e,ei)=>heatData[ei][ri]>=0)
                    const val=idx>=0?`${EVCTS_S[idx]}대`:'불가'
                    return <td key={ri} style={{padding:'5px 6px',background:idx>=0?'#dcfce7':'#fee2e2',color:idx>=0?'#166534':'#991b1b',fontWeight:700,fontSize:9,textAlign:'center'}}>{val}</td>
                  })}
                </tr></tbody>
              </table>
            </div>
          </div>,

        ]

        const costSlide = (
          <div ref={pdfSlideRef} style={SS}>
            {SB}{TL}
            {SH('비용 구조 (참고)','전체 운영기간 비용 구성 및 항목별 합계')}
            <div style={{marginLeft:17,display:'grid',gridTemplateColumns:'1fr 1fr',height:440}}>
              <div style={{display:'flex',alignItems:'center',justifyContent:'center'}}>
                <RPieChart width={440} height={400}>
                  <Pie data={costPieData} dataKey="value" nameKey="name" cx="50%" cy="48%"
                    outerRadius={145} innerRadius={68} isAnimationActive={false}
                    label={({name,percent}:{name?:string;percent?:number})=>`${name} ${((percent||0)*100).toFixed(0)}%`}
                    labelLine={false} fontSize={10}>
                    {costPieData.map((e,i)=><Cell key={i} fill={e.fill}/>)}
                  </Pie>
                </RPieChart>
              </div>
              <div style={{padding:'10px 20px 10px 10px'}}>
                <table style={{width:'100%',borderCollapse:'collapse',fontSize:11,marginBottom:12}}>
                  <thead><tr style={{background:NC}}>
                    {['비용 항목','합계 (만원)','비율'].map(h=>(
                      <th key={h} style={{padding:'6px 8px',color:'white',fontWeight:700,textAlign:h==='비용 항목'?'left':'right',fontSize:9}}>{h}</th>
                    ))}
                  </tr></thead>
                  <tbody>
                    {costPieData.map((e,i)=>{
                      const tot=costPieData.reduce((s,x)=>s+x.value,0)||1
                      return (
                        <tr key={i} style={{background:i%2===0?GY2:'white'}}>
                          <td style={{padding:'5px 8px',fontSize:10,color:e.fill,fontWeight:600}}>■ {e.name}</td>
                          <td style={{padding:'5px 8px',textAlign:'right',fontSize:10}}>{e.value.toLocaleString()}</td>
                          <td style={{padding:'5px 8px',textAlign:'right',fontSize:10}}>{((e.value/tot)*100).toFixed(1)}%</td>
                        </tr>
                      )
                    })}
                    <tr style={{background:LV}}>
                      <td style={{padding:'5px 8px',fontWeight:700,fontSize:10,color:NC}}>합  계</td>
                      <td style={{padding:'5px 8px',textAlign:'right',fontWeight:700,fontSize:10}}>{costPieData.reduce((s,e)=>s+e.value,0).toLocaleString()}</td>
                      <td style={{padding:'5px 8px',textAlign:'right',fontWeight:700,fontSize:10}}>100%</td>
                    </tr>
                  </tbody>
                </table>
                <div style={{fontSize:10,fontWeight:700,color:NC,marginBottom:6}}>전기요금 세부 내역 (1개월 기준)</div>
                <table style={{width:'100%',borderCollapse:'collapse',fontSize:10}}>
                  {[['기본료',r.records[0].elec_basic],['사용요금',r.records[0].elec_usage],['기금·부가세',r.records[0].elec_fund],['합  계',r.records[0].elec_cost]].map(([label,val],i)=>(
                    <tr key={i} style={{background:i===3?LV:i%2===0?GY2:'white'}}>
                      <td style={{padding:'4px 8px',fontSize:9,fontWeight:i===3?700:400}}>{label}</td>
                      <td style={{padding:'4px 8px',textAlign:'right',fontSize:9,fontWeight:i===3?700:400}}>{Number(val).toLocaleString()}원</td>
                    </tr>
                  ))}
                </table>
              </div>
            </div>
          </div>
        )

        if (scResults.length > 0) {
          // 5: 시나리오 비교 테이블
          const c0 = params.charger_configs[0]
          slides.push(
            <div ref={pdfSlideRef} style={SS}>
              {SB}{TL}
              {SH('시나리오별 리스크 분석','낙관/기본/보수 시나리오별 주요 지표 비교')}
              {/* 시나리오 조건 카드 */}
              <div style={{marginLeft:17,padding:'5px 20px 8px',display:'grid',gridTemplateColumns:'repeat(3,1fr)',gap:10}}>
                {scResults.map((sc,i)=>{
                  const def=scDefs.find(s=>s.label===sc.label)
                  const scColor=scColors[sc.label]||AC
                  const lastC=sc.result.records[sc.result.records.length-1].cumulative
                  return (
                    <div key={i} style={{border:`2px solid ${scColor}`,borderRadius:8,padding:'8px 12px',background:GY2}}>
                      <div style={{fontSize:12,fontWeight:800,color:scColor,marginBottom:5}}>{sc.label} 시나리오</div>
                      <div style={{fontSize:8,color:GY,marginBottom:2}}>일평균 차량: <b style={{color:'#1e293b'}}>{c0?((c0.daily_ev*(def?.dm||1)).toFixed(1)):'-'}대</b></div>
                      <div style={{fontSize:8,color:GY,marginBottom:2}}>충전요금: <b style={{color:'#1e293b'}}>{c0?Math.round(c0.rate*(def?.cm||1)):'-'}원</b></div>
                      <div style={{fontSize:8,color:GY,marginBottom:6}}>EV 증가율: <b style={{color:'#1e293b'}}>{(params.ev_growth_rate*(def?.gm||1)).toFixed(1)}%</b></div>
                      <div style={{fontSize:14,fontWeight:900,color:lastC>=0?GR:RD}}>{(lastC/10000).toFixed(0)}만원</div>
                      <div style={{fontSize:7.5,color:GY}}>총 순이익 · BEP {sc.result.bep_month?`${sc.result.bep_month}개월`:'N/A'} · ROI {sc.result.roi.toFixed(1)}%</div>
                    </div>
                  )
                })}
              </div>
              <div style={{marginLeft:17,padding:'0 20px'}}>
                <table style={{width:'100%',borderCollapse:'collapse',fontSize:11}}>
                  <thead><tr style={{background:NC}}>
                    {['시나리오','일평균 차량','충전요금','증가율','총 순이익','투자 회수','ROI'].map(h=>(
                      <th key={h} style={{padding:'6px 8px',color:'white',fontWeight:700,textAlign:'center',fontSize:9}}>{h}</th>
                    ))}
                  </tr></thead>
                  <tbody>
                    {scResults.map((sc,i)=>{
                      const lastC=sc.result.records[sc.result.records.length-1].cumulative
                      return (
                        <tr key={i} style={{background:i%2===0?GY2:'white'}}>
                          <td style={{padding:'7px 8px',fontWeight:700,color:scColors[sc.label]||AC,textAlign:'center'}}>{sc.label}</td>
                          <td style={{padding:'7px 8px',textAlign:'center'}}>{c0?(c0.daily_ev*(scDefs.find(s=>s.label===sc.label)?.dm||1)).toFixed(1):'-'}대</td>
                          <td style={{padding:'7px 8px',textAlign:'center'}}>{c0?Math.round(c0.rate*(scDefs.find(s=>s.label===sc.label)?.cm||1)):'-'}원</td>
                          <td style={{padding:'7px 8px',textAlign:'center'}}>{(params.ev_growth_rate*(scDefs.find(s=>s.label===sc.label)?.gm||1)).toFixed(1)}%</td>
                          <td style={{padding:'7px 8px',textAlign:'center',fontWeight:700,color:lastC>=0?GR:RD}}>{(lastC/10000).toFixed(0)}만</td>
                          <td style={{padding:'7px 8px',textAlign:'center'}}>{sc.result.bep_month?`${sc.result.bep_month}개월`:'N/A'}</td>
                          <td style={{padding:'7px 8px',textAlign:'center'}}>{sc.result.roi.toFixed(1)}%</td>
                        </tr>
                      )
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          )
          // 6: 시나리오 누적 손익 차트
          slides.push(
            <div ref={pdfSlideRef} style={SS}>
              {SB}{TL}
              {SH('시나리오별 누적 손익 추이','낙관/기본/보수 시나리오별 내 수익 누적 비교 (만원)')}
              {/* 최종 수치 요약 */}
              <div style={{marginLeft:17,padding:'4px 20px 0',display:'flex',gap:16}}>
                {scResults.map((sc,i)=>{
                  const lastC=Math.round(sc.result.records[sc.result.records.length-1].cumulative/10000)
                  return (
                    <div key={i} style={{display:'flex',alignItems:'center',gap:6}}>
                      <span style={{display:'inline-block',width:24,height:3,background:scColors[sc.label]||AC}}/>
                      <span style={{fontSize:10,fontWeight:700,color:scColors[sc.label]||AC}}>{sc.label}</span>
                      <span style={{fontSize:10,fontWeight:900,color:lastC>=0?GR:RD}}>{lastC>=0?'+':''}{lastC.toLocaleString()}만원</span>
                      {sc.result.bep_month && <span style={{fontSize:8,color:GY}}>BEP {sc.result.bep_month}개월</span>}
                    </div>
                  )
                })}
              </div>
              <div style={{marginLeft:17,padding:'0 16px',height:365}}>
                <ResponsiveContainer width="100%" height={365}>
                  <ComposedChart margin={{top:10,right:100,left:10,bottom:10}}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0"/>
                    <XAxis dataKey="month" type="category" allowDuplicatedCategory={false} tick={{fontSize:7}} interval={Math.floor(params.operation_months/8)}/>
                    <YAxis tick={{fontSize:7}} tickFormatter={v=>`${v}만`} width={50}/>
                    <ReferenceLine y={0} stroke="#94a3b8" strokeWidth={1.5} strokeDasharray="4 4"/>
                    {scResults.map(sc=>{
                      const recs=sc.result.records.map(rec=>({month:`${rec.month}M`,[sc.label]:Math.round(rec.cumulative/10000)}))
                      return (
                        <Line key={sc.label} type="monotone" isAnimationActive={false}
                          data={recs} dataKey={sc.label}
                          stroke={scColors[sc.label]||AC} strokeWidth={2.5}
                          dot={false}/>
                      )
                    })}
                  </ComposedChart>
                </ResponsiveContainer>
              </div>
            </div>
          )
        }

        // 비용 구조는 시나리오 뒤 (PPTX 슬라이드 7 순서 맞춤)
        slides.push(costSlide)

        // 슬라이드 8: 일시불·할부 비교
        {
          const cmpL = runSimulation({ ...params, payment_type: '일시불' })
          const cmpI = runSimulation({ ...params, payment_type: '할부' })
          const LCLR = '#6366F1', ICLR = '#EC4899'
          const cmpRows = [
            { label: '실 투자액', lv: cmpL.total_init_cost, iv: cmpI.total_init_cost, fmt: (v: number) => `${Math.round(v/10000)}만원` },
            { label: '월 할부금', lv: 0, iv: cmpI.monthly_installment, fmt: (v: number) => v > 0 ? `${Math.round(v).toLocaleString()}원` : '-' },
            { label: 'BEP', lv: cmpL.bep_month ?? 0, iv: cmpI.bep_month ?? 0, fmt: (v: number) => v ? `${v}개월` : 'N/A' },
            { label: 'ROI', lv: cmpL.roi, iv: cmpI.roi, fmt: (v: number) => `${v.toFixed(1)}%` },
            { label: '총 순이익', lv: cmpL.total_net, iv: cmpI.total_net, fmt: (v: number) => `${Math.round(v/10000)}만원` },
            { label: '월 평균 순이익', lv: cmpL.avg_monthly_net, iv: cmpI.avg_monthly_net, fmt: (v: number) => `${Math.round(v/10000)}만원` },
          ]
          slides.push(
            <div ref={pdfSlideRef} style={SS}>
              {SB}{TL}
              {SH('일시불 · 할부 비교', `할인율 ${params.discount_rate}% 일시불 vs 할부 ${params.operation_months}개월 분납`)}
              <div style={{marginLeft:17, padding:'4px 10px', display:'flex', flexDirection:'column', gap:8}}>
                {/* 상단 KPI 카드 */}
                <div style={{display:'grid', gridTemplateColumns:'1fr 1fr', gap:8}}>
                  {[
                    { label:'일시불', color:LCLR, res:cmpL, desc:`할인율 ${params.discount_rate}%` },
                    { label:'할부',   color:ICLR,  res:cmpI, desc:`월 ${Math.round(cmpI.monthly_installment).toLocaleString()}원 × ${params.operation_months}개월` },
                  ].map(c => (
                    <div key={c.label} style={{borderRadius:8, border:`2px solid ${c.color}`, padding:'8px 10px', background:'white'}}>
                      <div style={{display:'flex', alignItems:'center', gap:6, marginBottom:6}}>
                        <div style={{width:3, height:16, borderRadius:2, background:c.color}}/>
                        <span style={{fontSize:12, fontWeight:800, color:c.color}}>{c.label}</span>
                        <span style={{fontSize:8, color:'#9ca3af'}}>{c.desc}</span>
                      </div>
                      {[
                        {l:'실 투자액', v:`${Math.round(c.res.total_init_cost/10000)}만원`},
                        {l:'BEP',       v:c.res.bep_month ? `${c.res.bep_month}개월` : 'N/A'},
                        {l:'ROI',       v:`${c.res.roi.toFixed(1)}%`},
                        {l:'총 순이익', v:`${Math.round(c.res.total_net/10000)}만원`},
                      ].map(k => (
                        <div key={k.l} style={{display:'flex', justifyContent:'space-between', fontSize:9, marginBottom:3}}>
                          <span style={{color:'#9ca3af'}}>{k.l}</span>
                          <span style={{fontWeight:700, color:c.color}}>{k.v}</span>
                        </div>
                      ))}
                    </div>
                  ))}
                </div>
                {/* 비교 테이블 */}
                {(() => {
                  const pdfBepDiff = (cmpI.bep_month ?? 0) - (cmpL.bep_month ?? 0)
                  const pdfRoiDiff = cmpI.roi - cmpL.roi
                  const pdfNetDiff = cmpI.total_net - cmpL.total_net
                  const pdfAvgDiff = cmpI.avg_monthly_net - cmpL.avg_monthly_net
                  const pdfHints = [
                    `할인율 ${params.discount_rate}% 적용, 일시불이 ${Math.round((cmpI.total_init_cost - cmpL.total_init_cost)/10000)}만원 절감`,
                    `일시불 선택 시 매월 ${Math.round(cmpI.monthly_installment).toLocaleString()}원 할부 부담 없음`,
                    pdfBepDiff > 0 ? `일시불이 ${pdfBepDiff}개월 더 빨리 흑자 전환` : pdfBepDiff < 0 ? `할부가 ${Math.abs(pdfBepDiff)}개월 먼저 흑자 전환` : '두 방식 동일',
                    pdfRoiDiff < 0 ? `일시불 ROI가 ${Math.abs(pdfRoiDiff).toFixed(1)}%p 높음` : `할부 ROI가 ${pdfRoiDiff.toFixed(1)}%p 높음`,
                    pdfNetDiff < 0 ? `일시불이 ${Math.round(Math.abs(pdfNetDiff)/10000)}만원 더 수익` : `할부가 ${Math.round(pdfNetDiff/10000)}만원 더 수익`,
                    pdfAvgDiff < 0 ? `매월 일시불이 ${Math.round(Math.abs(pdfAvgDiff)/10000)}만원 더 수익` : `매월 할부가 ${Math.round(pdfAvgDiff/10000)}만원 더 수익`,
                  ]
                  return (
                    <table style={{width:'100%', borderCollapse:'collapse', fontSize:9}}>
                      <thead>
                        <tr style={{background:NC}}>
                          {['항목','일시불','할부','차이(할부-일시불)','설명'].map((h,hi) => (
                            <th key={h} style={{padding:'4px 6px', color:'white', fontWeight:700, textAlign: hi===0||hi===4 ? 'left' : 'right', fontSize:9}}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {cmpRows.map((row, i) => {
                          const diff = row.iv - row.lv
                          const diffStr = row.label === '월 할부금' ? '' :
                            row.label === 'ROI' ? `${diff>=0?'+':''}${diff.toFixed(1)}%` :
                            `${diff>=0?'+':''}${Math.round(diff/10000)}만`
                          const hint = pdfHints[i]
                          const hintFavLump = hint.includes('일시불이') || hint.includes('일시불 선택') || hint.includes('일시불 ROI')
                          return (
                            <tr key={row.label} style={{background: i%2===0 ? GY2 : 'white'}}>
                              <td style={{padding:'4px 6px', fontWeight:600, fontSize:9, color:'#374151'}}>{row.label}</td>
                              <td style={{padding:'4px 6px', textAlign:'right', fontWeight:700, color:LCLR, fontSize:9}}>{row.fmt(row.lv)}</td>
                              <td style={{padding:'4px 6px', textAlign:'right', fontWeight:700, color:ICLR, fontSize:9}}>{row.fmt(row.iv)}</td>
                              <td style={{padding:'4px 6px', textAlign:'right', fontWeight:700, fontSize:9, color: diffStr.startsWith('+') ? GR : diffStr.startsWith('-') ? RD : '#9ca3af'}}>{diffStr || '-'}</td>
                              <td style={{padding:'4px 6px', fontSize:8, color: hintFavLump ? '#15803d' : hint.includes('동일') ? '#9ca3af' : '#b45309'}}>
                                {hintFavLump ? '🟢 ' : hint.includes('동일') ? '⚪ ' : '🟠 '}{hint}
                              </td>
                            </tr>
                          )
                        })}
                      </tbody>
                    </table>
                  )
                })()}
              </div>
            </div>
          )
        }

        // 슬라이드 9: 월별 상세 데이터 (부록)
        const aTotal = r.records.length
        const aHEAD = 10, aTAIL = 5
        const aNeedEllipsis = aTotal > aHEAD + aTAIL
        const aRows = aNeedEllipsis
          ? [...r.records.slice(0, aHEAD), null, ...r.records.slice(aTotal - aTAIL)]
          : r.records
        const aFs = 7
        const aRh = '3px 3px'
        const aHdrs = ['월','충전량(kWh)','총매출','PG차감매출','운영사매출','전기요금','운영비','AS비','기타','총비용','순이익','누적손익']
        slides.push(
          <div ref={pdfSlideRef} style={SS}>
            {SB}{TL}
            {SH('월별 상세 데이터 (부록)',`전체 ${aTotal}개월 상세 내역 (단위: 만원)`)}
            <div style={{marginLeft:17,padding:'2px 10px'}}>
              <table style={{width:'100%',borderCollapse:'collapse',tableLayout:'fixed'}}>
                <thead><tr style={{background:NC}}>
                  {aHdrs.map((h,i)=>(
                    <th key={i} style={{padding:'2px 3px',color:'white',fontWeight:700,textAlign:i===0?'left':'right',fontSize:aFs,lineHeight:1.2,verticalAlign:'bottom'}}>{h}</th>
                  ))}
                </tr></thead>
                <tbody>
                  {aRows.map((rec,i)=> rec === null ? (
                    <tr key="ellipsis" style={{background:'#f0edff'}}>
                      {aHdrs.map((_,ci)=>(
                        <td key={ci} style={{padding:'2px 3px',textAlign:'center',fontSize:aFs,color:'#9ca3af',fontWeight:700}}>…</td>
                      ))}
                    </tr>
                  ) : (
                    <tr key={i} style={{background:i%2===0?GY2:'white'}}>
                      <td style={{padding:aRh,textAlign:'left',fontSize:aFs,fontWeight:600,color:AC}}>{rec.month}M</td>
                      <td style={{padding:aRh,textAlign:'right',fontSize:aFs}}>{Math.round(rec.kwh).toLocaleString()}</td>
                      <td style={{padding:aRh,textAlign:'right',fontSize:aFs}}>{Math.round(rec.gross_revenue/10000).toLocaleString()}만</td>
                      <td style={{padding:aRh,textAlign:'right',fontSize:aFs}}>{Math.round(rec.net_revenue/10000).toLocaleString()}만</td>
                      <td style={{padding:aRh,textAlign:'right',fontSize:aFs}}>{Math.round(rec.my_revenue/10000).toLocaleString()}만</td>
                      <td style={{padding:aRh,textAlign:'right',fontSize:aFs}}>{Math.round(rec.elec_cost/10000).toLocaleString()}만</td>
                      <td style={{padding:aRh,textAlign:'right',fontSize:aFs}}>{Math.round(rec.ops/10000).toLocaleString()}만</td>
                      <td style={{padding:aRh,textAlign:'right',fontSize:aFs}}>{Math.round(rec.as_cost/10000).toLocaleString()}만</td>
                      <td style={{padding:aRh,textAlign:'right',fontSize:aFs}}>{Math.round(rec.other/10000).toLocaleString()}만</td>
                      <td style={{padding:aRh,textAlign:'right',fontSize:aFs}}>{Math.round(rec.total_cost/10000).toLocaleString()}만</td>
                      <td style={{padding:aRh,textAlign:'right',fontSize:aFs,fontWeight:700,color:rec.net_profit>=0?GR:RD}}>{Math.round(rec.net_profit/10000).toLocaleString()}만</td>
                      <td style={{padding:aRh,textAlign:'right',fontSize:aFs,fontWeight:700,color:rec.cumulative>=0?GR:RD}}>{Math.round(rec.cumulative/10000).toLocaleString()}만</td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {aNeedEllipsis && <div style={{fontSize:6,color:'#9ca3af',marginTop:3,textAlign:'right'}}>처음 {aHEAD}개월 + 마지막 {aTAIL}개월 표시 | 전체 {aTotal}개월 데이터는 Excel 파일 참조</div>}
            </div>
          </div>
        )

        return (
          <>
            <div style={{position:'fixed',inset:0,zIndex:9998,background:'rgba(30,27,75,0.96)',display:'flex',flexDirection:'column',alignItems:'center',justifyContent:'center',gap:16}}>
              <div style={{fontSize:44}}>📄</div>
              <div style={{color:'white',fontSize:18,fontWeight:700}}>PDF 슬라이드 생성 중...</div>
              <div style={{color:'#a78bfa',fontSize:13}}>{pdfCaptureSlide!+1} / {totalSlides} 슬라이드</div>
            </div>
            <div style={{position:'fixed',top:0,left:0,zIndex:1,pointerEvents:'none'}}>
              {slides[pdfCaptureSlide!]}
            </div>
          </>
        )
      })()}

      {/* Excel 미리보기 모달 */}
      {xlsxPreview && result && (
        <div style={{ position: 'fixed', inset: 0, zIndex: 10000, background: 'rgba(0,0,0,0.85)', display: 'flex', flexDirection: 'column' }}>
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '14px 18px', background: '#1E1B4B', flexShrink: 0 }}>
            <span style={{ color: 'white', fontWeight: 700, fontSize: 15 }}>📊 Excel 미리보기</span>
            <div style={{ display: 'flex', gap: 8 }}>
              <button onClick={async () => { await exportExcel(); setXlsxPreview(false) }} style={{
                background: '#16a34a', color: 'white', border: 'none', borderRadius: 8,
                padding: '8px 18px', fontWeight: 700, fontSize: 13, cursor: 'pointer',
              }}>⬇ Excel 다운로드</button>
              <button onClick={() => setXlsxPreview(false)} style={{
                background: 'rgba(255,255,255,0.15)', color: 'white', border: 'none',
                borderRadius: 8, padding: '8px 14px', fontWeight: 600, fontSize: 13, cursor: 'pointer',
              }}>✕ 닫기</button>
            </div>
          </div>
          <div style={{ flex: 1, overflowY: 'auto', padding: '16px', background: '#f8fafc' }}>
            {/* KPI 요약 */}
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, 1fr)', gap: 10, marginBottom: 16 }}>
              {[
                { label: '손익분기점', value: result.bep_month ? `${result.bep_month}개월` : '미달성', color: result.bep_month ? '#16a34a' : '#DC2626' },
                { label: '월 평균 순이익', value: `${(result.avg_monthly_net/10000).toFixed(1)}만원`, color: C.primary },
                { label: '총 순이익', value: `${(result.records[result.records.length-1].cumulative/10000).toFixed(0)}만원`, color: C.primary },
                { label: 'ROI', value: `${result.roi.toFixed(1)}%`, color: C.primary },
              ].map(k => (
                <div key={k.label} style={{ background: 'white', borderRadius: 8, padding: '12px 14px', border: '1px solid #e5e7eb' }}>
                  <div style={{ fontSize: 10, color: '#9ca3af', marginBottom: 4 }}>{k.label}</div>
                  <div style={{ fontSize: 18, fontWeight: 700, color: k.color }}>{k.value}</div>
                </div>
              ))}
            </div>
            {/* 연간 요약 테이블 */}
            <div style={{ background: 'white', borderRadius: 8, border: '1px solid #e5e7eb', overflow: 'hidden', marginBottom: 14 }}>
              <div style={{ padding: '10px 14px', background: C.sidebar1, color: 'white', fontWeight: 700, fontSize: 12 }}>연도별 요약</div>
              <div style={{ overflowX: 'auto' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
                  <thead><tr style={{ background: C.light }}>
                    {['연도','운영사 매출','총 비용','순이익'].map(h => (
                      <th key={h} style={{ padding: '8px 10px', textAlign: h === '연도' ? 'left' : 'right', color: C.sidebar1, fontWeight: 700 }}>{h}</th>
                    ))}
                  </tr></thead>
                  <tbody>
                    {yearlyData.map((y, i) => (
                      <tr key={i} style={{ borderTop: '1px solid #f1f5f9' }}>
                        <td style={{ padding: '8px 10px', fontWeight: 600, color: C.primary }}>{y.yr}</td>
                        <td style={{ padding: '8px 10px', textAlign: 'right' }}>{Math.round(y.myRevenue/10000).toLocaleString()}만</td>
                        <td style={{ padding: '8px 10px', textAlign: 'right' }}>{Math.round(y.totalCost/10000).toLocaleString()}만</td>
                        <td style={{ padding: '8px 10px', textAlign: 'right', fontWeight: 700, color: y.netProfit >= 0 ? '#16a34a' : '#DC2626' }}>{Math.round(y.netProfit/10000).toLocaleString()}만</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
            <div style={{ fontSize: 11, color: '#9ca3af', textAlign: 'center' }}>
              Excel 파일에는 월별 상세 데이터 {result.records.length}개월치가 포함됩니다
            </div>
          </div>
        </div>
      )}

      {/* PDF 미리보기 모달 */}
      {pdfPreview && (
        <div style={{
          position: 'fixed', inset: 0, zIndex: 10000,
          background: 'rgba(0,0,0,0.85)',
          display: 'flex', flexDirection: 'column',
        }}>
          {/* 헤더 */}
          <div style={{
            display: 'flex', alignItems: 'center', justifyContent: 'space-between',
            padding: '14px 18px', background: '#1E1B4B', flexShrink: 0,
          }}>
            <span style={{ color: 'white', fontWeight: 700, fontSize: 15 }}>
              {isMobile ? '📄 PDF 미리보기' : '📊 차트 미리보기'}
            </span>
            <div style={{ display: 'flex', gap: 8 }}>
              {!isMobile && (
                <button onClick={async () => {
                  await exportPptx()
                  setPdfPreview(null)
                }} style={{
                  background: 'rgba(255,255,255,0.15)', color: 'white', border: '1px solid rgba(255,255,255,0.3)',
                  borderRadius: 8, padding: '8px 14px', fontWeight: 700, fontSize: 12, cursor: 'pointer',
                }}>⬇ PPTX</button>
              )}
              <button onClick={async () => {
                await triggerDownload(pdfPreview!.blob, `${rTitle || 'ego_시뮬레이션'}.pdf`)
                showToast('✅ PDF 다운로드가 시작됐습니다')
                setPdfPreview(null)
              }} style={{
                background: '#7C3AED', color: 'white', border: 'none', borderRadius: 8,
                padding: '8px 18px', fontWeight: 700, fontSize: 13, cursor: 'pointer',
              }}>⬇ PDF</button>
              <button onClick={() => setPdfPreview(null)} style={{
                background: 'rgba(255,255,255,0.15)', color: 'white', border: 'none',
                borderRadius: 8, padding: '8px 14px', fontWeight: 600, fontSize: 13, cursor: 'pointer',
              }}>✕ 닫기</button>
            </div>
          </div>
          {/* 이미지 스크롤 */}
          <div style={{ flex: 1, overflowY: 'auto', padding: '16px 12px', display: 'flex', flexDirection: 'column', gap: 12, alignItems: 'center' }}>
            {pdfPreview.images.map((src, i) => (
              <div key={i} style={{ width: '100%', maxWidth: 480 }}>
                <div style={{ fontSize: 11, color: '#a78bfa', marginBottom: 4, textAlign: 'center' }}>
                  {['손익 추이', '비용 분석', '시나리오'][i] || `페이지 ${i+1}`}
                </div>
                <img src={src} style={{ width: '100%', borderRadius: 6, boxShadow: '0 4px 20px rgba(0,0,0,0.5)', display: 'block' }} alt={`preview-${i}`}/>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Toast 알림 */}
      {toast && (
        <div style={{
          position: 'fixed', bottom: 80, left: '50%', transform: 'translateX(-50%)',
          zIndex: 9999, pointerEvents: 'none',
          background: toast.type === 'success' ? '#1a1a2e' : '#1e3a5f',
          color: 'white', borderRadius: 14, padding: '12px 20px',
          boxShadow: '0 6px 24px rgba(0,0,0,0.3)',
          fontSize: 13, fontWeight: 600, textAlign: 'center',
          maxWidth: '90vw', whiteSpace: 'pre-wrap',
          border: `1px solid ${toast.type === 'success' ? '#A855F7' : '#3B82F6'}55`,
          animation: 'fadeInUp 0.25s ease',
        }}>
          {toast.msg}
        </div>
      )}
      <div ref={stickyBarRef} style={{ position: isMobile ? 'relative' : 'sticky', top: 0, zIndex: 20, background: '#f5f3ff', flexShrink: 0 }}>

        {/* 헤더 배너 */}
        <div style={{ padding: isMobile ? '12px 12px 0' : '18px 60px 0' }}>
          <div style={{
            borderRadius: 12, padding: isMobile ? '12px 14px' : '18px 24px',
            background: `linear-gradient(135deg, ${C.sidebar1} 0%, ${C.navy} 100%)`,
            display: 'flex', alignItems: 'center', gap: 12,
          }}>
            <div style={{
              fontWeight: 900, fontSize: isMobile ? 14 : 18, color: 'white', padding: isMobile ? '3px 8px' : '4px 12px',
              borderRadius: 7, background: 'rgba(255,255,255,0.15)', letterSpacing: '-0.5px', flexShrink: 0,
            }}>ego</div>
            <div style={{ flex: 1, minWidth: 0 }}>
              <h1 style={{ color: 'white', fontWeight: 700, fontSize: isMobile ? 15 : 22, lineHeight: 1.3 }}>EV충전소 수익 시뮬레이터</h1>
              <p style={{ color: 'rgba(255,255,255,0.55)', fontSize: isMobile ? 11 : 13, marginTop: 2, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                {r.charger_summary} | {params.operation_months}개월 | 초기투자 {fmtM(r.total_init_cost)}만원
                {params.manager_name && <span style={{ color: '#FB923C', marginLeft: 6 }}>| 담당 {params.manager_name}{params.manager_discount > 0 ? ` · 할인 ${fmtM(params.manager_discount)}만원` : ''}</span>}
              </p>
            </div>
            {loading && <div style={{
              width: 16, height: 16, borderRadius: '50%', flexShrink: 0,
              border: `2px solid ${C.accent}`, borderTopColor: 'transparent',
              animation: 'spin 0.8s linear infinite',
            }}/>}
          </div>
        </div>

        {/* KPI 카드 */}
        {!kpiCollapsed && (() => {
          const KpiCard = ({ k, i, total }: { k: typeof kpiRow1[0]; i: number; total: number }) => (
            <div key={i} style={{
              background: 'white', borderRadius: 10, padding: isMobile ? '12px 12px' : '14px clamp(8px, 1vw, 16px)',
              border: '1px solid #e5e7eb', borderTop: `4px solid ${k.top}`,
              gridColumn: (isMobile && total % 2 !== 0 && i === total - 1) ? 'span 2' : undefined,
              minWidth: 0, display: 'flex', flexDirection: 'column', justifyContent: 'center',
              minHeight: isMobile ? 80 : 90, height: '100%', boxSizing: 'border-box',
            }}>
              <p style={{ fontSize: isMobile ? 10 : 'clamp(9px, 0.9vw, 11px)', color: '#9ca3af', fontWeight: 700, textTransform: 'uppercase', marginBottom: 5, letterSpacing: '0.03em', whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{k.label}</p>
              <p style={{ fontSize: isMobile ? 17 : 'clamp(13px, 1.5vw, 20px)', fontWeight: 700, color: '#1f2937', lineHeight: 1.2, wordBreak: 'keep-all', overflowWrap: 'anywhere' }}>{k.value}</p>
              {(k as any).badge && <span style={{ display: 'inline-block', marginTop: 5, padding: '2px 6px', borderRadius: 99, fontSize: 9, fontWeight: 700, color: 'white', background: (k as any).badgeBg }}>{(k as any).badge}</span>}
              {k.sub && <p style={{ fontSize: 10, color: '#9ca3af', marginTop: 3 }}>{k.sub}</p>}
            </div>
          )
          const outerPad = isMobile ? '10px 12px 0' : '12px min(60px, 3vw) 0'
          const gap = isMobile ? 8 : 10
          const colTemplate = isMobile ? 'repeat(2,1fr)' : 'repeat(4,minmax(0,1fr))'
          return (
            <div style={{ display: 'grid', gridTemplateRows: '1fr 1fr', gap, padding: outerPad }}>
              <div style={{ display: 'grid', gridTemplateColumns: colTemplate, gap, alignItems: 'stretch' }}>
                {kpiRow1.map((k, i) => <KpiCard key={i} k={k} i={i} total={kpiRow1.length}/>)}
              </div>
              <div style={isMobile
                ? { display: 'grid', gridTemplateColumns: colTemplate, gap, alignItems: 'stretch' }
                : { display: 'flex', justifyContent: 'center', gap, alignItems: 'stretch' }}>
                {kpiRow2.map((k, i) => (
                  <div key={i} style={isMobile ? undefined : { flex: `0 0 calc(25% - ${gap * 3 / 4}px)`, minWidth: 0 }}>
                    <KpiCard k={k} i={i} total={kpiRow2.length}/>
                  </div>
                ))}
              </div>
            </div>
          )
        })()}
        <div style={{ display: 'flex', justifyContent: 'center', padding: isMobile ? '4px 12px 2px' : '4px 60px 2px' }}>
          <button onClick={() => setKpiCollapsed(v => !v)} style={{
            background: 'rgba(107,114,128,0.08)', border: '1px solid rgba(107,114,128,0.2)',
            borderRadius: 7, color: '#6b7280', fontSize: 14, fontWeight: 600,
            padding: '3px 24px', cursor: 'pointer', lineHeight: 1,
          }}>{kpiCollapsed ? '▼' : '▲'}</button>
        </div>

        {/* 탭 바 */}
        <div style={{ padding: isMobile ? '8px 12px 0' : '8px 60px 0' }}>
          <div style={{ display: 'flex', gap: 3, padding: 3, borderRadius: 8, background: '#E5DFFE', overflowX: 'auto' }}>
            {TABS.map(t => {
              const active = tab === t.id
              return (
                <button key={t.id} onClick={() => { setTab(t.id); scrollContainerRef?.current?.scrollTo({ top: 0 }) }}
                  style={{
                    flex: 1, padding: isMobile ? '7px 2px' : '6px 4px', fontSize: isMobile ? 10 : 11, fontWeight: 600, borderRadius: 6,
                    cursor: 'pointer', border: 'none', transition: 'all 0.15s', whiteSpace: 'nowrap',
                    background: active ? C.primary : 'transparent',
                    color: active ? 'white' : '#6b7280',
                    boxShadow: active ? '0 1px 4px rgba(0,0,0,0.15)' : 'none',
                  }}>
                  {isMobile ? (t as any).mLabel : t.label}
                </button>
              )
            })}
          </div>
        </div>
      </div>{/* /상단 고정 */}

        {/* 탭 콘텐츠 */}
        <div ref={pdfAreaRef} style={{ padding: isMobile ? '10px 12px 80px' : '10px 60px 22px', display: 'flex', flexDirection: 'column', flex: tab === 'report' ? 1 : undefined }}>

          {/* ── 손익 추이 ── */}
          {tab === 'profit' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>

              {/* 일 사용량 조절 */}
              <div style={{
                background: `linear-gradient(135deg, ${C.primary}08 0%, ${C.accent}12 100%)`,
                borderRadius: 12, border: `1.5px solid ${C.primary}30`,
                padding: '6px 14px', marginTop: 4,
                boxShadow: `0 2px 12px ${C.primary}14`,
              }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 6 }}>
                  <span style={{ fontSize: 16 }}>🎛️</span>
                  <h3 style={{ fontSize: 13, fontWeight: 800, color: C.primary, whiteSpace: 'nowrap' }}>일 사용량 조절</h3>
                  {!isMobile && <span style={{ fontSize: 10, color: '#9ca3af' }}>↔ 드래그하면 그래프가 즉시 업데이트됩니다</span>}
                </div>
                <div style={{ display: 'grid', gridTemplateColumns: `repeat(${params.charger_configs.length}, 1fr)`, gap: 20 }}>
                  {params.charger_configs.map((cfg, i) => (
                    <div key={i}>
                      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 10 }}>
                        <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                          <span style={{ width: 8, height: 8, borderRadius: '50%', background: TYPE_COLORS[i], flexShrink: 0, display: 'inline-block', boxShadow: `0 0 0 3px ${TYPE_COLORS[i]}30` }}/>
                          <span style={{ fontSize: 12, color: '#374151', fontWeight: 700 }}>{cfg.label}</span>
                        </div>
                        <div style={{
                          background: C.primary, borderRadius: 8, padding: '3px 10px',
                          fontSize: 15, fontWeight: 900, color: 'white',
                          minWidth: 52, textAlign: 'center',
                          boxShadow: `0 2px 6px ${C.primary}40`,
                        }}>{cfg.daily_ev}<span style={{ fontSize: 10, fontWeight: 600, marginLeft: 2 }}>대</span></div>
                      </div>
                      <input type="range" min={0.5} max={20} step={0.5} value={cfg.daily_ev}
                        onChange={e => setParams({ charger_configs: params.charger_configs.map((c, idx) => idx === i ? { ...c, daily_ev: Number(e.target.value) } : c) })}
                        style={{ width: '100%', cursor: 'grab', accentColor: TYPE_COLORS[i], height: 6 }}/>
                      <div style={{ display: 'flex', justifyContent: 'space-between', marginTop: 4 }}>
                        <span style={{ fontSize: 9, color: '#9ca3af' }}>최소 0.5대</span>
                        <span style={{ fontSize: 9, color: '#9ca3af' }}>최대 20대</span>
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              <div style={{ background: 'white', borderRadius: 8, border: '1px solid #e5e7eb', padding: '14px 18px' }}>
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 10 }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 7 }}>
                    <div style={{ width: 3, height: 14, borderRadius: 2, background: C.primary }}/>
                    <h3 style={{ fontSize: 12, fontWeight: 700, color: '#374151' }}>월별 순이익 및 누적 손익 추이</h3>
                  </div>
                  {r.bep_month && !isMobile && (
                    <div style={{ display: 'flex', gap: 14, fontSize: 10, flexShrink: 0 }}>
                      <span style={{ color: C.red }}>◀ 회수 전</span>
                      <span style={{ color: C.primary }}>수익 구간 ▶</span>
                    </div>
                  )}
                </div>
                <div>
                <ResponsiveContainer width="100%" height={isMobile ? 220 : 370}>
                  <ComposedChart data={chartData} margin={{ top: 8, right: isMobile ? 15 : 55, left: isMobile ? -10 : 5, bottom: 5 }}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0"/>
                    <XAxis dataKey="month" tick={{ fontSize: 9 }} tickFormatter={v => `${v}M`}/>
                    <YAxis yAxisId="bar" tick={{ fontSize: 9 }} tickFormatter={v => `${Math.round(v)} 만`}/>
                    <YAxis yAxisId="line" orientation="right" tick={{ fontSize: 9 }} tickFormatter={v => `${Math.round(v)} 만`}
                      domain={[Math.min(minCum, 0) * 1.1, maxCum * 1.1]}/>
                    <Tooltip formatter={(v) => [`${Number(v ?? 0).toLocaleString()} 만원`]}
                      labelFormatter={v => `${v}개월`} contentStyle={{ fontSize: 11, borderRadius: 7 }}/>
                    <Legend wrapperStyle={{ fontSize: 10 }}/>
                    <ReferenceLine yAxisId="bar" y={0} stroke="#e5e7eb"/>
                    {r.bep_month && (
                      <ReferenceLine yAxisId="line" x={r.bep_month} stroke="#FB923C" strokeDasharray="0" strokeWidth={3.5}
                        label={{ value: `BEP (${r.bep_month}개월)`, fill: '#FB923C', fontSize: 12, fontWeight: 'bold', position: 'insideTopLeft', offset: 6 }}/>
                    )}
                    <Bar yAxisId="bar" dataKey="net" name="월 순이익" radius={[3,3,0,0]}>
                      {chartData.map((d, i) => {
                        const afterBep = r.bep_month ? i >= r.bep_month - 1 : false
                        return (
                          <Cell key={i}
                            fill={afterBep ? '#6366F1' : C.red}
                            fillOpacity={afterBep ? 0.88 : 0.82}/>
                        )
                      })}
                    </Bar>
                    <Line yAxisId="line" type="monotone" dataKey="cum" name="누적 손익"
                      stroke={C.accent} strokeWidth={3.5} dot={false} activeDot={{ r: 6 }}/>
                  </ComposedChart>
                </ResponsiveContainer>
                </div>
              </div>

              {/* 연도별 요약 */}
              <div style={{ background: 'white', borderRadius: 8, border: '1px solid #e5e7eb', overflow: 'hidden' }}>
                <div style={{ padding: '12px 18px 10px', display: 'flex', alignItems: 'center', gap: 8, borderBottom: '1px solid #f3f4f6' }}>
                  <div style={{ width: 3, height: 14, borderRadius: 2, background: C.primary }}/>
                  <h3 style={{ fontSize: 12, fontWeight: 700, color: '#374151' }}>연도별 요약</h3>
                </div>
                <div style={{ overflowX: 'auto' }}>
                  <table style={{ width: '100%', minWidth: 520, fontSize: 12, borderCollapse: 'collapse' }}>
                    <thead>
                      <tr style={{ background: C.light }}>
                        {['연도', '연간 운영사 매출', '연간 총비용', '연간 순이익', '월평균 순이익'].map(h => (
                          <th key={h} style={{ padding: '9px 14px', textAlign: h === '연도' ? 'left' : 'right', color: C.sidebar1, fontWeight: 700, whiteSpace: 'nowrap', fontSize: 11 }}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {yearlyData.map(y => (
                        <tr key={y.yr} style={{ borderTop: '1px solid #f3f4f6' }}>
                          <td style={{ padding: '9px 14px', fontWeight: 700, color: C.primary }}>{y.yr}</td>
                          <td style={{ padding: '9px 14px', textAlign: 'right', color: '#374151' }}>{fmt(y.myRevenue)}</td>
                          <td style={{ padding: '9px 14px', textAlign: 'right', color: '#374151' }}>{fmt(y.totalCost)}</td>
                          <td style={{ padding: '9px 14px', textAlign: 'right', fontWeight: 600, color: y.netProfit >= 0 ? C.green : C.red }}>{fmt(y.netProfit)}</td>
                          <td style={{ padding: '9px 14px', textAlign: 'right', fontWeight: 600, color: y.avgMonthly >= 0 ? C.green : C.red }}>{fmt(y.avgMonthly)}</td>
                        </tr>
                      ))}
                      {(() => {
                        const totRev = yearlyData.reduce((s, y) => s + y.myRevenue, 0)
                        const totCost = yearlyData.reduce((s, y) => s + y.totalCost, 0)
                        const totNet = yearlyData.reduce((s, y) => s + y.netProfit, 0)
                        const totAvg = yearlyData.length > 0 ? Math.round(totNet / (yearlyData.length * 12)) : 0
                        return (
                          <tr style={{ borderTop: '2px solid #d1d5db', background: '#f8fafc' }}>
                            <td style={{ padding: '10px 14px', fontWeight: 700, color: '#374151' }}>합계</td>
                            <td style={{ padding: '10px 14px', textAlign: 'right', fontWeight: 700, color: '#374151' }}>{fmt(totRev)}</td>
                            <td style={{ padding: '10px 14px', textAlign: 'right', fontWeight: 700, color: '#374151' }}>{fmt(totCost)}</td>
                            <td style={{ padding: '10px 14px', textAlign: 'right', fontWeight: 700, color: totNet >= 0 ? C.green : C.red }}>{fmt(totNet)}</td>
                            <td style={{ padding: '10px 14px', textAlign: 'right', fontWeight: 700, color: totAvg >= 0 ? C.green : C.red }}>{fmt(totAvg)}</td>
                          </tr>
                        )
                      })()}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          )}

          {/* ── 비용 분석 ── */}
          {tab === 'cost' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
            <div style={{ display: 'grid', gridTemplateColumns: isMobile ? '1fr' : '1fr 1fr', gap: 10 }}>
              <div style={{ background: 'white', borderRadius: 8, border: '1px solid #e5e7eb', padding: '14px 18px' }}>
                <h3 style={{ fontSize: 12, fontWeight: 700, color: '#374151', marginBottom: 10 }}>비용 구성 비율</h3>
                <div>
                <ResponsiveContainer width="100%" height={280}>
                  <RPieChart>
                    <Pie data={pieData} dataKey="value" nameKey="name" cx="50%" cy="50%"
                      outerRadius={105} innerRadius={50}
                      label={({ name, percent }: { name?: string; percent?: number }) => `${name ?? ''} ${((percent ?? 0) * 100).toFixed(0)}%`}
                      labelLine={false} fontSize={10}>
                      {pieData.map((d, i) => <Cell key={i} fill={d.fill}/>)}
                    </Pie>
                    <Tooltip formatter={(v) => [`${fmtM(Number(v ?? 0))} 만원`]}/>
                  </RPieChart>
                </ResponsiveContainer>
                </div>
              </div>
              <div style={{ background: 'white', borderRadius: 8, border: '1px solid #e5e7eb', padding: '14px 18px' }}>
                <h3 style={{ fontSize: 12, fontWeight: 700, color: '#374151', marginBottom: 10 }}>항목별 비용 합계</h3>
                <table style={{ width: '100%', fontSize: 11, borderCollapse: 'collapse' }}>
                  <thead><tr style={{ borderBottom: '1px solid #e5e7eb', color: '#9ca3af', fontSize: 10 }}>
                    <th style={{ textAlign: 'left', paddingBottom: 7 }}>항목</th>
                    <th style={{ textAlign: 'right', paddingBottom: 7 }}>합계</th>
                    <th style={{ textAlign: 'right', paddingBottom: 7 }}>비율</th>
                  </tr></thead>
                  <tbody>
                    {pieData.map((d, i) => (
                      <tr key={d.name} style={{ borderBottom: '1px solid #f9fafb' }}>
                        <td style={{ padding: '6px 0', display: 'flex', alignItems: 'center', gap: 5 }}>
                          <span style={{ width: 9, height: 9, borderRadius: '50%', background: d.fill, flexShrink: 0 }}/>
                          {d.name}
                        </td>
                        <td style={{ textAlign: 'right', color: C.primary, fontWeight: 600 }}>{fmtM(d.value)}만원</td>
                        <td style={{ textAlign: 'right', color: '#9ca3af' }}>{((d.value / totalCostAll) * 100).toFixed(1)}%</td>
                      </tr>
                    ))}
                    <tr style={{ fontWeight: 700, borderTop: '1px solid #e5e7eb' }}>
                      <td style={{ paddingTop: 7 }}>합계</td>
                      <td style={{ textAlign: 'right', paddingTop: 7 }}>{fmtM(totalCostAll)}만원</td>
                      <td style={{ textAlign: 'right', paddingTop: 7 }}>100%</td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </div>
            {/* 민감도 분석 */}

            {(() => {
              const cfg0 = params.charger_configs[0]
              const RATES = [200, 230, 260, 290, 320, 350, 380, 410, 440, 470, 500]
              const EVCTS = [1.0, 1.5, 2.0, 2.5, 3.0, 3.5, 4.0, 4.5, 5.0, 5.5, 6.0]
              const heatValues: number[][] = EVCTS.map(ev =>
                RATES.map(rate => {
                  const p2 = { ...params, charger_configs: params.charger_configs.map((c, i) => i === 0 ? { ...c, daily_ev: ev, rate } : c) }
                  return Math.round(runSimulation(p2).records[0].net_profit / 1000)
                })
              )
              const allVals = heatValues.flat()
              const minV = Math.min(...allVals), maxV = Math.max(...allVals)
              const getColor = (v: number) => {
                if (v >= 0) {
                  const t = Math.min(v / Math.max(maxV, 1), 1)
                  return `rgb(${Math.round(60 - t * 20)}, ${Math.round(180 + t * 75)}, ${Math.round(120 - t * 80)})`
                } else {
                  const t = Math.min(Math.abs(v) / Math.max(Math.abs(minV), 1), 1)
                  return `rgb(${Math.round(200 + t * 55)}, ${Math.round(80 - t * 40)}, ${Math.round(80 - t * 40)})`
                }
              }
              const getTextColor = (v: number) => {
                const t = v >= 0 ? Math.min(v / Math.max(maxV, 1), 1) : Math.min(Math.abs(v) / Math.max(Math.abs(minV), 1), 1)
                return t > 0.45 ? 'white' : '#374151'
              }
              const curRate = cfg0?.rate ?? 350
              const curEv = cfg0?.daily_ev ?? 3.0
              const curRateIdx = RATES.reduce((best, r2, i) => Math.abs(r2 - curRate) < Math.abs(RATES[best] - curRate) ? i : best, 0)
              const curEvIdx = EVCTS.reduce((best, e, i) => Math.abs(e - curEv) < Math.abs(EVCTS[best] - curEv) ? i : best, 0)
              const bepRow = RATES.map((_rate, ri) => {
                const idx = EVCTS.findIndex((_e, ei) => heatValues[ei][ri] >= 0)
                return idx >= 0 ? `${EVCTS[idx]}대` : '불가'
              })
              return (
                <div style={{ display: 'flex', flexDirection: 'column', gap: 10, marginTop: 10 }}>
                  {/* 섹션 헤더 */}
                  <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                    <div style={{ width: 3, height: 16, borderRadius: 2, background: C.primary }}/>
                    <h3 style={{ fontSize: 13, fontWeight: 700, color: '#374151' }}>민감도 분석 – 충전요금 × 일 이용 차량수 → 월 순이익</h3>
                  </div>
                  <p style={{ fontSize: 11, color: '#6b7280', marginLeft: 11 }}>충전요금과 차량수를 변화시켰을 때의 <strong>1개월 순이익(천원)</strong>. <span style={{ color: C.primary, fontWeight: 600 }}>보라 테두리</span>가 현재 설정입니다.</p>

                  {/* 현재 설정 요약 */}
                  <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 1fr)', gap: 8 }}>
                    {[
                      { label: '현재 충전요금', value: `${curRate}원/kWh`, icon: '⚡' },
                      { label: '현재 일 차량수', value: `${curEv}대/충전기`, icon: '🚗' },
                      { label: '현재 월 순이익', value: `${(heatValues[curEvIdx][curRateIdx] / 10).toFixed(1)}만원`, icon: '💰' },
                    ].map(item => (
                      <div key={item.label} style={{ background: 'white', borderRadius: 8, border: `1px solid ${C.primary}33`, padding: '10px 14px', textAlign: 'center' }}>
                        <div style={{ fontSize: 16, marginBottom: 3 }}>{item.icon}</div>
                        <div style={{ fontSize: 10, color: '#9ca3af', marginBottom: 2 }}>{item.label}</div>
                        <div style={{ fontSize: 13, fontWeight: 700, color: C.primary }}>{item.value}</div>
                      </div>
                    ))}
                  </div>

                  {/* 히트맵 */}
                  <div style={{ position: 'relative', background: 'white', borderRadius: 10, border: '1px solid #e5e7eb', padding: '14px' }}>
                    {isMobile && <div style={{ position: 'absolute', top: 0, right: 0, bottom: 0, width: 36, background: 'linear-gradient(to right, transparent, rgba(255,255,255,0.95))', pointerEvents: 'none', zIndex: 2, display: 'flex', alignItems: 'center', justifyContent: 'flex-end', paddingRight: 6, borderRadius: '0 10px 10px 0' }}>
                      <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 1 }}>
                        <span style={{ color: '#9333ea', fontSize: 14, fontWeight: 700, lineHeight: 1 }}>›</span>
                        <span style={{ color: '#9ca3af', fontSize: 7 }}>scroll</span>
                      </div>
                    </div>}
                    <div style={{ overflowX: 'auto' }}>
                    <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
                      <span style={{ fontSize: 11, fontWeight: 700, color: '#374151' }}>월 순이익 히트맵 (단위: 천원)</span>
                      <div style={{ display: 'flex', alignItems: 'center', gap: 5, fontSize: 10, color: '#6b7280' }}>
                        <div style={{ width: 10, height: 10, borderRadius: 2, background: 'rgb(235,70,70)' }}/>적자
                        <div style={{ width: 36, height: 10, borderRadius: 2, background: 'linear-gradient(to right, rgb(235,70,70), white, rgb(60,200,120))' }}/>
                        <div style={{ width: 10, height: 10, borderRadius: 2, background: 'rgb(60,200,120)' }}/>흑자
                      </div>
                    </div>
                    <div style={{ minWidth: 520 }}>
                      <div style={{ display: 'flex', marginLeft: 58 }}>
                        {RATES.map(rate => (
                          <div key={rate} style={{ flex: 1, textAlign: 'center', fontSize: 9, marginBottom: 3, fontWeight: rate === curRate ? 700 : 400, color: rate === curRate ? C.primary : '#6b7280' }}>{rate}</div>
                        ))}
                      </div>
                      <div style={{ textAlign: 'center', fontSize: 9, color: '#9ca3af', marginBottom: 6, marginLeft: 58 }}>충전 요금 (원/kWh)</div>
                      {EVCTS.map((ev, ei) => (
                        <div key={ev} style={{ display: 'flex', alignItems: 'center', marginBottom: 2 }}>
                          <div style={{ width: 50, textAlign: 'right', paddingRight: 7, fontSize: 9, fontWeight: ev === curEv ? 700 : 400, color: ev === curEv ? C.primary : '#6b7280', flexShrink: 0 }}>{ev}대</div>
                          {RATES.map((rate, ri) => {
                            const v = heatValues[ei][ri]
                            const isCurrent = ei === curEvIdx && ri === curRateIdx
                            return (
                              <div key={rate} style={{
                                flex: 1, height: 30, display: 'flex', alignItems: 'center', justifyContent: 'center',
                                background: getColor(v), color: getTextColor(v),
                                fontSize: 9, fontWeight: isCurrent ? 900 : 600,
                                borderRadius: 3, margin: '0 1px',
                                border: isCurrent ? `2px solid ${C.primary}` : 'none',
                                boxShadow: isCurrent ? `0 0 0 1px white, 0 0 0 3px ${C.primary}` : 'none',
                              }}>
                                {(v / 10).toFixed(1)}
                              </div>
                            )
                          })}
                        </div>
                      ))}
                    </div>
                    </div>
                  </div>

                  {/* 흑자전환 최소 차량수 */}
                  <div style={{ position: 'relative', background: 'white', borderRadius: 10, border: '1px solid #e5e7eb', overflow: 'hidden' }}>
                    {isMobile && <div style={{ position: 'absolute', top: 0, right: 0, bottom: 0, width: 36, background: 'linear-gradient(to right, transparent, rgba(255,255,255,0.95))', pointerEvents: 'none', zIndex: 2, display: 'flex', alignItems: 'center', justifyContent: 'flex-end', paddingRight: 6 }}>
                      <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 1 }}>
                        <span style={{ color: '#9333ea', fontSize: 14, fontWeight: 700, lineHeight: 1 }}>›</span>
                        <span style={{ color: '#9ca3af', fontSize: 7 }}>scroll</span>
                      </div>
                    </div>}
                    <div style={{ padding: '10px 14px', borderBottom: '1px solid #f3f4f6', display: 'flex', alignItems: 'center', gap: 7 }}>
                      <div style={{ width: 3, height: 13, borderRadius: 2, background: C.primary }}/>
                      <span style={{ fontSize: 11, fontWeight: 700, color: '#374151' }}>충전요금별 흑자 전환 최소 차량수</span>
                      <span style={{ fontSize: 10, color: '#9ca3af', marginLeft: 6 }}>(초기 투자비 제외)</span>
                    </div>
                    <div style={{ overflowX: 'auto' }}>
                      <table style={{ width: '100%', fontSize: 11, borderCollapse: 'collapse', minWidth: 520 }}>
                        <thead>
                          <tr style={{ background: C.light }}>
                            {RATES.map(rate => (
                              <th key={rate} style={{ padding: '7px 4px', textAlign: 'center', color: rate === curRate ? C.primary : C.sidebar1, fontWeight: rate === curRate ? 900 : 700, fontSize: 10 }}>{rate}원</th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          <tr>
                            {bepRow.map((b, i) => (
                              <td key={i} style={{ padding: '9px 4px', textAlign: 'center', fontWeight: 700, color: b === '불가' ? C.red : C.green, background: i === curRateIdx ? `${C.primary}11` : 'white' }}>{b}</td>
                            ))}
                          </tr>
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>
              )
            })()}
            </div>
          )}

          {/* ── 월별 데이터 ── */}
          {tab === 'detail' && (
            <>
            {params.payment_type === '일시불' && (
              <div style={{ display: 'flex', alignItems: 'center', gap: 6, padding: '6px 10px', marginBottom: 6, background: '#fef9c3', border: '1px solid #fde68a', borderRadius: 6, fontSize: 10, color: '#92400e' }}>
                <span style={{ fontSize: 12 }}>ℹ️</span>
                <span><strong>순이익</strong>은 월별 반복 수익 기준입니다. 초기 투자비(일시불 납부)는 <strong>누적손익</strong> 시작값에 반영됩니다.</span>
              </div>
            )}
            <div style={{ position: 'relative', background: 'white', borderRadius: 8, border: '1px solid #e5e7eb', overflow: 'auto', maxHeight: `calc(${isMobile ? '100dvh' : '(100vh / 1.2)'} - ${stickyBarH + (params.payment_type === '일시불' ? 68 : 32)}px)` }}>
              {isMobile && <div style={{ position: 'sticky', top: 0, right: 0, float: 'right', width: 36, height: 30, background: 'linear-gradient(to right, transparent, rgba(255,255,255,0.95))', pointerEvents: 'none', zIndex: 20, display: 'flex', alignItems: 'center', justifyContent: 'flex-end', paddingRight: 6 }}>
                <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 1 }}>
                  <span style={{ color: '#9333ea', fontSize: 14, fontWeight: 700, lineHeight: 1 }}>›</span>
                  <span style={{ color: '#9ca3af', fontSize: 7 }}>scroll</span>
                </div>
              </div>}
              <table style={{ fontSize: 10, borderCollapse: 'collapse', width: '100%', minWidth: 600, tableLayout: 'fixed' }}>
                <colgroup>
                  {[34, 48, 64, 54, 56, 68, 58, 50, 38, 38, ...(params.payment_type === '할부' ? [56] : []), 58, 54, 68].map((w, i) => (
                    <col key={i} style={{ width: w }}/>
                  ))}
                </colgroup>
                <thead>
                  <tr>
                    {[
                      { h: '월차', left: true },
                      { h: 'kWh' },
                      { h: '총매출' },
                      { h: 'PG수수료' },
                      { h: 'PG차감' },
                      { h: '운영사매출' },
                      { h: '전기요금' },
                      { h: '운영비' },
                      { h: 'AS비' },
                      { h: '기타' },
                      ...(params.payment_type === '할부' ? [{ h: '할부금' }] : []),
                      { h: '총비용' },
                      { h: '순이익' },
                      { h: '누적손익' },
                    ].map(({ h, left }) => (
                      <th key={h} style={{ position: 'sticky', top: 0, zIndex: 10, background: C.light, padding: '6px 3px', textAlign: left ? 'left' : 'right', color: C.sidebar1, fontWeight: 700, whiteSpace: 'nowrap', boxShadow: '0 1px 0 #e5e7eb' }}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {r.records.map(rec => (
                    <tr key={rec.month} style={{ borderTop: '1px solid #f9fafb' }}>
                      <td style={{ padding: '5px 3px', fontWeight: 600, color: C.primary }}>{rec.month}</td>
                      <td style={{ padding: '5px 3px', textAlign: 'right' }}>{fmt(rec.kwh)}</td>
                      <td style={{ padding: '5px 3px', textAlign: 'right' }}>{fmt(rec.gross_revenue)}</td>
                      <td style={{ padding: '5px 3px', textAlign: 'right' }}>{fmt(rec.pg_fee)}</td>
                      <td style={{ padding: '5px 3px', textAlign: 'right' }}>{fmt(rec.net_revenue)}</td>
                      <td style={{ padding: '5px 3px', textAlign: 'right' }}>{fmt(rec.my_revenue)}</td>
                      <td style={{ padding: '5px 3px', textAlign: 'right' }}>{fmt(rec.elec_cost)}</td>
                      <td style={{ padding: '5px 3px', textAlign: 'right' }}>{fmt(rec.ops)}</td>
                      <td style={{ padding: '5px 3px', textAlign: 'right' }}>{fmt(rec.as_cost)}</td>
                      <td style={{ padding: '5px 3px', textAlign: 'right' }}>{fmt(rec.other)}</td>
                      {params.payment_type === '할부' && <td style={{ padding: '5px 3px', textAlign: 'right', color: '#EC4899', fontWeight: 600 }}>{fmt(rec.installment)}</td>}
                      <td style={{ padding: '5px 3px', textAlign: 'right' }}>{fmt(rec.total_cost)}</td>
                      <td style={{ padding: '5px 3px', textAlign: 'right', fontWeight: 600, color: rec.net_profit >= 0 ? C.green : C.red }}>{fmt(rec.net_profit)}</td>
                      <td style={{ padding: '5px 3px', textAlign: 'right', fontWeight: 600, color: rec.cumulative >= 0 ? C.green : C.red }}>{fmt(rec.cumulative)}</td>
                    </tr>
                  ))}
                </tbody>
                <tfoot>
                  {(() => {
                    const tot = r.records.reduce((acc, rec) => ({
                      kwh: acc.kwh + rec.kwh,
                      gross_revenue: acc.gross_revenue + rec.gross_revenue,
                      pg_fee: acc.pg_fee + rec.pg_fee,
                      net_revenue: acc.net_revenue + rec.net_revenue,
                      my_revenue: acc.my_revenue + rec.my_revenue,
                      elec_cost: acc.elec_cost + rec.elec_cost,
                      ops: acc.ops + rec.ops,
                      as_cost: acc.as_cost + rec.as_cost,
                      other: acc.other + rec.other,
                      installment: acc.installment + rec.installment,
                      total_cost: acc.total_cost + rec.total_cost,
                      net_profit: acc.net_profit + rec.net_profit,
                    }), { kwh: 0, gross_revenue: 0, pg_fee: 0, net_revenue: 0, my_revenue: 0, elec_cost: 0, ops: 0, as_cost: 0, other: 0, installment: 0, total_cost: 0, net_profit: 0 })
                    const lastCumulative = r.records.length > 0 ? r.records[r.records.length - 1].cumulative : 0
                    const tfStyle: React.CSSProperties = { padding: '6px 3px', textAlign: 'right', fontWeight: 700, color: '#374151', background: '#f3f4f6', borderTop: '2px solid #d1d5db' }
                    return (
                      <tr>
                        <td style={{ ...tfStyle, textAlign: 'left' }}>합계</td>
                        <td style={tfStyle}>{fmt(Math.round(tot.kwh * 10) / 10)}</td>
                        <td style={tfStyle}>{fmt(tot.gross_revenue)}</td>
                        <td style={tfStyle}>{fmt(tot.pg_fee)}</td>
                        <td style={tfStyle}>{fmt(tot.net_revenue)}</td>
                        <td style={tfStyle}>{fmt(tot.my_revenue)}</td>
                        <td style={tfStyle}>{fmt(tot.elec_cost)}</td>
                        <td style={tfStyle}>{fmt(tot.ops)}</td>
                        <td style={tfStyle}>{fmt(tot.as_cost)}</td>
                        <td style={tfStyle}>{fmt(tot.other)}</td>
                        {params.payment_type === '할부' && <td style={{ ...tfStyle, color: '#EC4899' }}>{fmt(tot.installment)}</td>}
                        <td style={tfStyle}>{fmt(tot.total_cost)}</td>
                        <td style={{ ...tfStyle, color: tot.net_profit >= 0 ? C.green : C.red }}>{fmt(tot.net_profit)}</td>
                        <td style={{ ...tfStyle, color: lastCumulative >= 0 ? C.green : C.red }}>{fmt(lastCumulative)}</td>
                      </tr>
                    )
                  })()}
                </tfoot>
              </table>
            </div>
            </>
          )}

          {/* ── 리포트 ── */}
          {tab === 'report' && (
            <div style={{ display: 'grid', gridTemplateColumns: isMobile ? '1fr' : '1fr 1fr', gap: 10, alignItems: 'stretch', flex: 1 }}>
              <div style={{ background: 'white', borderRadius: 8, border: '1px solid #e5e7eb', padding: '18px 28px' }}>
                <h3 style={{ fontSize: 12, fontWeight: 700, color: '#374151', marginBottom: 10 }}>리포트 설정</h3>
                <div style={{ marginBottom: 8 }}>
                  <label style={{ fontSize: 10, color: '#9ca3af', display: 'block', marginBottom: 3 }}>리포트 제목</label>
                  <input value={rTitle} onChange={e => setRTitle(e.target.value)}
                    style={{ width: '100%', border: '1px solid #e5e7eb', borderRadius: 6, padding: '7px 10px', fontSize: 12, outline: 'none', boxSizing: 'border-box' }}/>
                </div>
                <div style={{ marginBottom: 8 }}>
                  <label style={{ fontSize: 10, color: '#9ca3af', display: 'block', marginBottom: 3 }}>기관명</label>
                  <input value={rCo} onChange={e => setRCo(e.target.value)} placeholder="(선택)"
                    style={{ width: '100%', border: '1px solid #e5e7eb', borderRadius: 6, padding: '7px 10px', fontSize: 12, outline: 'none', boxSizing: 'border-box' }}/>
                </div>
                <div style={{ marginBottom: 12 }}>
                  <label style={{ fontSize: 10, color: '#9ca3af', display: 'block', marginBottom: 3 }}>담당자</label>
                  <input value={rManager} onChange={e => setRManager(e.target.value)} placeholder="(선택)"
                    style={{ width: '100%', border: '1px solid #e5e7eb', borderRadius: 6, padding: '7px 10px', fontSize: 12, outline: 'none', boxSizing: 'border-box' }}/>
                </div>
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8 }}>
                  <div>
                    <button onClick={exportExcel} style={{
                      width: '100%', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 5,
                      padding: '10px 0', borderRadius: 8, color: 'white', fontWeight: 700, fontSize: 12, cursor: 'pointer', border: 'none',
                      background: `linear-gradient(135deg, ${C.sidebar1}, ${C.primary})`,
                    }}><FileDown size={13}/> Excel 저장</button>
                    <p style={{ fontSize: 10, color: '#9ca3af', textAlign: 'center', marginTop: 4 }}>월별데이터</p>
                  </div>
                  <div>
                  {isMobile ? (
                    <button onClick={exportPdf} style={{
                      width: '100%', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 5,
                      padding: '10px 0', borderRadius: 8, fontWeight: 700, fontSize: 12, cursor: 'pointer',
                      background: `linear-gradient(135deg, #7C3AED, #A855F7)`, color: 'white', border: 'none',
                    }}><FileDown size={13}/> PDF 미리보기</button>
                  ) : (
                    <button onClick={exportPptx} style={{
                      width: '100%', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 5,
                      padding: '10px 0', borderRadius: 8, fontWeight: 700, fontSize: 12, cursor: 'pointer',
                      background: `linear-gradient(135deg, #7C3AED, #A855F7)`, color: 'white', border: 'none',
                    }}><FileDown size={13}/> PPTX 저장</button>
                  )}
                    <p style={{ fontSize: 10, color: '#9ca3af', textAlign: 'center', marginTop: 4 }}>수익 분석</p>
                  </div>
                </div>
              </div>
              <div style={{ background: 'white', borderRadius: 8, border: '1px solid #e5e7eb', padding: '18px 28px' }}>
                <h3 style={{ fontSize: 12, fontWeight: 700, color: '#374151', marginBottom: 10 }}>연간 요약</h3>
                <table style={{ width: '100%', fontSize: 11, borderCollapse: 'collapse' }}>
                  <thead><tr style={{ background: C.light, color: C.sidebar1 }}>
                    {['연도','내 수익','총 비용','순이익','누적 손익'].map(h => (
                      <th key={h} style={{ padding: '7px 6px', textAlign: h === '연도' ? 'left' : 'right', fontWeight: 700 }}>{h}</th>
                    ))}
                  </tr></thead>
                  <tbody>
                    {yearlyData.map(y => (
                      <tr key={y.yr} style={{ borderTop: '1px solid #f9fafb' }}>
                        <td style={{ padding: '6px', fontWeight: 700, color: C.primary }}>{y.yr}</td>
                        <td style={{ padding: '6px', textAlign: 'right' }}>{fmtM(y.myRevenue)}만</td>
                        <td style={{ padding: '6px', textAlign: 'right' }}>{fmtM(y.totalCost)}만</td>
                        <td style={{ padding: '6px', textAlign: 'right', fontWeight: 600, color: y.netProfit >= 0 ? C.green : C.red }}>{fmtM(y.netProfit)}만</td>
                        <td style={{ padding: '6px', textAlign: 'right', fontWeight: 600, color: y.cumulative >= 0 ? C.green : C.red }}>{fmtM(y.cumulative)}만</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* ── 시나리오 ── */}
          {tab === 'scenario' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
              {/* 시나리오 설정 패널 */}
              {(() => {
                const cfg0 = params.charger_configs[0]
                const ROWS: { key: 'dm'|'cm'|'gm'|'km'; label: string; unit: string; baseLabel: string; baseVal: number; min: number; max: number; step: number; fmt: (abs: number) => string }[] = [
                  { key: 'dm', label: '일 이용 차량', unit: '대',     baseLabel: `${cfg0?.daily_ev ?? 0}대`,      baseVal: cfg0?.daily_ev ?? 1,         min: (cfg0?.daily_ev??1)*0.1, max: (cfg0?.daily_ev??1)*3.0, step: 0.5,  fmt: v => `${v.toFixed(1)}대` },
                  { key: 'cm', label: '충전 단가',    unit: '원/kWh', baseLabel: `${cfg0?.rate ?? 0}원/kWh`,     baseVal: cfg0?.rate ?? 1,             min: (cfg0?.rate??1)*0.5,     max: (cfg0?.rate??1)*2.0,    step: 10,   fmt: v => `${Math.round(v)}원/kWh` },
                  { key: 'gm', label: '연간 증가율',  unit: '%',      baseLabel: `${params.ev_growth_rate}%`,    baseVal: params.ev_growth_rate || 1,  min: 0,                       max: (params.ev_growth_rate||1)*3.0, step: 0.5, fmt: v => `${v.toFixed(1)}%` },
                  { key: 'km', label: '평균 충전량',  unit: 'kWh',   baseLabel: `${cfg0?.avg_kwh ?? 0}kWh`,     baseVal: cfg0?.avg_kwh ?? 1,          min: (cfg0?.avg_kwh??1)*0.3,  max: (cfg0?.avg_kwh??1)*2.0, step: 0.5,  fmt: v => `${v.toFixed(1)}kWh` },
                ]
                // slider value = absolute (baseVal × multiplier), onChange converts back to multiplier
                const updateParam = (key: 'dm'|'cm'|'gm'|'km', idx: number, absVal: number, baseVal: number) => {
                  const mult = baseVal > 0 ? absVal / baseVal : 1
                  if (!scLocked[key] && idx === 1) {
                    // 기본 시나리오 변경 시 → 다른 시나리오도 같은 비율로 스케일
                    setScDefs(prev => {
                      const oldBase = prev[1][key]
                      const ratio = oldBase > 0 ? mult / oldBase : 1
                      return prev.map(s => ({ ...s, [key]: Math.max(0.01, s[key] * ratio) }))
                    })
                  } else {
                    setScDefs(prev => prev.map((s, j) =>
                      scLocked[key] || j === idx ? { ...s, [key]: mult } : s
                    ))
                  }
                }
                return (
                  <div style={{ background: 'white', borderRadius: 10, border: '1px solid #e5e7eb', padding: '3px 12px' }}>
                    {/* 헤더 */}
                    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 2 }}>
                      <div style={{ display: 'flex', alignItems: 'center', gap: 5 }}>
                        <div style={{ width: 3, height: 8, borderRadius: 2, background: C.primary }}/>
                        <span style={{ fontSize: 9, fontWeight: 700, color: '#374151' }}>시나리오 설정</span>
                        <span style={{ fontSize: 7, color: '#9ca3af' }}>(배율 조절 시 자동 반영 · 🔒고정 시 전체 동일값)</span>
                      </div>
                      <button onClick={() => { setScDefs(SC_DEFAULTS.map(s => ({ ...s }))); setScLocked({ dm: false, cm: false, gm: false, km: false }) }} style={{
                        padding: '1px 6px', fontSize: 8, fontWeight: 600, borderRadius: 4, cursor: 'pointer',
                        background: 'white', border: `1px solid ${C.light}`, color: '#6b7280',
                      }}>초기화</button>
                    </div>

                    {/* 컬럼 헤더 */}
                    <div style={{ position: 'relative' }}>
                      {isMobile && <div style={{ position: 'absolute', top: 0, right: 0, bottom: 0, width: 36, background: 'linear-gradient(to right, transparent, rgba(255,255,255,0.95))', pointerEvents: 'none', zIndex: 2, display: 'flex', alignItems: 'center', justifyContent: 'flex-end', paddingRight: 6 }}>
                        <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 1 }}>
                          <span style={{ color: '#9333ea', fontSize: 14, fontWeight: 700, lineHeight: 1 }}>›</span>
                          <span style={{ color: '#9ca3af', fontSize: 7 }}>scroll</span>
                        </div>
                      </div>}
                    <div style={{ overflowX: 'auto' }}>
                    <div style={{ display: 'grid', gridTemplateColumns: '18px 80px 68px 1fr 1fr 1fr', gap: 5, marginBottom: 1, alignItems: 'center', minWidth: 500 }}>
                      <div/>
                      <div style={{ fontSize: 8, color: '#9ca3af', fontWeight: 600 }}>항목</div>
                      <div style={{ fontSize: 8, color: '#9ca3af', fontWeight: 600, textAlign: 'center' }}>현재 기준값</div>
                      {scDefs.map(sc => (
                        <div key={sc.label} style={{ textAlign: 'center', fontSize: 9, fontWeight: 700, color: sc.color }}>{sc.label}</div>
                      ))}
                    </div>

                    {/* 파라미터 행 */}
                    <div style={{ display: 'flex', flexDirection: 'column', gap: 0, minWidth: 500 }}>
                      {ROWS.map((row, ri) => {
                        const locked = scLocked[row.key]
                        return (
                          <React.Fragment key={row.key}>
                          {ri === 2 && (
                            <div style={{ borderTop: '1px dashed #e5e7eb', margin: '1px 0' }}/>
                          )}
                          <div style={{
                            display: 'grid', gridTemplateColumns: '18px 80px 68px 1fr 1fr 1fr', gap: 4, alignItems: 'center',
                            background: locked ? `${C.primary}08` : 'transparent',
                            borderRadius: 3, padding: '1px 3px',
                            border: locked ? `1px solid ${C.primary}22` : '1px solid transparent',
                          }}>
                            {/* 고정 버튼 */}
                            <button
                              onClick={() => {
                                const next = !locked
                                setScLocked(prev => ({ ...prev, [row.key]: next }))
                                if (next) {
                                  const baseVal = scDefs[1][row.key]
                                  setScDefs(prev => prev.map(s => ({ ...s, [row.key]: baseVal })))
                                }
                              }}
                              title={locked ? '고정 해제' : '전체 동일값 고정'}
                              style={{
                                width: 16, height: 16, borderRadius: 3, border: 'none', cursor: 'pointer', fontSize: 9,
                                background: locked ? C.primary : '#f3f4f6',
                                display: 'flex', alignItems: 'center', justifyContent: 'center',
                                transition: 'background 0.15s',
                              }}
                            >{locked ? '🔒' : '🔓'}</button>

                            {/* 항목명 */}
                            <div style={{ lineHeight: 1.2 }}>
                              <div style={{ fontSize: 9, fontWeight: 600, color: locked ? C.primary : '#374151' }}>{row.label}</div>
                              <div style={{ fontSize: 7, color: '#9ca3af' }}>{row.unit}</div>
                            </div>

                            {/* 현재 기준값 */}
                            <div style={{ textAlign: 'center', background: '#F5F3FF', borderRadius: 3, padding: '1px 4px' }}>
                              <div style={{ fontSize: 9, fontWeight: 700, color: C.primary }}>{row.baseLabel}</div>
                              <div style={{ fontSize: 6, color: '#9ca3af' }}>기준</div>
                            </div>

                            {/* 각 시나리오 슬라이더 */}
                            {scDefs.map((sc, i) => {
                              const absVal = row.baseVal * sc[row.key]
                              return (
                              <div key={sc.label} style={{ display: 'flex', alignItems: 'center', gap: 3 }}>
                                <input type="range" min={row.min} max={row.max} step={row.step}
                                  value={Math.min(Math.max(absVal, row.min), row.max)}
                                  onChange={e => updateParam(row.key, i, Number(e.target.value), row.baseVal)}
                                  style={{ flex: 1, accentColor: locked ? C.primary : sc.color, height: 2 }}/>
                                <div style={{ minWidth: 38, textAlign: 'right' }}>
                                  <div style={{ fontSize: 9, fontWeight: 700, color: locked ? C.primary : sc.color }}>×{sc[row.key].toFixed(2)}</div>
                                  <div style={{ fontSize: 7, color: '#9ca3af' }}>{row.fmt(absVal)}</div>
                                </div>
                              </div>
                            )})}
                          </div>
                          </React.Fragment>
                        )
                      })}
                    </div>
                    </div>
                    </div>
                  </div>
                )
              })()}

              {/* 그래프 먼저 */}
              {scRan && scResults.length > 0 && (
                <div style={{ background: 'white', borderRadius: 8, border: '1px solid #e5e7eb', padding: '14px 18px' }}>
                  <h3 style={{ fontSize: 12, fontWeight: 700, color: '#374151', marginBottom: 14 }}>누적 손익 비교 (만원)</h3>
                  <div>
                  <ResponsiveContainer width="100%" height={380}>
                    <ComposedChart data={scChartData} margin={{ top: 8, right: 50, left: 20, bottom: 8 }}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0"/>
                      <XAxis dataKey="month" tickFormatter={v => `${v}M`} tick={{ fontSize: 9 }}/>
                      <YAxis tickFormatter={v => v.toLocaleString()} tick={{ fontSize: 9 }}/>
                      <Tooltip formatter={(v) => [`${Number(v ?? 0).toLocaleString()} 만원`]}
                        labelFormatter={v => `${v}개월`} contentStyle={{ fontSize: 11, borderRadius: 7 }}
                        itemSorter={(item) => -(item.value as number)}/>
                      <Legend wrapperStyle={{ fontSize: 10 }}/>
                      <ReferenceLine y={0} stroke="#e5e7eb"/>
                      {scDefs.map(sc => <Line key={sc.label} type="monotone" dataKey={sc.label} stroke={sc.color} strokeWidth={2.5} dot={false} activeDot={{ r: 4 }}/>)}
                    </ComposedChart>
                  </ResponsiveContainer>
                  </div>
                </div>
              )}

              {/* 결과 카드 */}
              {scRan && scResults.length > 0 && (
                <div style={{ display: 'grid', gridTemplateColumns: isMobile ? '1fr' : 'repeat(3,1fr)', gap: 10 }}>
                  {scDefs.map(sc => {
                    const res = scResults.find(r => r.label === sc.label)
                    return (
                      <div key={sc.label} style={{ background: 'white', borderRadius: 8, padding: '12px 14px', border: '1px solid #e5e7eb', borderTop: `4px solid ${sc.color}` }}>
                        <p style={{ fontWeight: 700, fontSize: 12, color: sc.color, marginBottom: 3 }}>{sc.label} 시나리오</p>
                        <p style={{ fontSize: 10, color: '#9ca3af', marginBottom: 10 }}>일 차량 ×{sc.dm.toFixed(1)} · 단가 ×{sc.cm.toFixed(2)}</p>
                        {res ? [
                          { l: 'BEP', v: res.result.bep_month ? `${res.result.bep_month}개월` : 'N/A' },
                          { l: 'ROI', v: `${res.result.roi.toFixed(1)}%` },
                          { l: '최종 누적', v: `${fmtM(res.result.records[res.result.records.length - 1].cumulative)} 만원` },
                        ].map(k => (
                          <div key={k.l} style={{ display: 'flex', justifyContent: 'space-between', fontSize: 11, marginBottom: 4 }}>
                            <span style={{ color: '#9ca3af' }}>{k.l}</span>
                            <span style={{ fontWeight: 700, color: sc.color }}>{k.v}</span>
                          </div>
                        )) : null}
                      </div>
                    )
                  })}
                </div>
              )}

              {/* 수치 비교 테이블 */}
              {scRan && scResults.length > 0 && (
                <div style={{ position: 'relative', background: 'white', borderRadius: 10, border: '1px solid #e5e7eb', overflow: 'hidden' }}>
                  {isMobile && <div style={{ position: 'absolute', top: 0, right: 0, bottom: 0, width: 36, background: 'linear-gradient(to right, transparent, rgba(255,255,255,0.95))', pointerEvents: 'none', zIndex: 2, display: 'flex', alignItems: 'center', justifyContent: 'flex-end', paddingRight: 6 }}>
                    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 1 }}>
                      <span style={{ color: '#9333ea', fontSize: 14, fontWeight: 700, lineHeight: 1 }}>›</span>
                      <span style={{ color: '#9ca3af', fontSize: 7 }}>scroll</span>
                    </div>
                  </div>}
                  <div style={{ padding: '12px 18px', borderBottom: '1px solid #f3f4f6', display: 'flex', alignItems: 'center', gap: 8 }}>
                    <div style={{ width: 3, height: 14, borderRadius: 2, background: C.primary }}/>
                    <span style={{ fontSize: 12, fontWeight: 700, color: '#374151' }}>시나리오별 수치 비교</span>
                  </div>
                  <div style={{ overflowX: 'auto' }}>
                  <table style={{ width: '100%', minWidth: 640, fontSize: 11, borderCollapse: 'collapse' }}>
                    <thead>
                      <tr style={{ background: C.light }}>
                        {['시나리오','일평균 차량 (대)','충전요금 (원/kWh)','1회 충전량 (kWh)','차량 증가율 (%)','총 순이익 (만원)','투자 회수','ROI (%)'].map(h => (
                          <th key={h} style={{ padding: '9px 12px', textAlign: h === '시나리오' ? 'left' : 'right', color: C.sidebar1, fontWeight: 700, whiteSpace: 'nowrap', fontSize: 11 }}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {scDefs.map((sc, i) => {
                        const res = scResults.find(r => r.label === sc.label)
                        const cfg = params.charger_configs[0]
                        const ev = (cfg?.daily_ev ?? 0) * sc.dm
                        const rate = Math.round((cfg?.rate ?? 0) * sc.cm)
                        const kwh = ((cfg?.avg_kwh ?? 0) * sc.km).toFixed(1)
                        const growth = (params.ev_growth_rate * sc.gm).toFixed(1)
                        const lastCum = res ? res.result.records[res.result.records.length - 1].cumulative : null
                        const bep = res?.result.bep_month
                        const roi = res?.result.roi ?? 0
                        const isPos = (lastCum ?? 0) >= 0
                        return (
                          <tr key={sc.label} style={{ borderTop: '1px solid #f3f4f6', background: i % 2 === 0 ? 'white' : '#fafafa' }}>
                            <td style={{ padding: '10px 12px' }}>
                              <div style={{ display: 'flex', alignItems: 'center', gap: 7 }}>
                                <span style={{ width: 10, height: 10, borderRadius: '50%', background: sc.color, flexShrink: 0, display: 'inline-block' }}/>
                                <span style={{ fontWeight: 700, color: sc.color }}>{sc.label}</span>
                              </div>
                            </td>
                            <td style={{ padding: '10px 12px', textAlign: 'right' }}>{ev.toFixed(1)}</td>
                            <td style={{ padding: '10px 12px', textAlign: 'right' }}>{rate.toLocaleString()}</td>
                            <td style={{ padding: '10px 12px', textAlign: 'right' }}>{kwh}</td>
                            <td style={{ padding: '10px 12px', textAlign: 'right' }}>{growth}</td>
                            <td style={{ padding: '10px 12px', textAlign: 'right', fontWeight: 700, color: isPos ? C.green : C.red }}>{lastCum != null ? fmtM(lastCum) : '-'}</td>
                            <td style={{ padding: '10px 12px', textAlign: 'right', fontWeight: 600, color: bep ? C.green : C.red }}>{bep ? `${bep}개월차` : '미도달'}</td>
                            <td style={{ padding: '10px 12px', textAlign: 'right', fontWeight: 600, color: roi >= 0 ? C.green : C.red }}>{roi.toFixed(1)}</td>
                          </tr>
                        )
                      })}
                    </tbody>
                  </table>
                  </div>
                </div>
              )}
            </div>
          )}

          {/* ── 일시불·할부 비교 ── */}
          {/* ── 손익분기 kWh 계산기 ── */}
          {tab === 'bep_kwh' && (() => {
            // 충전기별 정보
            const configs = params.charger_configs
            const totalCount = configs.reduce((s, c) => s + c.count, 0)
            const totalKw = configs.reduce((s, c) => s + c.kw * c.count, 0)

            // 전기요금 단가 (kWh당)
            const elecKwhRate = (params.elec_kwh_rate + params.elec_climate_rate + params.elec_fuel_rate) * (1 + params.elec_fund_pct / 100) * (1 + params.elec_vat_pct / 100)

            // 충전기별 손익분기 계산
            type BepRow = {
              label: string; kw: number; count: number; rate: number
              revenuePerKwh: number; costPerKwh: number; marginPerKwh: number
              monthlyFixed: number; monthlyInitAmort: number; bepKwhMonth: number; bepKwhDay: number
              bepHoursDay: number; utilizationPct: number
            }

            const rows: BepRow[] = configs.map(cfg => {
              // 충전 단가
              const rate = cfg.rate
              // kWh당 순수입 = 단가 × (1-PG%) × 운영사% - 전기요금
              const revenuePerKwh = rate * (1 - params.pg_fee_pct / 100) * (params.revenue_share_pct / 100)
              const costPerKwh = elecKwhRate
              const marginPerKwh = revenuePerKwh - costPerKwh

              // 월 고정비용 (해당 충전기 비율로 안분)
              const ratio = totalCount > 0 ? cfg.count / totalCount : 1
              // 초기 투자비용 월 할당 = (충전기단가+설치비)*대수 + (기타초기+한전부담금+사용전검사)*비율 / 운영개월수
              const initCostThis = (cfg.cost_unit ?? params.cost_charger_unit) * cfg.count + (cfg.cost_install ?? params.cost_installation) * cfg.count
              const initCostShared = (params.cost_other_init + params.cost_kepco_burden + params.cost_safety_inspection) * ratio
              const monthlyInitAmort = (initCostThis + initCostShared) / params.operation_months
              const monthlyFixed =
                (params.monthly_ops + params.monthly_as + params.monthly_comm +
                 params.monthly_elec_safety + params.monthly_other) * ratio +
                params.elec_basic_rate * cfg.kw * cfg.count +
                (params.insurance_yearly / 12) * ratio +
                monthlyInitAmort

              // 손익분기 kWh (월)
              const bepKwhMonth = marginPerKwh > 0 ? monthlyFixed / marginPerKwh : Infinity
              const bepKwhDay = bepKwhMonth / 30

              // 충전기 1기당 하루 최대 충전 가능량 (kW × 24h × 이용률)
              const maxKwhDay = cfg.kw * 24 * cfg.count
              const bepHoursDay = cfg.kw * cfg.count > 0 ? bepKwhDay / (cfg.kw * cfg.count) : 0
              const utilizationPct = maxKwhDay > 0 ? (bepKwhDay / maxKwhDay) * 100 : 0

              return { label: cfg.label, kw: cfg.kw, count: cfg.count, rate,
                revenuePerKwh, costPerKwh, marginPerKwh,
                monthlyFixed, monthlyInitAmort, bepKwhMonth, bepKwhDay, bepHoursDay, utilizationPct }
            })

            // 전체 합산
            const totalMonthlyFixed = rows.reduce((s, r) => s + r.monthlyFixed, 0)
            const avgMarginPerKwh = rows.reduce((s, r) => s + r.marginPerKwh * r.count, 0) / (totalCount || 1)
            const totalBepKwhMonth = avgMarginPerKwh > 0 ? totalMonthlyFixed / avgMarginPerKwh : Infinity
            const totalBepKwhDay = totalBepKwhMonth / 30
            const totalMaxKwhDay = totalKw * 24
            const totalUtilization = totalMaxKwhDay > 0 ? (totalBepKwhDay / totalMaxKwhDay) * 100 : 0

            const cardStyle: React.CSSProperties = {
              background: 'white', borderRadius: 12, padding: '18px 20px',
              boxShadow: '0 1px 4px rgba(0,0,0,0.08)', border: '1px solid #f0f0f0',
            }
            const secTitle: React.CSSProperties = {
              fontSize: 13, fontWeight: 700, color: '#6b7280', letterSpacing: '0.05em',
              textTransform: 'uppercase', marginBottom: 14,
            }
            const kpiVal: React.CSSProperties = { fontSize: 28, fontWeight: 800, color: C.primary, lineHeight: 1.1 }
            const kpiLbl: React.CSSProperties = { fontSize: 12, color: '#9ca3af', marginTop: 4 }

            const fmtKwh = (v: number) => v === Infinity ? '∞' : v >= 10000 ? `${(v/10000).toFixed(1)}만` : v.toLocaleString('ko-KR', { maximumFractionDigits: 1 })
            const fmtHours = (h: number) => h >= 24 ? '24h 초과' : `${h.toFixed(1)}h`

            const barColor = (pct: number) => pct <= 20 ? '#22c55e' : pct <= 50 ? '#f59e0b' : pct <= 80 ? '#f97316' : '#ef4444'

            return (
              <div style={{ padding: isMobile ? '10px 0' : '0' }}>
                {/* 안내 배너 */}
                <div style={{ background: 'linear-gradient(135deg,#6D28D9,#4338CA)', borderRadius: 12, padding: '16px 20px', marginBottom: 20, color: 'white' }}>
                  <div style={{ fontSize: 15, fontWeight: 700, marginBottom: 6 }}>⚡ 손익분기 kWh 계산기</div>
                  <div style={{ fontSize: 13, opacity: 0.85, lineHeight: 1.6 }}>
                    좌측에서 충전기 구성·비용을 설정하면 <strong>하루 몇 kWh를 충전해야 수익이 나는지</strong> 자동으로 계산합니다.<br/>
                    시뮬레이션 실행 없이 현재 설정값 기준으로 즉시 계산됩니다.
                  </div>
                </div>

                {/* 전체 합산 KPI */}
                <div style={{ display: 'grid', gridTemplateColumns: isMobile ? '1fr 1fr' : 'repeat(4,1fr)', gap: 12, marginBottom: 20 }}>
                  {[
                    { label: '하루 필요 충전량', value: `${fmtKwh(totalBepKwhDay)} kWh`, sub: '전체 충전기 합산' },
                    { label: '월 필요 충전량', value: `${fmtKwh(totalBepKwhMonth)} kWh`, sub: '하루×30일' },
                    { label: '월 고정비용', value: `${Math.round(totalMonthlyFixed/10000).toLocaleString()}만원`, sub: '운영비+전기기본료+기타' },
                    { label: '설비 이용률', value: totalBepKwhDay === Infinity ? '불가' : `${totalUtilization.toFixed(1)}%`, sub: '24시간 대비 필요 가동률' },
                  ].map((k, i) => (
                    <div key={i} style={{ ...cardStyle, textAlign: 'center' }}>
                      <div style={kpiVal}>{k.value}</div>
                      <div style={{ fontSize: 12, fontWeight: 700, color: '#374151', marginTop: 6 }}>{k.label}</div>
                      <div style={kpiLbl}>{k.sub}</div>
                    </div>
                  ))}
                </div>

                {/* 충전기별 상세 */}
                <div style={cardStyle}>
                  <div style={secTitle}>충전기 유형별 손익분기</div>
                  <div style={{ overflowX: 'auto' }}>
                    <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
                      <thead>
                        <tr style={{ background: '#f9fafb' }}>
                          {['충전기', '수량', '단가(원/kWh)', 'kWh당 순이익', '월 고정비용', '하루 필요량', '가동 시간', '이용률', '판정'].map(h => (
                            <th key={h} style={{ padding: '10px 12px', textAlign: 'center', fontSize: 12, color: '#6b7280', fontWeight: 700, whiteSpace: 'nowrap', borderBottom: '2px solid #e5e7eb' }}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {rows.map((r, i) => {
                          const impossible = r.marginPerKwh <= 0
                          const easy = r.utilizationPct < 20
                          const verdict = impossible ? '❌ 손익 불가' : easy ? '✅ 달성 용이' : r.utilizationPct < 50 ? '🟡 보통' : r.utilizationPct < 80 ? '🟠 어려움' : '🔴 매우 어려움'
                          return (
                            <tr key={i} style={{ borderBottom: '1px solid #f3f4f6', background: i % 2 === 0 ? 'white' : '#fafafa' }}>
                              <td style={{ padding: '10px 12px', fontWeight: 700, color: C.primary, whiteSpace: 'nowrap' }}>{r.label}</td>
                              <td style={{ padding: '10px 12px', textAlign: 'center', color: '#374151' }}>{r.count}기</td>
                              <td style={{ padding: '10px 12px', textAlign: 'right', fontVariantNumeric: 'tabular-nums' }}>{r.rate.toLocaleString()}</td>
                              <td style={{ padding: '10px 12px', textAlign: 'right', fontVariantNumeric: 'tabular-nums', color: r.marginPerKwh > 0 ? '#16a34a' : '#dc2626', fontWeight: 700 }}>
                                {r.marginPerKwh > 0 ? `+${r.marginPerKwh.toFixed(1)}` : r.marginPerKwh.toFixed(1)}원
                              </td>
                              <td style={{ padding: '10px 12px', textAlign: 'right', fontVariantNumeric: 'tabular-nums' }}>{Math.round(r.monthlyFixed).toLocaleString()}원</td>
                              <td style={{ padding: '10px 12px', textAlign: 'right', fontVariantNumeric: 'tabular-nums', fontWeight: 700, color: impossible ? '#dc2626' : '#1d4ed8' }}>
                                {impossible ? '-' : `${fmtKwh(r.bepKwhDay)} kWh`}
                              </td>
                              <td style={{ padding: '10px 12px', textAlign: 'center', color: '#6b7280' }}>
                                {impossible ? '-' : fmtHours(r.bepHoursDay)}
                              </td>
                              <td style={{ padding: '10px 12px', textAlign: 'center' }}>
                                {!impossible && (
                                  <div style={{ display: 'flex', alignItems: 'center', gap: 6, justifyContent: 'center' }}>
                                    <div style={{ width: 48, height: 6, background: '#e5e7eb', borderRadius: 3, overflow: 'hidden' }}>
                                      <div style={{ width: `${Math.min(100, r.utilizationPct)}%`, height: '100%', background: barColor(r.utilizationPct), borderRadius: 3, transition: 'width 0.4s' }}/>
                                    </div>
                                    <span style={{ fontSize: 12, color: barColor(r.utilizationPct), fontWeight: 700 }}>{r.utilizationPct.toFixed(0)}%</span>
                                  </div>
                                )}
                              </td>
                              <td style={{ padding: '10px 12px', textAlign: 'center', fontSize: 12, whiteSpace: 'nowrap' }}>{verdict}</td>
                            </tr>
                          )
                        })}
                      </tbody>
                    </table>
                  </div>
                </div>

                {/* 계산 근거 */}
                <div style={{ ...cardStyle, marginTop: 16 }}>
                  <div style={secTitle}>계산 근거</div>
                  <div style={{ display: 'grid', gridTemplateColumns: isMobile ? '1fr' : '1fr 1fr', gap: 16 }}>
                    <div>
                      <div style={{ fontSize: 12, color: '#6b7280', marginBottom: 10, fontWeight: 700 }}>kWh당 순이익 산출</div>
                      {configs.map((cfg, i) => (
                        <div key={i} style={{ background: '#f9fafb', borderRadius: 8, padding: '10px 12px', marginBottom: 8, fontSize: 12 }}>
                          <div style={{ fontWeight: 700, color: C.primary, marginBottom: 6 }}>{cfg.label}</div>
                          <div style={{ color: '#6b7280', lineHeight: 1.8 }}>
                            충전단가 {cfg.rate}원<br/>
                            × (1 - PG {params.pg_fee_pct}%) × 운영사 {params.revenue_share_pct}%<br/>
                            = 수입 {rows[i]?.revenuePerKwh.toFixed(1)}원/kWh<br/>
                            - 전기요금 {rows[i]?.costPerKwh.toFixed(1)}원/kWh<br/>
                            <span style={{ fontWeight: 700, color: (rows[i]?.marginPerKwh ?? 0) > 0 ? '#16a34a' : '#dc2626' }}>
                              = 순이익 {rows[i]?.marginPerKwh.toFixed(1)}원/kWh
                            </span>
                          </div>
                        </div>
                      ))}
                    </div>
                    <div>
                      <div style={{ fontSize: 12, color: '#6b7280', marginBottom: 10, fontWeight: 700 }}>월 고정비용 항목</div>
                      <div style={{ background: '#f9fafb', borderRadius: 8, padding: '10px 12px', fontSize: 12 }}>
                        {[
                          { label: '운영비', val: params.monthly_ops },
                          { label: 'AS비', val: params.monthly_as },
                          { label: '통신비', val: params.monthly_comm },
                          { label: '전기안전관리대행비', val: params.monthly_elec_safety },
                          { label: '기타비용', val: params.monthly_other },
                          { label: `전기 기본료(${params.elec_basic_rate}원/kW × ${totalKw}kW)`, val: params.elec_basic_rate * totalKw },
                          { label: '보험료(월환산)', val: Math.round(params.insurance_yearly / 12) },
                          { label: `초기투자 월할당(÷${params.operation_months}개월)`, val: Math.round(rows.reduce((s,r)=>s+r.monthlyInitAmort,0)) },
                        ].filter(x => x.val > 0).map((x, i) => (
                          <div key={i} style={{ display: 'flex', justifyContent: 'space-between', padding: '4px 0', borderBottom: '1px solid #f0f0f0', color: '#374151' }}>
                            <span>{x.label}</span>
                            <span style={{ fontVariantNumeric: 'tabular-nums' }}>{x.val.toLocaleString()}원</span>
                          </div>
                        ))}
                        <div style={{ display: 'flex', justifyContent: 'space-between', padding: '6px 0 0', fontWeight: 700, color: C.primary, borderTop: '2px solid #e5e7eb', marginTop: 4 }}>
                          <span>합계</span>
                          <span style={{ fontVariantNumeric: 'tabular-nums' }}>{Math.round(totalMonthlyFixed).toLocaleString()}원</span>
                        </div>
                      </div>
                      <div style={{ marginTop: 10, background: '#eff6ff', borderRadius: 8, padding: '10px 12px', fontSize: 12, color: '#1d4ed8', lineHeight: 1.7 }}>
                        💡 <strong>이용률 판정 기준</strong><br/>
                        ✅ 20% 미만 — 달성 용이<br/>
                        🟡 20~50% — 보통 (현실적)<br/>
                        🟠 50~80% — 어려움 (고트래픽 필요)<br/>
                        🔴 80% 이상 — 사실상 24시간 만충
                      </div>
                    </div>
                  </div>
                </div>

                {/* 단가별 민감도 분석 */}
                <div style={{ ...cardStyle, marginTop: 16 }}>
                  <div style={secTitle}>충전 단가별 하루 필요 kWh (전체 합산)</div>
                  <div style={{ overflowX: 'auto' }}>
                    <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
                      <thead>
                        <tr style={{ background: '#f9fafb' }}>
                          <th style={{ padding: '8px 12px', textAlign: 'center', fontSize: 12, color: '#6b7280', borderBottom: '2px solid #e5e7eb' }}>단가(원/kWh)</th>
                          {[250,300,350,400,450,500,600,700].map(r => (
                            <th key={r} style={{ padding: '8px 12px', textAlign: 'center', fontSize: 12, color: '#6b7280', borderBottom: '2px solid #e5e7eb', whiteSpace: 'nowrap' }}>{r}원</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        <tr>
                          <td style={{ padding: '10px 12px', textAlign: 'center', fontWeight: 700, color: '#374151', background: '#f9fafb' }}>하루 필요량</td>
                          {[250,300,350,400,450,500,600,700].map(rate => {
                            const revPerKwh = rate * (1 - params.pg_fee_pct / 100) * (params.revenue_share_pct / 100)
                            const margin = revPerKwh - elecKwhRate
                            const bepDay = margin > 0 ? totalMonthlyFixed / margin / 30 : Infinity
                            const uPct = totalMaxKwhDay > 0 && bepDay < Infinity ? (bepDay / totalMaxKwhDay) * 100 : Infinity
                            const isCurrentRate = configs.every(c => c.rate === rate)
                            return (
                              <td key={rate} style={{
                                padding: '10px 12px', textAlign: 'center', fontVariantNumeric: 'tabular-nums',
                                fontWeight: isCurrentRate ? 800 : 400,
                                background: isCurrentRate ? '#ede9fe' : bepDay === Infinity ? '#fef2f2' : uPct < 20 ? '#f0fdf4' : uPct < 50 ? '#fefce8' : '#fff7ed',
                                color: bepDay === Infinity ? '#dc2626' : barColor(uPct),
                              }}>
                                {bepDay === Infinity ? '불가' : `${fmtKwh(bepDay)} kWh`}
                              </td>
                            )
                          })}
                        </tr>
                      </tbody>
                    </table>
                  </div>
                  <div style={{ fontSize: 11, color: '#9ca3af', marginTop: 8 }}>
                    * 현재 설정 단가와 일치하는 열은 보라색으로 표시됩니다. 배경색: 초록=달성 용이, 노랑=보통, 주황=어려움, 빨강=불가
                  </div>
                </div>
              </div>
            )
          })()}

          {tab === 'compare' && (() => {
            const rLump = runSimulation({ ...params, payment_type: '일시불' })
            const rInst = runSimulation({ ...params, payment_type: '할부' })
            const bepDiff = (rInst.bep_month ?? 0) - (rLump.bep_month ?? 0)
            const roiDiff = rInst.roi - rLump.roi
            const netDiff = rInst.total_net - rLump.total_net
            const avgDiff = rInst.avg_monthly_net - rLump.avg_monthly_net
            const rows = [
              { label: '초기 투자 (실납부)', lump: rLump.total_init_cost, inst: rInst.total_init_cost, unit: '원', fmt: (v: number) => v.toLocaleString('ko-KR'),
                hint: `할인율 ${params.discount_rate}% 적용으로 일시불 실납부액이 ${fmtM(rInst.total_init_cost - rLump.total_init_cost)}만원 절감됩니다.` },
              { label: '월 할부금', lump: 0, inst: rInst.monthly_installment, unit: '원', fmt: (v: number) => v > 0 ? v.toLocaleString('ko-KR') : '-',
                hint: `일시불 선택 시 매월 ${Math.round(rInst.monthly_installment).toLocaleString()}원의 할부 부담이 사라집니다.` },
              { label: '손익분기(BEP)', lump: rLump.bep_month, inst: rInst.bep_month, unit: '개월', fmt: (v: number | null) => v ? `${v}개월` : 'N/A',
                hint: bepDiff > 0 ? `일시불이 ${bepDiff}개월 더 빨리 흑자 전환됩니다.` : bepDiff < 0 ? `할부가 ${Math.abs(bepDiff)}개월 먼저 흑자 전환됩니다.` : '두 방식의 흑자 전환 시점이 동일합니다.' },
              { label: 'ROI', lump: rLump.roi, inst: rInst.roi, unit: '%', fmt: (v: number) => `${v.toFixed(1)}%`,
                hint: roiDiff < 0 ? `일시불의 투자 대비 수익률이 ${Math.abs(roiDiff).toFixed(1)}%p 높습니다.` : `할부의 ROI가 ${roiDiff.toFixed(1)}%p 높습니다.` },
              { label: '총 누적 순이익', lump: rLump.total_net, inst: rInst.total_net, unit: '원', fmt: (v: number) => `${fmtM(v)} 만원`,
                hint: netDiff < 0 ? `운영 기간 전체로 일시불이 ${fmtM(Math.abs(netDiff))}만원 더 수익이 납니다.` : `할부가 ${fmtM(netDiff)}만원 더 수익입니다.` },
              { label: '월 평균 순이익', lump: rLump.avg_monthly_net, inst: rInst.avg_monthly_net, unit: '원', fmt: (v: number) => `${fmtM(v)} 만원`,
                hint: avgDiff < 0 ? `일시불 선택 시 매월 평균 ${fmtM(Math.abs(avgDiff))}만원 더 수익이 납니다.` : `할부 선택 시 매월 평균 ${fmtM(avgDiff)}만원 더 수익입니다.` },
            ]
            const LUMP_C = '#6366F1'
            const INST_C = '#EC4899'
            // 누적 손익 비교 차트 데이터
            const chartData = rLump.records.map((rec, i) => ({
              month: rec.month,
              일시불: Math.round(rec.cumulative / 10000),
              할부: Math.round((rInst.records[i]?.cumulative ?? 0) / 10000),
            }))
            return (
              <div style={{ display: 'flex', flexDirection: 'column', gap: 14 }}>
                {/* 헤더 카드 */}
                <div style={{ display: 'grid', gridTemplateColumns: isMobile ? '1fr' : '1fr 1fr', gap: 10 }}>
                  {[
                    { label: '일시불', color: LUMP_C, desc: `할인율 ${params.discount_rate}% 적용`, init: rLump.total_init_cost, bep: rLump.bep_month, roi: rLump.roi, net: rLump.total_net },
                    { label: '할부', color: INST_C, desc: `월 ${Math.round(rInst.monthly_installment).toLocaleString()}원 × ${params.operation_months}개월`, init: rInst.total_init_cost, bep: rInst.bep_month, roi: rInst.roi, net: rInst.total_net },
                  ].map(c => (
                    <div key={c.label} style={{ background: 'white', borderRadius: 10, border: `2px solid ${c.color}`, padding: '16px 20px' }}>
                      <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 10 }}>
                        <div style={{ width: 4, height: 22, borderRadius: 2, background: c.color }}/>
                        <span style={{ fontSize: 16, fontWeight: 800, color: c.color }}>{c.label}</span>
                        <span style={{ fontSize: 10, color: '#9ca3af', marginLeft: 4 }}>{c.desc}</span>
                      </div>
                      {[
                        { l: '실 투자액', v: c.init.toLocaleString('ko-KR') + '원' },
                        { l: 'BEP', v: c.bep ? `${c.bep}개월차` : 'N/A' },
                        { l: 'ROI', v: `${c.roi.toFixed(1)}%` },
                        { l: '총 순이익', v: `${fmtM(c.net)} 만원` },
                      ].map(k => (
                        <div key={k.l} style={{ display: 'flex', justifyContent: 'space-between', fontSize: 12, marginBottom: 5 }}>
                          <span style={{ color: '#9ca3af' }}>{k.l}</span>
                          <span style={{ fontWeight: 700, color: c.color }}>{k.v}</span>
                        </div>
                      ))}
                    </div>
                  ))}
                </div>

                {/* 수치 비교 테이블 */}
                <div style={{ background: 'white', borderRadius: 10, border: '1px solid #e5e7eb', overflow: 'hidden' }}>
                  <div style={{ padding: '12px 18px', borderBottom: '1px solid #f3f4f6', display: 'flex', alignItems: 'center', gap: 8 }}>
                    <div style={{ width: 3, height: 14, borderRadius: 2, background: C.primary }}/>
                    <span style={{ fontSize: 12, fontWeight: 700, color: '#374151' }}>항목별 비교</span>
                  </div>
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
                    <thead>
                      <tr style={{ background: '#f9fafb' }}>
                        <th style={{ padding: '10px 16px', textAlign: 'left', color: '#6b7280', fontWeight: 600 }}>항목</th>
                        <th style={{ padding: '10px 16px', textAlign: 'right', color: LUMP_C, fontWeight: 700 }}>일시불</th>
                        <th style={{ padding: '10px 16px', textAlign: 'right', color: INST_C, fontWeight: 700 }}>할부</th>
                        <th style={{ padding: '10px 16px', textAlign: 'right', color: '#6b7280', fontWeight: 600 }}>차이 (할부-일시불)</th>
                        <th style={{ padding: '10px 16px', textAlign: 'left', color: '#6b7280', fontWeight: 600, minWidth: 160 }}>설명</th>
                      </tr>
                    </thead>
                    <tbody>
                      {rows.map((row, i) => {
                        const lv = typeof row.lump === 'number' ? row.lump : 0
                        const iv = typeof row.inst === 'number' ? row.inst : 0
                        const diff = iv - lv
                        const diffFmt = row.label === '손익분기(BEP)'
                          ? ((rInst.bep_month ?? 0) - (rLump.bep_month ?? 0) !== 0 ? `${(rInst.bep_month ?? 0) - (rLump.bep_month ?? 0)}개월` : '동일')
                          : row.label === '월 할부금' ? ''
                          : row.label === 'ROI' ? `${diff >= 0 ? '+' : ''}${diff.toFixed(1)}%`
                          : `${diff >= 0 ? '+' : ''}${fmtM(diff)} 만원`
                        const hintFavorLump = row.hint.includes('일시불') && (row.hint.includes('절감') || row.hint.includes('더 빨리') || row.hint.includes('높습니다') || row.hint.includes('더 수익') || row.hint.includes('사라집니다'))
                        return (
                          <tr key={row.label} style={{ borderTop: '1px solid #f3f4f6', background: i % 2 === 0 ? 'white' : '#fafafa' }}>
                            <td style={{ padding: '10px 16px', color: '#374151', fontWeight: 600 }}>{row.label}</td>
                            <td style={{ padding: '10px 16px', textAlign: 'right', color: LUMP_C, fontWeight: 600 }}>{row.fmt(row.lump as number)}</td>
                            <td style={{ padding: '10px 16px', textAlign: 'right', color: INST_C, fontWeight: 600 }}>{row.fmt(row.inst as number)}</td>
                            <td style={{ padding: '10px 16px', textAlign: 'right', color: diffFmt.startsWith('+') ? C.green : diffFmt === '동일' ? '#9ca3af' : C.red, fontWeight: 600 }}>{diffFmt}</td>
                            <td style={{ padding: '10px 16px', fontSize: 11, lineHeight: 1.4 }}>
                              {hintFavorLump
                                ? <span style={{ color: '#15803d' }}><span style={{ marginRight: 4 }}>🟢</span>{row.hint}</span>
                                : row.hint.includes('동일')
                                  ? <span style={{ color: '#6b7280' }}><span style={{ marginRight: 4 }}>⚪</span>{row.hint}</span>
                                  : <span style={{ color: '#b45309' }}><span style={{ marginRight: 4 }}>🟠</span>{row.hint}</span>
                              }
                            </td>
                          </tr>
                        )
                      })}
                    </tbody>
                  </table>
                </div>

                {/* 누적 손익 비교 차트 */}
                <div style={{ background: 'white', borderRadius: 10, border: '1px solid #e5e7eb', padding: '16px 20px' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 14 }}>
                    <div style={{ width: 3, height: 14, borderRadius: 2, background: C.primary }}/>
                    <span style={{ fontSize: 12, fontWeight: 700, color: '#374151' }}>누적 손익 비교 추이 (만원)</span>
                    <div style={{ display: 'flex', gap: 14, marginLeft: 'auto' }}>
                      {[{ label: '일시불', color: LUMP_C }, { label: '할부', color: INST_C }].map(l => (
                        <div key={l.label} style={{ display: 'flex', alignItems: 'center', gap: 5 }}>
                          <div style={{ width: 18, height: 3, background: l.color, borderRadius: 2 }}/>
                          <span style={{ fontSize: 10, color: '#6b7280' }}>{l.label}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                  <ResponsiveContainer width="100%" height={220}>
                    <ComposedChart data={chartData} margin={{ top: 4, right: 16, bottom: 4, left: 10 }}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0"/>
                      <XAxis dataKey="month" tick={{ fontSize: 10, fill: '#9ca3af' }} tickFormatter={(v: number) => `${v}M`}/>
                      <YAxis tick={{ fontSize: 10, fill: '#9ca3af' }} tickFormatter={(v: number) => `${v}만`}/>
                      <Tooltip formatter={(v: unknown) => [`${Number(v).toLocaleString()}만원`]} labelFormatter={(l: unknown) => `${l}개월`}/>
                      <ReferenceLine y={0} stroke="#e5e7eb" strokeWidth={1.5}/>
                      <Line type="monotone" dataKey="일시불" stroke={LUMP_C} strokeWidth={2} dot={false}/>
                      <Line type="monotone" dataKey="할부" stroke={INST_C} strokeWidth={2} dot={false}/>
                    </ComposedChart>
                  </ResponsiveContainer>
                </div>
              </div>
            )
          })()}

        </div>
    </div>
  )
}

// ── 루트 ─────────────────────────────────────────────────────
export default function SimDashboard() {
  const [params, setParamsState] = useState<SimParams>(DEFAULT_PARAMS)
  const [collapsed, setCollapsed] = useState(false)
  const [firstRec, setFirstRec] = useState<MonthRecord | null>(null)
  const [drawerOpen, setDrawerOpen] = useState(false)
  const isMobile = useIsMobile()
  const contentScrollRef = useRef<HTMLDivElement>(null)
  const setParams = (patch: Partial<SimParams>) => setParamsState(p => ({ ...p, ...patch }))

  return (
    <div style={{ display: 'flex', height: isMobile ? '100dvh' : 'calc(100vh / 1.2)', overflow: 'hidden', flexDirection: isMobile ? 'column' : 'row', zoom: isMobile ? 1 : 1.2 }}>
      {/* 데스크탑: 인플로우 사이드바 */}
      {!isMobile && (
        <SettingsSidebar params={params} setParams={setParams} collapsed={collapsed} setCollapsed={setCollapsed} firstRec={firstRec}/>
      )}
      {isMobile && drawerOpen && (
        <>
          {/* 배경 오버레이 — X 버튼으로만 닫기 */}
          <div style={{
            position: 'fixed', inset: 0, zIndex: 200,
            background: 'rgba(0,0,0,0.5)', backdropFilter: 'blur(2px)',
          }}/>
          {/* 하단 드로어 */}
          <div style={{
            position: 'fixed', bottom: 0, left: 0, right: 0, zIndex: 201,
            height: '85vh', borderRadius: '20px 20px 0 0', overflow: 'hidden',
            display: 'flex', flexDirection: 'column',
            background: `linear-gradient(180deg, ${C.sidebar1} 0%, ${C.sidebar2} 100%)`,
            boxShadow: '0 -8px 32px rgba(0,0,0,0.4)',
          }}>
            {/* 드로어 핸들 */}
            <div style={{ display: 'flex', justifyContent: 'center', padding: '12px 0 4px', flexShrink: 0 }}>
              <div style={{ width: 40, height: 4, borderRadius: 2, background: 'rgba(255,255,255,0.25)' }}/>
            </div>
            <SettingsSidebar params={params} setParams={setParams} collapsed={false} setCollapsed={() => setDrawerOpen(false)} firstRec={firstRec} onClose={() => setDrawerOpen(false)}/>
          </div>
        </>
      )}

      {/* 콘텐츠: 사이드바 닫히면 maxWidth로 중앙 정렬, 열리면 꽉 채움 */}
      <div ref={contentScrollRef} style={{
        flex: 1,
        minWidth: 0,
        minHeight: 0,
        overflowY: 'auto',
        overflowX: 'hidden',
        display: 'flex',
        flexDirection: 'column',
        background: '#f5f3ff',
      }}>
        <div style={{
          width: '100%',
          maxWidth: isMobile ? '100%' : 'min(1280px, calc(100vw / 1.2 - 225px))',
          marginLeft: 'auto',
          marginRight: 'auto',
          display: 'flex',
          flexDirection: 'column',
        }}>
          <MainContent params={params} setParams={setParams} onResult={r => setFirstRec(r ? r.records[0] : null)} isMobile={isMobile} scrollContainerRef={contentScrollRef}/>
        </div>
      </div>

      {/* 모바일 플로팅 설정 버튼 */}
      {isMobile && (
        <button onClick={() => setDrawerOpen(true)} style={{
          position: 'fixed', bottom: 20, right: 20, zIndex: 100,
          width: 54, height: 54, borderRadius: '50%', border: 'none',
          background: `linear-gradient(135deg, ${C.primary}, ${C.sidebar2})`,
          color: 'white', fontSize: 22, cursor: 'pointer',
          boxShadow: '0 4px 20px rgba(109,40,217,0.5)',
          display: 'flex', alignItems: 'center', justifyContent: 'center',
        }}>⚙️</button>
      )}
    </div>
  )
}
