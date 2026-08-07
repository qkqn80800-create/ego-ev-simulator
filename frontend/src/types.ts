export interface ChargerConfig {
  label: string
  kw: number
  count: number
  daily_ev: number
  avg_kwh: number
  rate: number
  vehicle_kwh?: number[] // 차량별 개별 충전량 (길이 = floor(daily_ev))
  cost_unit?: number          // 1기당 충전기 단가 (타입별 설정 시)
  cost_install?: number       // 1기당 설치비 (타입별 설정 시)
  monthly_ops_unit?: number   // 1기당 월 운영비
  monthly_ops_bulk?: number   // 2기 이상 시 1기당 월 운영비
}

export interface SimParams {
  charger_configs: ChargerConfig[]
  payment_type: '일시불' | '할부'
  discount_rate: number
  operation_months: number
  ev_growth_rate: number
  rate_increase: number
  elec_type: string
  elec_basic_rate: number
  elec_kwh_rate: number
  elec_climate_rate: number
  elec_fuel_rate: number
  elec_fund_pct: number
  elec_vat_pct: number
  cost_charger_unit: number
  cost_installation: number
  cost_other_init: number
  cost_kepco_burden: number       // 한전 시설부담금 (초기 1회)
  cost_safety_inspection: number  // 사용전검사·안전관리자·감리비 (초기 1회, 75kW 이상)
  monthly_ops: number
  monthly_as: number
  monthly_comm: number
  monthly_elec_safety: number     // 전기안전관리대행비 (월, 75kW 이상)
  monthly_tax_invoice: number     // 세금계산서 발행 수수료 (월)
  monthly_other: number
  pg_fee_pct: number
  revenue_share_pct: number
  manager_name: string
  manager_discount: number
}

export interface MonthRecord {
  month: number
  installment: number
  year_label: string
  kwh: number
  gross_revenue: number
  pg_fee: number
  net_revenue: number
  my_revenue: number
  elec_cost: number
  elec_basic: number
  elec_usage: number
  elec_fund: number
  elec_vat: number
  ops: number
  as_cost: number
  other: number
  total_cost: number
  net_profit: number
  cumulative: number
}

export interface SimResult {
  records: MonthRecord[]
  total_init_cost: number
  bep_month: number | null
  total_net: number
  roi: number
  avg_monthly_net: number
  kwh_total: number
  charger_summary: string
  monthly_installment: number
}

export const CHARGER_TYPES: Record<string, { kw: number; default_rate: number; default_cost_unit: number; default_cost_install: number; default_monthly_ops?: number }> = {
  '완속 7kW':     { kw: 7,   default_rate: 320, default_cost_unit: 1_100_000,  default_cost_install: 3_300_000,  default_monthly_ops: 33_000 },
  '중속 30kW':    { kw: 30,  default_rate: 350, default_cost_unit: 3_300_000,  default_cost_install: 5_500_000,  default_monthly_ops: 88_000 },
  '급속 50kW':    { kw: 50,  default_rate: 400, default_cost_unit: 11_000_000, default_cost_install: 7_700_000,  default_monthly_ops: 110_000 },
  '급속 100kW':   { kw: 100, default_rate: 450, default_cost_unit: 22_000_000, default_cost_install: 16_500_000, default_monthly_ops: 330_000 },
  '초급속 200kW': { kw: 200, default_rate: 500, default_cost_unit: 55_000_000, default_cost_install: 27_500_000, default_monthly_ops: 550_000 },
}

export const ELEC_PRESET: Record<string, { basic: number; kwh: number; climate: number; fuel: number; fund: number }> = {
  '저압': { basic: 2390,  kwh: 128.5, climate: 9, fuel: -5, fund: 2.7 },
  '고압': { basic: 2580,  kwh: 109.5, climate: 9, fuel: -5, fund: 2.7 },
}

export const DEFAULT_PARAMS: SimParams = {
  charger_configs: [
    { label: '완속 7kW', kw: 7, count: 1, daily_ev: 3, avg_kwh: 15, rate: 320 },
  ],
  payment_type: '일시불',
  discount_rate: 10,
  operation_months: 60,
  ev_growth_rate: 20,
  rate_increase: 3,
  elec_type: '저압',
  elec_basic_rate: 2390,
  elec_kwh_rate: 128.5,
  elec_climate_rate: 9,
  elec_fuel_rate: -5,
  elec_fund_pct: 2.7,
  elec_vat_pct: 10,
  cost_charger_unit: 1_100_000,
  cost_installation: 3_300_000,
  cost_other_init: 0,
  cost_kepco_burden: 0,
  cost_safety_inspection: 0,
  monthly_ops: 33_000,
  monthly_as: 0,
  monthly_comm: 5_000,
  monthly_elec_safety: 0,
  monthly_tax_invoice: 0,
  monthly_other: 0,
  pg_fee_pct: 2,
  revenue_share_pct: 90,
  manager_name: '',
  manager_discount: 0,
}
