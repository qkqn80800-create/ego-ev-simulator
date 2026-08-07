import type { SimParams, SimResult, MonthRecord } from './types'

export function runSimulation(p: SimParams): SimResult {
  const totalCount = p.charger_configs.reduce((s, c) => s + c.count, 0)
  const hasPerTypeCost = p.charger_configs.some(c => c.cost_unit !== undefined)
  // 충전기 구매·설치비: 할인율 적용 대상
  const chargerCost = hasPerTypeCost
    ? p.charger_configs.reduce((s, c) => s + ((c.cost_unit ?? p.cost_charger_unit) + (c.cost_install ?? 0)) * c.count, 0)
    : (p.cost_charger_unit + p.cost_installation) * totalCount
  // 외부 고정 비용: 할인율 미적용 (한전부담금, 사용전검사비, 부속시설물)
  const fixedExtraInit = (p.cost_other_init ?? 0) + (p.cost_kepco_burden ?? 0) + (p.cost_safety_inspection ?? 0)
  const totalInitCost = chargerCost + fixedExtraInit

  const isInstallment = p.payment_type === '할부'
  const instMonths = p.operation_months
  const monthlyInstallment = isInstallment && totalInitCost > 0 && instMonths > 0
    ? totalInitCost / instMonths
    : 0

  const managerDiscount = p.manager_name ? (p.manager_discount ?? 0) : 0
  // 할인율은 충전기 비용에만 적용, 외부 고정비는 전액 반영
  const discountedInitCost = isInstallment
    ? totalInitCost
    : chargerCost * (1 - (p.discount_rate ?? 0) / 100) - managerDiscount + fixedExtraInit

  const records: MonthRecord[] = []
  let cumulative = isInstallment ? 0 : -discountedInitCost

  for (let m = 1; m <= p.operation_months; m++) {
    const year = Math.floor((m - 1) / 12)
    const gf = Math.pow(1 + p.ev_growth_rate / 100, year)
    const rf = Math.pow(1 + p.rate_increase / 100, year)

    let grossRevenue = 0
    let totalKwh = 0
    let contractKw = 0

    for (const cfg of p.charger_configs) {
      const n = Math.floor(cfg.daily_ev)
      const dailyKwh = cfg.vehicle_kwh && cfg.vehicle_kwh.length === n
        ? cfg.vehicle_kwh.reduce((s, k) => s + k, 0)
        : cfg.daily_ev * cfg.avg_kwh
      const tKwh = dailyKwh * cfg.count * 30 * gf
      grossRevenue += tKwh * cfg.rate * rf
      totalKwh += tKwh
      contractKw += cfg.kw * cfg.count
    }

    const elecBasic = contractKw * p.elec_basic_rate
    const elecUsage = totalKwh * (p.elec_kwh_rate + p.elec_climate_rate + p.elec_fuel_rate)
    const elecSub = elecBasic + elecUsage
    const elecFund = elecSub * (p.elec_fund_pct / 100)
    const elecVat = (elecSub + elecFund) * (p.elec_vat_pct / 100)
    const elecCost = elecSub + elecFund + elecVat

    const pgFee = grossRevenue * (p.pg_fee_pct / 100)
    const netRevenue = grossRevenue - pgFee
    const myRevenue = netRevenue * (p.revenue_share_pct / 100)
    const hasPerTypeOps = p.charger_configs.some(c => c.monthly_ops_unit !== undefined)
    const perTypeOps = hasPerTypeOps
      ? p.charger_configs.reduce((s, c) => {
          if (c.monthly_ops_unit === undefined) return s
          const rate = (c.count >= 2 && c.monthly_ops_bulk !== undefined) ? c.monthly_ops_bulk : c.monthly_ops_unit
          return s + rate * c.count
        }, 0)
      : p.monthly_ops
    const instThisMonth = isInstallment && m <= instMonths ? monthlyInstallment : 0
    const totalCost = elecCost + perTypeOps + p.monthly_as + (p.monthly_comm ?? 5000) + (p.monthly_elec_safety ?? 0) + (p.monthly_insurance ?? 0) + p.monthly_other + instThisMonth
    const netProfit = myRevenue - totalCost
    cumulative += netProfit

    records.push({
      month: m,
      year_label: `${year + 1}년차`,
      kwh: Math.round(totalKwh * 10) / 10,
      gross_revenue: Math.round(grossRevenue),
      pg_fee: Math.round(pgFee),
      net_revenue: Math.round(netRevenue),
      my_revenue: Math.round(myRevenue),
      elec_cost: Math.round(elecCost),
      elec_basic: Math.round(elecBasic),
      elec_usage: Math.round(elecUsage),
      elec_fund: Math.round(elecFund),
      elec_vat: Math.round(elecVat),
      ops: perTypeOps,
      as_cost: p.monthly_as,
      other: p.monthly_other,
      installment: Math.round(instThisMonth),
      total_cost: Math.round(totalCost),
      net_profit: Math.round(netProfit),
      cumulative: Math.round(cumulative),
    })
  }

  const bepMonth = records.find(r => r.cumulative >= 0)?.month ?? null
  const totalNet = records.length > 0 ? records[records.length - 1].cumulative : 0
  const roi = discountedInitCost > 0 ? (totalNet / discountedInitCost) * 100 : 0
  const avgMonthlyNet = records.length > 0
    ? records.reduce((s, r) => s + r.net_profit, 0) / records.length
    : 0
  const kwhTotal = records.reduce((s, r) => s + r.kwh, 0)
  const chargerSummary = p.charger_configs.map(c => `${c.label} ×${c.count}`).join(' + ')

  return {
    records,
    total_init_cost: discountedInitCost,
    bep_month: bepMonth,
    total_net: totalNet,
    roi,
    avg_monthly_net: avgMonthlyNet,
    kwh_total: kwhTotal,
    charger_summary: chargerSummary,
    monthly_installment: Math.round(monthlyInstallment),
  }
}
