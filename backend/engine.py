"""
EV 충전소 수익 시뮬레이션 계산 엔진 (UI 의존성 없음)
"""
from __future__ import annotations
from dataclasses import dataclass, field
from typing import Optional


CHARGER_TYPES = {
    "완속 7kW":     {"kw": 7,   "default_rate": 320},
    "중속 30kW":    {"kw": 30,  "default_rate": 350},
    "급속 50kW":    {"kw": 50,  "default_rate": 400},
    "급속 100kW":   {"kw": 100, "default_rate": 450},
    "초급속 200kW": {"kw": 200, "default_rate": 500},
}

ELEC_PRESET = {
    "저압": {"basic": 2390,  "kwh": 128.5, "climate": 9, "fuel": -5, "fund": 2.7},
    "고압": {"basic": 2580,  "kwh": 109.5, "climate": 9, "fuel": -5, "fund": 2.7},
}


@dataclass
class ChargerConfig:
    label: str
    kw: float
    count: int
    daily_ev: float
    avg_kwh: float
    rate: float


@dataclass
class SimParams:
    charger_configs: list[ChargerConfig]
    operation_months: int = 60
    ev_growth_rate: float = 20
    rate_increase: float = 3
    elec_type: str = "저압"
    elec_basic_rate: float = 2390
    elec_kwh_rate: float = 128.5
    elec_climate_rate: float = 9
    elec_fuel_rate: float = -5
    elec_fund_pct: float = 2.7
    elec_vat_pct: float = 10.0
    cost_charger_unit: int = 1_000_000
    cost_installation: int = 2_000_000
    cost_other_init: int = 0
    monthly_ops: int = 33_000
    monthly_as: int = 0
    monthly_other: int = 0
    pg_fee_pct: float = 2.0
    revenue_share_pct: float = 90


@dataclass
class MonthRecord:
    month: int
    year_label: str
    kwh: float
    gross_revenue: float
    pg_fee: float
    net_revenue: float
    my_revenue: float
    elec_cost: float
    elec_basic: float
    elec_usage: float
    elec_fund: float
    elec_vat: float
    ops: float
    as_cost: float
    other: float
    total_cost: float
    net_profit: float
    cumulative: float


@dataclass
class SimResult:
    records: list[MonthRecord]
    total_init_cost: float
    bep_month: Optional[int]
    total_net: float
    roi: float
    avg_monthly_net: float
    kwh_total: float
    charger_summary: str


def run_simulation(p: SimParams) -> SimResult:
    total_init_cost = p.cost_charger_unit * sum(c.count for c in p.charger_configs) \
                      + p.cost_installation + p.cost_other_init

    records: list[MonthRecord] = []
    cumulative = -total_init_cost

    for m in range(1, p.operation_months + 1):
        year = (m - 1) // 12
        gf = (1 + p.ev_growth_rate / 100) ** year
        rf = (1 + p.rate_increase / 100) ** year

        gross_revenue = 0.0
        total_kwh = 0.0
        contract_kw = 0.0

        for cfg in p.charger_configs:
            t_kwh = cfg.daily_ev * cfg.count * 30 * gf * cfg.avg_kwh
            gross_revenue += t_kwh * cfg.rate * rf
            total_kwh += t_kwh
            contract_kw += cfg.kw * cfg.count

        elec_basic    = contract_kw * p.elec_basic_rate
        elec_usage    = total_kwh * (p.elec_kwh_rate + p.elec_climate_rate + p.elec_fuel_rate)
        elec_sub      = elec_basic + elec_usage
        elec_fund     = elec_sub * (p.elec_fund_pct / 100)
        elec_vat      = (elec_sub + elec_fund) * (p.elec_vat_pct / 100)
        elec_cost     = elec_sub + elec_fund + elec_vat

        pg_fee        = gross_revenue * (p.pg_fee_pct / 100)
        net_revenue   = gross_revenue - pg_fee
        my_revenue    = net_revenue * (p.revenue_share_pct / 100)
        total_cost    = elec_cost + p.monthly_ops + p.monthly_as + p.monthly_other
        net_profit    = my_revenue - total_cost
        cumulative   += net_profit

        records.append(MonthRecord(
            month=m, year_label=f"{year+1}년차",
            kwh=round(total_kwh, 1),
            gross_revenue=round(gross_revenue),
            pg_fee=round(pg_fee),
            net_revenue=round(net_revenue),
            my_revenue=round(my_revenue),
            elec_cost=round(elec_cost),
            elec_basic=round(elec_basic),
            elec_usage=round(elec_usage),
            elec_fund=round(elec_fund),
            elec_vat=round(elec_vat),
            ops=p.monthly_ops,
            as_cost=p.monthly_as,
            other=p.monthly_other,
            total_cost=round(total_cost),
            net_profit=round(net_profit),
            cumulative=round(cumulative),
        ))

    bep_month = next((r.month for r in records if r.cumulative >= 0), None)
    total_net = sum(r.net_profit for r in records) - total_init_cost
    roi = (total_net / total_init_cost * 100) if total_init_cost > 0 else 0
    avg_monthly_net = sum(r.net_profit for r in records) / len(records) if records else 0
    kwh_total = sum(r.kwh for r in records)
    charger_summary = " + ".join(f"{c.label} ×{c.count}" for c in p.charger_configs)

    return SimResult(
        records=records,
        total_init_cost=total_init_cost,
        bep_month=bep_month,
        total_net=total_net,
        roi=roi,
        avg_monthly_net=avg_monthly_net,
        kwh_total=kwh_total,
        charger_summary=charger_summary,
    )
