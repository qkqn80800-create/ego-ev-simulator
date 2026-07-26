from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import Optional
from dataclasses import asdict
from engine import (
    run_simulation, SimParams, ChargerConfig,
    CHARGER_TYPES, ELEC_PRESET,
)

app = FastAPI(title="EV Simulation API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


# ── Request / Response 스키마 ─────────────────────────────────
class ChargerConfigIn(BaseModel):
    label: str
    kw: float
    count: int
    daily_ev: float
    avg_kwh: float
    rate: float


class SimRequest(BaseModel):
    charger_configs: list[ChargerConfigIn]
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


# ── 엔드포인트 ─────────────────────────────────────────────────
@app.get("/api/constants")
def get_constants():
    return {"charger_types": CHARGER_TYPES, "elec_preset": ELEC_PRESET}


@app.post("/api/simulate")
def simulate(req: SimRequest):
    params = SimParams(
        charger_configs=[ChargerConfig(**c.model_dump()) for c in req.charger_configs],
        **{k: v for k, v in req.model_dump().items() if k != "charger_configs"},
    )
    result = run_simulation(params)
    return {
        "records": [asdict(r) for r in result.records],
        "total_init_cost": result.total_init_cost,
        "bep_month": result.bep_month,
        "total_net": result.total_net,
        "roi": result.roi,
        "avg_monthly_net": result.avg_monthly_net,
        "kwh_total": result.kwh_total,
        "charger_summary": result.charger_summary,
    }
