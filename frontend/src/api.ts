import type { SimParams, SimResult } from './types'
import { runSimulation } from './engine'

export async function simulate(params: SimParams): Promise<SimResult> {
  return runSimulation(params)
}
