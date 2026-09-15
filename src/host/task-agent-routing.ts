import { closeSync, openSync, readSync, statSync } from 'node:fs'
import type { ResolvedTaskAgentModelConfig, ResolvedTaskAgentRoutingConfig, TaskAgentRoutingConfig } from './protocol.ts'

export type { ResolvedTaskAgentRoutingConfig, TaskAgentRoutingConfig }

export const DEFAULT_OFFICE_QUOTA_ADVICE_PATH = '/home/zdy/额度建议.md'
export const OFFICE_QUOTA_ADVICE_MAX_BYTES = 1_048_576

export type TaskAgentOperation =
  | 'write'
  | 'answer'
  | 'review'
  | 'placement'
  | 'migration'
  | 'compaction'
  | 'document-archive'
  | 'metadata-maintenance'

export type OfficeQuotaCandidateId = 'spark' | 'kimi' | 'composer' | 'luna' | 'grok' | 'flash'

export interface OfficeQuotaAdviceRow {
  product: string
  remaining: number
  todayPercent: number | null
  spent: string
}

export interface OfficeQuotaRoute {
  id: OfficeQuotaCandidateId
  provider: string
  model: string
}

interface OfficeQuotaCandidate {
  id: OfficeQuotaCandidateId
  provider: string
  model: string
  product?: 'spark' | 'kimi' | 'cursor' | 'codex' | 'grok'
  missing: 'eligible' | 'skip' | 'final'
}

const SPARK: OfficeQuotaCandidate = {
  id: 'spark',
  provider: 'openai-codex',
  model: 'gpt-5.3-codex-spark',
  product: 'spark',
  missing: 'eligible',
}

const COMPOSER: OfficeQuotaCandidate = {
  id: 'composer',
  provider: 'cursor-subscription',
  model: 'composer-2.5',
  product: 'cursor',
  missing: 'skip',
}

const LUNA: OfficeQuotaCandidate = {
  id: 'luna',
  provider: 'openai-codex',
  model: 'gpt-5.6-luna',
  product: 'codex',
  missing: 'skip',
}

const GROK: OfficeQuotaCandidate = {
  id: 'grok',
  provider: 'xai',
  model: 'grok-4.3',
  product: 'grok',
  missing: 'skip',
}

const KIMI: OfficeQuotaCandidate = {
  id: 'kimi',
  provider: 'kimi',
  model: 'k2.8',
  product: 'kimi',
  missing: 'skip',
}

const FLASH: OfficeQuotaCandidate = {
  id: 'flash',
  provider: 'deepseek-official',
  model: 'deepseek-v4-flash',
  missing: 'final',
}

const ANSWER_CANDIDATES: readonly OfficeQuotaCandidate[] = [SPARK, KIMI, COMPOSER, LUNA, GROK, FLASH]
const OTHER_CANDIDATES: readonly OfficeQuotaCandidate[] = [KIMI, LUNA, GROK, FLASH]
const STOPPED_SPENT = new Set(['已过线', '到线', '已过软线'])

let adviceCache: { path: string; mtimeMs: number; size: number; text: string } | undefined

export function resetOfficeQuotaAdviceCache(): void {
  adviceCache = undefined
}

export function resolveTaskAgentRouting(value: TaskAgentRoutingConfig | undefined): ResolvedTaskAgentRoutingConfig {
  const mode = value?.mode ?? 'default'
  if (mode !== 'default' && mode !== 'office-quota') throw new Error(`dsh-mnemon: unsupported task Agent routing mode: ${String(mode)}`)
  const advicePath = value?.advicePath?.trim()
  if (advicePath === undefined || advicePath === '') return { mode }
  if (advicePath.includes('\0')) throw new Error('dsh-mnemon: taskAgentRouting.advicePath must not contain a null byte')
  if (advicePath.length > 4_096) throw new Error('dsh-mnemon: taskAgentRouting.advicePath is too long')
  return { mode, advicePath }
}

export function resolveOfficeQuotaTaskAgentModel(
  routing: ResolvedTaskAgentRoutingConfig | undefined,
  operation: TaskAgentOperation = 'write',
): OfficeQuotaRoute | undefined {
  if (routing?.mode !== 'office-quota') return undefined
  return selectOfficeQuotaRoute(operation, routing.advicePath === undefined ? {} : { advicePath: routing.advicePath })
}

export function resolveConfiguredTaskAgentModel(
  config: { taskAgentModel: ResolvedTaskAgentModelConfig; taskAgentRouting: ResolvedTaskAgentRoutingConfig },
  operation: TaskAgentOperation = 'write',
): { provider: string; model: string; source: 'office-quota' | 'fixed' } | undefined {
  const officeQuota = resolveOfficeQuotaTaskAgentModel(config.taskAgentRouting, operation)
  if (officeQuota !== undefined) return { provider: officeQuota.provider, model: officeQuota.model, source: 'office-quota' }
  if (config.taskAgentModel.mode !== 'fixed') return undefined
  const provider = config.taskAgentModel.provider?.trim()
  const model = config.taskAgentModel.model?.trim()
  if (provider === undefined || provider === '' || model === undefined || model === '') return undefined
  return { provider, model, source: 'fixed' }
}

export function selectOfficeQuotaRoute(
  operation: TaskAgentOperation,
  options: { advicePath?: string; markdown?: string } = {},
): OfficeQuotaRoute {
  const markdown = options.markdown ?? readOfficeQuotaAdvice(options.advicePath)
  return selectOfficeQuotaCandidate(operation, parseAdviceTable(markdown))
}

export function parseAdviceTable(markdown: string): OfficeQuotaAdviceRow[] {
  const rows: OfficeQuotaAdviceRow[] = []
  for (const line of markdown.split(/\r?\n/)) {
    if (!line.startsWith('|')) continue
    const cells = line.split('|').map(cell => cell.trim()).filter((cell, index, all) => !(index === 0 && cell === '') && !(index === all.length - 1 && cell === ''))
    if (cells.length < 4) continue
    if (cells[0] === '顺序' || /^-+$/.test(cells[0] ?? '')) continue
    const remainingMatch = plainText(cells[3]).match(/(\d+(?:\.\d+)?)\s*%/)
    if (remainingMatch === null) continue
    rows.push({
      product: plainText(cells[1]) || '建议',
      remaining: Number(remainingMatch[1]),
      todayPercent: percentIn(cells[4]),
      spent: plainText(cells[5]),
    })
  }
  return rows
}

export function isAdviceStopped(spent: string, remaining: number, todayPercent: number | null): boolean {
  if (STOPPED_SPENT.has(plainText(spent))) return true
  if (Number.isFinite(remaining) && remaining <= 0) return true
  return Number.isFinite(remaining) && todayPercent !== null && Number.isFinite(todayPercent) && remaining <= todayPercent
}

export function selectOfficeQuotaCandidate(operation: TaskAgentOperation, rows: readonly OfficeQuotaAdviceRow[]): OfficeQuotaRoute {
  for (const candidate of operation === 'answer' ? ANSWER_CANDIDATES : OTHER_CANDIDATES) {
    if (candidate.missing === 'final' || isCandidateEligible(candidate, rows)) {
      return { id: candidate.id, provider: candidate.provider, model: candidate.model }
    }
  }
  return { id: FLASH.id, provider: FLASH.provider, model: FLASH.model }
}

export function readOfficeQuotaAdvice(configuredPath?: string): string {
  const path = resolveAdvicePath(configuredPath)
  try {
    const stat = statSync(path)
    if (adviceCache !== undefined && adviceCache.path === path && adviceCache.mtimeMs === stat.mtimeMs && adviceCache.size === stat.size) {
      return adviceCache.text
    }
    const text = readAdviceFile(path, stat.size)
    adviceCache = { path, mtimeMs: stat.mtimeMs, size: stat.size, text }
    return text
  } catch {
    return ''
  }
}

function isCandidateEligible(candidate: OfficeQuotaCandidate, rows: readonly OfficeQuotaAdviceRow[]): boolean {
  if (candidate.product === undefined) return true
  const row = rows.find(entry => productKey(entry.product) === candidate.product)
  if (row === undefined) return candidate.missing === 'eligible'
  return !isAdviceStopped(row.spent, row.remaining, row.todayPercent)
}

function productKey(value: string): OfficeQuotaCandidate['product'] | undefined {
  const normalized = value.toLowerCase().trim()
  if (normalized === 'spark' || normalized === 'codex-spark' || normalized === 'codex spark' || normalized === 'codex_bengalfox') return 'spark'
  if (normalized === 'cursor' || normalized.startsWith('cursor ')) return 'cursor'
  if (normalized === 'codex' || normalized.startsWith('codex ')) return 'codex'
  if (normalized === 'grok' || normalized.startsWith('grok ')) return 'grok'
  if (normalized === 'kimi' || normalized.startsWith('kimi')) return 'kimi'
  return undefined
}

function resolveAdvicePath(configuredPath?: string): string {
  const fromConfig = configuredPath?.trim()
  if (fromConfig !== undefined && fromConfig !== '') return fromConfig
  const fromEnv = process.env.DSH_MNEMON_ADVICE_PATH?.trim()
  if (fromEnv !== undefined && fromEnv !== '') return fromEnv
  return DEFAULT_OFFICE_QUOTA_ADVICE_PATH
}

function readAdviceFile(path: string, size: number): string {
  const length = Math.min(Math.max(0, size), OFFICE_QUOTA_ADVICE_MAX_BYTES)
  if (length === 0) return ''
  const fd = openSync(path, 'r')
  try {
    const buffer = Buffer.alloc(length)
    const read = readSync(fd, buffer, 0, length, 0)
    return buffer.subarray(0, read).toString('utf8')
  } finally {
    closeSync(fd)
  }
}

function percentIn(value: string | undefined): number | null {
  const match = plainText(value).match(/(\d+(?:\.\d+)?)\s*%/)
  return match === null ? null : Number(match[1])
}

function plainText(value: string | undefined): string {
  return (value ?? '').replace(/\*+/g, '').trim()
}
