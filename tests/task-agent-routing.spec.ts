import { mkdtempSync, rmSync, statSync, utimesSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import {
  isAdviceStopped,
  OFFICE_QUOTA_ADVICE_MAX_BYTES,
  parseAdviceTable,
  readOfficeQuotaAdvice,
  resetOfficeQuotaAdviceCache,
  resolveConfiguredTaskAgentModel,
  resolveTaskAgentRouting,
  selectOfficeQuotaCandidate,
  selectOfficeQuotaRoute,
} from '../src/host/task-agent-routing.ts'

const directories: string[] = []

afterEach(() => {
  resetOfficeQuotaAdviceCache()
  delete process.env.DSH_MNEMON_ADVICE_PATH
  for (const directory of directories.splice(0)) rmSync(directory, { recursive: true, force: true })
})

function adviceFile(markdown: string): string {
  const directory = mkdtempSync(join(tmpdir(), 'dsh-mnemon-advice-'))
  directories.push(directory)
  const path = join(directory, 'advice.md')
  writeFileSync(path, markdown)
  return path
}

function table(rows: Array<[string, string, string, string]>): string {
  return [
    '| 顺序 | 产品 | 设备 | 现在还剩 | 今天用到 | 大约花掉 | 重置 |',
    '| --- | --- | --- | ---: | ---: | ---: | --- |',
    ...rows.map(([product, remaining, today, spent], index) => `| ${index + 1} | ${product} | 全设备 | ${remaining} | ${today} | ${spent} | 9/15 |`),
  ].join('\n')
}

describe('office quota advice parser', () => {
  it('parses markdown table rows and applies legacy stop semantics', () => {
    const rows = parseAdviceTable(table([
      ['Cursor', '**4%**', '**6%** 软', '已过软线'],
      ['Grok', '6%', '5%', '1 点'],
      ['Codex', '0%', '62%', '3 点'],
      ['Spark', '80%', '10%', '到线'],
    ]))
    expect(rows).toEqual([
      { product: 'Cursor', remaining: 4, todayPercent: 6, spent: '已过软线' },
      { product: 'Grok', remaining: 6, todayPercent: 5, spent: '1 点' },
      { product: 'Codex', remaining: 0, todayPercent: 62, spent: '3 点' },
      { product: 'Spark', remaining: 80, todayPercent: 10, spent: '到线' },
    ])
    expect(isAdviceStopped('已过线', 80, 10)).toBe(true)
    expect(isAdviceStopped('到线', 80, 10)).toBe(true)
    expect(isAdviceStopped('已过软线', 4, 6)).toBe(true)
    expect(isAdviceStopped('1 点', 0, 10)).toBe(true)
    expect(isAdviceStopped('1 点', 10, 10)).toBe(true)
    expect(isAdviceStopped('1 点', 11, 10)).toBe(false)
  })

  it('reads at most 1 MiB and reuses a small mtime cache', () => {
    const grok = '| 顺序 | 产品 | 设备 | 现在还剩 | 今天用到 | 大约花掉 | 重置 |\n| --- | --- | --- | ---: | ---: | ---: | --- |\n| 1 | Grok | 全设备 | 90% | 10% | 1 点 | 9/15 |\n'
    const path = adviceFile(grok)
    const original = statSync(path)
    expect(readOfficeQuotaAdvice(path)).toContain('Grok')
    writeFileSync(path, grok)
    utimesSync(path, new Date(original.atimeMs), new Date(original.mtimeMs))
    expect(readOfficeQuotaAdvice(path)).toContain('Grok')
    writeFileSync(path, '| 顺序 | 产品 | 设备 | 现在还剩 | 今天用到 | 大约花掉 | 重置 |\n| --- | --- | --- | ---: | ---: | ---: | --- |\n| 1 | Codex | 全设备 | 90% | 10% | 1 点 | 9/15 |\n')
    utimesSync(path, new Date(original.atimeMs + 2_000), new Date(original.mtimeMs + 2_000))
    expect(readOfficeQuotaAdvice(path)).toContain('Codex')

    const oversized = adviceFile(`${'x'.repeat(OFFICE_QUOTA_ADVICE_MAX_BYTES)}| 1 | Codex | 全设备 | 90% | 10% | 1 点 | 9/15 |\n`)
    expect(readOfficeQuotaAdvice(oversized)).toHaveLength(OFFICE_QUOTA_ADVICE_MAX_BYTES)
    expect(readOfficeQuotaAdvice(oversized)).not.toContain('Codex')
  })
})

describe('office quota candidate selection', () => {
  it('uses Spark, Composer, Luna, Grok, then Flash for answer', () => {
    expect(selectOfficeQuotaCandidate('answer', parseAdviceTable(table([
      ['Cursor', '4%', '6%', '已过软线'],
      ['Codex', '0%', '10%', '1 点'],
      ['Grok', '0%', '10%', '1 点'],
    ])))).toMatchObject({ id: 'spark', provider: 'openai-codex', model: 'gpt-5.3-codex-spark' })
    expect(selectOfficeQuotaCandidate('answer', parseAdviceTable(table([
      ['Codex Spark', '1%', '10%', '到线'],
      ['Cursor', '80%', '10%', '1 点'],
    ])))).toMatchObject({ id: 'composer', provider: 'cursor-subscription', model: 'composer-2.5' })
    expect(selectOfficeQuotaCandidate('answer', parseAdviceTable(table([
      ['Spark', '1%', '10%', '已过线'],
      ['Cursor', '4%', '6%', '已过软线'],
      ['Codex', '90%', '10%', '1 点'],
    ])))).toMatchObject({ id: 'luna', provider: 'openai-codex', model: 'gpt-5.6-luna' })
    expect(selectOfficeQuotaCandidate('answer', parseAdviceTable(table([
      ['Spark', '1%', '10%', '到线'],
      ['Cursor', '4%', '6%', '已过软线'],
      ['Codex', '10%', '10%', '1 点'],
      ['Grok', '20%', '5%', '1 点'],
    ])))).toMatchObject({ id: 'grok', provider: 'xai', model: 'grok-4.3' })
    expect(selectOfficeQuotaCandidate('answer', parseAdviceTable(table([
      ['Spark', '1%', '10%', '到线'],
      ['Cursor', '4%', '6%', '已过软线'],
      ['Codex', '0%', '10%', '1 点'],
      ['Grok', '5%', '10%', '1 点'],
    ])))).toMatchObject({ id: 'flash', provider: 'deepseek-official', model: 'deepseek-v4-flash' })
  })

  it('uses Luna, Grok, then Flash for every other existing operation', () => {
    expect(selectOfficeQuotaCandidate('write', parseAdviceTable(table([
      ['Cursor', '90%', '10%', '1 点'],
      ['Spark', '90%', '10%', '1 点'],
      ['Codex', '90%', '10%', '1 点'],
    ])))).toMatchObject({ id: 'luna', provider: 'openai-codex', model: 'gpt-5.6-luna' })
    expect(selectOfficeQuotaCandidate('review', parseAdviceTable(table([
      ['Codex', '10%', '10%', '1 点'],
      ['Grok', '20%', '5%', '1 点'],
    ])))).toMatchObject({ id: 'grok', provider: 'xai', model: 'grok-4.3' })
    expect(selectOfficeQuotaCandidate('placement', [])).toMatchObject({ id: 'flash' })
    expect(selectOfficeQuotaCandidate('migration', parseAdviceTable(table([
      ['Codex', '0%', '10%', '1 点'],
      ['Grok', '0%', '10%', '1 点'],
    ])))).toMatchObject({ id: 'flash', provider: 'deepseek-official', model: 'deepseek-v4-flash' })
  })

  it('keeps a missing Spark row eligible, skips missing Cursor/Codex/Grok rows, and always finishes on DeepSeek', () => {
    expect(selectOfficeQuotaCandidate('answer', parseAdviceTable(table([
      ['Cursor', '4%', '6%', '已过软线'],
      ['Codex', '0%', '10%', '1 点'],
      ['Grok', '0%', '10%', '1 点'],
    ])))).toMatchObject({ id: 'spark' })
    expect(selectOfficeQuotaCandidate('answer', parseAdviceTable(table([
      ['Spark', '1%', '10%', '到线'],
      ['Codex', '0%', '10%', '1 点'],
      ['Grok', '0%', '10%', '1 点'],
    ])))).toMatchObject({ id: 'flash' })
    expect(selectOfficeQuotaCandidate('write', parseAdviceTable(table([
      ['Cursor', '90%', '10%', '1 点'],
    ])))).toMatchObject({ id: 'flash' })
  })

  it('reads DSH_MNEMON_ADVICE_PATH when advicePath is omitted', () => {
    process.env.DSH_MNEMON_ADVICE_PATH = adviceFile(table([
      ['Spark', '1%', '10%', '到线'],
      ['Cursor', '80%', '10%', '1 点'],
    ]))
    expect(selectOfficeQuotaRoute('answer')).toMatchObject({ id: 'composer' })
  })
})

describe('configured task Agent resolver', () => {
  it('leaves inherit/fixed defaults unchanged until office-quota is opted in', () => {
    expect(resolveTaskAgentRouting(undefined)).toEqual({ mode: 'default' })
    expect(resolveConfiguredTaskAgentModel({
      taskAgentModel: { mode: 'inherit' },
      taskAgentRouting: { mode: 'default' },
    })).toBeUndefined()
    expect(resolveConfiguredTaskAgentModel({
      taskAgentModel: { mode: 'fixed', provider: 'openai', model: 'gpt-5' },
      taskAgentRouting: { mode: 'default' },
    })).toEqual({ provider: 'openai', model: 'gpt-5', source: 'fixed' })
  })

  it('selects office-quota once per operation even when a fixed model is also configured', () => {
    const markdown = table([
      ['Spark', '1%', '10%', '到线'],
      ['Cursor', '80%', '10%', '1 点'],
      ['Codex', '90%', '10%', '1 点'],
    ])
    const path = adviceFile(markdown)
    const routing = resolveTaskAgentRouting({ mode: 'office-quota', advicePath: path })
    const config = {
      taskAgentModel: { mode: 'fixed' as const, provider: 'openai', model: 'gpt-5' },
      taskAgentRouting: routing,
    }
    expect(resolveConfiguredTaskAgentModel(config, 'answer')).toEqual({
      provider: 'cursor-subscription',
      model: 'composer-2.5',
      source: 'office-quota',
    })
    expect(resolveConfiguredTaskAgentModel(config, 'write')).toEqual({
      provider: 'openai-codex',
      model: 'gpt-5.6-luna',
      source: 'office-quota',
    })
  })
})
