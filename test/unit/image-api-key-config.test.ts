import { afterEach, describe, expect, it } from 'vitest'
import type { Volatile } from '@deepseek-ai/cordis'
import { resolveConfig } from '../../src/host/config.ts'

describe('resolveConfig image api key', () => {
  const previous = process.env.ARK_API_KEY

  afterEach(() => {
    if (previous === undefined) delete process.env.ARK_API_KEY
    else process.env.ARK_API_KEY = previous
  })

  it('reads a plain key and follows a live getter', () => {
    delete process.env.ARK_API_KEY
    expect(resolveConfig({ imageApiKey: 'test-key' }).readImageApiKey()).toBe('test-key')

    let current: string | undefined = 'first'
    const live: Volatile<string | undefined> = { get: () => current } as Volatile<
      string | undefined
    >
    const resolved = resolveConfig({ imageApiKey: live })
    expect(resolved.readImageApiKey()).toBe('first')
    current = 'second'
    expect(resolved.readImageApiKey()).toBe('second')
    current = undefined
    expect(resolved.readImageApiKey()).toBeUndefined()
    process.env.ARK_API_KEY = 'from-env'
    expect(resolved.readImageApiKey()).toBe('from-env')
  })
})
