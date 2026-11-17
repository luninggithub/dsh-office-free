import { useState } from 'react'
import type { CSSProperties, FormEvent, ReactElement } from 'react'
import type { PropsLocale, PropsRuntime } from '@deepseek-ai/dsh-client-ui-slots'

/**
 * The Volcengine key for `univer_image_edit`, edited on this plugin's row in the DSH plugin list.
 *
 * The Host keeps the key out of the form value, so the page only learns whether one is already
 * stored. A blank field saves nothing and leaves that key in place.
 */
export type ImageApiKeyFormProps = PropsRuntime<'plugins.row.config'> &
  PropsLocale<'univer'> & {
    readonly apiKeyConfigured: boolean
  }

export function ImageApiKeyForm(props: ImageApiKeyFormProps): ReactElement {
  const { t, view, form, apiKeyConfigured } = props
  const [draft, setDraft] = useState('')
  const [saving, setSaving] = useState(false)
  const [notice, setNotice] = useState<'saved' | 'failed' | null>(null)
  if (view === 'summary') return <span>{t('plugin.apiKeySummary')}</span>
  if (form === undefined || form.state.status === 'unavailable') {
    return <p style={hintStyle}>{t('plugin.apiKeyUnavailable')}</p>
  }
  const ready = form.state.status === 'ready'
  const disabled = !ready || form.state.writable !== true || saving
  const save = (event: FormEvent<HTMLFormElement>): void => {
    event.preventDefault()
    if (form.state.revision === undefined) return
    const key = draft.trim()
    if (key.length === 0) return
    setSaving(true)
    setNotice(null)
    void settle(
      form.mutate([{ op: 'set', path: ['imageApiKey'], value: key }], form.state.revision),
      true
    )
  }
  const clear = (): void => {
    if (form.state.revision === undefined) return
    setSaving(true)
    setNotice(null)
    void settle(form.mutate([{ op: 'unset', path: ['imageApiKey'] }], form.state.revision), false)
  }
  const settle = async (result: Promise<boolean>, clearDraft: boolean): Promise<void> => {
    try {
      const accepted = await result
      setSaving(false)
      if (!accepted) {
        setNotice('failed')
        return
      }
      if (clearDraft) setDraft('')
      setNotice('saved')
    } catch {
      // The settings write can fail after the page has already closed the form.
      setSaving(false)
      setNotice('failed')
    }
  }
  return (
    <form onSubmit={save} style={formStyle}>
      <div style={headStyle}>
        <label htmlFor="univer-image-api-key" style={labelStyle}>
          {t('plugin.apiKeyLabel')}
        </label>
        <span style={stateStyle}>
          {t(apiKeyConfigured ? 'plugin.apiKeySet' : 'plugin.apiKeyUnset')}
        </span>
      </div>
      <input
        id="univer-image-api-key"
        type="password"
        autoComplete="new-password"
        value={draft}
        disabled={disabled}
        placeholder={t('plugin.apiKeyHint')}
        style={inputStyle}
        onChange={(event) => {
          setDraft(event.target.value)
          setNotice(null)
        }}
      />
      <div style={actionsStyle}>
        <button type="submit" disabled={disabled || draft.trim().length === 0} style={saveStyle}>
          {t(saving ? 'plugin.apiKeySaving' : 'plugin.apiKeySave')}
        </button>
        {apiKeyConfigured && (
          <button type="button" disabled={disabled} style={clearStyle} onClick={clear}>
            {t('plugin.apiKeyClear')}
          </button>
        )}
        {notice !== null && (
          <span style={noticeStyle}>
            {t(notice === 'saved' ? 'plugin.apiKeySaved' : 'plugin.apiKeyFailed')}
          </span>
        )}
      </div>
    </form>
  )
}

const formStyle: CSSProperties = { display: 'flex', flexDirection: 'column', gap: 8, maxWidth: 420 }
const headStyle: CSSProperties = {
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'space-between',
  gap: 12
}
const labelStyle: CSSProperties = { color: '#00141A', fontSize: 14, lineHeight: '22px' }
const stateStyle: CSSProperties = {
  color: 'rgba(0, 20, 26, 0.45)',
  fontSize: 12,
  lineHeight: '20px'
}
const inputStyle: CSSProperties = {
  height: 36,
  padding: '0 12px',
  border: '1px solid rgba(0, 20, 26, 0.08)',
  borderRadius: 8,
  fontSize: 14,
  color: '#00141A'
}
const actionsStyle: CSSProperties = { display: 'flex', alignItems: 'center', gap: 8 }
const saveStyle: CSSProperties = {
  height: 32,
  padding: '0 14px',
  border: 'none',
  borderRadius: 16,
  background: '#00141A',
  color: '#fff',
  fontSize: 13,
  cursor: 'pointer'
}
const clearStyle: CSSProperties = {
  height: 32,
  padding: '0 14px',
  border: '1px solid rgba(0, 20, 26, 0.08)',
  borderRadius: 16,
  background: '#fff',
  color: '#00141A',
  fontSize: 13,
  cursor: 'pointer'
}
const noticeStyle: CSSProperties = { color: 'rgba(0, 20, 26, 0.7)', fontSize: 12 }
const hintStyle: CSSProperties = { margin: 0, color: 'rgba(0, 20, 26, 0.45)', fontSize: 12 }
