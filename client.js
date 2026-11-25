window.__ModuleLoader__.load({
  id: 'dsh-univer-office',
  factory: (require) => {
    const module = { exports: {} }
    const exports = module.exports
    const React = require('react')
    const { useEffect, useRef } = React

    function fileName(address) {
      const slash = address.lastIndexOf('/')
      const raw = slash < 0 ? address : address.slice(slash + 1)
      try {
        return decodeURIComponent(raw)
      } catch {
        return raw
      }
    }

    function copyBytes(bytes) {
      return bytes.slice().buffer
    }

    function PptxBody(props) {
      const frame = useRef(null)
      const content = props.content
      const bytes = content !== undefined && content.kind === 'bytes' ? content.data : undefined
      const name = fileName(props.resourceAddress)
      useEffect(() => {
        const iframe = frame.current
        if (iframe === null || bytes === undefined) return undefined
        const send = () => {
          const target = iframe.contentWindow
          if (target === null) return
          target.postMessage(
            { type: 'dsh-office-open', name, bytes: copyBytes(bytes) },
            location.origin,
          )
        }
        const onReady = (event) => {
          if (event.source !== iframe.contentWindow || event.origin !== location.origin) return
          if (event.data?.type !== 'dsh-office-ready') return
          send()
        }
        window.addEventListener('message', onReady)
        iframe.addEventListener('load', send)
        if (iframe.contentDocument?.readyState === 'complete') send()
        return () => {
          window.removeEventListener('message', onReady)
          iframe.removeEventListener('load', send)
        }
      }, [bytes, name])
      return React.createElement(
        'div',
        {
          'data-pptx-preview': '',
          style: {
            display: 'flex',
            width: '100%',
            height: '100%',
            minWidth: 0,
            minHeight: 0,
            background: '#fff',
          },
        },
        bytes === undefined
          ? React.createElement('p', { style: { margin: '12px', color: '#667085' } }, '正在打开演示文稿')
          : React.createElement('iframe', {
            ref: frame,
            title: name,
            src: '/office/?embed=1',
            style: { display: 'block', flex: '1 1 auto', width: '100%', height: '100%', border: 'none', background: '#fff' },
          }),
      )
    }

    const id = 'dsh-univer-office/pptx'

    function apply(ctx) {
      ctx.effect(() => ctx.documentPreviews.register({
        id,
        extensions: ['pptx'],
        binaryExtensions: ['pptx'],
        priority: 'extension',
        title: () => 'PPT',
        loading: 'bytes-complete',
        wrap: false,
      }))
      ctx.effect(() => ctx.slots.inject('sidebar.right.tab.document', () => ctx.slots.register({
        name: 'sidebar.right.tab.document',
        key: id,
      }, PptxBody)))
    }

    exports.name = 'dsh-univer-office'
    exports.inject = ['documentPreviews', 'slots']
    exports.apply = apply
    return module.exports
  },
})
