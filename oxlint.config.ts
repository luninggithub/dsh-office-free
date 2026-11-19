import amamo from '@amamo/oxlint-config'

export default amamo({
  node: true,
  typescript: true,
  rules: {
    'eslint/no-underscore-dangle': 'off'
  }
})
