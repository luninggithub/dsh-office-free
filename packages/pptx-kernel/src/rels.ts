import { attribute, childElements } from './xml.ts'

export interface Relation {
  readonly type: string
  readonly target: string
}

export function readRelations(xml: string): Map<string, Relation> {
  const relations = new Map<string, Relation>()
  for (const element of childElements(xml)) {
    if (element.name !== 'Relationship' && element.name !== 'Relationships') continue
    const nested =
      element.name === 'Relationships'
        ? childElements(xml, element.openEnd, element.closeStart)
        : [element]
    for (const relation of nested) {
      if (relation.name !== 'Relationship') continue
      const id = attribute(relation.attrs, 'Id')
      const type = attribute(relation.attrs, 'Type')
      const target = attribute(relation.attrs, 'Target')
      if (id !== undefined && type !== undefined && target !== undefined) {
        relations.set(id, { type, target })
      }
    }
  }
  return relations
}

export function resolvePart(base: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1)
  const stack = base.split('/').slice(0, -1)
  for (const piece of target.split('/')) {
    if (piece === '..') stack.pop()
    else if (piece !== '.' && piece !== '') stack.push(piece)
  }
  return stack.join('/')
}
