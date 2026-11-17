// The jsdom Client suite against the built bundle and a fake Host API. The Client no longer has
// Settings-backed preferences, so one suite covers every supported host shape.
import { defineClientSuite } from './client-suite.ts'

defineClientSuite()
