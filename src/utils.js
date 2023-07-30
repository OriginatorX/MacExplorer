'use strict'

import APIClient from '@logocomune/maclookup'
import { parseArgs } from 'node:util'
import config from 'config'

export function API(key) {
    if (key) {
        new APIClient(key)
    } else if (config.has('apiKey')) {
        return new APIClient(config.get('apiKey'))
    } else {
        return new APIClient()
    }
}

export function exist(object, checkingElem) {
    for (let prop of object) {
        if (prop === checkingElem) {
            return true
        } 
    }
    return false
}

export function argsHandle(args) {
    const options = {
        file: {
            type: 'string',
            short: 'f'
        },
        sheet: {
            type: 'string',
            short: 's'
        }
    }
    const { values } = parseArgs({ args, options, allowPositionals: true })
    return values
}
