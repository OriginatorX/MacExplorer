'use strict'

import XlSX from 'xlsx'
import { exist } from './utils.js'

export class Book {
    #workBookFile
    #workBook
    #sheetRows
    #insertions

    constructor(pathToFile) {
        this.#workBookFile = pathToFile
        this.#workBook = XlSX.readFile(pathToFile)
        this.#sheetRows = []
        this.#insertions = []
    }

    rowsFromSheet(name) {
        if (!exist(this.#workBook.SheetNames, name)) {
            throw new Error(`Sheet with name ${name} doesn't exist`)
        }
        for (const [key, val] of Object.entries(this.#workBook.Sheets[name])) {
            if (val?.v !== undefined) {
                this.#sheetRows.push(val?.v)
            }
        }
        return this.#sheetRows
    }

    addEntry(entry, sheetName) {
        this.#insertions.push([entry.id, entry.mac, entry.vendor])
        if (this.#sheetRows.length === this.#insertions.length) {
            this.#insertions.sort((a, b) => {
                if (a[0] < b[0]) {
                    return -1
                }
                if (a[0] > b[0]) {
                    return 1
                }
                return 0
            })
            const workSheet = XlSX.utils.aoa_to_sheet(this.#insertions)
            XlSX.utils.book_append_sheet(this.#workBook, workSheet, sheetName)
            XlSX.writeFile(this.#workBook, this.#workBookFile)
        }
    }
}
