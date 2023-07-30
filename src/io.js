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
        this.#insertions.push([entry.id, entry.val])
        if (this.#sheetRows.length === this.#insertions.length) {
            const workSheet = XlSX.utils.aoa_to_sheet(this.#insertions)
            XlSX.utils.book_append_sheet(this.#workBook, workSheet, sheetName)
            XlSX.writeFile(this.#workBook, this.#workBookFile)
        }
    }
}
