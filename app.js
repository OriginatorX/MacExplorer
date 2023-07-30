'use strict'

import { Book } from './src/io.js'
import { API, argsHandle } from './src/utils.js'

(function app() {

    try {
        const { file: fileName, sheet: sheetName } = argsHandle(process.argv)

        const newSheetName = 'Vendor'
        const apiClient = API()
        const book = new Book(fileName)
        const rows = book.rowsFromSheet(sheetName)

        book.addEntry({id: 'Mac-address', val: 'Vendor'}, newSheetName)
        
        let timeSkip = 100
        for (let i = 1; i < rows.length; i++) {
           setTimeout(function vendorResolve(idx) {
                apiClient.getMacInfo(rows[idx], (json) => {
                    const entry = {id: rows[idx], val: json['macInfo']['company']}
                    this.addEntry(entry, newSheetName)
                    console.log(`# ${idx}: ${entry.id} - ${entry.val}`)
                })
            }.bind(book, i), timeSkip += 100)
        }    
    } catch (error) {
        console.log(error.message)
    }
})()
