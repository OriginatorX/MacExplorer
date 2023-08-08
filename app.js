'use strict'

import { Book } from './src/io.js'
import { API, argsHandle } from './src/utils.js'

(function app() {

    try {
        const { 
            file: fileName, 
            sheet: sheetName,
            newSheet: newSheetName
        } = argsHandle(process.argv)

        const apiClient = API()
        const book = new Book(fileName)
        const rows = book.rowsFromSheet(sheetName)

        book.addEntry({id: 0, mac: 'Mac-address', vendor: 'Vendor'}, newSheetName)
        
        let timeSkip = 100
        for (let i = 1; i < rows.length; i++) {
           setTimeout(function vendorResolve(idx) {
                apiClient.getMacInfo(rows[idx], (json) => {
                    const entry = {id: idx, mac: rows[idx], vendor: json['macInfo']['company']}
                    this.addEntry(entry, newSheetName)
                    console.log(`# ${entry.id}: ${entry.mac} - ${entry.vendor}`)
                })
            }.bind(book, i), timeSkip += 100)
        }    
    } catch (error) {
        console.log(error.message)
    }
})()
