(function() {
  'use strict'

  const Sheet = window.Sheet

  // Model class
  class Model {
    // Constructor
    constructor(colLength = 0, rowLength = 0, defaultColWidth = 0, defaultRowHeight = 0) {
      this.colLength = colLength
      this.rowLength = rowLength
      this.cols = []
      this.rows = []
      this.defaultCol = { width: defaultColWidth }
      this.defaultRow = { height: defaultRowHeight }
    }

    // Destroys model
    dispose() {
      this.cols = null
      this.rows = null
      this.defaultCol = null
      this.defaultRow = null
    }

    // Returns column data
    getColData(colNum, name) {
      const data = this.cols[colNum - 1] || this.defaultCol

      return data[name]
    }

    // Sets column data
    setColData(colNum, name, value) {
      let data = this.cols[colNum - 1]

      if (!data) {
        this.cols[colNum - 1] = data = {}
      }

      data[name] = value
    }

    // Returns row data
    getRowData(rowNum, name) {
      const data = this.rows[rowNum - 1] || this.defaultRow

      return data[name]
    }

    // Sets row data
    setRowData(rowNum, name, value) {
      let data = this.rows[rowNum - 1]

      if (!data) {
        this.rows[rowNum - 1] = data = {}
      }

      data[name] = value
    }

    // Returns cell data
    getCellData(colNum, rowNum, name) {
      // TODO: Return useful data

      return (name === 'content' ? Sheet.getCellName(colNum, rowNum) : '')
    }

    // Sets cell data
    setCellData(colNum, rowNum, name, value) {

      // TODO: Set data
    }
  }

  // Export
  window.Model = Model
})()
