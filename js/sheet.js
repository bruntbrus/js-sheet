(function(utils) {
  'use strict'

  const colNames = makeColNames()
  // const maxColLength = colNames.length // 702
  // const maxRowLength = 9999
  const minColWidth = 48
  const minRowHeight = 24
  const maxColWidth = 480
  const maxRowHeight = 240

  // Sheet class
  class Sheet {
    // Constructor
    constructor(config = {}) {
      const sideWidth = config.sideWidth || 64
      const headHeight = config.headHeight || 24
      const model = config.model

      this.container = createSheetContainer()
      this.model = model
      this.selection = null

      this._sideWidth = 0
      this._headHeight = 0
      this._tableWidth = 0
      this._tableHeight = 0
      this._colLength = utils.limit(Math.ceil((screen.availWidth - sideWidth) / minColWidth), 1, model.colLength)
      this._rowLength = utils.limit(Math.ceil((screen.availHeight - headHeight) / minRowHeight), 1, model.rowLength)
      this._visColLength = 0
      this._visRowLength = 0
      this._leftColNum = 1
      this._topRowNum = 1
      this._debug = config.debug || false

      initHead(this)
      initSide(this)
      initBody(this)

      this.setSideWidth(sideWidth)
      this.setHeadHeight(headHeight)
      this.setColWidth()
      this.setRowHeight()

      initEvents(this)
    }

    // Destroys sheet
    dispose() {
      // TODO: Remove event listeners

      this.container = null
      this.model = null
      this.selection = null
    }

    // Returns parent element of container
    getParent() {
      return this.container.parentNode
    }

    // Returns head element
    getHead() {
      return utils.query('.sheet-head', this.container)
    }

    // Returns head table element
    getHeadTable() {
      return utils.query('table', this.getHead())
    }

    // Returns head row element
    getHeadRow() {
      return utils.query('tr', this.getHeadTable())
    }

    // Returns side element
    getSide() {
      return utils.query('.sheet-side', this.container)
    }

    // Returns side table element
    getSideTable() {
      return utils.query('table', this.getSide())
    }

    // Returns body element
    getBody() {
      return utils.query('.sheet-body', this.container)
    }

    // Returns body table element
    getBodyTable() {
      return utils.query('table', this.getBody())
    }

    // Returns filler element
    getFiller() {
      return utils.query('.sheet-filler', this.container)
    }

    // Sets side element width
    setSideWidth(width) {
      width = limitColWidth(width)

      if (width === this._sideWidth) {
        return false
      }

      utils.setLeft(this.getHead(), width)
      utils.setLeft(this.getBody(), width)
      utils.setWidth(this.getSideTable(), width)
      utils.setWidth(this.getFiller(), width)

      this._sideWidth = width

      return true
    }

    // Sets head element height
    setHeadHeight(height) {
      height = limitRowHeight(height)

      if (height === this._headHeight) {
        return false
      }

      utils.setTop(this.getSide(), height)
      utils.setTop(this.getBody(), height)
      utils.setHeight(this.getHeadRow(), height)
      utils.setHeight(this.getFiller(), height)

      this._headHeight = height

      return true
    }

    // Sets column width
    setColWidth(colNum = 0, width = 0) {
      if (utils.isString(colNum)) {
        colNum = getColNumByName(colNum)
      }

      const headTable = this.getHeadTable()
      const headRow = this.getHeadRow()
      const bodyTable = this.getBodyTable()
      const colLength = this._colLength
      const rowLength = this._rowLength
      const model = this.model

      let tableWidth

      if (colNum > 0) {
        const colWidth = getColWidth(colNum)

        if (width === colWidth) {
          return false
        }

        const headCell = getCell(headRow, colNum)

        utils.setWidth(headCell, width)

        eachRow(bodyTable, 1, rowLength, (bodyRow) => {
          const bodyCell = getCell(bodyRow, colNum)

          utils.setWidth(bodyCell, width)
        })

        tableWidth = this._tableWidth + width - colWidth
        model.setColData(colNum, 'width', width)
      } else {
        tableWidth = 0

        eachCell(headRow, 1, colLength, (headCell, colNum) => {
          const colWidth = getColWidth(colNum)

          utils.setWidth(headCell, colWidth)
          tableWidth += colWidth
        })

        eachRow(bodyTable, 1, rowLength, (bodyRow) => {
          eachCell(bodyRow, 1, colLength, (bodyCell, colNum) => {
            const colWidth = getColWidth(colNum)

            utils.setWidth(bodyCell, colWidth)
          })
        })
      }

      utils.setWidth(headTable, tableWidth)
      utils.setWidth(bodyTable, tableWidth)

      this._tableWidth = tableWidth

      function getColWidth(colNum) {
        return model.getColData(colNum, 'width')
      }

      return true
    }

    // Sets row height
    setRowHeight(rowNum = 0, height = 0) {
      const sideTable = this.getSideTable()
      const bodyTable = this.getBodyTable()
      // const colLength = this._colLength
      const rowLength = this._rowLength
      const model = this.model

      if (rowNum > 0) {
        const rowHeight = getRowHeight(rowNum)

        if (height === rowHeight) {
          return false
        }

        const sideRow = getRow(sideTable, rowNum)
        const bodyRow = getRow(bodyTable, rowNum)

        utils.setHeight(sideRow, height)
        utils.setHeight(bodyRow, height)

        model.setRowData(rowNum, 'height', height)
      } else {
        eachRow(bodyTable, 1, rowLength, (bodyRow, rowNum) => {
          const sideRow = getRow(sideTable, rowNum)
          const rowHeight = getRowHeight(rowNum)

          utils.setHeight(sideRow, rowHeight)
          utils.setHeight(bodyRow, rowHeight)
        })
      }

      function getRowHeight(rowNum) {
        return model.getRowData(rowNum, 'height')
      }

      return true
    }

    // Returns row element (if visible)
    getRow(rowNum) {
      return getRow(this.getBodyTable(), rowNum - this._topRowNum + 1)
    }

    // Returns cell element (if visible)
    getCell(colNum, rowNum) {
      const row = this.getRow(rowNum)

      return (row ? getCell(row, colNum - this._leftColNum + 1) : null)
    }

    // Returns cell element by name
    getCellByName(name) {
      const pos = getCellPositionByName(name)

      return this.getCell(pos.col, pos.row)
    }

    // Returns selected cell element (if visible)
    getSelectedCell() {
      const selection = this.selection

      return (selection ? this.getCell(selection.col, selection.row) : null)
    }

    // Appends container to parent element
    appendTo(parent) {
      parent.appendChild(this.container)
      setTableSize(this)
      renderSheet(this)
    }

    // Sets top left cell position
    moveTo(colNum = 0, rowNum = 0, update = true) {
      const model = this.model

      let moved = false
      let movedLeft = false
      let movedTop = false

      colNum = utils.limit(colNum, 1, model.colLength)
      rowNum = utils.limit(rowNum, 1, model.rowLength)

      if (colNum !== this._leftColNum) {
        moved = true
        movedLeft = true
      }

      if (rowNum !== this._topRowNum) {
        moved = true
        movedTop = true
      }

      if (moved) {
        let selectedCell = this.getSelectedCell()

        if (selectedCell) {
          deselectElement(selectedCell)
        }

        this._leftColNum = colNum
        this._topRowNum = rowNum

        selectedCell = this.getSelectedCell()

        if (selectedCell) {
          selectElement(selectedCell)
        }

        if (update) {
          setTableSize(this, movedLeft, movedTop)
          renderSheet(this)
        }
      }

      return moved
    }

    // Moves top left cell position
    moveBy(deltaColNum = 0, deltaRowNum = 0, render = true) {
      return this.moveTo(this._leftColNum + deltaColNum, this._topRowNum + deltaRowNum, render)
    }

    // Selects cell
    selectCell(colNum, rowNum) {
      const cell = this.getCell(colNum, rowNum)

      this.deselectCell()

      if (cell) {
        selectElement(cell)
      }

      this.selection = {
        col: colNum,
        row: rowNum,
      }
    }

    // Removes cell selection
    deselectCell() {
      const selection = this.selection

      if (selection) {
        const cell = this.getCell(selection.col, selection.row)

        if (cell) {
          deselectElement(cell)
        }

        this.selection = null
      }
    }
  }

  // Creates sheet container element
  function createSheetContainer() {
    return utils.element('div', {
      class: 'sheet',
      _html:
        '<header class="sheet-head">' +
          '<table></table>' +
        '</header>' +
        '<section class="sheet-side">' +
          '<table></table>' +
        '</section>' +
        '<section class="sheet-body">' +
          '<table></table>' +
        '</section>' +
        '<div class="sheet-filler"></div>',
    })
  }

  // Creates row element
  function createRow(num = 0) {
    const row = utils.element('tr')

    if (num > 0) {
      row.dataset.number = num
    }

    return row
  }

  // Creates cell element
  function createCell(num = 0, heading = false) {
    const cell = utils.element(heading ? 'th' : 'td')

    if (num > 0) {
      cell.dataset.number = num
    }

    cell.appendChild(utils.text())
    cell._cleared = true

    return cell
  }

  // Clears cell element
  function clearCell(cell) {
    if (!cell._cleared) {
      clearCellStyle(cell)
      cell.firstChild.textContent = ''
      cell._cleared = true
    }
  }

  // Sets cell style
  function setCellStyle(cell, data) {
    const style = cell.style

    if (data.color) {
      style.color = data.color
    }

    if (data.background) {
      style.backgroundColor = data.background
    }

    if (data.weight) {
      style.fontWeight = data.weight
    }

    if (data.align) {
      style.textAlign = data.align
    }
  }

  // Clears cell style
  function clearCellStyle(cell) {
    const style = cell.style

    if (style.color) {
      style.color = ''
    }

    if (style.backgroundColor) {
      style.backgroundColor = ''
    }

    if (style.fontWeight) {
      style.fontWeight = ''
    }

    if (style.textAlign) {
      style.textAlign = ''
    }
  }

  // Returns content of cell element
  /*
  function getCellContent(cell) {
    return cell.firstChild.textContent
  }
  */

  // Updates cell element
  function renderCell(cell, data) {
    let content

    if (utils.isObject(data)) {
      if (data.style) {
        setCellStyle(cell, data.style)
      }

      content = data.content
    } else {
      content = '' + data
    }

    cell.firstChild.textContent = content
    cell._cleared = false
  }

  // Initializes head element
  function initHead(sheet) {
    const table = sheet.getHeadTable()
    const row = createRow()
    const colLength = sheet._colLength

    for (let colNum = 1; colNum <= colLength; colNum++) {
      row.appendChild(createCell(colNum, true))
    }

    table.appendChild(row)
  }

  // Initializes side element
  function initSide(sheet) {
    const table = sheet.getSideTable()
    const rowLength = sheet._rowLength

    for (let rowNum = 1; rowNum <= rowLength; rowNum++) {
      const row = createRow(rowNum)

      row.appendChild(createCell(0, true))
      table.appendChild(row)
    }
  }

  // Initializes body element
  function initBody(sheet) {
    const table = sheet.getBodyTable()
    const rowLength = sheet._rowLength

    for (let rowNum = 1; rowNum <= rowLength; rowNum++) {
      const row = createRow(rowNum)
      const colLength = sheet._colLength

      for (let colNum = 1; colNum <= colLength; colNum++) {
        row.appendChild(createCell(colNum, false))
      }

      table.appendChild(row)
    }
  }

  // Initializes event listeners
  function initEvents(sheet) {
    utils.on(sheet.getBodyTable(), 'click', (event) => {
      const target = event.target

      if (target.matches('td')) {
        const logColNum = getLogicalColNum(sheet, getColNum(target))
        const logRowNum = getLogicalRowNum(sheet, getRowNum(target))

        sheet.selectCell(logColNum, logRowNum)
      }
    })

    utils.on(document.body, 'keydown', (event) => {
      switch (event.key) {
        case 'ArrowDown': sheet.moveBy(0, 1); break
        case 'ArrowUp': sheet.moveBy(0, -1); break
        case 'ArrowRight': sheet.moveBy(1, 0); break
        case 'ArrowLeft': sheet.moveBy(-1, 0); break
        case 'Escape': sheet.deselectCell(); break
        default:
      }
    })

    utils.on(sheet.container, 'wheel', (event) => {
      sheet.moveBy(0, 10 * Math.sign(event.deltaY))
    })

    utils.on(window, 'resize', utils.debounce(500, () => {
      setTableSize(sheet)
      renderSheet(sheet)
    }))
  }

  // Sets table size (visible columns and rows)
  function setTableSize(sheet, updateWidth = true, updateHeight = true) {
    const headTable = sheet.getHeadTable()
    const headRow = sheet.getHeadRow()
    const sideTable = sheet.getSideTable()
    const body = sheet.getBody()
    const bodyTable = sheet.getBodyTable()
    const colLength = sheet._colLength
    const rowLength = sheet._rowLength
    const model = sheet.model

    let width = 0
    let height = 0
    let visColLength = colLength
    let visRowLength = rowLength

    if (updateWidth) {
      const bodyWidth = utils.getWidth(body)

      eachCell(headRow, 1, colLength, (headCell, colNum) => {
        const logColNum = getLogicalColNum(sheet, colNum)

        width += model.getColData(logColNum, 'width')

        if (logColNum === model.colLength || width >= bodyWidth) {
          visColLength = colNum
          return true
        }
      })

      eachCell(headRow, 1, colLength, (headCell, colNum) => {
        if (colNum <= visColLength) {
          showElement(headCell)
        } else {
          hideElement(headCell)
          clearCell(headCell)
        }
      })

      sheet._visColLength = visColLength
    } else {
      visColLength = sheet._visColLength
    }

    if (updateHeight) {
      const bodyHeight = utils.getHeight(body)

      eachRow(bodyTable, 1, rowLength, (bodyRow, rowNum) => {
        const logRowNum = getLogicalRowNum(sheet, rowNum)

        height += model.getRowData(logRowNum, 'height')

        if (logRowNum === model.rowLength || height >= bodyHeight) {
          visRowLength = rowNum
          return true
        }
      })

      sheet._visRowLength = visRowLength
    } else {
      visRowLength = sheet._visRowLength
    }

    eachRow(bodyTable, 1, rowLength, (bodyRow, rowNum) => {
      if (updateHeight) {
        const sideRow = getRow(sideTable, rowNum)

        if (rowNum <= visRowLength) {
          showElement(sideRow)
          showElement(bodyRow)
        } else {
          hideElement(sideRow)
          hideElement(bodyRow)
        }
      }

      if (updateWidth) {
        eachCell(bodyRow, 1, colLength, (bodyCell, colNum) => {
          if (colNum <= visColLength) {
            showElement(bodyCell)
          } else {
            hideElement(bodyCell)
            clearCell(bodyCell)
          }
        })
      }
    })

    if (updateWidth && width !== sheet._tableWidth) {
      utils.setWidth(headTable, width)
      utils.setWidth(bodyTable, width)
      sheet._tableWidth = width
    }

    if (updateHeight && height !== sheet._tableHeight) {
      utils.setHeight(bodyTable, height)
      sheet._tableHeight = height
    }
  }

  // Updates cell elements
  function renderSheet(sheet) {
    let time = (sheet._debug ? performance.now() : 0)

    const headRow = sheet.getHeadRow()
    const sideTable = sheet.getSideTable()
    const bodyTable = sheet.getBodyTable()
    const model = sheet.model
    const colLength = Math.min(sheet._visColLength, model.colLength - sheet._leftColNum + 1)
    const rowLength = Math.min(sheet._visRowLength, model.rowLength - sheet._topRowNum + 1)

    eachCell(headRow, 1, colLength, (headCell, colNum) => {
      const logColNum = getLogicalColNum(sheet, colNum)

      renderCell(headCell, getColName(logColNum))
    })

    eachRow(bodyTable, 1, rowLength, (bodyRow, rowNum) => {
      const logRowNum = getLogicalRowNum(sheet, rowNum)
      const sideRow = getRow(sideTable, rowNum)
      const sideCell = getCell(sideRow, 1)

      renderCell(sideCell, logRowNum)

      eachCell(bodyRow, 1, colLength, (bodyCell, colNum) => {
        const logColNum = getLogicalColNum(sheet, colNum)

        clearCell(bodyCell)
        renderCell(bodyCell, model.getCellData(logColNum, logRowNum, 'content'))
      })
    })

    if (time) {
      time = utils.round3(performance.now() - time)
      console.log(`Render time: ${time} ms`)
    }
  }

  // Returns row element
  function getRow(table, rowNum) {
    return (table.children[rowNum - 1] || null)
  }

  // Returns cell element
  function getCell(row, colNum) {
    return (row.children[colNum - 1] || null)
  }

  // Returns column number of cell element
  function getColNum(cell) {
    const number = cell.dataset.number

    return (number ? parseInt(number, 10) : 0)
  }

  // Returns row number of cell element
  function getRowNum(cell) {
    const number = cell.parentNode.dataset.number

    return (number ? parseInt(number, 10) : 0)
  }

  // Returns logical column number
  function getLogicalColNum(sheet, colNum) {
    return (sheet._leftColNum + colNum - 1)
  }

  // Returns logical row number
  function getLogicalRowNum(sheet, rowNum) {
    return (sheet._topRowNum + rowNum - 1)
  }

  // Invokes callback for each row element
  function eachRow(table, fromRowNum, toRowNum, callback) {
    for (let rowNum = fromRowNum; rowNum <= toRowNum; rowNum++) {
      const row = getRow(table, rowNum)

      if (callback(row, rowNum)) {
        break
      }
    }
  }

  // Invokes callback for each cell element
  function eachCell(row, fromColNum, toColNum, callback) {
    for (let colNum = fromColNum; colNum <= toColNum; colNum++) {
      const cell = getCell(row, colNum)

      if (callback(cell, colNum)) {
        break
      }
    }
  }

  // Returns column name by number
  function getColName(colNum) {
    return (colNames[colNum - 1] || '')
  }

  // Returns column number by name
  function getColNumByName(colName) {
    return (colNames.indexOf(colName) + 1)
  }

  // Returns cell name by position
  function getCellName(colNum, rowNum) {
    return (getColName(colNum) + rowNum)
  }

  // Returns logical cell position by name
  function getCellPositionByName(name) {
    const matches = name.match(/^([A-Z]+)(\d+)$/)

    if (!matches) {
      return null
    }

    return {
      col: getColNumByName(matches[1]),
      row: parseInt(matches[2], 10),
    }
  }

  // Makes column names (A-ZZ, 702 names)
  function makeColNames() {
    const names = []
    const length = 26 // A-Z
    const first = 'A'.charCodeAt(0)

    for (let i = 0; i < length; i++) {
      names.push(toChar(i))
    }

    for (let i = 0; i < length; i++) {
      const name = names[i]

      for (let j = 0; j < length; j++) {
        names.push(name + toChar(j))
      }
    }

    function toChar(num) {
      return String.fromCharCode(first + num)
    }

    return names
  }

  // Returns column width within limited range
  function limitColWidth(width) {
    return utils.limit(width, minColWidth, maxColWidth)
  }

  // Returns row height within limited range
  function limitRowHeight(height) {
    return utils.limit(height, minRowHeight, maxRowHeight)
  }

  // Hides element
  function hideElement(element) {
    if (!element._hidden) {
      element.classList.add('hidden')
      element._hidden = true
    }
  }

  // Shows element
  function showElement(element) {
    if (element._hidden) {
      element.classList.remove('hidden')
      element._hidden = false
    }
  }

  // Selects element
  function selectElement(element) {
    if (!element._selected) {
      element.classList.add('selected')
      element._selected = true
    }
  }

  // Deselects element
  function deselectElement(element) {
    if (element._selected) {
      element.classList.remove('selected')
      element._selected = false
    }
  }

  // Export
  Sheet.getColName = getColName
  Sheet.getColNumByName = getColNumByName
  Sheet.getCellName = getCellName
  Sheet.getCellPositionByName = getCellPositionByName

  window.Sheet = Sheet
})(window.utils)
