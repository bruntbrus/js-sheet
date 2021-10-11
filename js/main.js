(function main() {
  'use strict'

  const Sheet = window.Sheet
  const Model = window.Model
  const utils = window.utils

  // Create sheet
  const sheet = new Sheet({

    sideWidth: 64,
    headHeight: 24,
    model: new Model(26, 100, 120, 24),
    debug: true,
  })

  // Append sheet to main container and render
  sheet.appendTo(utils.query('#main'))

  // Export
  window.sheet = sheet
})()
