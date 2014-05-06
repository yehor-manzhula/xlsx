class Sheet

  constructor:->
    @name = "Sheet"
    @data = []

  cellToNumber: (cell_string, ret_also_column) ->
    cellNumber = 0
    cellIndex = 0
    cellMax = cell_string.length
    rowId = 0

    # Converted from C++ (from DuckWriteC++):
    while cellIndex < cellMax
      curChar = cell_string.charCodeAt(cellIndex)
      if (curChar >= 0x30) and (curChar <= 0x39)
        rowId = parseInt(cell_string.slice(cellIndex), 10)
        rowId = (if (rowId > 0) then (rowId - 1) else 0)
        break
      else if (curChar >= 0x41) and (curChar <= 0x5A)
        if cellIndex > 0
          cellNumber++
          cellNumber *= (0x5B - 0x41)
        cellNumber += (curChar - 0x41)
      else if (curChar >= 0x61) and (curChar <= 0x7A)
        if cellIndex > 0
          cellNumber++
          cellNumber *= (0x5B - 0x41)
        cellNumber += (curChar - 0x61)
      cellIndex++
    if ret_also_column
      return (
        row: rowId
        column: cellNumber
      )
    cellNumber

  # Converts cell number to cell name
  # i.e. 3 to D
  numberToCell: (cell_number) ->
    outCell = ""
    curCell = cell_number

    while curCell >= 0
      outCell = String.fromCharCode((curCell % (0x5B - 0x41)) + 0x41) + outCell

      if curCell >= (0x5B - 0x41)
        curCell = Math.floor(curCell / (0x5B - 0x41)) - 1
      else
        break

    outCell

  setCell:(position, data)->
    cellPositionNumber = Sheet::cellToNumber position, true
    @data[cellPositionNumber.row] = []  unless @data[cellPositionNumber.row]
    @data[cellPositionNumber.row][cellPositionNumber.column] = data

  setRow:(number, value, style)->

module.exports = Sheet