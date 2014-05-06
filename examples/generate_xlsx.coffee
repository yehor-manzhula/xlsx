Xlsx = require '../lib/xlsx'

# Create document
xlsxDocument = new Xlsx()
xlsxDocument.creator = "XLSX Generator"

# First sheet
sheet = xlsxDocument.createSheet()

sheet.name = "Excel 1"

sheet.data[0] = []
sheet.data[0][0] = 1
sheet.data[1] = []
sheet.data[1][3] = {value:"abc", style:"BOLD"}
sheet.data[1][4] = {value:"More", style: "bOld"}
sheet.data[1][5] = "Text fffdaffdfdadfasdfasdfaf"
sheet.data[1][6] = "Here"
sheet.data[2] = []
sheet.data[2][5] = "abc"
sheet.data[2][6] = 900
sheet.data[6] = []
sheet.data[6][2] = 1972

sheet.setCell "E7", 340
sheet.setCell "I1", -3
sheet.setCell "I2", 31.12
sheet.setCell "G102", {value:"Hello World!", style:"bold"}

# Second sheet
sheet2 = xlsxDocument.createSheet()

sheet2.name = "Excel 2"

# The direct option - two-dimensional array:
sheet2.data[0] = []
sheet2.data[0][0] = 1
sheet2.data[1] = []
sheet2.data[1][3] =
  value: "abc"
  style: "BOLD"

sheet2.data[1][4] =
  value: "More 1"
  style: "bOld"

sheet2.data[1][5] = "Text 2"
sheet2.data[1][6] = "Here 2"
sheet2.data[2] = []
sheet2.data[2][5] = "abc 4"
sheet2.data[2][6] = 900444
sheet2.data[6] = []
sheet2.data[6][2] = 19724213

# Using setCell:
sheet2.setCell "E7", 3404
sheet2.setCell "I1", -34
sheet2.setCell "I2", 31.125
sheet2.setCell "G102",
  value: "Hello World!"
  style: "bold"

generateDeferred = xlsxDocument.generate __dirname + "/test.xlsx"

generateDeferred.then (totalSize) ->
  console.log "#{totalSize} total bytes"
  console.log "XLSX generator has been finalized and the output file descriptor has closed."
, (error)->
  console.error "XLSX generator has been failed with error #{error}"

# TODO: Implement setting style and data for row
# with setRow method
#sheet.setRow ( "3", [] , {style:"bold"})
