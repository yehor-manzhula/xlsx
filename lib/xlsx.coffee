fs = require "fs"
_  = require "underscore"

moment   = require "moment"
mustache = require "mustache"

Sheet = require "./sheet"
ObservableMixin = require "./mixins/observableMixin"

# Extend require to load xml files as string
require.extensions[".xml"] = (module, filename)->
  module.exports = fs.readFileSync filename, "utf8"

#Templates
appTemplate    = require "./templates/app.xml"
sheetTemplate  = require "./templates/sheet.xml"
stylesTemplate = require "./templates/styles.xml"
workBookTemplate = require "./templates/workbook.xml"
sharedStringsTemplate = require "./templates/sharedStrings.xml"
themeTemplate = require "./templates/theme.xml"
relationsTemplate = require "./templates/relations.xml"
coreTemplate = require "./templates/core.xml"
contentTypesTemplate = require "./templates/contentTypes.xml"

archiver = require 'archiver'

class Xlsx

  # @mixes EventDispatcherMixin
  _.extend @prototype, ObservableMixin

  XMLDOCTYPE: "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"

  MAIN_RELATIONS:
    APP:
      TYPE:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
      TARGET: "docProps/app.xml"

    CORE:
      TYPE: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
      TARGET: "docProps/core.xml"

    WORKBOOK:
      TYPE:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
      TARGET: "xl/workbook.xml"

  APP_RELATIONS:
    STYLES:
      TYPE:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
      TARGET: "styles.xml"

    WORKSHEET:
      TYPE:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
      TARGET: (sheetNumber)-> "worksheets/sheet#{sheetNumber}.xml"

    THEME:
      TYPE:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
      TARGET: "theme/theme1.xml"

    SHAREDSTRINGS:
      TYPE:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
      TARGET: "sharedStrings.xml"

  CONTENT_TYPES:
    APP:
      TYPE: "application/vnd.openxmlformats-officedocument.extended-properties+xml"
      TARGET: "/docProps/app.xml"

    CORE:
      TYPE:   "application/vnd.openxmlformats-package.core-properties+xml"
      TARGET: "/docProps/core.xml"

    STYLES:
      TYPE: "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
      TARGET: "/xl/styles.xml"

    WORKBOOK:
      TYPE: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
      TARGET: "/xl/workbook.xml"

    WORKSHEET:
      TYPE: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
      TARGET: (pageNumber)-> "/xl/worksheets/sheet#{pageNumber}.xml"

    SHAREDSTRINGS:
      TYPE:   "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
      TARGET: "/xl/sharedStrings.xml"

    THEME:
      TYPE:   "application/vnd.openxmlformats-officedocument.theme+xml"
      TARGET: (pageNumber)-> "/xl/theme/theme#{pageNumber}.xml"

  FONT_STYLES:
    "default": 0
    "normal" : 0
    "bold"   : 1

  TYPES_CODES:
    "number" : "n"
    "string" : "s"
    "default": "s"

  constructor:->
    @sheets = []

    @_sharedStrings = []
    @_contentTypes  = []
    @_totalStrings  = 0

    @_appRelations  = []
    @_mainRelations = []

    @creator = "test"

    @_addMainRelation @MAIN_RELATIONS.APP.TYPE,  @MAIN_RELATIONS.APP.TARGET
    @_addMainRelation @MAIN_RELATIONS.CORE.TYPE,      @MAIN_RELATIONS.CORE.TARGET
    @_addMainRelation @MAIN_RELATIONS.WORKBOOK.TYPE,  @MAIN_RELATIONS.WORKBOOK.TARGET

  createSheet:(name)->
    name?= "Sheet #{@sheets.length}"
    sheet = new Sheet name
    @sheets.push sheet

    sheet

  # Returns cell value
  # in case if value is object returns
  # value field
  _getCellValue: (cell) ->
    return cell["value"] if typeof cell is "object" and cell.value?
    cell

  _generateSharedStrings:->
    # Add to relation shared strings
    @_addAppRelation @APP_RELATIONS.SHAREDSTRINGS.TYPE, @APP_RELATIONS.SHAREDSTRINGS.TARGET
    @_addContentType @CONTENT_TYPES.SHAREDSTRINGS.TYPE, @CONTENT_TYPES.SHAREDSTRINGS.TARGET

    @sheets.forEach (sheet) =>
      sheetData = sheet.data

      sheetData.forEach (row) =>
        row.forEach (column) =>
          value = @_getCellValue column

          return unless typeof(value) is "string"
          @_totalStrings++
          @_sharedStrings.push(value) if @_sharedStrings.indexOf(value) is -1

    result = mustache.render(sharedStringsTemplate,
      xmlDocType : @XMLDOCTYPE
      uniqueCount: @_sharedStrings.length
      count: @_totalStrings
      items: @_sharedStrings
    )

    result

  _addAppRelation:(type, target)->
    @_appRelations.push
      id: @_appRelations.length + 1
      type:type
      target: target

  _addMainRelation:(type, target)->
    @_mainRelations.push
      id: @_mainRelations.length + 1
      type:type
      target: target

  _generateAppRelations:->
    result = mustache.render(relationsTemplate,
      xmlDocType: @XMLDOCTYPE,
      relations: @_appRelations
    )

    result

  _generateMainRelations:->
    result = mustache.render(relationsTemplate,
      xmlDocType: @XMLDOCTYPE,
      relations: @_mainRelations
    )

    result

  # Generates styles.xml
  # TODO: Remove result variable
  _generateXlsStyles:(data) ->
    # Add to relation styles strings
    @_addAppRelation @APP_RELATIONS.STYLES.TYPE, @APP_RELATIONS.STYLES.TARGET
    @_addContentType @CONTENT_TYPES.STYLES.TYPE, @CONTENT_TYPES.STYLES.TARGET

    result = mustache.render(stylesTemplate,
      xmlDocType: @XMLDOCTYPE
    )

    result

  #
  # Generates app.xml
  _generateXlsApp:(data) ->

    @_addContentType @CONTENT_TYPES.APP.TYPE, @CONTENT_TYPES.APP.TARGET

    result = mustache.render(appTemplate,
      xmlDocType: @XMLDOCTYPE
      userName: @creator or "officegen"
      pagesCount: @sheets.length
      sheets: ((totalPages) ->
          result = []
          i = 0

          while i < totalPages
            result.push i + 1
            i++
          result
        )(@sheets.length)
    )

    result

  # @param[in] data Ignored by this callback function.
  # @return Text string.
  #
  _generateXlsWorkbook:() ->
    @_addContentType @CONTENT_TYPES.WORKBOOK.TYPE, @CONTENT_TYPES.WORKBOOK.TARGET

    sheets = []
    @sheets.forEach (sheet, index) =>

      sheetName = sheet.name or "Sheet" + (index + 1)

      relationItem = _.find @_appRelations, (element)=>
        element.type is @APP_RELATIONS.WORKSHEET.TYPE and
        element.target is @APP_RELATIONS.WORKSHEET.TARGET(index + 1)

      sheets.push
        name: sheetName
        sheetId: (index + 1)
        rId: relationItem.id

    result = mustache.render(workBookTemplate,
      xmlDocType: @XMLDOCTYPE
      sheets: sheets
    )
    result

  # Returns cell style
  # in case if value is object returns
  # style field
  # otherwise return default style
  _getCellStyleId: (cell) ->
    if typeof(cell) is "object" and cell.style?
      style = cell.style.toLowerCase()
      return @FONT_STYLES[style] if @FONT_STYLES[style]?

    @FONT_STYLES["default"]

  _generateXlsSheets: ->
    @sheets.forEach (sheet, index)=>
      fs.writeFileSync "./tmp/xl/worksheets/sheet#{index + 1}.xml", @_generateXlsSheet(sheet)

      #TODO: Improve pathes issue
      @_addAppRelation @APP_RELATIONS.WORKSHEET.TYPE, @APP_RELATIONS.WORKSHEET.TARGET(index + 1)
      @_addContentType @CONTENT_TYPES.WORKSHEET.TYPE, @CONTENT_TYPES.WORKSHEET.TARGET(index + 1)

  _generateXlsSheet: (sheet) ->
    xSize = 0
    ySize = 0
    rows = []

    # Find the maximum cells area:
    ySize = sheet.data.length - 1 if sheet.data.length

    sheet.data.forEach (row)->
      currentColumnSize = row.length - 1 if row.length
      xSize = Math.max currentColumnSize, xSize

    sheet.data.forEach (row, rowIndex)=>
      rowLines = 1

      currentRow =
        columns: []
        rowId: rowIndex + 1
        height: rowLines * 15
        spansDimension: "1:#{row.length}"

      row.forEach (column, columnIndex) =>
        columnValue = @_getCellValue column

        if typeof columnValue isnt "undefined"
          value = undefined
          type = @TYPES_CODES["default"]

          cellValueType = typeof columnValue
          type = @TYPES_CODES[cellValueType] if @TYPES_CODES[cellValueType]

          switch typeof columnValue

            when "number"
              value = columnValue

            when "string"
              # Calculate row height
              # depend on max lines in cell
              # TODO: move height calculation to another method
              candidate = columnValue.split("\n").length
              rowLines = Math.max rowLines, candidate
              currentRow.height = rowLines * 15

              value = @_sharedStrings.indexOf columnValue

          currentColumn =
            cellName: "#{Sheet::numberToCell(columnIndex)}#{rowIndex + 1}"
            type: type
            value: value
            styleId: @_getCellStyleId(column)

          currentRow.columns.push currentColumn

        rows.push currentRow

    result = mustache.render(sheetTemplate,
      xmlDocType: @XMLDOCTYPE
      dimension: "A1:#{Sheet::numberToCell(xSize)}#{(ySize + 1)}"
      rows: rows
    )

    result

  _generateTheme:->
    # Add to relation themes
    @_addAppRelation @APP_RELATIONS.THEME.TYPE, @APP_RELATIONS.THEME.TARGET
    @_addContentType @CONTENT_TYPES.THEME.TYPE, @CONTENT_TYPES.THEME.TARGET(1)

    mustache.render(themeTemplate,
      xmlDocType: @XMLDOCTYPE
    )

  _generateCore:->
    @_addContentType @CONTENT_TYPES.CORE.TYPE, @CONTENT_TYPES.CORE.TARGET

    mustache.render(coreTemplate,
      xmlDocType : @XMLDOCTYPE
      creator : "Egor Manjula"
      extraFields: []
      currentDateTime: moment().format("YYYY-MM-DD[T]HH[:]mm[:]ss[Z]") #"2014-04-28T20:36:44Z"
      revisionNumber: 1
    )

  # TODO: Finish this part
  _addContentType:(type, target)->
    @_contentTypes.push
      type: type
      target: target

  _generateContentTypes:->
    mustache.render(contentTypesTemplate,
      xmlDocType  : @XMLDOCTYPE
      contentTypes: @_contentTypes
    )

  generate:() ->

    # TODO: Temporary folders
    fs.mkdirSync "./tmp"       unless fs.existsSync "./tmp"
    fs.mkdirSync "./tmp/xl"    unless fs.existsSync "./tmp/xl"
    fs.mkdirSync "./tmp/xl/theme" unless fs.existsSync "./tmp/xl/theme"
    fs.mkdirSync "./tmp/xl/_rels" unless fs.existsSync "./tmp/xl/_rels"
    fs.mkdirSync "./tmp/xl/worksheets" unless fs.existsSync "./tmp/xl/worksheets"

    fs.mkdirSync "./tmp/_rels" unless fs.existsSync "./tmp/_rels"
    fs.mkdirSync "./tmp/docProps" unless fs.existsSync "./tmp/docProps"

    fs.writeFileSync "./tmp/docProps/app.xml",  @_generateXlsApp()
    fs.writeFileSync "./tmp/docProps/core.xml", @_generateCore()

    fs.writeFileSync "./tmp/xl/sharedStrings.xml", @_generateSharedStrings()

    fs.writeFileSync "./tmp/xl/styles.xml", @_generateXlsStyles()

    fs.writeFileSync "./tmp/xl/theme/theme1.xml", @_generateTheme()

    @_generateXlsSheets()
    fs.writeFileSync "./tmp/xl/workbook.xml", @_generateXlsWorkbook()

    fs.writeFileSync "./tmp/[Content_Types].xml", @_generateContentTypes()

    fs.writeFileSync "./tmp/xl/_rels/workbook.xml.rels", @_generateAppRelations()
    fs.writeFileSync "./tmp/_rels/.rels", @_generateMainRelations()


    output = fs.createWriteStream __dirname + '/bulk-output.xlsx'
    archive = archiver 'zip'

    output.on 'close', ()->
      console.log archive.pointer() + ' total bytes'
      console.log 'archiver has been finalized and the output file descriptor has closed.'

    archive.on 'error', (err)->
      throw err

    archive.pipe(output)
    archive.bulk([
      { expand: true, cwd: './tmp', src: ['_rels/.rels']},
      { expand: true, cwd: './tmp', src: ['**/*'] }])

    archive.finalize()

module.exports = Xlsx

# test
xlsxDocument = new Xlsx()

# First sheet
sheet = xlsxDocument.createSheet()

sheet.name = "Excel 1"

sheet.data[0] = []
sheet.data[0][0] = 1
sheet.data[1] = []
sheet.data[1][3] = {value:'abc', style:'BOLD'}
sheet.data[1][4] = {value:'More', style: 'bOld'}
sheet.data[1][5] = 'Text fffdaffdfdadfasdfasdfaf'
sheet.data[1][6] = 'Here'
sheet.data[2] = []
sheet.data[2][5] = 'abc'
sheet.data[2][6] = 900
sheet.data[6] = []
sheet.data[6][2] = 1972

sheet.setCell 'E7', 340
sheet.setCell 'I1', -3
sheet.setCell 'I2', 31.12
sheet.setCell 'G102', {value:'Hello World!', style:'bold'}

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

xlsxDocument.generate()

#sheet.setRow ( "3", [] , {style:"bold"})