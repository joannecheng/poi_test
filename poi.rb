require 'cell'
require 'java'
require 'jar/poi-3.9-20131003.jar'
require 'jar/poi-excelant-3.9-20131003.jar'
require 'jar/poi-ooxml-3.9-20131003.jar'

import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import java.io.FileInputStream


WB = Java::OrgApachePoiSsUsermodel::Workbook
WorkbookFactory = Java::OrgApachePoiSsUsermodel::WorkbookFactory
HSSFWorkbook = Java::OrgApachePoiHssfUsermodel::HSSFWorkbook

class Poi
  attr_reader :workbook

  def initialize(filename)
    inp = FileInputStream.new(filename)
    @workbook = WorkbookFactory.create(inp)
  end

  def worksheets
    @workbook.get_number_of_sheets.times.map do |i|
      Worksheet.new @workbook.getSheetAt i
    end
  end

  def worksheet(index)
    Worksheet.new @workbook.getSheetAt(index)
  end
end

class Worksheet
  attr_reader :worksheet
  def initialize(worksheet)
    @worksheet = worksheet
  end

  def rows
    row_count.times.map do |i|
      row(i)
    end
  end

  def row(index)
    Row.new @worksheet.getRow(index)
  end

  def name
    @worksheet.getSheetName()
  end

  private

  def row_count
    @worksheet.getLastRowNum
  end

end

class Row
  attr_reader :row
  def initialize(row)
    @row = row
  end

  def cells
    cell_count.times.map do |i|
      cell(i)
    end
  end

  def cell(index)
    cell = @row.getCell(index)
    Cell.cell_by_type(cell)
  end

  private

  def cell_count
    @row.getLastCellNum
  end

end

