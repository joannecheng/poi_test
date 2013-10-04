require 'row'
require 'cell'
require 'worksheet'
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

  def evaluator
    @evaluator ||= @workbook.getCreationHelper().createFormulaEvaluator()
  end
end

