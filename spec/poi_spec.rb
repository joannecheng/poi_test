require 'rspec'
require_relative '../poi'

describe Poi do

  def filename
    'spec/USA_Country_MetaData_en_EXCEL.xls'
  end

  describe '#initialize' do

    it 'creates a workbook' do
      p = Poi.new filename
      expect(p.workbook.class).to eq HSSFWorkbook
    end
  end

  describe '#worksheets' do
    it 'returns all worksheets' do
      p = Poi.new filename
      worksheets = p.worksheets
      # I need better tests
      expect(worksheets[0].class).to eq Worksheet
      expect(worksheets[1].class).to eq Worksheet
    end
  end

  describe '#worksheet' do
    it 'returns a single worksheet at index' do
      p = Poi.new filename
      worksheet = p.worksheet 1
    end
  end
end
