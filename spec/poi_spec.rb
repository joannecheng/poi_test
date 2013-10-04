require 'rspec'
require_relative '../poi'

  def filename
    'spec/USA_Country_MetaData_en_EXCEL.xls'
  end

describe Poi do

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
      expect(worksheets.length).to eq 2
      expect(worksheets[0].class).to eq Worksheet
      expect(worksheets[1].class).to eq Worksheet
    end
  end

  describe '#worksheet' do
    it 'returns a single worksheet at index' do
      p = Poi.new filename
      worksheet = p.worksheet 1
      expect(worksheet.class).to eq Worksheet
    end
  end
end

describe Worksheet do

  describe '#rows' do

    it 'returns all row in worksheet' do
      p = Poi.new filename
      ws = p.worksheet(1)
      expect(ws.rows.count).to eq 1289
    end

  end
end

describe Row do

  describe '#cells' do
    it 'returns all cells in row' do
      p = Poi.new filename
      ws = p.worksheet(1)
      row = ws.row(5)
      expect(row.cells.count).to eq 4
    end
  end
end
